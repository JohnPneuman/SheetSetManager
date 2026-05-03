using System.Xml.Linq;
using SheetSet.Core.Interop;
using SheetSet.Core.Models;
using SheetSet.Core.Parsing;

namespace SheetSet.Core.Writing;

public static class SheetSetWriter
{
    // ─── Sheet property updates ───────────────────────────────────────────────

    public static void Apply(IEnumerable<SheetUpdateModel> updates)
    {
        foreach (var u in updates)
        {
            var el = u.Sheet.Element
                ?? throw new InvalidOperationException(
                    $"Sheet '{u.Sheet.Number}' heeft geen XElement referentie. Gebruik ParseForEditing().");

            ApplyIfSet(el, "Number",         u.NewNumber,      v => u.Sheet.Number         = v);
            ApplyIfSet(el, "Title",          u.Title,          v => u.Sheet.Title          = v);
            ApplyIfSet(el, "Desc",           u.Description,    v => u.Sheet.Description    = v);
            ApplyIfSet(el, "Category",       u.Category,       v => u.Sheet.Category       = v);
            ApplyIfSet(el, "RevisionNumber", u.RevisionNumber, v => u.Sheet.RevisionNumber = v);
            ApplyIfSet(el, "RevisionDate",   u.RevisionDate,   v => u.Sheet.RevisionDate   = v);
            ApplyIfSet(el, "IssuePurpose",   u.IssuePurpose,   v => u.Sheet.IssuePurpose   = v);

            if (u.DoNotPlot.HasValue)
            {
                XmlUtil.SetPropValue(el, "DoNotPlot", u.DoNotPlot.Value ? "-1" : "0");
                u.Sheet.DoNotPlot = u.DoNotPlot.Value;
            }

            if (u.CustomProperties != null)
                foreach (var (key, value) in u.CustomProperties)
                {
                    XmlUtil.SetCustomProperty(el, key, value);
                    u.Sheet.CustomProperties[key] = value;
                }
        }
    }

    // ─── Save ─────────────────────────────────────────────────────────────────

    public static void Save(SheetSetDocument doc, string? outputPath = null)
    {
        if (doc.SourceDocument == null)
            throw new InvalidOperationException("Geen SourceDocument. Gebruik ParseForEditing() om te laden.");

        var path = outputPath ?? doc.SourceDstPath
            ?? throw new InvalidOperationException("Geen uitvoerpad opgegeven.");

        if (outputPath != null)
            doc.SourceDstPath = outputPath;

        BackupHelper.BackupFile(path);
        DstCodec.SaveXmlAsDst(doc.SourceDocument, path);
    }

    // ─── SheetSet rename ──────────────────────────────────────────────────────

    // ─── SheetSet custom property definitions ─────────────────────────────────

    // flags: 2 = per-sheet veld (zichtbaar per blad), 1 = sheetset veld (één waarde voor de set)
    public static void AddSheetSetCustomProperty(SheetSetDocument doc, string name, int flags, string defaultValue = "")
    {
        var sheetSetEl = doc.Info.Element
            ?? throw new InvalidOperationException("SheetSetDocument heeft geen XElement referentie.");

        XmlUtil.SetCustomProperty(sheetSetEl, name, defaultValue, flags);

        var def = new CustomPropertyDefinition { Name = name, Value = defaultValue, Flags = flags };
        doc.Info.CustomPropertyDefinitions.Add(def);
        doc.Info.CustomProperties[name] = defaultValue;

        // Sheet-level property: zet ook een standaard waarde op alle bestaande sheets
        if (flags == 2)
        {
            foreach (var sheet in doc.GetAllSheets())
            {
                if (sheet.Info.Element == null) continue;
                XmlUtil.SetCustomProperty(sheet.Info.Element, name, defaultValue, flags);
                sheet.Info.CustomProperties[name] = defaultValue;
            }
        }
    }

    public static void DeleteSheetSetCustomProperty(SheetSetDocument doc, string name)
    {
        var sheetSetEl = doc.Info.Element
            ?? throw new InvalidOperationException("SheetSetDocument heeft geen XElement referentie.");

        XmlUtil.DeleteCustomProperty(sheetSetEl, name);

        // Verwijder ook de per-sheet waarden bij blad-velden (Flags=2)
        foreach (var sheet in doc.GetAllSheets())
        {
            if (sheet.Info.Element == null) continue;
            XmlUtil.DeleteCustomProperty(sheet.Info.Element, name);
            sheet.Info.CustomProperties.Remove(name);
        }

        doc.Info.CustomPropertyDefinitions.RemoveAll(d =>
            string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase));
        doc.Info.CustomProperties.Remove(name);
    }

    // ─── New sheetset from template ───────────────────────────────────────────

    public static void CreateFromTemplate(
        string templateDstPath, string outputPath, string name, string? description,
        Dictionary<string, string>? valueOverrides = null,
        ISet<string>? propertyNamesToKeep = null)
    {
        var tempXml = DstCodec.DecodeDstToTempXml(templateDstPath);
        try
        {
            var xdoc = XDocument.Load(tempXml, LoadOptions.PreserveWhitespace);

            var dbEl = xdoc.Descendants("AcSmDatabase").FirstOrDefault();
            var sheetSetEl = xdoc.Descendants("AcSmSheetSet").FirstOrDefault()
                ?? throw new InvalidOperationException("Geen AcSmSheetSet gevonden in template.");

            // Naam + omschrijving
            XmlUtil.SetPropValue(sheetSetEl, "Name", name);
            if (!string.IsNullOrEmpty(description))
                XmlUtil.SetPropValue(sheetSetEl, "Desc", description);

            // Standaard waarde overrides vanuit de wizard
            if (valueOverrides != null)
                foreach (var (propName, propValue) in valueOverrides)
                    XmlUtil.SetCustomProperty(sheetSetEl, propName, propValue);

            // Verwijder properties die de gebruiker uit de wizard heeft verwijderd
            if (propertyNamesToKeep != null)
            {
                var allDefs = XmlUtil.ReadCustomPropertyDefinitions(sheetSetEl);
                foreach (var def in allDefs)
                {
                    if (!propertyNamesToKeep.Contains(def.Name))
                    {
                        XmlUtil.DeleteCustomProperty(sheetSetEl, def.Name);
                        foreach (var sheetEl in xdoc.Descendants("AcSmSheet"))
                            XmlUtil.DeleteCustomProperty(sheetEl, def.Name);
                    }
                }
            }

            // Nieuwe unieke vingerafdruk zodat AutoCAD het als een nieuw bestand herkent
            if (dbEl != null)
            {
                var fp = dbEl.Elements("AcSmProp")
                    .FirstOrDefault(x => string.Equals((string?)x.Attribute("propname"), "DbFingerPrint",
                        StringComparison.OrdinalIgnoreCase));
                if (fp != null) fp.Value = "g" + Guid.NewGuid().ToString("D").ToUpperInvariant();

                var rev = dbEl.Elements("AcSmProp")
                    .FirstOrDefault(x => string.Equals((string?)x.Attribute("propname"), "FileRevision",
                        StringComparison.OrdinalIgnoreCase));
                if (rev != null) rev.Value = "0";
            }

            DstCodec.SaveXmlAsDst(xdoc, outputPath);
        }
        finally
        {
            if (File.Exists(tempXml)) File.Delete(tempXml);
        }
    }

    // ─── SheetSet rename ──────────────────────────────────────────────────────

    public static void RenameSheetSet(SheetSetDocument doc, string newName)
    {
        var el = doc.Info.Element
            ?? throw new InvalidOperationException("SheetSetDocument heeft geen XElement referentie.");
        XmlUtil.SetPropValue(el, "Name", newName);
        doc.Info.Name = newName;
    }

    public static void UpdateSheetSetInfo(SheetSetDocument doc,
        string? description, string? projectName, string? projectNumber,
        string? projectPhase, string? projectMilestone)
    {
        var el = doc.Info.Element
            ?? throw new InvalidOperationException("SheetSetDocument heeft geen XElement referentie.");
        ApplyIfSet(el, "Desc",             description,      v => doc.Info.Description     = v);
        ApplyIfSet(el, "ProjectName",      projectName,      v => doc.Info.ProjectName     = v);
        ApplyIfSet(el, "ProjectNumber",    projectNumber,    v => doc.Info.ProjectNumber   = v);
        ApplyIfSet(el, "ProjectPhase",     projectPhase,     v => doc.Info.ProjectPhase    = v);
        ApplyIfSet(el, "ProjectMilestone", projectMilestone, v => doc.Info.ProjectMilestone = v);
    }

    // ─── Subset operations ────────────────────────────────────────────────────

    public static void UpdateSubsetInfo(SubsetNode node, string? description)
    {
        var el = node.Info.Element
            ?? throw new InvalidOperationException($"Subset '{node.Name}' heeft geen XElement referentie.");
        ApplyIfSet(el, "Desc", description, v => node.Info.Description = v);
    }

    public static void RenameSheet(SheetNode node, string newTitle)
    {
        var el = node.Info.Element
            ?? throw new InvalidOperationException($"Sheet '{node.Info.Number}' heeft geen XElement referentie.");
        XmlUtil.SetPropValue(el, "Title", newTitle);
        node.Info.Title = newTitle;
    }

    public static void RenameSubset(SubsetNode node, string newName)
    {
        var el = node.Info.Element
            ?? throw new InvalidOperationException($"Subset '{node.Name}' heeft geen XElement referentie.");
        XmlUtil.SetPropValue(el, "Name", newName);
        node.Info.Name = newName;
    }

    public static SheetNode AddSheet(SheetSetDocument doc, SubsetNode? parent, string number, string title)
    {
        var parentEl = parent?.Info.Element ?? doc.Info.Element
            ?? throw new InvalidOperationException("Geen parent element beschikbaar voor het toevoegen van een sheet.");

        var newEl = new XElement("AcSmSheet",
            new XElement("AcSmProp", new XAttribute("propname", "Number"), new XAttribute("vt", "8"), number),
            new XElement("AcSmProp", new XAttribute("propname", "Title"),  new XAttribute("vt", "8"), title));

        parentEl.Add(newEl);

        var info = new SheetInfo { Number = number, Title = title, Element = newEl };
        var node = new SheetNode(info);
        (parent?.Sheets ?? doc.RootSheets).Add(node);
        return node;
    }

    public static SubsetNode AddSubset(SheetSetDocument doc, SubsetNode? parent, string name)
    {
        var parentEl = parent?.Info.Element ?? doc.Info.Element
            ?? throw new InvalidOperationException("Geen parent element beschikbaar voor het toevoegen van een subset.");

        var newEl = new XElement("AcSmSubset",
            new XElement("AcSmProp", new XAttribute("propname", "Name"), name),
            new XElement("AcSmProp", new XAttribute("propname", "Desc"), ""));

        var firstSheet = parentEl.Elements("AcSmSheet").FirstOrDefault();
        if (firstSheet != null)
            firstSheet.AddBeforeSelf(newEl);
        else
            parentEl.Add(newEl);

        var info = new SubsetInfo { Name = name, Element = newEl };
        var node = new SubsetNode(info, [], []);
        (parent?.Subsets ?? doc.RootSubsets).Add(node);
        return node;
    }

    public static void DeleteSheet(SheetSetDocument doc, SheetNode node)
    {
        node.Info.Element?.Remove();
        if (doc.RootSheets.Remove(node)) return;
        foreach (var sub in doc.RootSubsets)
            if (RemoveSheetFromSubset(sub, node)) return;
    }

    private static bool RemoveSheetFromSubset(SubsetNode parent, SheetNode target)
    {
        if (parent.Sheets.Remove(target)) return true;
        foreach (var sub in parent.Subsets)
            if (RemoveSheetFromSubset(sub, target)) return true;
        return false;
    }

    // ─── Drag & drop / reorder ────────────────────────────────────────────────

    /// <summary>Move element directly before reference in the XML tree.</summary>
    public static void MoveElementBefore(XElement element, XElement reference)
    {
        element.Remove();
        reference.AddBeforeSelf(element);
    }

    /// <summary>Move element directly after reference in the XML tree.</summary>
    public static void MoveElementAfter(XElement element, XElement reference)
    {
        element.Remove();
        reference.AddAfterSelf(element);
    }

    /// <summary>Move element as last child of newParent in the XML tree.</summary>
    public static void MoveElementInto(XElement element, XElement newParent)
    {
        element.Remove();
        newParent.Add(element);
    }

    // ─── New sheet set ────────────────────────────────────────────────────────

    /// <summary>
    /// Create an empty SheetSetDocument in memory (not yet saved to disk).
    /// Call Save() to write the .dst file.
    /// </summary>
    public static SheetSetDocument CreateNew(string name, string dstPath)
    {
        var xdoc = XDocument.Parse(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<AcSmSheetSetDB>" +
            "<AcSmSheetSet>" +
            $"<AcSmProp propname=\"Name\">{SecurityEncode(name)}</AcSmProp>" +
            "<AcSmProp propname=\"Desc\"></AcSmProp>" +
            "<AcSmProp propname=\"ProjectName\"></AcSmProp>" +
            "<AcSmProp propname=\"ProjectNumber\"></AcSmProp>" +
            "<AcSmProp propname=\"ProjectPhase\"></AcSmProp>" +
            "<AcSmProp propname=\"ProjectMilestone\"></AcSmProp>" +
            "<AcSmCustomPropertyBag/>" +
            "</AcSmSheetSet>" +
            "</AcSmSheetSetDB>");

        var sheetSetEl = xdoc.Descendants("AcSmSheetSet").First();
        var info = new SheetSetInfo { Name = name, Element = sheetSetEl };
        var doc = new SheetSetDocument(info, [], []) { SourceDocument = xdoc, SourceDstPath = dstPath };
        return doc;
    }

    // ─── Private helpers ──────────────────────────────────────────────────────

    private static void ApplyIfSet(XElement element, string propName, string? value, Action<string> updateModel)
    {
        if (string.IsNullOrEmpty(value)) return;
        XmlUtil.SetPropValue(element, propName, value);
        updateModel(value);
    }

    private static string SecurityEncode(string s)
        => s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
}
