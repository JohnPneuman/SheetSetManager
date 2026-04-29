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

        DstCodec.SaveXmlAsDst(doc.SourceDocument, path);
    }

    // ─── SheetSet rename ──────────────────────────────────────────────────────

    public static void RenameSheetSet(SheetSetDocument doc, string newName)
    {
        var el = doc.Info.Element
            ?? throw new InvalidOperationException("SheetSetDocument heeft geen XElement referentie.");
        XmlUtil.SetPropValue(el, "Name", newName);
        doc.Info.Name = newName;
    }

    // ─── Subset operations ────────────────────────────────────────────────────

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
            new XElement("AcSmProp", new XAttribute("propname", "Number"), number),
            new XElement("AcSmProp", new XAttribute("propname", "Title"), title),
            new XElement("AcSmProp", new XAttribute("propname", "Desc"), ""));

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
        if (value == null) return;
        XmlUtil.SetPropValue(element, propName, value);
        updateModel(value);
    }

    private static string SecurityEncode(string s)
        => s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
}
