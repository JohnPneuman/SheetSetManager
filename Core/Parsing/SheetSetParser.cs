using System.Xml.Linq;
using SheetSet.Core.Models;

namespace SheetSet.Core.Parsing;

public sealed class SheetSetParser
{
    /// <summary>
    /// Parse for read-only display.
    /// </summary>
    public SheetSetDocument Parse(string xmlPath)
        => ParseInternal(xmlPath, keepDocument: false);

    /// <summary>
    /// Parse while keeping the XDocument in memory so SheetSetWriter can modify and re-save.
    /// </summary>
    public SheetSetDocument ParseForEditing(string dstPath, string xmlPath)
    {
        var doc = ParseInternal(xmlPath, keepDocument: true);
        doc.SourceDstPath = dstPath;
        return doc;
    }

    private SheetSetDocument ParseInternal(string xmlPath, bool keepDocument)
    {
        if (!File.Exists(xmlPath))
            throw new FileNotFoundException("XML-bestand niet gevonden.", xmlPath);

        var xdoc = XDocument.Load(xmlPath, LoadOptions.PreserveWhitespace);

        var sheetSetElement = xdoc.Descendants("AcSmSheetSet").FirstOrDefault()
            ?? throw new InvalidOperationException("Geen AcSmSheetSet gevonden.");

        var baseFolder = Path.GetDirectoryName(xmlPath) ?? Environment.CurrentDirectory;

        var info = new SheetSetInfo
        {
            Name = XmlUtil.GetPropValue(sheetSetElement, "Name"),
            Description = XmlUtil.GetPropValue(sheetSetElement, "Desc"),
            ProjectName = XmlUtil.GetPropValue(sheetSetElement, "ProjectName"),
            ProjectNumber = XmlUtil.GetPropValue(sheetSetElement, "ProjectNumber"),
            ProjectPhase = XmlUtil.GetPropValue(sheetSetElement, "ProjectPhase"),
            ProjectMilestone = XmlUtil.GetPropValue(sheetSetElement, "ProjectMilestone"),
            BaseFolder = baseFolder,
            CustomProperties = XmlUtil.ReadCustomProperties(sheetSetElement),
            Element = sheetSetElement
        };

        var rootSubsets = sheetSetElement.Elements("AcSmSubset")
            .Select(e => ParseSubset(e, baseFolder)).ToList();
        var rootSheets = sheetSetElement.Elements("AcSmSheet")
            .Select(e => ParseSheet(e, baseFolder)).ToList();

        var document = new SheetSetDocument(info, rootSubsets, rootSheets);
        if (keepDocument)
            document.SourceDocument = xdoc;

        return document;
    }

    private static SubsetNode ParseSubset(XElement element, string baseFolder)
    {
        var info = new SubsetInfo
        {
            Name = XmlUtil.GetPropValue(element, "Name"),
            Description = XmlUtil.GetPropValue(element, "Desc"),
            NewSheetLocation = XmlUtil.ReadFileReference(element, "NewSheetLocation", baseFolder),
            CustomProperties = XmlUtil.ReadCustomProperties(element),
            Element = element
        };

        var subsets = element.Elements("AcSmSubset").Select(e => ParseSubset(e, baseFolder)).ToList();
        var sheets = element.Elements("AcSmSheet").Select(e => ParseSheet(e, baseFolder)).ToList();

        return new SubsetNode(info, subsets, sheets);
    }

    private static SheetNode ParseSheet(XElement element, string baseFolder)
    {
        var layoutElement = element.Elements("AcSmAcDbLayoutReference")
            .FirstOrDefault(e => string.Equals((string?)e.Attribute("propname"), "Layout",
                StringComparison.OrdinalIgnoreCase));

        var layoutRef = XmlUtil.ReadLayoutReference(layoutElement, baseFolder);

        var doNotPlotRaw = XmlUtil.GetPropValue(element, "DoNotPlot");
        bool? doNotPlot = doNotPlotRaw switch
        {
            "1" or "-1" => true,
            "0" => false,
            _ => null
        };

        var info = new SheetInfo
        {
            Number = XmlUtil.GetPropValue(element, "Number"),
            Title = XmlUtil.GetPropValue(element, "Title"),
            Description = XmlUtil.GetPropValue(element, "Desc"),
            Category = XmlUtil.GetPropValue(element, "Category"),
            RevisionNumber = XmlUtil.GetPropValue(element, "RevisionNumber"),
            RevisionDate = XmlUtil.GetPropValue(element, "RevisionDate"),
            IssuePurpose = XmlUtil.GetPropValue(element, "IssuePurpose"),
            DoNotPlot = doNotPlot,
            LayoutName = layoutRef.Name,
            DwgFileName = layoutRef.FileName,
            RelativeDwgFileName = layoutRef.RelativeFileName,
            ResolvedDwgPath = XmlUtil.ResolveBestPath(layoutRef.FileName, layoutRef.RelativeFileName, baseFolder),
            FolderPath = XmlUtil.ResolveFolder(layoutRef.FileName, layoutRef.RelativeFileName, baseFolder),
            CustomProperties = XmlUtil.ReadCustomProperties(element),
            Element = element
        };

        return new SheetNode(info);
    }
}
