using System.Xml.Linq;
using SheetSet.Core.Models;

namespace SheetSet.Core.Parsing;

public static class XmlUtil
{
    public static string? GetPropValue(XElement parent, string propName)
        => parent.Elements("AcSmProp")
            .FirstOrDefault(x => Eq((string?)x.Attribute("propname"), propName))
            ?.Value;

    public static void SetPropValue(XElement parent, string propName, string value)
    {
        var prop = parent.Elements("AcSmProp")
            .FirstOrDefault(x => Eq((string?)x.Attribute("propname"), propName));
        if (prop != null)
            prop.Value = value;
        else
            parent.AddFirst(new XElement("AcSmProp", new XAttribute("propname", propName), value));
    }

    public static Dictionary<string, string> ReadCustomProperties(XElement parent)
    {
        var bag = parent.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "AcSmCustomPropertyBag");

        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (bag == null) return dict;

        foreach (var item in bag.Elements().Where(e => e.Name.LocalName == "AcSmCustomPropertyValue"))
        {
            var key = (string?)item.Attribute("propname");
            var value = item.Elements("AcSmProp")
                .FirstOrDefault(x => Eq((string?)x.Attribute("propname"), "Value"))
                ?.Value;

            if (!string.IsNullOrWhiteSpace(key))
                dict[key!] = value ?? string.Empty;
        }

        return dict;
    }

    public static void SetCustomProperty(XElement parent, string propName, string value)
    {
        var bag = parent.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "AcSmCustomPropertyBag");
        if (bag == null) return;

        var item = bag.Elements()
            .Where(e => e.Name.LocalName == "AcSmCustomPropertyValue")
            .FirstOrDefault(e => Eq((string?)e.Attribute("propname"), propName));

        if (item == null) return;

        var valueProp = item.Elements("AcSmProp")
            .FirstOrDefault(x => Eq((string?)x.Attribute("propname"), "Value"));
        if (valueProp != null)
            valueProp.Value = value;
    }

    public static FileReferenceInfo? ReadFileReference(XElement parent, string propName, string baseFolder)
    {
        var element = parent.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "AcSmFileReference"
                && Eq((string?)e.Attribute("propname"), propName));

        if (element == null) return null;

        var fileName = GetPropValue(element, "FileName");
        var relative = GetPropValue(element, "Relative_FileName");

        return new FileReferenceInfo
        {
            FileName = fileName,
            RelativeFileName = relative,
            ResolvedPath = ResolveBestPath(fileName, relative, baseFolder)
        };
    }

    public static LayoutReferenceInfo ReadLayoutReference(XElement? element, string baseFolder)
    {
        if (element == null) return new LayoutReferenceInfo();

        var fileName = GetPropValue(element, "FileName");
        var relative = GetPropValue(element, "Relative_FileName");

        return new LayoutReferenceInfo
        {
            Name = GetPropValue(element, "Name"),
            FileName = fileName,
            RelativeFileName = relative,
            ResolvedPath = ResolveBestPath(fileName, relative, baseFolder)
        };
    }

    public static string? ResolveBestPath(string? absolutePath, string? relativePath, string baseFolder)
    {
        if (!string.IsNullOrWhiteSpace(absolutePath) && File.Exists(absolutePath))
            return absolutePath;

        if (!string.IsNullOrWhiteSpace(relativePath))
        {
            try
            {
                var combined = Path.GetFullPath(Path.Combine(baseFolder, relativePath));
                return combined;
            }
            catch { }
        }

        return absolutePath ?? relativePath;
    }

    public static string? ResolveFolder(string? absolutePath, string? relativePath, string baseFolder)
    {
        var best = ResolveBestPath(absolutePath, relativePath, baseFolder);
        if (string.IsNullOrWhiteSpace(best)) return null;
        if (Directory.Exists(best)) return best;
        try { return Path.GetDirectoryName(best); }
        catch { return null; }
    }

    private static bool Eq(string? a, string? b)
        => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
}
