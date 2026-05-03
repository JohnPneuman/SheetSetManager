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
            // Use Add (append at end) — never AddFirst which inserts before
            // AcSmCustomPropertyBag and corrupts the element order.
            // Include vt="8" (string type) as AutoCAD writes it.
            parent.Add(new XElement("AcSmProp",
                new XAttribute("propname", propName),
                new XAttribute("vt", "8"),
                value));
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

    public static void DeleteCustomProperty(XElement parent, string propName)
    {
        var bag = parent.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "AcSmCustomPropertyBag");
        if (bag == null) return;
        bag.Elements()
            .Where(e => e.Name.LocalName == "AcSmCustomPropertyValue")
            .FirstOrDefault(e => Eq((string?)e.Attribute("propname"), propName))
            ?.Remove();
    }

    // flags: 2 = per-sheet veld (default), 1 = sheetset veld
    public static void SetCustomProperty(XElement parent, string propName, string value, int flags = 2)
    {
        var bag = GetOrCreateBag(parent);
        if (bag == null) return;

        var item = bag.Elements()
            .Where(e => e.Name.LocalName == "AcSmCustomPropertyValue")
            .FirstOrDefault(e => Eq((string?)e.Attribute("propname"), propName));

        if (item != null)
        {
            var valueProp = item.Elements("AcSmProp")
                .FirstOrDefault(x => Eq((string?)x.Attribute("propname"), "Value"));
            if (valueProp != null)
                valueProp.Value = value;
            return;
        }

        // Property bestaat nog niet — aanmaken met AutoCAD-structuur
        bag.Add(new XElement("AcSmCustomPropertyValue",
            new XAttribute("clsid", "g8D22A2A4-1777-4D78-84CC-69EF741FE954"),
            new XAttribute("ID", NewAcSmId()),
            new XAttribute("propname", propName),
            new XAttribute("vt", "13"),
            new XElement("AcSmProp", new XAttribute("propname", "Flags"), new XAttribute("vt", "3"), flags.ToString()),
            new XElement("AcSmProp", new XAttribute("propname", "Value"), new XAttribute("vt", "8"), value)));
    }

    public static List<CustomPropertyDefinition> ReadCustomPropertyDefinitions(XElement parent)
    {
        var bag = parent.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "AcSmCustomPropertyBag");
        var list = new List<CustomPropertyDefinition>();
        if (bag == null) return list;

        foreach (var item in bag.Elements().Where(e => e.Name.LocalName == "AcSmCustomPropertyValue"))
        {
            var key = (string?)item.Attribute("propname");
            if (string.IsNullOrWhiteSpace(key)) continue;

            var value = item.Elements("AcSmProp")
                .FirstOrDefault(x => Eq((string?)x.Attribute("propname"), "Value"))?.Value ?? string.Empty;
            var flagsRaw = item.Elements("AcSmProp")
                .FirstOrDefault(x => Eq((string?)x.Attribute("propname"), "Flags"))?.Value;
            int.TryParse(flagsRaw, out var flags);

            list.Add(new CustomPropertyDefinition
            {
                Name  = key!,
                Value = value,
                Flags = flags == 0 ? 2 : flags
            });
        }
        return list;
    }

    private static XElement? GetOrCreateBag(XElement parent)
    {
        var bag = parent.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "AcSmCustomPropertyBag");
        if (bag != null) return bag;

        // Nieuw blad (toegevoegd via de editor) heeft nog geen bag — aanmaken
        bag = new XElement("AcSmCustomPropertyBag",
            new XAttribute("clsid", "g4D103908-8C86-4D95-BBF4-68B9A7B00731"),
            new XAttribute("ID", NewAcSmId()),
            new XAttribute("propname", "CustomPropertyBag"),
            new XAttribute("vt", "13"));
        parent.AddFirst(bag);
        return bag;
    }

    private static string NewAcSmId()
        => "g" + Guid.NewGuid().ToString("D").ToUpperInvariant();

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
