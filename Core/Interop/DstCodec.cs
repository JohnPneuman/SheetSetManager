using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace SheetSet.Core.Interop;

public static class DstCodec
{
    // The substitution pairs: (encodedByte, plainByte) — must be declared before DecodeMap/EncodeMap
    private static readonly (byte Enc, byte Plain)[] Pairs =
    [
        (131, 10), (134, 9), (172, 32), (175, 33), (174, 34), (169, 35), (168, 36), (171, 37),
        (170, 38), (165, 39), (164, 40), (167, 41), (166, 42), (161, 43), (160, 44), (163, 45),
        (162, 46), (221, 47), (220, 48), (223, 49), (222, 50), (217, 51), (216, 52), (219, 53),
        (218, 54), (213, 55), (212, 56), (215, 57), (214, 58), (209, 59), (208, 60), (211, 61),
        (210, 62), (205, 63), (204, 64), (207, 65), (206, 66), (201, 67), (200, 68), (203, 69),
        (202, 70), (197, 71), (196, 72), (199, 73), (198, 74), (193, 75), (192, 76), (195, 77),
        (194, 78), (253, 79), (252, 80), (255, 81), (254, 82), (249, 83), (248, 84), (251, 85),
        (250, 86), (245, 87), (244, 88), (247, 89), (246, 90), (241, 91), (240, 92), (243, 93),
        (242, 94), (237, 95), (236, 96), (239, 97), (238, 98), (233, 99), (232, 100), (235, 101),
        (234, 102), (229, 103), (228, 104), (231, 105), (230, 106), (225, 107), (224, 108), (227, 109),
        (226, 110), (29, 111), (28, 112), (31, 113), (30, 114), (25, 115), (24, 116), (27, 117),
        (26, 118), (21, 119), (20, 120), (23, 121), (22, 122)
    ];

    private static readonly byte[] DecodeMap = BuildDecodeMap();
    private static readonly byte[] EncodeMap = BuildEncodeMap();

    public static byte[] DecodeDst(byte[] dstBytes) => Transform(dstBytes, DecodeMap);
    public static byte[] EncodeToDst(byte[] xmlBytes) => Transform(xmlBytes, EncodeMap);

    public static string DecodeDstToTempXml(string dstPath)
    {
        var dstBytes = File.ReadAllBytes(dstPath);
        var xmlBytes = DecodeDst(dstBytes);
        var tempXml = Path.Combine(Path.GetTempPath(), $"sse_{Guid.NewGuid():N}.xml");
        File.WriteAllBytes(tempXml, xmlBytes);
        return tempXml;
    }

    public static void SaveXmlAsDst(XDocument doc, string dstPath)
    {
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            NewLineHandling = NewLineHandling.None,
            OmitXmlDeclaration = false
        };

        using var ms = new MemoryStream();
        using (var writer = XmlWriter.Create(ms, settings))
            doc.Save(writer);

        var dstBytes = EncodeToDst(ms.ToArray());
        File.WriteAllBytes(dstPath, dstBytes);
    }

    private static byte[] Transform(byte[] input, byte[] map)
    {
        var output = new byte[input.Length];
        for (var i = 0; i < input.Length; i++)
            output[i] = map[input[i]];
        return output;
    }

    private static byte[] BuildDecodeMap()
    {
        var map = Enumerable.Range(0, 256).Select(b => (byte)b).ToArray();
        foreach (var (enc, plain) in Pairs)
            map[enc] = plain;
        return map;
    }

    private static byte[] BuildEncodeMap()
    {
        var map = Enumerable.Range(0, 256).Select(b => (byte)b).ToArray();
        foreach (var (enc, plain) in Pairs)
            map[plain] = enc;
        return map;
    }
}
