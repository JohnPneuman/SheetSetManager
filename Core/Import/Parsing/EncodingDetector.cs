using System.Text;

namespace SheetSet.Core.Import.Parsing;

public static class EncodingDetector
{
    public static Encoding Detect(string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var bom = new byte[4];
        var read = fs.Read(bom, 0, 4);

        if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
        if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
            return Encoding.Unicode;        // UTF-16 LE
        if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
            return Encoding.BigEndianUnicode;

        if (TryReadAsUtf8(filePath))
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        // Fallback: Windows-1252 is common for Dutch Excel/Business Central exports
        return Encoding.GetEncoding(1252);
    }

    private static bool TryReadAsUtf8(string filePath)
    {
        try
        {
            File.ReadAllText(filePath, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true));
            return true;
        }
        catch (DecoderFallbackException)
        {
            return false;
        }
    }
}
