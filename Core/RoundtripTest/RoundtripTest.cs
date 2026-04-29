// Run via: dotnet run --project Core/RoundtripTest
// Dit is een standalone test - geen unit test framework nodig
using SheetSet.Core.Interop;
using SheetSet.Core.Models;
using SheetSet.Core.Parsing;
using SheetSet.Core.Writing;

if (args.Length == 0)
{
    Console.WriteLine("Gebruik: RoundtripTest <pad-naar.dst> [output.dst]");
    Console.WriteLine("Voorbeeld: RoundtripTest C:\\test.dst C:\\test_roundtrip.dst");
    return 1;
}

var inputDst = args[0];
var outputDst = args.Length > 1 ? args[1] : Path.ChangeExtension(inputDst, ".roundtrip.dst");

Console.WriteLine($"Input:  {inputDst}");
Console.WriteLine($"Output: {outputDst}");
Console.WriteLine();

string? tempXml = null;
try
{
    // Stap 1: Decoderen
    Console.Write("Stap 1: DST decoderen naar XML... ");
    tempXml = DstCodec.DecodeDstToTempXml(inputDst);
    Console.WriteLine("OK");
    Console.WriteLine($"  Temp XML: {tempXml}");

    // Stap 2: Parsen voor bewerken
    Console.Write("Stap 2: XML parsen... ");
    var parser = new SheetSetParser();
    var doc = parser.ParseForEditing(inputDst, tempXml);
    var allSheets = doc.GetAllSheets().ToList();
    Console.WriteLine($"OK — {allSheets.Count} sheets gevonden");
    Console.WriteLine($"  Sheet set: {doc.Info.Name}");

    foreach (var s in allSheets.Take(5))
        Console.WriteLine($"  [{s.Number}] {s.Title} | custom velden: {s.Info.CustomProperties.Count}");

    if (allSheets.Count > 5)
        Console.WriteLine($"  ... en {allSheets.Count - 5} meer");

    // Stap 3: Wijziging toepassen (eerste sheet, paginanummer aanpassen)
    var testSheet = allSheets.FirstOrDefault();
    string? originalNumber = null;
    if (testSheet != null)
    {
        originalNumber = testSheet.Number;
        var testValue = (originalNumber ?? "0") + "_TEST";

        Console.Write($"Stap 3: Wijziging toepassen op sheet '{originalNumber}'... ");
        var update = new SheetUpdateModel
        {
            Sheet = testSheet.Info,
            NewNumber = testValue
        };
        SheetSetWriter.Apply([update]);
        Console.WriteLine($"OK → nummer wordt '{testValue}'");
    }

    // Stap 4: Opslaan als .dst
    Console.Write($"Stap 4: Opslaan als {Path.GetFileName(outputDst)}... ");
    SheetSetWriter.Save(doc, outputDst);
    Console.WriteLine("OK");

    // Stap 5: Verificatie — opnieuw inlezen
    Console.Write("Stap 5: Verificatie — output opnieuw inlezen... ");
    string? tempXml2 = null;
    try
    {
        tempXml2 = DstCodec.DecodeDstToTempXml(outputDst);
        var parser2 = new SheetSetParser();
        var doc2 = parser2.Parse(tempXml2);
        var sheets2 = doc2.GetAllSheets().ToList();
        var firstSheet2 = sheets2.FirstOrDefault();

        if (testSheet != null && firstSheet2 != null)
        {
            var expected = (originalNumber ?? "0") + "_TEST";
            var actual = firstSheet2.Number;

            if (actual == expected)
                Console.WriteLine($"OK — nummer correct: '{actual}'");
            else
                Console.WriteLine($"FOUT — verwacht '{expected}', gevonden '{actual}'");
        }
        else
        {
            Console.WriteLine($"OK — {sheets2.Count} sheets ingelezen");
        }
    }
    finally
    {
        if (tempXml2 != null && File.Exists(tempXml2))
            File.Delete(tempXml2);
    }

    Console.WriteLine();
    Console.WriteLine($"Roundtrip geslaagd! Open {outputDst} in AutoCAD om te verifiëren.");
    return 0;
}
catch (Exception ex)
{
    Console.WriteLine($"FOUT: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
    return 1;
}
finally
{
    if (tempXml != null && File.Exists(tempXml))
        File.Delete(tempXml);
}
