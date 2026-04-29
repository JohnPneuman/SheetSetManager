# SheetSetEditor — Voortgang & Vervolgplan

**Datum:** 22 april 2026  
**Status:** PoC roundtrip geslaagd ✓ · nieuwe solution bouwt ✓ · oud project definitief afgeschreven

---

## Belangrijke ontdekking (einde sessie)

**AutoCAD 2027** is geïnstalleerd op `C:\Program Files\Autodesk\AutoCAD 2027\`.  
De managed DLLs daarin (`accoremgd.dll`, `acdbmgd.dll`, `acmgd.dll`) zijn **gebouwd op .NET 10**.  
Dit maakt het oude .NET 4.8 plugin project permanent onbruikbaar — er zijn onoplosbare assembly-conflicten tussen .NET Framework 4.8 types en .NET 10 types.

**Conclusie:** het oude `BoekSolutions.SheetSetEditor.csproj` is hiermee definitief afgeschreven. Geen actie nodig, alleen negeren.

**Goed nieuws voor de plugin-wrapper:**  
De AutoCAD 2027 DLLs zijn .NET 10 — dat betekent dat de nieuwe plugin straks gewoon `net10.0-windows` target en direct de DLLs uit `C:\Program Files\Autodesk\AutoCAD 2027\` gebruikt. Geen libs-map nodig, gewoon `HintPath` naar die folder.

**Solution terugzet-probleem:**  
VS heeft de `.sln` teruggezet naar alleen het oude project. Dit is aan het einde van de sessie gecorrigeerd. Als VS dit opnieuw doet: **niet opslaan via VS**, de `.sln` handmatig bewerken of via `dotnet sln` commando's.

---

## Wat er is gedaan (deze sessie)

### 1. Architectuurbeslissing
De COM-laag (AutoCAD ACSMCOMPONENTS) wordt volledig vervangen door directe XML-manipulatie.  
Het .dst bestand is encoded XML via een byte-substitutie cipher — deze is gedocumenteerd en geïmplementeerd.  
De nieuwe app werkt **zonder AutoCAD geïnstalleerd** (standalone). AutoCAD-integratie komt via een dunne plugin-wrapper later.

### 2. Nieuwe solution structuur
`BoekSolutions.SheetSetEditor.sln` bevat nu twee projecten (oud .csproj blijft staan voor referentie):

```
Core/SheetSet.Core.csproj          .NET 10 class library — geen AutoCAD dependency
App/SheetSetEditor.App.csproj      .NET 10 WPF standalone app
Core/RoundtripTest/                Console PoC test (niet in solution, los draaien)
libs/AutoCAD/net10.0/              Hier komen AutoCAD .NET 10 DLLs (voor plugin later)
```

### 3. Core library (SheetSet.Core) — volledig gebouwd

| Bestand | Wat het doet |
|---|---|
| `Interop/DstCodec.cs` | `DecodeDst()` en `EncodeToDst()` via 256-entry lookup table + `SaveXmlAsDst()` |
| `Parsing/SheetSetParser.cs` | `Parse()` (read-only) en `ParseForEditing()` (houdt XDocument + XElement refs) |
| `Parsing/XmlUtil.cs` | `GetPropValue`, `SetPropValue`, `ReadCustomProperties`, `SetCustomProperty` |
| `Writing/SheetSetWriter.cs` | `Apply(updates)` past wijzigingen in-memory toe + `Save(doc, path)` |
| `Models/SheetSetModels.cs` | `SheetSetDocument`, `SheetInfo`, `SubsetNode`, etc. — alle met `XElement? Element` |
| `Models/SheetUpdateModel.cs` | Batch update container (Number, Title, Desc, Category, Revision, CustomProps) |
| `Import/CsvTsvImporter.cs` | CSV/TSV → `List<SheetUpdateModel>`, matcht op paginanummer |

### 4. App (SheetSetEditor.App) — gebouwd, nog minimale UI

Bouwt succesvol (`dotnet build`). Bevat:
- Toolbar: Openen, CSV/TSV import, Toepassen, Opslaan, Opslaan als
- Boom links (subsets + sheets)
- Property editor rechts (standaard velden + custom props)
- DataGrid onderaan (alle sheets, multi-select)

Nog NIET gemigreerd vanuit de oude editor:
- Shift+Ctrl multi-select in boom
- AutoNumber / renummer dialoog
- `[gedeeld]` indicator bij multi-select
- DoNotPlot checkbox
- AutoCAD plugin commando (`SSM_UI`)

### 5. Roundtrip PoC — GESLAAGD ✓

```
testset2.dst → decode → parse (5 sheets, 13 custom velden elk) 
→ wijziging (nr 1 → 1_TEST) → encode → test_roundtrip.dst 
→ AutoCAD: paginanummer is 1_TEST ✓
```

---

## Build situatie

### Nieuwe projecten (Core + App)
```bash
cd C:\Users\veerm\source\repos\BoekSolutions.SheetSetEditor
dotnet build Core/SheetSet.Core.csproj       # werkt
dotnet build App/SheetSetEditor.App.csproj   # werkt
dotnet run --project App/SheetSetEditor.App.csproj  # start standalone app
```

### Roundtrip test uitvoeren
```bash
dotnet run --project Core/RoundtripTest -- testset2.dst output.dst
```
Kopieer eerst een .dst bestand naar de solution root, of geef een volledig pad mee.

### Oud project (BoekSolutions.SheetSetEditor.csproj — .NET 4.8)
Dit bouwt nog steeds via VS 2022 / x64 Release. De build-meldingen die je zag zijn **geen fouten maar verwachte waarschuwingen**:
- `"Solution is not saved"` → in VS: File → Save All, dan pas NuGet beheren
- COM marshal warnings → normaal bij ACSMCOMPONENTS24Lib, geen actie nodig
- `MSIL vs AMD64` → zet Platform op x64 in Configuration Manager (was al zo ingesteld)

Het oude project hoeft NIET meer te bouwen voor de nieuwe aanpak. Het staat er alleen nog als referentie.

---

## Wat er morgen moet gebeuren (in volgorde)

### Stap 1 — App UI compleet maken (prioriteit)
De huidige MainWindow is een werkende shell maar mist functionaliteit van de oude editor.

**Te migreren vanuit het oude project:**

1. **Multi-select in boom via Shift+Ctrl**  
   Zie `Helpers/TreeViewSelectionHelper.cs` en `Helpers/SheetTreeHelper.cs` in het oude project.  
   WPF TreeView heeft geen ingebouwde multi-select — de logica zit in MouseDown events.

2. **`[gedeeld]` indicator**  
   Al geïmplementeerd in `MainWindow.xaml.cs` (`Mixed()` methode), maar de TextBox bindings zijn nog niet readonly-gedrag bewust.

3. **AutoNumber / renummer**  
   Zie `UI/Controls/AutoNumberControl.xaml` + `Helpers/SheetEditorHelper.cs` (AutoNumber logica) in het oude project.  
   Dit moet in de nieuwe Core komen als `AutoNumberService` die `List<SheetUpdateModel>` genereert.

4. **DoNotPlot checkbox**  
   Is al in `SheetInfo.DoNotPlot` en `SheetUpdateModel.DoNotPlot` — alleen nog XAML toevoegen.

5. **Recent files menu**  
   `Helpers/RecentFilesHelper.cs` in oud project — simpele JSON lijst in `%APPDATA%`.

### Stap 2 — AutoCAD plugin wrapper
Dunne .NET 10 plugin die alleen:
- `[CommandMethod("SSM_UI")]` registreert
- Het huidige `.dst` pad uit AutoCAD ophaalt
- `SheetSetEditor.App.exe <dstPath>` als extern proces start

**AutoCAD 2027 DLLs staan op:**  
`C:\Program Files\Autodesk\AutoCAD 2027\accoremgd.dll` (en acdbmgd, acmgd)  
→ Geen apart libs-mapje nodig, gewoon `HintPath` in de .csproj.

Plugin project aanmaken: `Plugin/SheetSetEditor.Plugin.csproj`
```xml
<TargetFramework>net10.0-windows</TargetFramework>
<PlatformTarget>x64</PlatformTarget>
<Reference Include="accoremgd">
  <HintPath>C:\Program Files\Autodesk\AutoCAD 2027\accoremgd.dll</HintPath>
  <Private>false</Private>
</Reference>
```

### Stap 3 — Bug fixes vs oud project
Bugs die opgelost zijn door de nieuwe aanpak (geen actie nodig):
- ✓ Lege custom velden worden nu correct ingelezen (XML leest altijd, ook als waarde leeg is)
- ✓ COM crashes / vastlopen verdwenen
- ✓ Kan nu ook zonder AutoCAD draaien

Bugs die nog getest moeten worden:
- Custom properties aanmaken die nog niet bestaan (huidig: alleen bestaande props worden overschreven)
- Encoding van speciale tekens (ë, ü, à etc.) in custom veld waarden → verwacht OK want UTF-8

### Stap 4 — CSV/TSV import verfijnen
De importer werkt al. Nog te doen:
- Feedbackvenster: "X sheets bijgewerkt, Y niet gevonden"
- Optie om velden leeg te laten (lege cel = niet aanpassen vs. leeg maken)
- Export naar CSV/TSV (huidige waarden downloaden als template)

---

## Sleutelpaden

| Wat | Pad |
|---|---|
| Nieuwe solution | `BoekSolutions.SheetSetEditor\BoekSolutions.SheetSetEditor.sln` |
| Core library | `BoekSolutions.SheetSetEditor\Core\` |
| App (WPF) | `BoekSolutions.SheetSetEditor\App\` |
| Oud project (referentie) | `BoekSolutions.SheetSetEditor\BoekSolutions.SheetSetEditor.csproj` |
| Viewer (decoder origine) | `sheetsetviewer\SheetSetViewer\Program.cs` |
| Test DST bestanden | `sheetsetviewer\testset2.dst` (heeft 5 sheets met custom velden) |
| AutoCAD 2027 DLLs | `C:\Program Files\Autodesk\AutoCAD 2027\` (accoremgd, acdbmgd, acmgd) |

---

## Technische notities

### Decode/encode cipher
- 121 byte-paren: `encoded_byte → plain_byte`
- Plain bytes zijn standaard ASCII (9=tab, 10=newline, 32-122=printable)
- Encode is gewoon de inverse map
- UTF-8 multi-byte sequences (128+) die niet in de tabel staan → identity (pass-through)

### XML structuur .dst
```xml
<AcSmDatabase>
  <AcSmSheetSet>
    <AcSmCustomPropertyBag>          ← sheet set custom properties
      <AcSmCustomPropertyValue propname="Veldnaam">
        <AcSmProp propname="Value">waarde</AcSmProp>
      </AcSmCustomPropertyValue>
    </AcSmCustomPropertyBag>
    <AcSmSubset>                     ← recursief nestbaar
      <AcSmSheet>
        <AcSmProp propname="Number">1</AcSmProp>
        <AcSmProp propname="Title">Titel</AcSmProp>
        <AcSmProp propname="RevisionNumber">A</AcSmProp>
        <AcSmProp propname="RevisionDate">01-01-2026</AcSmProp>
        <AcSmProp propname="Category">...</AcSmProp>
        <AcSmProp propname="IssuePurpose">...</AcSmProp>
        <AcSmCustomPropertyBag>      ← sheet-niveau custom properties
          ...
        </AcSmCustomPropertyBag>
        <AcSmAcDbLayoutReference propname="Layout">
          <AcSmProp propname="FileName">C:\pad\naar.dwg</AcSmProp>
        </AcSmAcDbLayoutReference>
      </AcSmSheet>
    </AcSmSubset>
  </AcSmSheetSet>
</AcSmDatabase>
```

### SheetSetWriter.Apply() — gedrag
- `null` waarde in `SheetUpdateModel` = veld NIET aanpassen (skip)
- Lege string `""` = veld leegmaken
- `[gedeeld]` in UI = null meesturen = skip
- Custom prop die niet bestaat in XML = stille skip (geen crash, geen aanmaak)
