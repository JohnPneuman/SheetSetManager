# SSM_BOOT.md

Deze prompt wordt gebruikt om bij elke sessie opnieuw de juiste context in te laden voor het Sheet Set Manager project.

## Projectstructuur

- .NET Framework 4.8
- AutoCAD plugin met COM interop (`ACSMCOMPONENTS24Lib`)
- COM-gebruik is strikt gereguleerd via ComScope, TryHelper, Logger
- Sheetwijzigingen verlopen via `SheetUpdateModel` + `ApplyUpdates(...)`
- Geen Xceed meer – UI gebruikt eigen ViewModels
- Multiselect Apply is actief via TreeView
- CSV-import en AutoNumber gebruiken dezelfde applylogica
- Alles moet gelogd worden via `Logger.Log`, nooit direct naar Editor

## Verplichtingen

- Gebruik `TryHelper.Run(...)` rond elke COM-call
- COM-objecten mogen nooit gecachet of buiten scope worden gehouden
- UI werkt alleen met `IAcSmObjectId` – nooit met COM-objecten
- `Apply`, `AutoNumber`, `CSV` → allemaal via `SheetEditorHelper.ApplyUpdates(...)`
- Logging moet via `Logger.Log.Info/Debug/Error` – géén WriteMessage()

## Componenten

- `ComScope.cs` – COM object cleanup
- `TryHelper.cs` – foutafhandeling
- `Logger.cs` / `LoggerInitializer.cs` – log4net integratie
- `PropertyBagHelper.cs` – veilige read/write op custom props
- `SheetEditorHelper.cs` – centrale applylogica
- `SheetUpdateModel.cs` – complete model voor wijzigingen

✋ Verboden handelingen
❌ db.Save(null)
❌ COM-objecten vasthouden buiten lock-scope
❌ COM-sets via GetEnumerator() zonder ComScope
❌ Directe calls zonder TryHelper.Run(...)

### ✅ Gerealiseerd (april/mei 2025)

- ✔️ Multiselect werkt via Shift/Control (WPF-conform, geen checkboxes meer)
- ✔️ `Apply()` ondersteunt meerdere sheets tegelijk (ook custom props)
- ✔️ `"..."` / lege waarden worden genegeerd bij `ApplyUpdates(...)`
- ✔️ `ApplyUpdates(...)` zorgt ook voor realtime visuele boomupdate
- ✔️ `AutoNumber` gebruikt exact dezelfde logica als `Apply()`
- ✔️ PropertyGrid toont `[shared]` bij gemengde selectie (veldverschillen)
- ✔️ COM-setters worden niet meer aangeroepen bij `null` of ongewijzigd
- ✔️ Bescherming tegen `SetTitle(null)` fouten (COM-safe)

---

## **Belangrijkste lessen / don'ts (mei 2025):**

- **Zet nooit DataContext in C# op een XAML-binding-string.**
- **Altijd DataContext = _standardViewModel** bij nieuwe sheetselectie.
- **Bindings in XAML altijd direct op het veld** (`Text="{Binding NewRevisionNumber}"`), nooit via `StandardViewModel` pad of DataContext-binding.
- **Maak ALTIJD een nieuwe SheetEditViewModel bij sheetselectie** (ook bij multiselect/shared!).
- **SheetUpdateModel moet ALLE velden bevatten** die in de UI kunnen worden bewerkt.
- **Description altijd als standaardveld verwerken (GetDesc/SetDesc), NIET als custom property!**
- **Bij Apply altijd alle velden wegschrijven** indien gewijzigd, nooit alleen een subset.
- **Test altijd eerst met een enkele sheet of alles wordt weggeschreven.**
- **Bij twijfel: altijd loggen!**

📌 **Volgende stap (gepland):**
- SheetSet-niveau ondersteuning voor propertyweergave bij root-selectie (`IAcSmSheetSet`)
- `SheetPropertyGrid` uitbreiden met herkenning van SheetSet vs Sheet
- CSV-import backport naar nieuwe Apply-structuur
