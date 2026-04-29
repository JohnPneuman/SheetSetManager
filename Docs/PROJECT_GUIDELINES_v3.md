# PROJECT_GUIDELINES_v3.md

## Doel
Deze richtlijnen beschrijven de structuur, logica en COM-geheugenbeheer voor de AutoCAD Sheet Set Manager plugin (SSM Editor). Ze vervangen v2 en houden rekening met het nieuwe SheetUpdateModel en multiselect ondersteuning.

---

## Belangrijkste wijzigingen t.o.v. v2

- ❌ Geen Xceed PropertyGrid meer
- ✅ Alle wijzigingen aan sheets gaan via `SheetUpdateModel`
- ✅ `SheetEditorHelper.ApplyUpdates(...)` is de centrale logica voor Apply, AutoNumber, CSV-import
- ✅ TreeView gebruikt geen checkboxen meer – shift-selectie is verplicht
- ✅ `SheetPropertyGrid` toont standaard- en custom properties zonder Xceed
- ✅ Meerdere sheets tegelijk aanpassen via `Apply()`

---

## Technische structuur

### SheetUpdateModel
```csharp
public class SheetUpdateModel {
    public IAcSmObjectId SheetId { get; set; }
    public string NewNumber { get; set; }
    public string Title { get; set; }
    public string NewDescription { get; set; }
    public Dictionary<string, string> CustomProperties { get; set; }
}
```

### Apply logica
```csharp
SheetEditorHelper.ApplyUpdates(List<SheetUpdateModel>);
```
- Sheets worden per stuk COM-safe verwerkt via `ComScope`
- Alleen gewijzigde velden worden aangepast
- Custom properties verlopen via `PropertyBagHelper.SetCustomPropertyIfValid`

---

## UI-gedrag

### SheetPropertyGrid
- Geen Xceed meer
- Maakt gebruik van `SheetEditViewModel` en `ObservableCollection<CustomPropertyViewModel>`
- `Apply()` bouwt een `SheetUpdateModel` en stuurt die naar `ApplyUpdates(...)`

### TreeView gedrag
- Checkboxen zijn verwijderd
- Shift + klik is de enige selectie-methode
- `GetSelectedSheetIds()` retourneert lijst van `IAcSmObjectId`

---

## AutoNumber & CSV
- Gebruiken exact dezelfde `SheetUpdateModel` structuur
- Geen losse `SetNumber(...)` of `SetTitle(...)` meer
- Alles verloopt via `ApplyUpdates(...)` backend

---

## COM- en foutafhandeling
- `ComScope` wordt standaard gebruikt in alle helpers
- `TryHelper.Run(...)` verplicht bij alle COM-calls
- `Logger.Log` wordt gebruikt voor alle belangrijke acties

---

## Best practices
- COM-objecten mogen nooit opgeslagen of gecachet worden
- Alleen `IAcSmObjectId` mag worden gebruikt in ViewModels of UI
- Gebruik `GetPersistObject()` altijd binnen een `ComScope`
- Nooit `db.Save(null)` gebruiken – opslaan gebeurt via `SaveHelper.SaveDatabase()`

---

## Toekomstige uitbreidingen
- Uniforme styling voor alle vensters (PropertyGrid, AutoNumber, Settings)
- Mogelijkheid om meerdere SheetSet-bestanden tegelijk te openen
- Batch import/export via CSV blijft uitbreidbaar via `SheetUpdateModel`

---

Laatste update: v3 – April 2025
Auteur: John + GPT

## 🔄 Updates april 2025

### ➕ Toegevoegd aan `SheetUpdateModel`

```csharp
public class SheetUpdateModel {
    public IAcSmObjectId SheetId { get; set; }
    public string NewNumber { get; set; }
    public string Title { get; set; }
    public string Description { get; set; } // (vanaf april 2025 direct ondersteund)
    public Dictionary<string, string> CustomProperties { get; set; }
}
```

### ✅ Gedrag in UI

- `SheetEditorHelper.GetSharedValue(...)` toont `[shared]` voor velden bij gemengde selectie
- `MainWindow.OnApplyClicked()` ondersteunt meerdere sheets, filtert `"..."`, null en lege strings
- PropertyGrid herkent automatisch enkelvoudige/meervoudige selectie
- TreeView headers worden realtime geüpdatet met fallback op bestaande waarden

### 📌 Best practices

- Gebruik `Safe(...)` helper om `"..."` of lege strings te filteren vóór update
- UI mag geen `"..."` doorsturen naar COM via `SetTitle`, `SetNumber`, etc.
- Gebruik bestaande COM-waarden (`GetNumber()`, `GetTitle()`) als fallback bij visuele updates