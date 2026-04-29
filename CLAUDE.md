# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**BoekSolutions.SheetSetEditor** is an AutoCAD plugin (WPF-based) for editing Sheet Set Manager (SSM) properties in bulk. It allows users to manage sheet properties, apply batch updates, and perform auto-numbering operations via a visual tree interface.

### Tech Stack

- **.NET Framework 4.8** (Desktop Framework)
- **WPF** (Windows Presentation Foundation) for UI
- **AutoCAD 2024** interop via COM (ACSMCOMPONENTS24Lib)
- **log4net** for logging
- **Newtonsoft.Json** for JSON serialization
- **Extended.Wpf.Toolkit** for WPF controls

### Architecture Pattern

The codebase follows a **ViewModel + Helper pattern** (MVVM-adjacent):
- **ViewModels** (SheetEditViewModel, CustomPropertyViewModel) hold UI state and sheet data
- **Helpers** contain business logic (ComScope, TryHelper, SheetEditorHelper, PropertyBagHelper, etc.)
- **UI Controls** (XAML + C# codebehind) bind to ViewModels
- **COM interop** is strictly regulated through ComScope and TryHelper

---

## Building and Running

### Build Command

Use Visual Studio 2022 or MSBuild:

```bash
msbuild BoekSolutions.SheetSetEditor.sln /p:Configuration=Release /p:Platform=x64
```

In Visual Studio: **Ctrl + Shift + B**

### Build Configurations

- **Debug|x64** - Debug symbols, no optimization, outputs to `bin\x64\Debug\`
- **Release|x64** - Optimized, PDB only, outputs to `bin\x64\Release\`

### Post-Build Deployment

The .csproj file automatically copies the built DLL to AutoCAD plugins folder. Restart AutoCAD 2024 to load.

### Loading and Testing in AutoCAD

1. Launch AutoCAD 2024 (plugin auto-loads)
2. Run command: `SSM_UI`
3. Opens the main WPF window
4. Open a `.dst` Sheet Set file via the UI
5. Test tree rendering, selection, Apply, and Save

There is **no unit test framework**. All testing is manual via AutoCAD.

### Logs

Plugin logs to: `%APPDATA%\BoekSolutions\Logs\ssm.log` (rolling file, max 1MB with 5 backups)

---

## Directory Structure and Key Components

### Core Directories

Helpers/ - Business logic and utilities including ComScope, TryHelper, SheetEditorHelper, PropertyBagHelper, SheetHelper, TreeBuilder, SaveHelper, and DatabaseHelper.

UI/ - WPF XAML/C# views including MainWindow, Controls (SheetTreeControl, SheetPropertyGrid, AutoNumberControl, SheetSetMenuButton), and Dialogs.

ViewModels/ - UI state: SheetEditViewModel (single/multi-sheet state with Original+New pairs) and CustomPropertyViewModel.

Models/ - SheetUpdateModel (batch update container - CENTRAL DATA STRUCTURE) and AutoNumberOptions.

Config/ - log4net.config for logging configuration.

Docs/ - Project documentation in Dutch: PROJECT_GUIDELINES_v3.md, SSM_BOOT.md, and CheatSheet.md.

---

## Central Concepts and Workflows

### 1. COM Safety: ComScope and TryHelper (MANDATORY)

**Every COM interaction must follow this pattern:**

```csharp
using (var scope = new ComScope())
{
    var sheet = scope.Track(ComHelper.GetObject<IAcSmSheet>(id));
    if (sheet != null)
    {
        var number = sheet.GetNumber();
    }
}
```

**All operations must wrap errors:**

```csharp
TryHelper.Run(() =>
{
    // Your COM operation
}, "Descriptive context");

var result = TryHelper.Run(() => sheet.GetNumber(), "Get number", fallback: "");
```

**Why:** AutoCAD COM is fragile; ComScope prevents memory leaks and crashes.

### 2. SheetUpdateModel: The Central Data Structure

All property changes flow through this model. It contains SheetId, NewNumber, Title, Description, RevisionNumber, RevisionDate, IssuePurpose, Category, DoNotPlot (bool?), and CustomProperties (Dictionary).

**Apply workflow:**

1. Collect edited values into List<SheetUpdateModel>
2. Filter out null, empty, and "..." values using Safe() helper
3. Call SheetEditorHelper.ApplyUpdates(updates)
4. ApplyUpdates iterates sheets, applies changes, logs, refreshes tree
5. SaveHelper.SaveDatabase() persists to .dst file

**Key rule:** UI → SheetUpdateModel → ApplyUpdates() → COM → Save

### 3. UI Data Flow and Selection

- TreeView displays SheetSet hierarchy (built by TreeBuilder)
- Selection: Only Shift+Click and Ctrl+Click (no checkboxes)
- Single sheet: SheetEditViewModel populated with current values
- Multi-select: Shows merged values; [shared] indicates field differences
- Apply button: Calls SheetEditorHelper.ApplyUpdates()
- Save button: SaveHelper.SaveDatabase()

### 4. Logging

Logs go to two places:
- File: %APPDATA%\BoekSolutions\Logs\ssm.log (log4net)
- AutoCAD Editor: Command line

```csharp
Log.Info("message");
Log.Debug("message");
Log.Warn("message");
Log.Error("message");
```

Never call Ed.WriteMessage() directly; use Log.*.

---

## Key Constraints and Best Practices

### DO's

- Always use ComScope around COM object operations
- Always wrap COM calls in TryHelper.Run()
- Store only IAcSmObjectId in ViewModels (never COM objects)
- Create fresh SheetEditViewModel on selection change
- Use Safe() helper to filter before ApplyUpdates
- Handle IAcSmSheet2 for revision fields
- Log important operations (selections, applies, errors)

### DON'Ts

- Never cache COM objects outside a ComScope
- Never call db.Save(null) - use SaveHelper.SaveDatabase()
- Never iterate COM collections without ComScope
- Never set sheet properties to null without checking first
- Never bind COM objects directly in XAML - use ViewModels
- Never suppress exceptions - let TryHelper catch them

---

## Adding a New Sheet Property

1. Add field to SheetUpdateModel.cs
2. Add Original+New pair to SheetEditViewModel.cs
3. Update SheetEditorHelper.cs: GetSheetValues() and ApplyUpdates()
4. Add XAML binding in SheetPropertyGrid.xaml
5. Test single and multi-select scenarios

For custom properties via COM property bag:
- Read via PropertyBagHelper.GetCustomPropertyIfValid()
- Write via PropertyBagHelper.SetCustomPropertyIfValid()

---

## Recent Changes (April-June 2025)

- Removed Xceed PropertyGrid (custom XAML UI)
- Added multi-select support (Shift+Click, Ctrl+Click)
- Centralized all Apply logic via SheetUpdateModel
- Auto-numbering reuses Apply logic
- Real-time tree refresh after Apply
- [shared] indicator for mixed values
- Support for IAcSmSheet2 revision fields
- Safe filtering of null/"..." values

---

## Documentation Files

- **PROJECT_GUIDELINES_v3.md** - Detailed patterns (Dutch)
- **SSM_BOOT.md** - Context & constraints (Dutch)
- **CheatSheet.md** - VS shortcuts (Dutch)

---

## Debugging Tips

1. Check %APPDATA%\BoekSolutions\Logs\ssm.log for detailed errors
2. Logs also appear in AutoCAD Command Line
3. Log.Error() shows popup in DEBUG builds
4. Attach debugger to acad.exe (Debug → Attach to Process)
5. Set breakpoints in SheetEditorHelper.ApplyUpdates()
6. If COM objects not released: check ComScope wrapping

