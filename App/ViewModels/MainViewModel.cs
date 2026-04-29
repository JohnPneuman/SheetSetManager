using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using SheetSet.Core.Models;
using SheetSetEditor.Models;

namespace SheetSetEditor.ViewModels;

public sealed class MainViewModel : INotifyPropertyChanged
{
    private DocumentContext? _activeDocument;
    private string _statusText = "Geen bestand geladen";

    // All open documents
    public ObservableCollection<DocumentContext> OpenDocuments { get; } = [];

    public DocumentContext? ActiveDocument
    {
        get => _activeDocument;
        set
        {
            _activeDocument = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(FileTitle));
            OnPropertyChanged(nameof(Title));
            OnPropertyChanged(nameof(IsDocumentLoaded));
            OnPropertyChanged(nameof(IsDirty));
        }
    }

    public bool IsDocumentLoaded => _activeDocument != null;

    public bool IsDirty
    {
        get => _activeDocument?.IsDirty ?? false;
        set
        {
            if (_activeDocument == null) return;
            _activeDocument.IsDirty = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(Title));
        }
    }

    public string StatusText
    {
        get => _statusText;
        set { _statusText = value; OnPropertyChanged(); }
    }

    public string FileTitle
    {
        get
        {
            if (_activeDocument == null) return "Geen bestand ▼";
            var dirty = _activeDocument.IsDirty ? " *" : "";
            return $"{_activeDocument.DisplayName}{dirty} ▼";
        }
    }

    public string Title => _activeDocument == null
        ? "SheetSet Editor"
        : $"SheetSet Editor — {_activeDocument.DisplayName}{(_activeDocument.IsDirty ? " *" : string.Empty)}";

    // ── Document management ───────────────────────────────────────────────────

    public DocumentContext AddDocument(SheetSetDocument doc)
    {
        var ctx = new DocumentContext(doc);
        OpenDocuments.Add(ctx);
        ActiveDocument = ctx;
        var count = doc.GetAllSheets().Count();
        StatusText = $"{count} sheets geladen uit {Path.GetFileName(doc.SourceDstPath)}";
        return ctx;
    }

    public void CloseDocument(DocumentContext ctx)
    {
        OpenDocuments.Remove(ctx);
        if (ActiveDocument == ctx)
            ActiveDocument = OpenDocuments.LastOrDefault();
        if (OpenDocuments.Count == 0)
            StatusText = "Geen bestand geladen";
    }

    public void SetActive(DocumentContext ctx)
    {
        if (!OpenDocuments.Contains(ctx)) return;
        ActiveDocument = ctx;
    }

    public void MarkDirty()
    {
        IsDirty = true;
        OnPropertyChanged(nameof(FileTitle));
        OnPropertyChanged(nameof(Title));
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
