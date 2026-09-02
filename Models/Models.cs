using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace OneNoteExporter.Models;

public class NotebookInfo
{
    public string Id           { get; set; } = "";
    public string Name         { get; set; } = "";
    public string Path         { get; set; } = "";
    public string LastModified { get; set; } = "";
    public bool IsCurrentlyViewed { get; set; }
}

public class SectionInfo
{
    public string Id           { get; set; } = "";
    public string Name         { get; set; } = "";
    public string NotebookId   { get; set; } = "";
    public string NotebookName { get; set; } = "";
    public string NotebookPath { get; set; } = "";
    public string GroupName      { get; set; } = ""; // empty when section is directly in notebook
    public string GroupExportPath { get; set; } = ""; // collision-safe HTML group path
    public string ExportFileStem { get; set; } = ""; // collision-safe; UI continues to show Name
    public bool   IsCloud        { get; set; }
}

public class ExportResult
{
    public bool   Success      { get; set; }
    public string Message      { get; set; } = "";
    public string ExportedPath { get; set; } = "";
    public int    AttemptCount { get; set; } = 1;
    public bool   RecoveredAfterRetry { get; set; }
}

public class VersionInfo
{
    public bool   OneNoteInstalled  { get; set; }
    public string OneNoteVersion    { get; set; } = "";
}

public class BackupAvailability
{
    public bool   Available      { get; set; }
    public string Reason         { get; set; } = "";
    public string Message        { get; set; } = "";
    public string Path           { get; set; } = "";
    public int    NotebookCount  { get; set; }
}

/// <summary>
/// Wraps SectionInfo for the WPF section list with selection state.
/// </summary>
public class SectionViewModel : INotifyPropertyChanged
{
    private bool _isSelected;

    public SectionInfo Info { get; }
    public string Id   => Info.Id;
    public string Name => Info.Name;

    public string DisplayName => string.IsNullOrEmpty(Info.GroupName)
        ? $"📄 {Info.Name}"
        : $"📂 {Info.GroupName}  /  📄 {Info.Name}";

    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; OnPropertyChanged(); }
    }

    public SectionViewModel(SectionInfo info) => Info = info;

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

/// <summary>
/// Wraps NotebookInfo for the WPF notebook list with selection and section-expand state.
/// </summary>
public class NotebookViewModel : INotifyPropertyChanged
{
    private bool _isSelected;
    private bool _isEnabled = true;
    private bool _isExpanded;
    private bool _isLoadingSections;

    public NotebookInfo Info { get; }

    /// <summary>Sections loaded lazily when the user first expands this notebook.</summary>
    public ObservableCollection<SectionViewModel> Sections { get; } = new();

    /// <summary>True once GetSections() has been called for this notebook.</summary>
    public bool HasSectionsLoaded { get; set; } = false;

    public string Id   => Info.Id;
    public string Name => Info.Name;

    public string DisplayName =>
        Info.IsCurrentlyViewed
            ? $"📓 {Info.Name}  (currently open)"
            : $"📓 {Info.Name}";

    public string LastModifiedFormatted
    {
        get
        {
            if (DateTime.TryParse(Info.LastModified, null,
                    System.Globalization.DateTimeStyles.RoundtripKind, out var dt))
                return $"Last modified: {dt.ToLocalTime():dd.MM.yyyy HH:mm}";
            return "Last modified: Unknown";
        }
    }

    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; OnPropertyChanged(); }
    }

    public bool IsEnabled
    {
        get => _isEnabled;
        set { _isEnabled = value; OnPropertyChanged(); }
    }

    public bool IsExpanded
    {
        get => _isExpanded;
        set
        {
            _isExpanded = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsSectionsContentVisible));
            OnPropertyChanged(nameof(SectionsEmpty));
        }
    }

    public bool IsLoadingSections
    {
        get => _isLoadingSections;
        set
        {
            _isLoadingSections = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsSectionsContentVisible));
            OnPropertyChanged(nameof(SectionsEmpty));
        }
    }

    /// <summary>True when sections panel should show the list (expanded + loaded + not loading).</summary>
    public bool IsSectionsContentVisible => IsExpanded && HasSectionsLoaded && !IsLoadingSections;

    /// <summary>True when expanded and loaded but no sections found.</summary>
    public bool SectionsEmpty => IsExpanded && HasSectionsLoaded && !IsLoadingSections && Sections.Count == 0;

    public NotebookViewModel(NotebookInfo info) => Info = info;

    public event PropertyChangedEventHandler? PropertyChanged;
    public void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
