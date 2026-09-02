using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using OneNoteExporter.Helpers;
using OneNoteExporter.Models;
using OneNoteExporter.Services;

namespace OneNoteExporter;

public partial class MainWindow : Window
{
    // ── State ────────────────────────────────────────────────────────────────

    private OneNoteComWorker?     _service;
    private CancellationTokenSource? _exportCts;
    private bool                  _exportInProgress = false;
    private bool                  _isLocalBackupMode = false;

    private readonly ObservableCollection<NotebookViewModel> _notebooks = new();

    // ── Startup / Shutdown ───────────────────────────────────────────────────

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        NotebookList.ItemsSource = _notebooks;

        await InitOneNoteServiceAsync();
        UpdateLocalBackupOption();
        await LoadNotebooksAsync();

        var defaultPath = FileHelper.GetDefaultDownloadsPath();
        DestPathBox.Text = Path.Combine(defaultPath, "OneNote-Export");
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        _exportCts?.Cancel();
        _service?.Dispose();
    }

    // ── OneNote initialisation ───────────────────────────────────────────────

    private async Task InitOneNoteServiceAsync()
    {
        SetVersionBadge("Checking OneNote installation...", StatusKind.Info);

        try
        {
            _service = await OneNoteComWorker.CreateAsync();
            var info = await _service.GetVersionInfoAsync();

            if (info.OneNoteInstalled)
                SetVersionBadge($"✓ {info.OneNoteVersion}", StatusKind.Success);
            else
                SetVersionBadge("⚠  OneNote Desktop not found", StatusKind.Warning);
        }
        catch (Exception ex)
        {
            _service = null;
            SetVersionBadge(
                $"⚠  {UserFacingError.Describe(ex, "OneNote Desktop connection failed.")}",
                StatusKind.Error);
        }
    }

    // ── Notebook loading ─────────────────────────────────────────────────────

    private async Task LoadNotebooksAsync()
    {
        _notebooks.Clear();
        ShowNotebookOverlay(OverlayKind.Loading);

        try
        {
            if (_service == null)
                throw new InvalidOperationException("OneNote Helper is not available.");

            var list = await _service.GetNotebooksAsync();

            ShowNotebookOverlay(OverlayKind.None);

            if (list.Count == 0)
            {
                ShowNotebookOverlay(OverlayKind.Empty);
                return;
            }

            foreach (var nb in list)
                _notebooks.Add(new NotebookViewModel(nb));

            // Re-apply localbackup UI state if that format is selected
            if (GetSelectedFormat() == "localbackup")
                ApplyLocalBackupMode(enable: true);
        }
        catch (Exception ex)
        {
            ShowNotebookOverlay(
                OverlayKind.Error,
                UserFacingError.Describe(
                    ex,
                    "Notebooks could not be loaded. Open OneNote Desktop and try again."));
        }

        UpdateExportButtonState();
    }

    // ── UI helpers ───────────────────────────────────────────────────────────

    private string GetSelectedFormat()
    {
        if (FormatCombo.SelectedItem is ComboBoxItem item)
            return item.Tag?.ToString() ?? "onepkg";
        return "onepkg";
    }

    private static bool IsRenderedFormat(string format)
        => format.Equals("pdf", StringComparison.OrdinalIgnoreCase) ||
           format.Equals("xps", StringComparison.OrdinalIgnoreCase);

    private void UpdateRenderedFormatWarning(string format)
    {
        bool isRenderedFormat = IsRenderedFormat(format);
        RenderedFormatWarning.Visibility = isRenderedFormat
            ? Visibility.Visible
            : Visibility.Collapsed;

        if (!isRenderedFormat) return;

        string formatName = format.Equals("pdf", StringComparison.OrdinalIgnoreCase)
            ? "PDF"
            : "XPS";
        RenderedFormatWarningText.Text =
            $"⚠ Whole-notebook {formatName} exports can fail in OneNote for large notebooks. " +
            "Before export, the app will offer the safer section-by-section mode.";
    }

    private void UpdateExportButtonState()
    {
        if (GetSelectedFormat() == "localbackup") return; // handled by ApplyLocalBackupMode

        // Notebooks selected for whole-notebook export
        // (a notebook that has any section selected is handled at section level instead)
        int notebookCount = _notebooks.Count(nb => nb.IsSelected && !nb.Sections.Any(s => s.IsSelected));
        int sectionCount  = _notebooks.Sum(nb => nb.Sections.Count(s => s.IsSelected));

        bool hasAny = notebookCount > 0 || sectionCount > 0;
        ExportButton.IsEnabled = hasAny && !_exportInProgress;

        ExportButton.Content = (notebookCount, sectionCount) switch
        {
            (0, 0)   => "Export Selected",
            (> 0, 0) => $"Export {notebookCount} notebook(s)",
            (0, > 0) => $"Export {sectionCount} section(s)",
            _        => $"Export {notebookCount} notebook(s)  +  {sectionCount} section(s)"
        };
    }

    private void UpdateLocalBackupOption()
    {
        var avail = FileHelper.CheckLocalBackupAvailable();

        if (avail.Available)
        {
            LocalBackupItem.Content  = $"Local Backup Copy ({avail.NotebookCount} notebooks) – Fast, direct file copy";
            LocalBackupItem.IsEnabled = true;
        }
        else
        {
            LocalBackupItem.Content  = $"Local Backup Copy – Not available ({avail.Message})";
            LocalBackupItem.IsEnabled = false;
        }
    }

    private void ApplyLocalBackupMode(bool enable)
    {
        if (enable)
        {
            _isLocalBackupMode = true;
            BackupWarning.Visibility = Visibility.Visible;

            foreach (var nb in _notebooks)
            {
                nb.IsSelected = true;
                nb.IsEnabled  = false;
                nb.IsExpanded = false;
                foreach (var s in nb.Sections) s.IsSelected = false;
            }

            var avail = FileHelper.CheckLocalBackupAvailable();
            ExportButton.Content   = avail.Available
                ? $"Export ALL Notebooks ({avail.NotebookCount} total)"
                : "Export ALL Notebooks";
            ExportButton.IsEnabled = avail.Available;
        }
        else
        {
            bool wasLocalBackupMode = _isLocalBackupMode;
            _isLocalBackupMode = false;
            BackupWarning.Visibility = Visibility.Collapsed;

            if (wasLocalBackupMode)
            {
                foreach (var nb in _notebooks)
                {
                    nb.IsSelected = false;
                    nb.IsEnabled  = true;
                }
            }

            UpdateExportButtonState();
        }
    }

    private void SetButtonsEnabled(bool enabled)
    {
        ExportButton.IsEnabled  = enabled;
        RefreshButton.IsEnabled = enabled;
        BrowseButton.IsEnabled  = enabled;
    }

    // ── Status display ───────────────────────────────────────────────────────

    private enum StatusKind { Info, Success, Warning, Error }
    private enum OverlayKind { None, Loading, Empty, Error }

    private void SetVersionBadge(string text, StatusKind kind)
    {
        VersionText.Text = text;
        (VersionBadge.Background, VersionBadge.BorderBrush, VersionText.Foreground) = kind switch
        {
            StatusKind.Success => (Brush(StaticRes("SuccessBg")), Brush(StaticRes("SuccessBorder")), Brush(StaticRes("SuccessFg"))),
            StatusKind.Warning => (Brush(StaticRes("WarningBg")), Brush(StaticRes("WarningBorder")), Brush(StaticRes("WarningFg"))),
            StatusKind.Error   => (Brush(StaticRes("ErrorBg")),   Brush(StaticRes("ErrorBorder")),   Brush(StaticRes("ErrorFg"))),
            _                  => (Brush(StaticRes("InfoBg")),    Brush(StaticRes("InfoBorder")),    Brush(StaticRes("InfoFg")))
        };
    }

    private void ShowNotebookOverlay(OverlayKind kind, string? errorMsg = null)
    {
        LoadingText.Visibility = kind == OverlayKind.Loading ? Visibility.Visible : Visibility.Collapsed;
        EmptyText.Visibility   = kind == OverlayKind.Empty   ? Visibility.Visible : Visibility.Collapsed;
        ErrorText.Visibility   = kind == OverlayKind.Error   ? Visibility.Visible : Visibility.Collapsed;

        if (kind == OverlayKind.Error && errorMsg != null)
            ErrorText.Text = errorMsg;
    }

    private void ShowProgress(string text)
    {
        ExportProgressBar.Visibility = Visibility.Visible;
        ProgressArea.Visibility      = Visibility.Visible;
        ProgressText.Text            = text;
        StatusBorder.Visibility      = Visibility.Collapsed;
    }

    private void UpdateProgressText(string text) => ProgressText.Text = text;

    private void HideProgress()
    {
        ExportProgressBar.Visibility = Visibility.Collapsed;
        ProgressArea.Visibility      = Visibility.Collapsed;
    }

    private void ShowStatus(string text, StatusKind kind)
    {
        StatusText.Text              = text;
        StatusBorder.Visibility      = Visibility.Visible;

        (StatusBorder.Background, StatusBorder.BorderBrush, StatusText.Foreground) = kind switch
        {
            StatusKind.Success => (Brush(StaticRes("SuccessBg")), Brush(StaticRes("SuccessBorder")), Brush(StaticRes("SuccessFg"))),
            StatusKind.Warning => (Brush(StaticRes("WarningBg")), Brush(StaticRes("WarningBorder")), Brush(StaticRes("WarningFg"))),
            StatusKind.Error   => (Brush(StaticRes("ErrorBg")),   Brush(StaticRes("ErrorBorder")),   Brush(StaticRes("ErrorFg"))),
            _                  => (Brush(StaticRes("InfoBg")),    Brush(StaticRes("InfoBorder")),    Brush(StaticRes("InfoFg")))
        };
    }

    private void ClearStatus()
    {
        StatusBorder.Visibility = Visibility.Collapsed;
        StatusText.Text         = "";
    }

    // Resource helpers
    private object StaticRes(string key) => FindResource(key);
    private static SolidColorBrush Brush(object res) => (SolidColorBrush)res;

    // ── Event handlers ───────────────────────────────────────────────────────

    private async void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var nb in _notebooks) nb.IsSelected = false;
        await LoadNotebooksAsync();
    }

    private async void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        // OpenFolderDialog is available in WPF on .NET 8+ (Windows)
        var dialog = new OpenFolderDialog
        {
            Title            = "Select Destination Folder",
            InitialDirectory = FileHelper.GetDefaultDownloadsPath()
        };

        if (dialog.ShowDialog(this) == true)
        {
            DestPathBox.Text = dialog.FolderName;
            ShowStatus($"Path selected: {dialog.FolderName}", StatusKind.Info);
        }
    }

    private void FormatCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Guard: may fire before UI is fully initialised
        if (BackupWarning == null) return;

        string format = GetSelectedFormat();
        bool isLocalBackup = format == "localbackup";
        ApplyLocalBackupMode(isLocalBackup);
        UpdateRenderedFormatWarning(format);
    }

    private void Notebook_SelectionChanged(object sender, RoutedEventArgs e)
        => UpdateExportButtonState();

    private void Section_SelectionChanged(object sender, RoutedEventArgs e)
        => UpdateExportButtonState();

    private void AboutButton_Click(object sender, RoutedEventArgs e)
    {
        var aboutWindow = new AboutWindow { Owner = this };
        aboutWindow.ShowDialog();
    }

    private async void ExpandSections_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.DataContext is not NotebookViewModel nb) return;

        nb.IsExpanded = !nb.IsExpanded;

        if (nb.IsExpanded && !nb.HasSectionsLoaded)
            await LoadSectionsAsync(nb);
    }

    private async Task<bool> LoadSectionsAsync(
        NotebookViewModel nb,
        CancellationToken ct = default)
    {
        if (_service == null) return false;

        nb.IsLoadingSections = true;

        try
        {
            var sections = await _service.GetSectionsAsync(nb.Id, ct);
            ct.ThrowIfCancellationRequested();

            nb.Sections.Clear();
            foreach (var s in sections)
                nb.Sections.Add(new SectionViewModel(s));

            nb.HasSectionsLoaded = true;
            return true;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            ShowStatus(
                UserFacingError.Describe(
                    ex,
                    $"Sections for '{nb.Name}' could not be loaded."),
                StatusKind.Error);
            return false;
        }
        finally
        {
            nb.IsLoadingSections = false; // fires SectionsEmpty + IsSectionsContentVisible
        }
    }

    private async Task<bool> ConvertNotebookJobsToSectionsAsync(
        List<NotebookViewModel> notebookJobs,
        List<(NotebookViewModel Notebook, SectionViewModel Section)> sectionJobs,
        string format,
        CancellationToken ct)
    {
        int notebookIndex = 0;

        foreach (var notebook in notebookJobs)
        {
            ct.ThrowIfCancellationRequested();
            notebookIndex++;
            ShowProgress(
                $"Preparing safer {format.ToUpperInvariant()} export...\n" +
                $"Loading sections for notebook {notebookIndex}/{notebookJobs.Count}: " +
                notebook.Name);

            if (!notebook.HasSectionsLoaded && !await LoadSectionsAsync(notebook, ct))
            {
                HideProgress();
                return false;
            }

            if (notebook.Sections.Count == 0)
            {
                HideProgress();
                ShowStatus(
                    $"No exportable sections were found in '{notebook.Name}'. " +
                    "Use a OneNote package for a complete backup.",
                    StatusKind.Warning);
                return false;
            }

            ct.ThrowIfCancellationRequested();

            foreach (var section in notebook.Sections)
            {
                bool alreadyQueued = sectionJobs.Any(job => job.Section.Id == section.Id);
                if (!alreadyQueued)
                    sectionJobs.Add((notebook, section));
            }
        }

        notebookJobs.Clear();
        return true;
    }

    private async Task<(bool Continue, string Format)> ApplyRenderedExportGuardAsync(
        List<NotebookViewModel> notebookJobs,
        List<(NotebookViewModel Notebook, SectionViewModel Section)> sectionJobs,
        string format,
        CancellationToken ct)
    {
        if (!IsRenderedFormat(format) || notebookJobs.Count == 0)
            return (true, format);

        var warningWindow = new RenderedExportWarningWindow(
            format,
            sectionJobs.Count > 0)
        {
            Owner = this
        };
        warningWindow.ShowDialog();

        switch (warningWindow.Choice)
        {
            case RenderedExportChoice.IndividualSections:
                bool converted = await ConvertNotebookJobsToSectionsAsync(
                    notebookJobs, sectionJobs, format, ct);
                return (converted, format);

            case RenderedExportChoice.OneNotePackage:
                SelectFormat("onepkg");
                return (true, "onepkg");

            case RenderedExportChoice.WholeNotebook:
                return (true, format);

            default:
                return (false, format);
        }
    }

    private void SelectFormat(string format)
    {
        var item = FormatCombo.Items
            .OfType<ComboBoxItem>()
            .FirstOrDefault(candidate => string.Equals(
                candidate.Tag?.ToString(), format, StringComparison.OrdinalIgnoreCase));

        if (item != null)
            FormatCombo.SelectedItem = item;
    }

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        if (_exportInProgress)
        {
            ShowStatus("⚠  An export is already running! Please wait until it is completed.", StatusKind.Warning);
            return;
        }

        var destPath = DestPathBox.Text.Trim();
        if (string.IsNullOrEmpty(destPath))
        {
            ShowStatus("Please specify a destination folder.", StatusKind.Error);
            return;
        }

        var format = GetSelectedFormat();

        // For COM export: build separate job lists for notebooks and sections.
        // Rule: if a notebook has any selected sections, export those sections individually
        //       (the notebook-level checkbox is ignored for that notebook).
        List<NotebookViewModel>? notebookJobs = null;
        List<(NotebookViewModel Notebook, SectionViewModel Section)>? sectionJobs = null;

        if (format != "localbackup")
        {
            notebookJobs = _notebooks
                .Where(nb => nb.IsSelected && !nb.Sections.Any(s => s.IsSelected))
                .ToList();

            sectionJobs = _notebooks
                .SelectMany(nb => nb.Sections
                    .Where(s => s.IsSelected)
                    .Select(s => (Notebook: nb, Section: s)))
                .ToList();

            if (notebookJobs.Count == 0 && sectionJobs.Count == 0)
            {
                ShowStatus("Please select at least one notebook or section.", StatusKind.Error);
                return;
            }
        }

        // Lock UI
        _exportInProgress       = true;
        _exportCts              = new CancellationTokenSource();
        CancelButton.Visibility = Visibility.Visible;
        SetButtonsEnabled(false);
        ClearStatus();

        try
        {
            if (format == "localbackup")
                await RunLocalBackupAsync(destPath, _exportCts.Token);
            else
            {
                var prepared = await ApplyRenderedExportGuardAsync(
                    notebookJobs!, sectionJobs!, format, _exportCts.Token);

                if (!prepared.Continue)
                    return;

                _exportCts.Token.ThrowIfCancellationRequested();
                await RunComExportAsync(
                    notebookJobs!, sectionJobs!, destPath, prepared.Format, _exportCts.Token);
            }
        }
        catch (OperationCanceledException)
        {
            HideProgress();
            ShowStatus("Export cancelled.", StatusKind.Warning);
        }
        finally
        {
            _exportInProgress       = false;
            CancelButton.Visibility = Visibility.Collapsed;
            SetButtonsEnabled(true);
            UpdateExportButtonState();
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_exportInProgress) return;

        string cancellationDetails = _isLocalBackupMode
            ? "The export will stop safely after the current file copy finishes."
            : "The export will stop safely without closing OneNote. If OneNote is currently " +
              "rendering, cancellation completes as soon as OneNote returns control.";

        var confirm = MessageBox.Show(
            $"Do you really want to cancel the export?\n\n{cancellationDetails}",
            "Cancel Export",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (confirm != MessageBoxResult.Yes) return;

        ShowStatus("Cancelling export...", StatusKind.Warning);

        _exportCts?.Cancel();
    }

    // ── Export implementations ───────────────────────────────────────────────

    private async Task RunLocalBackupAsync(string destPath, CancellationToken ct)
    {
        ShowProgress("Copying local backup files...");

        try
        {
            var result = await Task.Run(() => FileHelper.CopyLocalBackup(destPath, ct), ct);
            HideProgress();

            if (result.Success)
            {
                ShowStatus($"✓ Local backup successfully copied!\n\nExported to: {destPath}", StatusKind.Success);
                FileHelper.OpenFolder(destPath);
            }
            else
            {
                ShowStatus($"❌ Export failed: {result.Message}", StatusKind.Error);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            HideProgress();
            ShowStatus(
                $"❌ {UserFacingError.Describe(ex, "The local backup export could not be completed.")}",
                StatusKind.Error);
        }
    }

    private async Task RunComExportAsync(
        List<NotebookViewModel> notebookJobs,
        List<(NotebookViewModel Notebook, SectionViewModel Section)> sectionJobs,
        string destPath,
        string format,
        CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();

        if (_service == null)
        {
            ShowStatus("OneNote Helper is not available. Please ensure OneNote Desktop is installed.", StatusKind.Error);
            return;
        }

        // Auto-dismiss the OneNote sync-warning dialog ("Ja" / "Nein") if it appears
        StartDialogWatcher(ct);

        int successCount = 0, failCount = 0;
        var messages     = new List<string>();
        int totalJobs    = notebookJobs.Count + sectionJobs.Count;
        int jobIndex     = 0;

        // ── Notebook-level exports ────────────────────────────────────────────
        foreach (var nb in notebookJobs)
        {
            ct.ThrowIfCancellationRequested();
            jobIndex++;

            var label    = $"Notebook {jobIndex}/{totalJobs}: {nb.Name}";
            var progress = new Progress<string>(msg =>
                UpdateProgressText($"{label}\n{msg}\nPlease be patient, this may take several minutes..."));

            ShowProgress($"{label}\nStarting export...\nPlease be patient, this may take several minutes...");

            try
            {
                var result = await _service.ExportNotebookAsync(
                    nb.Id, destPath, format, progress, ct);

                if (result.Success)
                {
                    successCount++;
                    string retryNote = result.RecoveredAfterRetry
                        ? " (recovered after COM retry)"
                        : "";
                    messages.Add($"✓ {nb.Name}{retryNote}");
                }
                else                { failCount++;    messages.Add($"✗ {nb.Name}: {result.Message}"); }
            }
            catch (OperationCanceledException)
            {
                HideProgress();
                ShowStatus("Export cancelled.", StatusKind.Warning);
                return;
            }
            catch (Exception ex)
            {
                failCount++;
                messages.Add(
                    $"✗ {nb.Name}: " +
                    UserFacingError.Describe(ex, "The export could not be completed."));
            }

            await Task.Delay(300, CancellationToken.None);
        }

        // ── Section-level exports ─────────────────────────────────────────────
        foreach (var (nb, sec) in sectionJobs)
        {
            ct.ThrowIfCancellationRequested();
            jobIndex++;

            var label    = $"Section {jobIndex}/{totalJobs}: {sec.Name}  ({nb.Name})";
            var progress = new Progress<string>(msg =>
                UpdateProgressText($"{label}\n{msg}\nPlease be patient, this may take several minutes..."));

            ShowProgress($"{label}\nStarting export...\nPlease be patient, this may take several minutes...");

            try
            {
                var result = await _service.ExportSectionAsync(
                    sec.Info, destPath, format, progress, ct);

                if (result.Success)
                {
                    successCount++;
                    string retryNote = result.RecoveredAfterRetry
                        ? " (recovered after COM retry)"
                        : "";
                    messages.Add($"✓ {nb.Name}  /  {sec.Name}{retryNote}");
                }
                else                { failCount++;    messages.Add($"✗ {nb.Name}  /  {sec.Name}: {result.Message}"); }
            }
            catch (OperationCanceledException)
            {
                HideProgress();
                ShowStatus("Export cancelled.", StatusKind.Warning);
                return;
            }
            catch (Exception ex)
            {
                failCount++;
                messages.Add(
                    $"✗ {nb.Name}  /  {sec.Name}: " +
                    UserFacingError.Describe(ex, "The export could not be completed."));
            }

            await Task.Delay(300, CancellationToken.None);
        }

        HideProgress();

        var finalMsg = $"Export completed: {successCount} successful, {failCount} failed\n\n" +
                       string.Join("\n", messages);
        ShowStatus(finalMsg, failCount == 0 ? StatusKind.Success : StatusKind.Warning);

        if (successCount > 0)
            FileHelper.OpenFolder(destPath);
    }

    // ── Utilities ────────────────────────────────────────────────────────────

    private void StartDialogWatcher(CancellationToken ct)
    {
        _ = Task.Run(() =>
        {
            while (!ct.IsCancellationRequested && _exportInProgress)
            {
                try
                {
                    var allWindows = AutomationElement.RootElement.FindAll(
                        TreeScope.Children,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                    foreach (AutomationElement window in allWindows)
                    {
                        if (!window.Current.Name.Contains("OneNote", StringComparison.OrdinalIgnoreCase))
                            continue;

                        var jaButton = window.FindFirst(
                            TreeScope.Descendants,
                            new AndCondition(
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                                new PropertyCondition(AutomationElement.NameProperty, "Ja")));

                        if (jaButton != null &&
                            jaButton.TryGetCurrentPattern(InvokePattern.Pattern, out var pattern))
                        {
                            ((InvokePattern)pattern).Invoke();
                            Debug.WriteLine("OneNote sync-warning dialog auto-dismissed.");
                        }
                    }
                }
                catch { /* dialog may have closed between find and invoke */ }

                ct.WaitHandle.WaitOne(300);
            }
        });
    }

}
