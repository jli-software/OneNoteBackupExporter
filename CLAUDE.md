# CLAUDE.md

This file provides guidance to Claude Code when working with code in this repository.

## Project Overview

OneNote Backup Exporter (C# WPF Edition) – a pure C# WPF desktop application for exporting OneNote notebooks via the COM API to multiple formats (.onepkg, .xps, .pdf). This is a migration of the original Go/Wails project; the JSON-RPC subprocess layer has been eliminated – the WPF app calls `OneNoteService` directly in-process.

## Architecture

Single-project WPF app, no subprocesses:

```
MainWindow.xaml / .xaml.cs   (UI + code-behind)
    ↕ async calls
Services/OneNoteComWorker.cs (serialized STA thread + message pump)
    ↕ synchronous calls on the owning STA thread
Services/OneNoteService.cs   (low-level OneNote COM API wrapper)
    ↕ COM API
OneNote Desktop
```

### File Structure

```
OneNoteExporterC#/
├── OneNoteExporter.csproj   Target: net10.0-windows, UseWPF, PlatformTarget=x64
├── OneNoteExporter.sln
├── App.xaml                  Global styles (colors, buttons, inputs)
├── App.xaml.cs               Global exception handler
├── AboutWindow.xaml/.cs      About dialog with runtime version/build metadata
├── MainWindow.xaml           Full UI layout
├── MainWindow.xaml.cs        All UI logic, async export flow
├── Models/
│   └── Models.cs             NotebookInfo, ExportResult, VersionInfo,
│                             BackupAvailability, NotebookViewModel
├── Services/
│   ├── OneNoteComWorker.cs   Dedicated STA owner, serialized async facade
│   └── OneNoteService.cs     Low-level hierarchy and publish COM calls
└── Helpers/
    └── FileHelper.cs         Local backup path, CopyLocalBackup, OpenFolder,
                              GetDefaultDownloadsPath
```

## Development Commands

```powershell
# Build
dotnet build

# Run (debug)
dotnet run

# Publish self-contained x64
dotnet publish -c Release -r win-x64 --self-contained true
```

## Critical Technical Details

### COM Compatibility
- **`PlatformTarget=x64` is mandatory** – this build targets current 64-bit OneNote Desktop installations. Changing the platform can cause "Class not registered" COM errors.
- **NuGet package:** `Interop.Microsoft.Office.Interop.OneNote` v1.1.0.2 (NOT `Microsoft.Office.Interop.OneNote`)
- **Do NOT add `<UseWindowsForms>true</UseWindowsForms>`** – causes namespace conflict with the COM `Application` class

### Threading Model
- `OneNoteComWorker` owns `OneNoteService` and its COM object on one dedicated STA thread
- All OneNote COM calls are dispatched through that thread's WPF `Dispatcher` message pump
- Public worker methods are async and serialized; do not wrap them in `Task.Run(...)`
- Export file polling runs asynchronously outside the STA thread while the serialization gate remains held
- Use `IProgress<string>` for progress reporting (marshals back to UI thread automatically via `Progress<T>`)
- Use `CancellationToken` for cooperative cancellation and `Task.Delay(..., ct)` in polling loops

### Export Process
1. `OneNoteComWorker.CreateAsync()` starts the STA thread and creates `OneNoteService` there
2. `GetNotebooksAsync()` retrieves all notebooks via `GetHierarchy()` XML parse on the worker
3. `ExportNotebookAsync()` dispatches `OpenHierarchy()` and `Publish()` to the STA worker
4. After `Publish()` returns, async file polling waits until size is stable for 10 seconds while the STA message pump stays responsive
5. Timeout: 20 min for local notebooks, 30 min for cloud (SharePoint/OneDrive)

### Export Formats
| UI value | PublishFormat | Extension |
|---|---|---|
| `"onepkg"` | `pfOneNotePackage` | `.onepkg` |
| `"xps"` | `pfXPS` | `.xps` |
| `"pdf"` | `pfPDF` | `.pdf` |
| `"localbackup"` | n/a (file copy) | n/a |

### Local Backup Mode
- Copies `%LOCALAPPDATA%\Microsoft\OneNote\16.0\Sicherung` to destination
- Implemented in `FileHelper.CopyLocalBackup()`
- When selected in dropdown: all notebooks auto-checked, checkboxes disabled, warning shown

### OneNote COM Cleanup
- `OneNoteComWorker.Dispose()` is called in `MainWindow.Window_Closing`
- The worker releases the COM object on its owning STA thread and does NOT kill OneNote
- OneNote continues running in background after app closes (by design, avoids "could not be started on last attempt" error)

### Known COM Errors
| HResult | Cause | Action |
|---|---|---|
| `0x8004201A` | Password-protected sections or offline | Show message, suggest unlocking |
| `0x800706BA` | RPC timeout (large notebook, PDF/XPS) | Suggest .onepkg format |
| `0x80070005` | Access denied | Check permissions |
| `0x80042010` | Notebook not accessible | Open in OneNote first |

## UI Design

### Layout (3-row Grid)
```
Row 0 (Auto):   Header – title (#0078d7) + version badge (dynamic color)
Row 1 (*):      Notebook list (ScrollViewer + ItemsControl with CheckBoxes)
Row 2 (Auto):   Export section (path, format dropdown, warning, buttons, status)
```

### Color Scheme (defined as SolidColorBrush in App.xaml)
- Title: `#0078D7` (Microsoft blue)
- Primary button: `#007BFF`
- Cancel button: `#DC3545` (red)
- Browse button: `#5A5A5A` (gray)
- Success: green tones (`#D4EDDA` / `#155724`)
- Warning: yellow tones (`#FFF3CD` / `#856404`)
- Error: red tones (`#F8D7DA` / `#721C24`)
- Info/progress: cyan tones (`#D1ECF1` / `#0C5460`)
- Page background: `#F5F5F5`
- Card background: `White`
- List background: `#F8F9FA`

### UI Language
**English** – all UI strings are in English (matching the original Go/Wails frontend).

### Key Behaviors
- Export button disabled until ≥1 notebook selected; text = `"Export N notebook(s)"`
- Cancel button hidden normally, shown during export
- `localbackup` format: auto-checks all notebooks, disables checkboxes, shows warning banner
- Progress: `ProgressBar` (IsIndeterminate) + `TextBlock` during export
- Final status: color-coded border (green/yellow/red)
- Explorer opens once after all successful exports

## Window Configuration
- Width: 900, Height: 800
- MinWidth: 700, MinHeight: 500
- MaxWidth: 1280, MaxHeight: 1100

## Platform
Windows-only (requires OneNote Desktop 2016, .NET 10 Windows, COM infrastructure).
