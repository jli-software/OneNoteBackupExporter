# CLAUDE.md

This file provides guidance to Claude Code when working with code in this repository.

## Project Overview

OneNote Backup Exporter (C# WPF Edition) – a pure C# WPF desktop application for exporting OneNote notebooks via the COM API to multiple formats (.onepkg, .xps, .pdf). This is a migration of the original Go/Wails project; the JSON-RPC subprocess layer has been eliminated – the WPF app calls `OneNoteService` directly in-process.

## Architecture

Single-project WPF app, no subprocesses:

```
MainWindow.xaml / .xaml.cs   (UI + code-behind)
    ↕ direct method calls
Services/OneNoteService.cs   (OneNote COM API wrapper)
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
├── MainWindow.xaml           Full UI layout
├── MainWindow.xaml.cs        All UI logic, async export flow
├── Models/
│   └── Models.cs             NotebookInfo, ExportResult, VersionInfo,
│                             BackupAvailability, NotebookViewModel
├── Services/
│   └── OneNoteService.cs     COM wrapper: GetNotebooks, ExportNotebook,
│                             ExportAllNotebooks, GetVersionInfo, Dispose
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
- `OneNoteService` methods are synchronous and blocking (COM calls + polling with `Thread.Sleep`)
- Always call service methods via `Task.Run(...)` from the UI layer
- Use `IProgress<string>` for progress reporting (marshals back to UI thread automatically via `Progress<T>`)
- Use `CancellationToken` for cooperative cancellation; `ct.WaitHandle.WaitOne(ms)` replaces `Thread.Sleep` in polling loops for cancellation-aware waits

### Export Process
1. `OneNoteService` is initialized at window load via `Task.Run(() => new OneNoteService())`
2. `GetNotebooks()` retrieves all notebooks via `GetHierarchy()` XML parse
3. `ExportNotebook()` calls `OpenHierarchy()` then `Publish()` with the appropriate `PublishFormat` enum
4. After `Publish()` returns (immediately), the service polls the output file until size is stable for 10 seconds (file write monitoring)
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
- `OneNoteService.Dispose()` is called in `MainWindow.Window_Closing`
- `Dispose()` calls `Marshal.ReleaseComObject()` – does NOT kill the OneNote process
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
