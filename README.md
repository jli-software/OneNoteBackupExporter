<div align="center">

<img src="icon.ico" alt="OneNote Backup Exporter" width="100" />

# OneNote Backup Exporter

A Windows desktop application to export OneNote notebooks to portable formats.  
Talks directly to OneNote Desktop via the COM API.

![version](https://img.shields.io/badge/version-1.3.0-blueviolet?style=flat-square)
![platform](https://img.shields.io/badge/platform-Windows%2010%2F11-0078D7?style=flat-square&logo=windows&logoColor=white)
![license](https://img.shields.io/badge/license-MIT-brightgreen?style=flat-square)
![built with](https://img.shields.io/badge/built%20with-.NET%2010%20%2F%20WPF-512BD4?style=flat-square&logo=dotnet&logoColor=white)

</div>

---

![OneNote Backup Exporter main window](demoScreenshot.png)

## Download

[Download the latest installer](https://github.com/jli-software/OneNoteBackupExporter/releases/latest/download/OneNoteBackupExporter_Setup.exe)
or view the [latest release](../../releases/latest). The public asset names remain stable
across releases:

- `OneNoteBackupExporter_Setup.exe`
- `OneNoteBackupExporter_Setup.exe.sha256`

The matching [SHA-256 checksum](https://github.com/jli-software/OneNoteBackupExporter/releases/latest/download/OneNoteBackupExporter_Setup.exe.sha256)
can be used to verify the download. The installed application version is available in the
About dialog.

## Automated builds and releases

GitHub Actions runs the automated tests, builds a self-contained Windows x64 application,
and compiles the Inno Setup installer for every pull request and every push to `main`. The
installer and its SHA-256 checksum are available as workflow artifacts for 14 days.

To publish a release:

1. Update `<Version>` in `OneNoteExporter.csproj` and commit the change.
2. Create and push a matching tag, for example `v1.3.0`.
3. GitHub Actions builds the installer, creates the GitHub release with generated notes, and
   uploads the installer plus checksum automatically.

The tag and project versions must match; otherwise the release build fails without publishing.

## Requirements

- Windows 10 or 11 (64-bit)
- OneNote Desktop 2016 or later (including Microsoft 365)

The Microsoft Store version of OneNote is not supported.

## Export Formats

| Format | Notes |
|---|---|
| OneNote Package (.onepkg) | Recommended. Full-fidelity, re-importable into OneNote. |
| HTML Web Pages | Browser-readable page folders plus generated notebook/section indexes. OneNote writes each page as `.htm` with any companion assets beside it. |
| XPS Document (.xps) | Good layout preservation. Large whole notebooks can be split into safer per-section exports. |
| PDF Document (.pdf) | Universal, slightly lower quality on complex pages. Large whole notebooks can be split into safer per-section exports. |
| Local Backup Copy | Copies the raw backup folder from `%LOCALAPPDATA%\Microsoft\OneNote\16.0\Sicherung`. Fastest option, all notebooks at once. |

## Notes

- Password-protected sections cannot be exported. Unlock them in OneNote first.
- Cloud notebooks (OneDrive, SharePoint) must be synced before export. Large notebooks can take several minutes.
- HTML exports create a `<Notebook>__<stable-id>_html` folder with an `index.html` start
  page. The stable suffix prevents same-named notebooks from overwriting each other. Complete
  notebook exports link directly to every page; selected-section exports create or refresh
  the notebook start page with links to the available section indexes. Each OneNote page is
  published separately so notebook and section selections work even though OneNote's COM
  API only provides reliable HTML publishing at page level. Section groups remain folders;
  OneNote subpage levels are currently listed as normal pages. HTML is intended for offline
  browser access and is not a re-importable replacement for `.onepkg`.
- OneNote does not expose notebook size through its COM interface. Before a whole-notebook
  PDF/XPS export, the app therefore offers to export each section separately (recommended),
  create a complete `.onepkg` backup, or consciously try the whole rendered export. Split
  exports include all exportable, unlocked sections and preserve section groups as folders
  to avoid filename collisions. Use `.onepkg` when a single complete backup is required.
