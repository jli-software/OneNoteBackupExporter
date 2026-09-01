<div align="center">

<img src="icon.ico" alt="OneNote Backup Exporter" width="100" />

# OneNote Backup Exporter

A Windows desktop application to export OneNote notebooks to portable formats.  
Talks directly to OneNote Desktop via the COM AP.

![version](https://img.shields.io/badge/version-1.2.0-blueviolet?style=flat-square)
![platform](https://img.shields.io/badge/platform-Windows%2010%2F11-0078D7?style=flat-square&logo=windows&logoColor=white)
![license](https://img.shields.io/badge/license-MIT-brightgreen?style=flat-square)
![built with](https://img.shields.io/badge/built%20with-.NET%2010%20%2F%20WPF-512BD4?style=flat-square&logo=dotnet&logoColor=white)

</div>

---

![OneNote Backup Exporter main window](demoScreenshot.png)

## Download

Go to [Releases](../../releases/latest) and download the latest
`OneNoteBackupExporter_Setup_<version>.exe`. Run the installer, done. A matching
`.sha256` file is included so the download can be verified.

## Automated builds and releases

GitHub Actions builds a self-contained Windows x64 application and compiles the Inno Setup
installer for every pull request and every push to `main`. The installer and its SHA-256
checksum are available as workflow artifacts for 14 days.

To publish a release:

1. Update `<Version>` in `OneNoteExporter.csproj` and commit the change.
2. Create and push a matching tag, for example `v1.2.0`.
3. GitHub Actions builds the installer, creates the GitHub release with generated notes, and
   uploads the installer plus checksum automatically.

The tag and project versions must match; otherwise the release build fails without publishing.

## Requirements

- Windows 10 or 11 (64-bit)
- OneNote Desktop 2016 or later (including Microsoft 365)

The Microsoft Store version of OneNote is not supported.

## Export Formats

![Available export formats](FormatDemoBild.png)

| Format | Notes |
|---|---|
| OneNote Package (.onepkg) | Recommended. Full-fidelity, re-importable into OneNote. |
| XPS Document (.xps) | Good layout preservation. Large whole notebooks can be split into safer per-section exports. |
| PDF Document (.pdf) | Universal, slightly lower quality on complex pages. Large whole notebooks can be split into safer per-section exports. |
| Local Backup Copy | Copies the raw backup folder from `%LOCALAPPDATA%\Microsoft\OneNote\16.0\Sicherung`. Fastest option, all notebooks at once. |

## Notes

- Password-protected sections cannot be exported. Unlock them in OneNote first.
- Cloud notebooks (OneDrive, SharePoint) must be synced before export. Large notebooks can take several minutes.
- OneNote does not expose notebook size through its COM interface. Before a whole-notebook
  PDF/XPS export, the app therefore offers to export each section separately (recommended),
  create a complete `.onepkg` backup, or consciously try the whole rendered export. Split
  exports include all exportable, unlocked sections and preserve section groups as folders
  to avoid filename collisions. Use `.onepkg` when a single complete backup is required.
