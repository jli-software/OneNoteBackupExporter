<div align="center">

<img src="icon.ico" alt="OneNote Backup Exporter" width="100" />

# OneNote Backup Exporter

A Windows desktop application to export OneNote notebooks to portable formats.  
Talks directly to OneNote Desktop via the COM API — no cloud services, no account required, no internet connection.

![version](https://img.shields.io/badge/version-1.1.0-blueviolet?style=flat-square)
![platform](https://img.shields.io/badge/platform-Windows%2010%2F11-0078D7?style=flat-square&logo=windows&logoColor=white)
![license](https://img.shields.io/badge/license-MIT-brightgreen?style=flat-square)
![built with](https://img.shields.io/badge/built%20with-.NET%2010%20%2F%20WPF-512BD4?style=flat-square&logo=dotnet&logoColor=white)

</div>

---

![OneNote Backup Exporter main window](demoScreenshot.png)

## Download

Go to [Releases](../../releases/latest) and download `OneNoteBackupExporter_Setup_1.0.0.exe`. Run the installer, done.

## Requirements

- Windows 10 or 11 (64-bit)
- OneNote Desktop 2016 or later (including Microsoft 365)

The Microsoft Store version of OneNote is not supported.

## Export Formats

![Available export formats](FormatDemoBild.png)

| Format | Notes |
|---|---|
| OneNote Package (.onepkg) | Recommended. Full-fidelity, re-importable into OneNote. |
| XPS Document (.xps) | Good layout preservation. |
| PDF Document (.pdf) | Universal, slightly lower quality on complex pages. |
| Local Backup Copy | Copies the raw backup folder from `%LOCALAPPDATA%\Microsoft\OneNote\16.0\Sicherung`. Fastest option, all notebooks at once. |

## Notes

- Password-protected sections cannot be exported. Unlock them in OneNote first.
- Cloud notebooks (OneDrive, SharePoint) must be synced before export. Large notebooks can take several minutes.
- If PDF or XPS export times out on a large notebook, use the OneNote Package format instead.
