# Changelog

## 1.2.0

- Improved reliability with a serialized OneNote COM worker, safe cancellation, and bounded connection recovery.
- Protected existing backups with staging files and atomic replacement.
- Added safer PDF/XPS handling for large notebooks, including automatic per-section export.
- Prevented filename collisions while preserving OneNote section-group folders.
- Added an About dialog with version and build information.
- Standardized all user-facing errors, progress messages, and release tooling in English.
