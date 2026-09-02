# Changelog

## 1.2.1

- Enlarged the large PDF/XPS export dialog so every export option is fully visible.
- Added text wrapping and window resizing for improved display scaling compatibility.
- Tightened safe cancellation before finalizing an export, made local-backup copying
  cancellable between files, and clarified cancellation behavior.

## 1.2.0

- Improved reliability with a serialized OneNote COM worker, safe cancellation, and bounded connection recovery.
- Protected existing backups with staging files and atomic replacement.
- Added safer PDF/XPS handling for large notebooks, including automatic per-section export.
- Prevented filename collisions while preserving OneNote section-group folders.
- Added an About dialog with version and build information.
- Standardized all user-facing errors, progress messages, and release tooling in English.
