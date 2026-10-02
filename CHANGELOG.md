# Changelog

What changed for users, grouped by release version. Not a commit log.

## 1.0.0 - 2026-09-30

- The pane follows a Save As: a workbook renamed to Contracts.xlsx gets the pane, one renamed away loses it.
- The pane comes back when its window closes while the file stays open in another window.
- Floating height is capped at 1200 points, like the 400-point width cap.
- Every caught error is logged; a repeating one fills a single line.
- Initial release.
- Worksheets pane opens whenever Contracts.xlsx is opened and lists its visible sheets; click, or move to a name with the arrow keys or its first letter and press Enter or Space, to jump.
- Closing the pane hides it until the file is opened again.
- Pane position (left, right or floating) and size are remembered.
- Can be switched off with the Enabled setting in user.config.
- Stays idle while Sheet Navigator is installed.
