# Changelog

What changed for users, grouped by release version. Not a commit log.

## 1.0.0 - 2026-09-30

- Enter or a click on a sheet name hands focus back to the worksheet, so Excel's keys work right away.
- A resize made just before the file closes or Excel quits is kept.
- Clicking the blank space under the sheet names no longer jumps to the last sheet.
- A saved pane position or size that Excel rejects falls back to the default instead of blocking the pane.
- The list pauses its rebuilds while a macro has screen updating off.
- Message boxes stay in front of Excel.
- The pane follows a Save As: a workbook renamed to Contracts.xlsx gets the pane, one renamed away loses it.
- The pane comes back when its window closes while the file stays open in another window.
- Floating height is capped at 1200 points, like the 400-point width cap.
- Every caught error is logged; a repeating one fills two lines.
- Initial release.
- Worksheets pane opens whenever Contracts.xlsx is opened and lists its visible sheets; click, or move to a name with the arrow keys or its first letter and press Enter, to jump.
- Closing the pane hides it until the file is opened again.
- Pane position (left, right or floating) and size are remembered.
- Can be switched off with the Enabled setting in user.config.
- Stays idle while Sheet Navigator is installed.
