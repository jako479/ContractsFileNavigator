# Status

Code complete and compiling, never run in Excel yet. Next: Clean Solution in Sheet Navigator, then Ctrl+F5 here and check that the pane opens with Contracts.xlsx, closes with its X, and comes back when the file is reopened.

## Decisions

- Only a saved workbook named Contracts.xlsx gets the pane: one pane, no per-window tracking, no Ribbon button.
- The pane opens on every open of the file; closing it hides it until the file is opened again.
- State lives in the user's settings, never in the workbook, so opening the pane never dirties a file.
- Settings are Enabled, LastWidth and UpgradeRequired only. LastWidth 0 means "use the screen-scaled default"; all three are written on first run so the file can be edited.
- Stays idle, with a message, while Sheet Navigator is registered in Excel; ClickOnce has no mainstream install-time block, so the check runs at startup.
- A single click activates a sheet; arrow keys still move the highlight and jump (Sheet Navigator now blocks keys; not ported).
- The signing key stays out of the repo; contributors create their own test certificate.
- The release ZIP is made by hand; zipping is not a VSTO publish convention.
