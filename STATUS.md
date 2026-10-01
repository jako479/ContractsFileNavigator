# Status

Code complete and compiling, never run in Excel yet. Next: Clean Solution in Sheet Navigator, then Ctrl+F5 here and check that the pane opens with Contracts.xlsx, closes with its X, and comes back where it was, floating included, when the file is reopened.

## Decisions

- Excel rejects any pane property set inside a pane event handler ("cannot be set during the object's event handler"), so the floating height is applied by code queued with BeginInvoke to run right after the dock event returns.
- A floating pane is remembered as floating, at its width and height; its screen position is not saved because the pane API exposes none, so Excel places it. Pane behavior per event is in DESIGN.md.
- No automated tests: the testable logic is private to ThisAddIn, which only exists inside Excel, and extracting it is not worth it for an add-in this size.
- Only unexpected errors are logged, to %TEMP%\ContractsFileNavigator.log; catches for expected conditions stay silent.
- Only a saved workbook named Contracts.xlsx gets the pane: one pane, no per-window tracking, no Ribbon button.
- The pane opens on every open of the file; closing it hides it until the file is opened again.
- State lives in the user's settings, never in the workbook, so opening the pane never dirties a file.
- Settings are Enabled, Width, Height, DockPosition and UpgradeRequired only. Width is in points with no floor and a 400 ceiling, default 150; Height is in points with no floor, default 400, applied whenever the pane floats and saved only while it is floating, since a docked pane takes the window's height; DockPosition is Left, Right or Floating, default Right. All five are written on first run so the file can be edited.
- Stays idle, with a message, while Sheet Navigator is registered in Excel; ClickOnce has no mainstream install-time block, so the check runs at startup.
- A single click activates a sheet; keyboard navigation in the pane is blocked, since Ctrl+PgUp/PgDn already covers it.
- The signing key stays out of the repo; contributors create their own test certificate.
- The release ZIP is made by hand; zipping is not a VSTO publish convention.
