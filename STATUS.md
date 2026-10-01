# Status

Compiles, and the size rules pass their tests. The add-in was started in Excel once, before the floating-height rework; the fixes from the 2026-09-30 code review are not yet tried in Excel. Next: Clean Solution in Sheet Navigator, then Ctrl+F5 here and run the first-run check in TODO.md.

## Decisions

- A settings file is deleted and recreated only if it is the add-in's own user.config under the user's profile and .NET reports a parse error; any other settings failure is logged and the file left alone.
- No automatic update check; a new version is installed by running the new release.
- Excel rejects any pane property set inside a pane event handler ("cannot be set during the object's event handler"), so the floating height is applied by the resize-save timer, 500 ms after the last layout event, which also lets Excel's own resize at the end of the drag settle first.
- A floating pane is remembered as floating, at its width and height; its screen position is not saved because the pane API exposes none, so Excel places it. Pane behavior per event is in DESIGN.md.
- Automated tests cover only the pure size rules (PaneSizeRules, compiled into ContractsFileNavigator.Tests as a linked file); everything else needs Excel and is verified by hand.
- Every caught error is logged to %TEMP%\ContractsFileNavigator.log; a message identical to the previous one is dropped, so a failure repeating every second fills one line.
- Only a saved workbook named Contracts.xlsx gets the pane: one pane, no per-window tracking, no Ribbon button.
- The pane opens on every open of the file; closing it hides it until the file is opened again.
- State lives in the user's settings, never in the workbook, so opening the pane never dirties a file.
- Settings are Enabled, Width, Height, DockPosition and UpgradeRequired only. Width is in points with no floor and a 400 ceiling, default 150; Height is in points with no floor and a 1200 ceiling, default 400, applied whenever the pane floats and saved only while it is floating, since a docked pane takes the window's height; DockPosition is Left, Right or Floating, default Right. All five are written on first run so the file can be edited.
- Stays idle, with a message, while Sheet Navigator is registered in Excel; ClickOnce has no mainstream install-time block, so the check runs at startup.
- A single click activates a sheet; keyboard navigation in the pane is blocked, since Ctrl+PgUp/PgDn already covers it.
- The signing key stays out of the repo; contributors create their own test certificate.
- The release ZIP is made by hand; zipping is not a VSTO publish convention.
