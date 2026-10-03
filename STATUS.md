# Status

Compiles, and the size rules pass their tests. The add-in was started in Excel once, before the floating-height rework; the fixes from the 2026-09-30 code review, the keyboard-accessible list (matched to Sheet Navigator's on 2026-10-02), the fixes ported from the Sheet Navigator reviews on 2026-10-02 and the dock, closing and shutdown handling matched to Sheet Navigator's the same day are not yet tried in Excel. Next: Clean Solution in Sheet Navigator, then Ctrl+F5 here.

## Decisions

- Dock, float and resize all go through the 500 ms resize save, since Excel may change the width while docking; a change queued while the file is closing is re-checked up to four times before it is written or dropped. Both match Sheet Navigator.
- A saved dock position, width or height Excel rejects falls back to its default rather than blocking the pane; if the default is rejected too, the pane keeps Excel's value.
- A resize still waiting for its delay is written when the file closes or Excel quits; it happened before the close, so it is the user's.
- List rebuilds are skipped while a macro has screen updating off; the pane catches up on the next tick.
- A pane docked top or bottom has no saved form: a layout event ignores it and any other save of it throws, so Excel ignoring the dock restriction shows up as a failure instead of being saved as Left.
- The list follows the Windows list pattern, as WCAG 2.1.1 requires keyboard operation: arrows and typed letters move the highlight, Enter or a click activates, and focus then goes back to the pane's window so Excel's keys work right away, as after a click on Excel's own sheet tabs. While the list has keyboard focus, refreshes leave the highlight alone. Screen readers get the name Worksheets.
- A settings file is deleted and recreated only if it is the add-in's own user.config under the user's profile and .NET reports a parse error; any other settings failure is logged and the file left alone.
- No automatic update check; a new version is installed by running the new release.
- Excel rejects any pane property set inside a pane event handler ("cannot be set during the object's event handler"), so the floating height is applied by the resize-save timer, 500 ms after the last layout event, which also lets Excel's own resize at the end of the drag settle first.
- A floating pane is remembered as floating, at its width and height; its screen position is not saved because the pane API exposes none, so Excel places it. Pane behavior per event is in DESIGN.md.
- Automated tests cover only the pure size rules (PaneSizeRules, compiled into ContractsFileNavigator.Tests as a linked file); everything else needs Excel and is verified by hand.
- Every caught error is logged; identical consecutive messages are written once, followed by "Last message repeated N times" when a different message arrives or Excel closes, so a failure repeating every second fills two lines. Only two failures stay unlogged: a failure of the log itself, and Excel's expected refusal when Enter or a click lands mid-edit.
- Only a saved workbook named Contracts.xlsx gets the pane: one pane, no per-window tracking, no Ribbon button.
- The pane opens on every open of the file; closing it hides it until the file is opened again.
- State lives in the user's settings, never in the workbook, so opening the pane never dirties a file.
- Settings are Enabled, Width, Height, DockPosition and UpgradeRequired only. Width is in points with no floor and a 400 ceiling, default 150; Height is in points with no floor and a 1200 ceiling, default 400, applied whenever the pane floats and saved only while it is floating, since a docked pane takes the window's height; DockPosition is Left, Right or Floating, default Right. All five are written on first run so the file can be edited.
- Stays idle, with a message, while Sheet Navigator is registered in Excel; ClickOnce has no mainstream install-time block, so the check runs at startup.
- The signing key stays out of the repo; contributors create their own test certificate.
- The release ZIP is made by hand; zipping is not a VSTO publish convention.
