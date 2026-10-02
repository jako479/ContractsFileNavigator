# Design

How the Worksheets pane behaves, and what the add-in does on each Excel event.

## Saved state

Everything lives in the user's settings, never in the workbook.

- **DockPosition**: where the pane was last left, `Left`, `Right` or
  `Floating` (Excel's own names). Starts at `Right`.
- **Width**: the width in points the pane was last left at, capped at 400.
  Starts at 150.
- **Height**: the height in points the pane was last left at while floating,
  capped at 1200. Starts at 400. A docked pane is stretched to the window, so
  docking leaves this alone.
- **Enabled**: `False` switches the add-in off.
- A value that cannot be read falls back to its default: dock position to
  `Right`, width to 150, height to 400. A settings file .NET reports as
  corrupt is deleted and starts over, provided it is the add-in's own
  user.config in the user's profile; any other settings failure is logged and
  leaves the file alone.
- After an Office update the settings are carried over from the previous
  Excel build's folder.
- Every setting is written on the first run, so the file lists them all for
  editing.

## Pane

- One pane, on the first window of `Contracts.xlsx`. Only a saved workbook
  with that name gets one.
- Opening the file creates the pane at the saved dock position and width,
  and at the saved height if that position is floating, replacing any pane
  left from an earlier open.
- A saved dock position, width or height that Excel rejects falls back to
  its default; if Excel rejects that too, the pane keeps Excel's own value.
- A floating pane's screen position is Excel's; the add-in only knows it is
  floating. A `Floating` setting comes back floating at its saved width and
  height, wherever Excel puts it.

## When settings are applied and saved

**Applied to the pane**: dock position and width when the pane is created on
file open, plus height if it starts floating; and height again whenever the
pane floats, once the drag has settled. Nothing else is ever written from the
settings back to the pane.

**Saved from the pane**: dock position and width on every dock, float or
resize, and height too while the pane is floating. Docking never touches the
saved height. Hiding the pane saves nothing; closing the file saves only a
resize still waiting for its delay.

## Pane events

**Pane hidden** (the pane's X)

Nothing is written. The pane stays hidden until the file is opened again.

**Pane docked or floated**

A pane docked left or right is written at once: its dock position and width.
Width is written too because Excel may change it when docking or floating.
A pane that just floated is not written yet; Excel rejects property sets
inside the dock event and may still resize the pane as the drag ends, so it
is handled by the resize rule below once the drag settles.

**Pane resized** (written once the drag settles, floating or docked)

A pane that has just floated is given the saved height first; a height Excel
rejects is logged and the pane keeps Excel's, which is then what is written.
Then the pane's dock position and width are written, and its height if it is
floating.

A dock, float or resize event that left the pane exactly as last written
(for example the layout event Excel raises right after the pane is shown)
writes nothing. A pane docked top or bottom, only possible if Excel rejected
the add-in's dock restriction, is never written: a layout event ignores it,
and any other attempt to save it is reported as a failure. A pane reporting a width or
height of zero or less keeps the saved value. A docked pane's height is
Excel's and is never written, so the saved height is always the last floating
height.

## Workbook and window events

- **Excel starts**: if `Contracts.xlsx` is already the active workbook, its
  pane is created.
- **Workbook opened**: `Contracts.xlsx` gets a fresh pane; any other workbook
  is ignored.
- **Workbook activated**: `Contracts.xlsx` gets a pane if it has none; a pane
  the user closed stays closed. A cancelled close is forgotten.
- **Workbook or window deactivated**: nothing.
- **Pane window closed** while the file stays open in another window: the
  dead pane is dropped on the next refresh tick, and the file's first window
  gets a fresh pane if the file is active.
- **Workbook closing**: a resize still waiting for its delay is written, then
  the pane is flagged as closing, so the layout events of the teardown write
  nothing. The pane is replaced when the file is next
  opened. If Excel's save prompt cancels the close, the flag is cleared by the
  next refresh tick once Excel answers again, or sooner by a workbook or sheet
  activation.
- **Workbook saved**: a Save As that gives a workbook the Contracts name gets
  it a pane (a pane the user closed stays closed); a Save As that takes the
  name away from the pane's workbook removes the pane. A plain save changes
  nothing.
- **Sheet activated**: the pane re-highlights its window's active sheet; if
  its window is the active one and no macro has screen updating off, it also
  refreshes its list.
- **Excel closes**: a resize still waiting for its delay is written; then
  timers stop and events unhook.

## Worksheet list

The list shows the workbook's visible worksheets, in tab order. Hidden and
very hidden sheets are not listed, and neither are chart sheets; while a chart
sheet is active nothing is highlighted.

**When it refreshes**

- A sheet is activated, unless a macro has screen updating off.
- Once a second, while the pane is shown. If Excel is mid-edit (typing in a
  cell or a tab name) or a macro has screen updating off, only the highlight
  is updated.
- The pointer enters the list.
- The pane is created.
- Activating a name could not jump (the sheet was renamed or removed).

A refresh rebuilds the list only when the visible names changed, so it
catches sheets that were added, deleted, renamed, moved, hidden or unhidden,
even though Excel raises no event for a rename or a move. The highlight then
moves to the window's active sheet, unless the keyboard is using the list.

**Keyboard and mouse**

- Arrows, Home, End, PgUp, PgDn and typed letters move the highlight without
  switching sheets, as in any Windows list.
- Enter, Space or a click on a name activates that sheet in the pane's own
  window. A click on the blank space under the names does nothing.
- While the list has keyboard focus the highlight is the user's cursor:
  refreshes leave it where it is, and a rebuild puts it back on the same name
  if that sheet still exists. When focus leaves the list, the highlight
  returns to the active sheet.
- The pane handles no other keys.
- Screen readers announce the list as "Worksheets".
