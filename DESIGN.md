# Design

How the Worksheets pane behaves, and what the add-in does on each Excel event.

## Saved state

Everything lives in the user's settings, never in the workbook.

- **DockPosition**: where the pane was last left, `Left`, `Right` or
  `Floating` (Excel's own names). Starts at `Right`.
- **Width**: the width in points the pane was last left at, capped at 400.
  Starts at 150.
- **Height**: the height in points the pane was last left at while floating.
  Starts at 400. A docked pane is stretched to the window, so docking leaves
  this alone.
- **Enabled**: `False` switches the add-in off.
- A value that cannot be read falls back to its default: dock position to
  `Right`, width to 150, height to 400. A settings file .NET cannot read at
  all is deleted and starts over.
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
- A floating pane's screen position is Excel's; the add-in only knows it is
  floating. A `Floating` setting comes back floating at its saved width and
  height, wherever Excel puts it.

## When settings are applied and saved

**Applied to the pane**: dock position and width when the pane is created on
file open, plus height if it starts floating; and height again whenever the
pane floats. Nothing else is ever written from the settings back to the pane.

**Saved from the pane**: dock position and width on every dock, float or
resize, and height too while the pane is floating. Docking never touches the
saved height. Nothing is saved when the pane is hidden or the file closes.

## Pane events

**Pane hidden** (the pane's X)

Nothing is written. The pane stays hidden until the file is opened again.

**Pane docked or floated**

A pane that just floated is given the saved height first. Then the pane's
dock position and width are written, and its height if it is now floating.
Width is written too because Excel may change it when docking or floating.

**Pane resized** (written once the drag settles, floating or docked)

Writes the pane's dock position and width, and its height if it is floating.

A dock, float or resize event that left the pane exactly as last written
(for example the layout event Excel raises right after the pane is shown)
writes nothing. A pane docked top or bottom, only possible if Excel rejected
the add-in's dock restriction, is not written. A pane reporting a width or
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
- **Workbook closing**: the pane is flagged as closing, so the layout events
  of the teardown write nothing. The pane is replaced when the file is next
  opened.
- **Workbook saved**: nothing.
- **Sheet activated**: the pane re-highlights its window's active sheet; if
  its window is the active one, it also refreshes its list.
- **Excel closes**: timers stop, events unhook, nothing is written.

## Worksheet list

The list shows the workbook's visible sheets, in tab order. Hidden and very
hidden sheets are not listed.

**When it refreshes**

- A sheet is activated.
- Once a second, while the pane is shown. If Excel is mid-edit (typing in a
  cell or a tab name) only the highlight is updated.
- The pointer enters the list.
- The pane is created.
- A click on a sheet name could not jump (the sheet was renamed or removed).

A refresh rebuilds the list only when the visible names changed, so it
catches sheets that were added, deleted, renamed, moved, hidden or unhidden,
even though Excel raises no event for a rename or a move. The highlight then
moves to the window's active sheet.

**Clicking a name** activates that sheet in the pane's own window. Keyboard
navigation in the list is blocked; Ctrl+PgUp/PgDn already covers it.
