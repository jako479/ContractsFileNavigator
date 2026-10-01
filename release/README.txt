============================================================
CONTRACTS FILE NAVIGATOR - INSTALLATION
============================================================

Contracts File Navigator adds a Worksheets pane to Excel that
lists the sheets in Contracts.xlsx. The pane opens every time
that file is opened. Click a name to jump to it.

Install this OR Sheet Navigator, not both. Contracts File
Navigator stays idle and shows a message while Sheet Navigator
is installed.

STEP 1: EXTRACT THE FILES
-------------------------
* Do not run the installer from inside the ZIP window.
* Extract the whole folder first, for example to your Desktop.

STEP 2: RUN THE INSTALLER
-------------------------
* Close Microsoft Excel completely.
* Double-click setup.exe.
* When the Microsoft Office Customization security alert appears,
  click "Install".

STEP 3: USE IT
--------------
* Open Contracts.xlsx. The pane appears on the right.
* Close the pane with its X; it comes back the next time the
  file is opened. Excel remembers the pane's position (left,
  right or floating) and size.
* To switch the add-in off without uninstalling it, set Enabled
  to False in its user.config (under %LOCALAPPDATA%\
  Microsoft_Corporation, in the folder whose name starts with
  ContractsFileNavigator.vs) and restart Excel.

REQUIREMENTS
------------
* Windows with desktop Microsoft Excel 2013 or later.
* .NET Framework 4.7.2 and the Visual Studio 2010 Tools for Office
  Runtime. The installer adds both if they are missing.

REMOVAL
-------
Windows Settings > Apps, find "Contracts File Navigator", and
click Uninstall.

LICENSE
-------
MIT. See LICENSE in this folder.
============================================================
