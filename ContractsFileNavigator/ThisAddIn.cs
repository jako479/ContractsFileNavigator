using Microsoft.Office.Tools;
using System;
using System.Configuration;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ContractsFileNavigator
{
    /// <summary>
    /// Excel add-in that shows a "Worksheets" task pane listing the sheets of Contracts.xlsx.
    /// The pane opens every time that file is opened. Closing it hides it until the file is
    /// opened again. The pane's position (left, right or floating) and size are remembered per user,
    /// and the add-in can be switched off by setting <c>Enabled</c> to False in the user settings file
    /// (user.config under %LOCALAPPDATA%\Microsoft_Corporation, in the folder whose name starts with
    /// ContractsFileNavigator.vs; .NET cuts the name short). See DESIGN.md.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>Only a saved workbook with this file name gets the pane.</summary>
        private const string ContractsFileName = "Contracts.xlsx";

        /// <summary>
        /// Fallback pane width in points when the saved width is not a positive number. The normal
        /// default is the Width setting's own default value; points already scale with DPI.
        /// </summary>
        private const int DefaultPaneWidth = 150;

        /// <summary>
        /// Fallback floating height in points when the saved height is not a positive number. Only a
        /// floating pane has a height of its own; a docked pane is stretched to the window.
        /// </summary>
        private const int DefaultPaneHeight = 400;

        /// <summary>Fallback position when the saved DockPosition is not Left, Right or Floating.</summary>
        private const Office.MsoCTPDockPosition DefaultDockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;

        /// <summary>The settings file .NET keeps per user: the only file settings recovery may delete.</summary>
        private const string UserConfigFileName = "user.config";

        /// <summary>
        /// Registry name of the general-purpose Sheet Navigator add-in. Both add-ins manage the
        /// same kind of pane, so this one stays idle while that one is installed.
        /// </summary>
        private const string SheetNavigatorProgId = "SheetNavigator";

        /// <summary>The pane and its list, created on the Contracts window while the file is open.</summary>
        private CustomTaskPane pane;
        private ContractsFileNavigatorControl control;

        /// <summary>
        /// Delays the save until the user has stopped dragging the pane border. It also runs after the
        /// pane floats, once the drag that floated it has settled.
        /// </summary>
        private readonly Timer resizeSaveTimer = new Timer { Interval = 500 };

        /// <summary>
        /// Excel raises no event when a sheet tab is dragged to a new position, so the visible pane
        /// re-checks the sheet list once a second. Only a changed list triggers a rebuild.
        /// </summary>
        private readonly Timer refreshTimer = new Timer { Interval = 1000 };

        /// <summary>
        /// Set while the Contracts file is closing, so a layout event raised by the teardown is not saved.
        /// </summary>
        private bool isPaneClosing = false;

        /// <summary>
        /// Set when the pane floats. The saved height is applied by the next resize save, because Excel
        /// rejects a property set inside the dock event and may resize the pane again as the drag ends.
        /// </summary>
        private bool isFloatingHeightPending = false;

        /// <summary>
        /// The dock position, width and floating height last written to the settings (set when the pane
        /// is created, updated on every write), so a layout event that changed nothing is not written again.
        /// </summary>
        private Office.MsoCTPDockPosition recordedDockPosition;
        private int recordedWidth;
        private int recordedHeight;

        /// <summary>
        /// Wires up Excel events, unless the add-in is disabled or Sheet Navigator is installed.
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Diagnostics.HookUnhandledExceptions();

                UpgradeSettingsIfNeeded();
                if (!ReadSetting(() => Properties.Settings.Default.Enabled)) return;

                if (IsSheetNavigatorInstalled())
                {
                    MessageBox.Show("Contracts File Navigator is not running because Sheet Navigator is installed.\n\n" +
                                    "Only one of the two add-ins can be installed. Uninstall one of them in Windows Settings, Apps.",
                                    "Contracts File Navigator", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                resizeSaveTimer.Tick += new EventHandler(ResizeSaveTimer_Tick);
                refreshTimer.Tick += new EventHandler(RefreshTimer_Tick);
                refreshTimer.Start();

                this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
                this.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookActivate);
                this.Application.WorkbookAfterSave += new Excel.AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);
                this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookBeforeClose);
                this.Application.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(Application_SheetActivate);

                InitialCheckOnLoad();
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Startup failed: " + ex);
                MessageBox.Show($"Contracts File Navigator failed to initialize components.\n\nError Details: {ex.Message}",
                                "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Carries settings forward after an Office update, since the settings folder is named after
        /// the Excel build. On that first run every setting is also written, so the file lists them
        /// all for editing.
        /// </summary>
        private void UpgradeSettingsIfNeeded()
        {
            try
            {
                if (!ReadSetting(() => Properties.Settings.Default.UpgradeRequired)) return;

                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.UpgradeRequired = false;

                // Assigning a setting marks it changed, which is what makes Save write it out
                bool enabled = ReadSetting(() => Properties.Settings.Default.Enabled);
                Properties.Settings.Default.Enabled = enabled;
                Properties.Settings.Default.Width = SavedWidth();
                Properties.Settings.Default.Height = SavedHeight();
                Properties.Settings.Default.DockPosition = DockPositionName(SavedDockPosition());

                SaveSettings();
            }
            catch (Exception ex) { Diagnostics.Write("Settings carry-over failed: " + ex); }
        }

        /// <summary>
        /// Reads a setting. If .NET reports the settings file corrupt, the file is reset and the read retried.
        /// </summary>
        private static T ReadSetting<T>(Func<T> read)
        {
            try
            {
                return read();
            }
            catch (ConfigurationErrorsException ex)
            {
                if (!RecoverSettings(ex)) throw;
                Properties.Settings.Default.Reload();
                return read();
            }
        }

        /// <summary>
        /// Saves the settings. If .NET reports the settings file corrupt, the file is reset and the save
        /// retried; the values being saved stay in memory, so none are lost.
        /// </summary>
        private static void SaveSettings()
        {
            try
            {
                Properties.Settings.Default.Save();
            }
            catch (ConfigurationErrorsException ex)
            {
                if (!RecoverSettings(ex)) throw;
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// A user.config left truncated by a crash makes every settings call throw until it is deleted.
        /// Deletes it, but only the add-in's own user.config inside the user's profile, and only when
        /// .NET reports a parse error; anything else (another file, a passing "file in use") is logged
        /// and left alone. True if the file was deleted.
        /// </summary>
        private static bool RecoverSettings(ConfigurationErrorsException ex)
        {
            string file = ConfigFileNamedBy(ex);
            if (!IsOwnUserConfig(file) || !IsParseError(ex))
            {
                Diagnostics.Write("Settings unreadable, left as is: " + ex);
                return false;
            }

            Diagnostics.Write("Settings file corrupt, deleting it: " + file + " | " + ex.Message);
            try
            {
                File.Delete(file);
                return true;
            }
            catch (Exception deleteException)
            {
                Diagnostics.Write("Settings file delete failed: " + deleteException);
                return false;
            }
        }

        /// <summary>The first file name in the exception chain, or null if .NET named none.</summary>
        private static string ConfigFileNamedBy(ConfigurationErrorsException ex)
        {
            for (Exception current = ex; current != null; current = current.InnerException)
            {
                string file = (current as ConfigurationException)?.Filename;
                if (!string.IsNullOrEmpty(file)) return file;
            }
            return null;
        }

        /// <summary>
        /// True only for a file named user.config inside the user's local or roaming application data,
        /// which is where .NET keeps this add-in's settings and never where a workbook lives.
        /// </summary>
        private static bool IsOwnUserConfig(string file)
        {
            try
            {
                if (string.IsNullOrEmpty(file)) return false;
                if (!string.Equals(Path.GetFileName(file), UserConfigFileName, StringComparison.OrdinalIgnoreCase)) return false;

                string fullPath = Path.GetFullPath(file);
                return IsInside(fullPath, Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
                    || IsInside(fullPath, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Settings path check failed: " + ex);
                return false;
            }
        }

        /// <summary>True if the path lies inside the folder; the folder itself does not count.</summary>
        private static bool IsInside(string fullPath, string folder)
        {
            if (string.IsNullOrEmpty(folder)) return false;
            string prefix = Path.GetFullPath(folder).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            return fullPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// True if the failure is a parse error in the file's contents: an XML error, or a configuration
        /// error that names a line. A locked or missing file reports neither.
        /// </summary>
        private static bool IsParseError(ConfigurationErrorsException ex)
        {
            for (Exception current = ex; current != null; current = current.InnerException)
            {
                if (current is XmlException) return true;
                if (current is ConfigurationException configurationError && configurationError.Line > 0) return true;
            }
            return false;
        }

        /// <summary>
        /// True if Sheet Navigator is registered with Excel, whether or not it is loaded.
        /// </summary>
        private bool IsSheetNavigatorInstalled()
        {
            try
            {
                foreach (Office.COMAddIn addIn in this.Application.COMAddIns)
                {
                    if (string.Equals(addIn.ProgId, SheetNavigatorProgId, StringComparison.OrdinalIgnoreCase)) return true;
                }
            }
            catch (Exception ex) { Diagnostics.Write("Add-in list unreadable, assuming no conflict: " + ex); }

            return false;
        }

        /// <summary>
        /// The saved pane width, clamped, or the default width if none has been saved yet.
        /// </summary>
        private static int SavedWidth()
        {
            return PaneSizeRules.ClampWidth(ReadSetting(() => Properties.Settings.Default.Width), DefaultPaneWidth);
        }

        /// <summary>
        /// The saved floating height, clamped, or the default height if none has been saved yet.
        /// </summary>
        private static int SavedHeight()
        {
            return PaneSizeRules.ClampHeight(ReadSetting(() => Properties.Settings.Default.Height), DefaultPaneHeight);
        }

        /// <summary>
        /// The saved position, Left, Right or Floating; anything else falls back to the default.
        /// </summary>
        private static Office.MsoCTPDockPosition SavedDockPosition()
        {
            string saved = ReadSetting(() => Properties.Settings.Default.DockPosition);
            if (string.Equals(saved, "Left", StringComparison.OrdinalIgnoreCase))
            {
                return Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            }
            if (string.Equals(saved, "Floating", StringComparison.OrdinalIgnoreCase))
            {
                return Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
            }
            return DefaultDockPosition;
        }

        /// <summary>
        /// The dock position as written in settings: "Left", "Right" or "Floating".
        /// </summary>
        private static string DockPositionName(Office.MsoCTPDockPosition position)
        {
            switch (position)
            {
                case Office.MsoCTPDockPosition.msoCTPDockPositionRight: return "Right";
                case Office.MsoCTPDockPosition.msoCTPDockPositionFloating: return "Floating";
                default: return "Left";
            }
        }

        /// <summary>
        /// Shows the pane if the Contracts file is already open when Excel starts, because its
        /// open and activate events fired before the add-in subscribed.
        /// </summary>
        private void InitialCheckOnLoad()
        {
            try
            {
                RestorePane(this.Application.ActiveWorkbook);
            }
            catch (Exception ex) { Diagnostics.Write("Initial check failed: " + ex); }
        }

        /// <summary>
        /// Opening the Contracts file always brings the pane back, replacing any pane left from an earlier open.
        /// </summary>
        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            try
            {
                if (!IsContractsFile(workbook)) return;

                RemovePane();
                CreatePane(workbook);
            }
            catch (Exception ex) { Diagnostics.Write("WorkbookOpen failed: " + ex); }
        }

        /// <summary>
        /// Creates the pane for a Contracts file that has none yet. A pane the user closed stays closed.
        /// </summary>
        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            try
            {
                if (!IsContractsFile(workbook)) return;

                // Working in the file again means any pending close was cancelled
                isPaneClosing = false;

                RestorePane(workbook);
            }
            catch (Exception ex) { Diagnostics.Write("WorkbookActivate failed: " + ex); }
        }

        /// <summary>
        /// A Save As can give a workbook the Contracts name or take it away, so the pane follows the
        /// name: a workbook that now has it gets a pane, and the pane's own workbook losing it drops the pane.
        /// </summary>
        private void Application_WorkbookAfterSave(Excel.Workbook workbook, bool success)
        {
            try
            {
                if (!success) return;

                if (IsContractsFile(workbook))
                {
                    RestorePane(workbook);
                }
                else if (IsPaneWorkbook(workbook))
                {
                    RemovePane();
                }
            }
            catch (Exception ex) { Diagnostics.Write("WorkbookAfterSave failed: " + ex); }
        }

        /// <summary>
        /// Flags the pane as closing. Excel asks about unsaved changes after this event, so the close
        /// may still be cancelled; the pane itself is replaced when the file is next opened.
        /// </summary>
        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            try
            {
                if (pane != null && IsContractsFile(workbook)) isPaneClosing = true;
            }
            catch (Exception ex) { Diagnostics.Write("WorkbookBeforeClose failed: " + ex); }
        }

        /// <summary>
        /// Keeps the list's highlight on the active sheet. Each window has its own active sheet,
        /// so only a change in the pane's own window can also change the list.
        /// </summary>
        private void Application_SheetActivate(object sheet)
        {
            try
            {
                if (pane == null) return;

                Excel.Workbook workbook = this.Application.ActiveWorkbook;
                if (!IsContractsFile(workbook)) return;

                // Working in the file again means any pending close was cancelled
                isPaneClosing = false;

                if (!IsAlive() || !pane.Visible) return;

                if (IsPaneWindowActive())
                {
                    control.RefreshWorksheets(workbook);
                }
                else
                {
                    control.HighlightActiveSheet();
                }
            }
            catch (Exception ex) { Diagnostics.Write("SheetActivate failed: " + ex); }
        }

        /// <summary>
        /// Once a second: replaces a pane whose window closed while the file stayed open, forgets a
        /// close that Excel's save prompt cancelled, and lets the visible pane catch sheet reorders
        /// that raise no Excel event.
        /// </summary>
        private void RefreshTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (pane == null) return;

                if (!IsAlive())
                {
                    // The pane's window closed. If another window of the file is active, it gets a pane.
                    RemovePane();
                    RestorePane(this.Application.ActiveWorkbook);
                    return;
                }

                if (this.Application.Ready == false) return;

                if (isPaneClosing)
                {
                    // Excel answers again and the file is still open, so the close was cancelled
                    if (!IsWorkbookOpen(control.Workbook)) return;
                    isPaneClosing = false;
                }

                if (!pane.Visible) return;

                control.RefreshQuietly();
            }
            catch (Exception ex) { Diagnostics.Write("Refresh tick failed: " + ex); }
        }

        /// <summary>
        /// Creates and shows the pane if the workbook is the Contracts file and no live pane exists.
        /// A pane whose window has gone counts as none.
        /// </summary>
        private void RestorePane(Excel.Workbook workbook)
        {
            if (pane != null && !IsAlive()) RemovePane();
            if (pane != null || !IsContractsFile(workbook)) return;
            CreatePane(workbook);
        }

        /// <summary>
        /// Creates the pane on the workbook's window at the saved position and size, fills the list, and
        /// shows it. The fields are set only once the pane is shown, so a failure part-way leaves no
        /// hidden pane behind and the next activation tries again.
        /// </summary>
        private void CreatePane(Excel.Workbook workbook)
        {
            Excel.Window window = FirstWindowOf(workbook);
            if (window == null) return;

            ContractsFileNavigatorControl newControl = new ContractsFileNavigatorControl
            {
                Window = window,
                Workbook = workbook
            };

            Office.MsoCTPDockPosition savedDockPosition = SavedDockPosition();
            int savedWidth = SavedWidth();
            int savedHeight = SavedHeight();
            CustomTaskPane newPane = this.CustomTaskPanes.Add(newControl, "Worksheets", window);

            Office.MsoCTPDockPosition dock;
            int width;
            int height;
            try
            {
                // Dock position before size: the pane API expects that order. Only a floating pane takes a height.
                newPane.DockPosition = savedDockPosition;
                newPane.Width = savedWidth;
                if (savedDockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating) newPane.Height = savedHeight;

                // Width is meaningless when docked top or bottom, so keep the pane on a side or floating. Excel's
                // "NoHorizontal" is the restriction compatible with a side-docked pane, despite its name.
                // Optional: a rejected restriction must never stop the pane from working.
                try { newPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal; }
                catch (Exception ex) { Diagnostics.Write("Dock restriction rejected: " + ex); }

                newControl.Resize += new EventHandler(NavigatorControl_Resize);
                newPane.DockPositionChanged += new EventHandler(Pane_DockPositionChanged);

                // What Excel actually gave the pane is the baseline for change detection
                dock = newPane.DockPosition;
                width = PaneSizeRules.ClampWidth(newPane.Width, savedWidth);
                height = dock == Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                    ? PaneSizeRules.ClampHeight(newPane.Height, savedHeight)
                    : savedHeight;

                // The quiet fill tolerates a busy Excel; the refresh timer fills the list a moment later
                newControl.RefreshQuietly();
                newPane.Visible = true;
            }
            catch
            {
                // Don't leave a half-configured pane behind
                try { newControl.Resize -= NavigatorControl_Resize; } catch (Exception ex) { Diagnostics.Write("Resize unhook after failed creation failed: " + ex); }
                try { newPane.DockPositionChanged -= Pane_DockPositionChanged; } catch (Exception ex) { Diagnostics.Write("Dock unhook after failed creation failed: " + ex); }
                try { this.CustomTaskPanes.Remove(newPane); } catch (Exception ex) { Diagnostics.Write("Pane removal after failed creation failed: " + ex); }
                throw;
            }

            pane = newPane;
            control = newControl;
            isPaneClosing = false;
            isFloatingHeightPending = false;
            recordedDockPosition = dock;
            recordedWidth = width;
            recordedHeight = height;
        }

        /// <summary>
        /// Drops the pane. Its handlers are unhooked first so the removal saves nothing.
        /// </summary>
        private void RemovePane()
        {
            if (pane == null) return;

            CustomTaskPane oldPane = pane;
            ContractsFileNavigatorControl oldControl = control;
            pane = null;
            control = null;
            isPaneClosing = false;
            isFloatingHeightPending = false;
            resizeSaveTimer.Stop();

            try { oldControl.Resize -= NavigatorControl_Resize; } catch (Exception ex) { Diagnostics.Write("Resize unhook failed: " + ex); }
            try { oldPane.DockPositionChanged -= Pane_DockPositionChanged; } catch (Exception ex) { Diagnostics.Write("Dock unhook failed: " + ex); }
            try { this.CustomTaskPanes.Remove(oldPane); } catch (Exception ex) { Diagnostics.Write("Pane removal failed (already disposed with its window?): " + ex); }
        }

        /// <summary>
        /// False once Excel or VSTO has disposed the pane, for example because its window closed.
        /// </summary>
        private bool IsAlive()
        {
            try
            {
                bool unused = pane.Visible;
                return !control.IsDisposed;
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Pane is gone: " + ex);
                return false;
            }
        }

        /// <summary>
        /// True while the pane's own window is Excel's active window.
        /// </summary>
        private bool IsPaneWindowActive()
        {
            try
            {
                Excel.Window active = this.Application.ActiveWindow;
                return active != null && control.Window != null && active.Hwnd == control.Window.Hwnd;
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Active window check failed: " + ex);
                return false;
            }
        }

        /// <summary>
        /// True if the workbook is the one the pane was created for.
        /// </summary>
        private bool IsPaneWorkbook(Excel.Workbook workbook)
        {
            try
            {
                return control != null && control.Workbook != null && workbook != null
                    && string.Equals(control.Workbook.FullName, workbook.FullName, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Workbook comparison failed: " + ex);
                return false;
            }
        }

        /// <summary>
        /// True while the workbook can still be asked its name; a closed one throws.
        /// </summary>
        private static bool IsWorkbookOpen(Excel.Workbook workbook)
        {
            try
            {
                return workbook != null && !string.IsNullOrEmpty(workbook.Name);
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Workbook is gone: " + ex);
                return false;
            }
        }

        /// <summary>
        /// True while the pane exists, is not being torn down, and is shown: the only time a layout
        /// event is worth recording.
        /// </summary>
        private bool CanRecordPane()
        {
            return pane != null && !isPaneClosing && IsAlive() && pane.Visible;
        }

        /// <summary>
        /// Restarts the save delay on every resize, so a drag is written once when it ends.
        /// </summary>
        private void NavigatorControl_Resize(object sender, EventArgs e)
        {
            try
            {
                if (!CanRecordPane()) return;

                resizeSaveTimer.Stop();
                resizeSaveTimer.Start();
            }
            catch (Exception ex) { Diagnostics.Write("Resize failed: " + ex); }
        }

        /// <summary>
        /// Records the pane once resizing has settled, after first giving a pane that just floated its saved height.
        /// </summary>
        private void ResizeSaveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                resizeSaveTimer.Stop();
                if (!CanRecordPane()) return;

                if (isFloatingHeightPending) ApplyFloatingHeight();
                RecordPaneChange();
            }
            catch (Exception ex) { Diagnostics.Write("Resize save failed: " + ex); }
        }

        /// <summary>
        /// Records the pane when the user docks it left or right. A pane that just floated is not
        /// recorded yet: Excel rejects property sets inside this handler and may still resize the pane
        /// as the drag ends, so the saved height is applied, and the pane recorded, by the resize save
        /// once the drag has settled.
        /// </summary>
        private void Pane_DockPositionChanged(object sender, EventArgs e)
        {
            try
            {
                if (!CanRecordPane()) return;

                if (pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating)
                {
                    isFloatingHeightPending = true;
                    resizeSaveTimer.Stop();
                    resizeSaveTimer.Start();
                    return;
                }

                RecordPaneChange();
            }
            catch (Exception ex) { Diagnostics.Write("DockPositionChanged failed: " + ex); }
        }

        /// <summary>
        /// Gives a pane that just floated its saved height. A rejected height is logged and the pane
        /// keeps Excel's, which is then what gets recorded; the dock position and width are recorded either way.
        /// </summary>
        private void ApplyFloatingHeight()
        {
            isFloatingHeightPending = false;
            if (pane.DockPosition != Office.MsoCTPDockPosition.msoCTPDockPositionFloating) return;

            try
            {
                pane.Height = recordedHeight;
            }
            catch (Exception ex) { Diagnostics.Write("Floating height failed: " + ex); }
        }

        /// <summary>
        /// After the user docks, floats or resizes the pane: writes its dock position, width and, while
        /// floating, height to the settings. A layout event that changed nothing since the last write is
        /// ignored, and so is a pane docked top or bottom. A width or height of zero or less is not a
        /// change; the saved value stays. A docked pane's height is Excel's, so the saved height stays too.
        /// </summary>
        private void RecordPaneChange()
        {
            Office.MsoCTPDockPosition dock = pane.DockPosition;
            int width = PaneSizeRules.ClampWidth(pane.Width, recordedWidth);
            int height = dock == Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                ? PaneSizeRules.ClampHeight(pane.Height, recordedHeight)
                : recordedHeight;
            if (dock == recordedDockPosition && width == recordedWidth && height == recordedHeight) return;

            // Top and bottom are only reachable if Excel rejected the dock restriction; they have no saved form
            if (dock == Office.MsoCTPDockPosition.msoCTPDockPositionTop || dock == Office.MsoCTPDockPosition.msoCTPDockPositionBottom) return;

            Properties.Settings.Default.DockPosition = DockPositionName(dock);
            Properties.Settings.Default.Width = width;
            Properties.Settings.Default.Height = height;
            SaveSettings();

            recordedDockPosition = dock;
            recordedWidth = width;
            recordedHeight = height;
        }

        /// <summary>
        /// True for a saved workbook named Contracts.xlsx, in any folder.
        /// </summary>
        private static bool IsContractsFile(Excel.Workbook workbook)
        {
            try
            {
                return workbook != null
                    && !string.IsNullOrEmpty(workbook.Path)
                    && string.Equals(workbook.Name, ContractsFileName, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Workbook name check failed: " + ex);
                return false;
            }
        }

        /// <summary>
        /// The workbook's first window, which is the one Excel shows it in.
        /// </summary>
        private static Excel.Window FirstWindowOf(Excel.Workbook workbook)
        {
            try
            {
                return workbook.Windows.Count > 0 ? workbook.Windows[1] : null;
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Window lookup failed: " + ex);
                return null;
            }
        }

        /// <summary>
        /// Stops the timers and unhooks Excel events, each step on its own so one failure cannot skip
        /// the rest. The VSTO runtime has already disposed the pane by now.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            RunLogged("Resize timer stop", () => { resizeSaveTimer.Stop(); resizeSaveTimer.Dispose(); });
            RunLogged("Refresh timer stop", () => { refreshTimer.Stop(); refreshTimer.Dispose(); });

            RunLogged("WorkbookOpen unhook", () => this.Application.WorkbookOpen -= Application_WorkbookOpen);
            RunLogged("WorkbookActivate unhook", () => this.Application.WorkbookActivate -= Application_WorkbookActivate);
            RunLogged("WorkbookAfterSave unhook", () => this.Application.WorkbookAfterSave -= Application_WorkbookAfterSave);
            RunLogged("WorkbookBeforeClose unhook", () => this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose);
            RunLogged("SheetActivate unhook", () => this.Application.SheetActivate -= Application_SheetActivate);

            pane = null;
            control = null;
        }

        /// <summary>
        /// Runs one shutdown step, logging a failure instead of letting it stop the steps after it.
        /// </summary>
        private static void RunLogged(string step, Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex) { Diagnostics.Write(step + " failed: " + ex); }
        }

        #region VSTO Generated Code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
