using Microsoft.Office.Tools;
using System;
using System.Configuration;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ContractsFileNavigator
{
    /// <summary>
    /// Excel add-in that shows a "Worksheets" task pane listing the sheets of Contracts.xlsx.
    /// The pane opens every time that file is opened. Closing it hides it until the file is
    /// opened again. The pane's position (left, right or floating) and size are remembered per user,
    /// and the add-in can be switched off by setting <c>Enabled</c> to False in the user settings file
    /// (user.config under %LOCALAPPDATA%\Microsoft_Corporation, in the folder named after
    /// ContractsFileNavigator.vsto). See DESIGN.md.
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

        /// <summary>
        /// Ceiling for saved widths, on both save and load, to reject a garbage value in the settings file.
        /// There is no floor: Excel enforces its own minimum whenever a width is set.
        /// </summary>
        private const int MaxPaneWidth = 400;

        /// <summary>
        /// Registry name of the general-purpose Sheet Navigator add-in. Both add-ins manage the
        /// same kind of pane, so this one stays idle while that one is installed.
        /// </summary>
        private const string SheetNavigatorProgId = "SheetNavigator";

        /// <summary>The pane and its list, created on the Contracts window while the file is open.</summary>
        private CustomTaskPane pane;
        private ContractsFileNavigatorControl control;

        /// <summary>Delays the save until the user has stopped dragging the pane border.</summary>
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
            catch { /* Nothing to carry over */ }
        }

        /// <summary>
        /// Reads a setting, resetting the settings file first if .NET reports it unreadable.
        /// </summary>
        private static T ReadSetting<T>(Func<T> read)
        {
            try
            {
                return read();
            }
            catch (ConfigurationErrorsException ex)
            {
                RecoverSettings(ex);
                return read();
            }
        }

        /// <summary>
        /// Saves the settings, resetting the settings file first if .NET reports it unreadable.
        /// </summary>
        private static void SaveSettings()
        {
            try
            {
                Properties.Settings.Default.Save();
            }
            catch (ConfigurationErrorsException ex)
            {
                RecoverSettings(ex);
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// A user.config left truncated by a crash makes every settings call throw until it is deleted.
        /// Deletes it and reloads the defaults.
        /// </summary>
        private static void RecoverSettings(ConfigurationErrorsException ex)
        {
            string file = ex.Filename ?? (ex.InnerException as ConfigurationErrorsException)?.Filename;
            Diagnostics.Write("Settings file unreadable, resetting it: " + (file ?? "(unknown path)") + " | " + ex.Message);

            try
            {
                if (!string.IsNullOrEmpty(file) && File.Exists(file)) File.Delete(file);
            }
            catch { }

            Properties.Settings.Default.Reload();
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
            catch { /* An unreadable list is treated as no conflict */ }

            return false;
        }

        /// <summary>
        /// The saved pane width, clamped, or the default width if none has been saved yet.
        /// </summary>
        private static int SavedWidth()
        {
            return ClampPaneWidth(ReadSetting(() => Properties.Settings.Default.Width), DefaultPaneWidth);
        }

        /// <summary>
        /// Keeps a width within the saved range; a width of zero or less (a garbage value) is the fallback.
        /// </summary>
        private static int ClampPaneWidth(int width, int fallback)
        {
            if (width <= 0) return fallback;
            return Math.Min(MaxPaneWidth, width);
        }

        /// <summary>
        /// The saved floating height, or the default height if the saved one is not a positive number.
        /// </summary>
        private static int SavedHeight()
        {
            return PositiveHeight(ReadSetting(() => Properties.Settings.Default.Height), DefaultPaneHeight);
        }

        /// <summary>
        /// A height of zero or less (a garbage value) is the fallback. There is no ceiling: Excel keeps
        /// a floating pane on screen.
        /// </summary>
        private static int PositiveHeight(int height, int fallback)
        {
            return height <= 0 ? fallback : height;
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
            catch { /* Excel may not be ready yet */ }
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
            catch { /* Leave the pane hidden */ }
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
            catch { /* Leave the pane hidden */ }
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
            catch { /* Workbook may already be gone */ }
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
            catch { /* Chart sheets and templates can refuse the refresh */ }
        }

        /// <summary>
        /// Once a second, lets the visible pane catch sheet reorders that raise no Excel event.
        /// </summary>
        private void RefreshTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (pane == null || isPaneClosing || this.Application.Ready == false) return;
                if (!IsAlive() || !pane.Visible) return;

                control.RefreshQuietly();
            }
            catch { /* Excel is busy; try again next tick */ }
        }

        /// <summary>
        /// Creates and shows the pane if the workbook is the Contracts file and no pane exists.
        /// </summary>
        private void RestorePane(Excel.Workbook workbook)
        {
            if (pane != null || !IsContractsFile(workbook)) return;
            CreatePane(workbook);
        }

        /// <summary>
        /// Creates the pane on the workbook's window at the saved position and size, fills the list, and shows it.
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
                catch { }

                newControl.Resize += new EventHandler(NavigatorControl_Resize);
                newPane.DockPositionChanged += new EventHandler(Pane_DockPositionChanged);
            }
            catch
            {
                // Don't leave a half-configured pane behind
                try { this.CustomTaskPanes.Remove(newPane); } catch { }
                throw;
            }

            pane = newPane;
            control = newControl;
            isPaneClosing = false;
            recordedDockPosition = pane.DockPosition;
            recordedWidth = ClampPaneWidth(pane.Width, savedWidth);
            recordedHeight = recordedDockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                ? PositiveHeight(pane.Height, savedHeight)
                : savedHeight;

            control.RefreshWorksheets(workbook);
            pane.Visible = true;
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
            resizeSaveTimer.Stop();

            try { oldControl.Resize -= NavigatorControl_Resize; } catch { }
            try { oldPane.DockPositionChanged -= Pane_DockPositionChanged; } catch { }
            try { this.CustomTaskPanes.Remove(oldPane); } catch { /* Already disposed with its window */ }
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
            catch { return false; }
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
            catch { return false; }
        }

        /// <summary>
        /// Restarts the save delay on every resize, so a drag is written once when it ends.
        /// </summary>
        private void NavigatorControl_Resize(object sender, EventArgs e)
        {
            try
            {
                if (pane == null || isPaneClosing || !IsAlive() || !pane.Visible) return;

                resizeSaveTimer.Stop();
                resizeSaveTimer.Start();
            }
            catch { /* Excel is busy */ }
        }

        /// <summary>
        /// Records the resized pane once resizing has settled.
        /// </summary>
        private void ResizeSaveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                resizeSaveTimer.Stop();
                if (pane == null || isPaneClosing || !IsAlive() || !pane.Visible) return;

                RecordPaneChange();
            }
            catch { /* Excel is busy */ }
        }

        /// <summary>
        /// Records the pane when the user docks it left or right or floats it. A pane that just floated
        /// is given the saved height first, since only a floating pane has a height of its own. Floating
        /// fires as the pane detaches; Excel may change the width at the same time.
        /// </summary>
        private void Pane_DockPositionChanged(object sender, EventArgs e)
        {
            try
            {
                if (pane == null || isPaneClosing || !IsAlive() || !pane.Visible) return;

                if (pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating) pane.Height = recordedHeight;

                RecordPaneChange();
            }
            catch { /* Excel is busy */ }
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
            int width = ClampPaneWidth(pane.Width, recordedWidth);
            int height = dock == Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                ? PositiveHeight(pane.Height, recordedHeight)
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
            catch { return false; }
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
            catch { return null; }
        }

        /// <summary>
        /// Unhooks Excel events. The VSTO runtime has already disposed the pane by now.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            resizeSaveTimer.Stop();
            resizeSaveTimer.Dispose();
            refreshTimer.Stop();
            refreshTimer.Dispose();

            this.Application.WorkbookOpen -= Application_WorkbookOpen;
            this.Application.WorkbookActivate -= Application_WorkbookActivate;
            this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            this.Application.SheetActivate -= Application_SheetActivate;

            pane = null;
            control = null;
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
