using Microsoft.Office.Tools;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ContractsFileNavigator
{
    /// <summary>
    /// Excel add-in that shows a "Worksheets" task pane listing the sheets of Contracts.xlsx.
    /// The pane opens every time that file is opened. Closing it hides it until the file is
    /// opened again. The pane width is remembered per user, and the add-in can be switched off
    /// by setting <c>Enabled</c> to False in the user settings file (user.config under
    /// %LOCALAPPDATA%\Microsoft_Corporation, in the folder named after ContractsFileNavigator.vsto).
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>Only a saved workbook with this file name gets the pane.</summary>
        private const string ContractsFileName = "Contracts.xlsx";

        /// <summary>Saved widths are clamped to this range on both save and load.</summary>
        private const int MinPaneWidth = 100;
        private const int MaxPaneWidth = 800;

        /// <summary>
        /// Registry name of the general-purpose Sheet Navigator add-in. Both add-ins manage the
        /// same kind of pane, so this one stays idle while that one is installed.
        /// </summary>
        private const string SheetNavigatorProgId = "SheetNavigator";

        private int? cachedDefaultWidth = null;

        /// <summary>The pane and its list, created on the Contracts window while the file is open.</summary>
        private CustomTaskPane pane;
        private ContractsFileNavigatorControl control;

        /// <summary>Delays the width save until the user has stopped dragging the pane border.</summary>
        private readonly Timer resizeSaveTimer = new Timer { Interval = 500 };

        /// <summary>
        /// Set while the Contracts file is closing, so a resize raised by the teardown is not saved.
        /// </summary>
        private bool isPaneClosing = false;

        /// <summary>
        /// Wires up Excel events, unless the add-in is disabled or Sheet Navigator is installed.
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                UpgradeSettingsIfNeeded();
                if (!Properties.Settings.Default.Enabled) return;

                if (IsSheetNavigatorInstalled())
                {
                    MessageBox.Show("Contracts File Navigator is not running because Sheet Navigator is installed.\n\n" +
                                    "Only one of the two add-ins can be installed. Uninstall one of them in Windows Settings, Apps.",
                                    "Contracts File Navigator", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                resizeSaveTimer.Tick += new EventHandler(ResizeSaveTimer_Tick);

                this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
                this.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookActivate);
                this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookBeforeClose);
                this.Application.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(Application_SheetActivate);

                InitialCheckOnLoad();
            }
            catch (Exception ex)
            {
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
                Properties.Settings settings = Properties.Settings.Default;
                if (!settings.UpgradeRequired) return;

                settings.Upgrade();
                settings.UpgradeRequired = false;

                // Assigning a setting marks it changed, which is what makes Save write it out
                bool enabled = settings.Enabled;
                settings.Enabled = enabled;
                settings.LastWidth = SavedWidth();

                settings.Save();
            }
            catch { /* Nothing to carry over */ }
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
        /// Default pane width, scaled from 150 on a 1920-wide screen and kept within 120-400.
        /// </summary>
        private int EnsureResponsiveDefaultWidth()
        {
            if (cachedDefaultWidth.HasValue) return cachedDefaultWidth.Value;

            int defaultWidth = 150;

            try
            {
                int screenWidth = Screen.PrimaryScreen.Bounds.Width;
                double targetScalePercentage = defaultWidth / 1920.0;
                int calculatedWidth = (int)Math.Round(screenWidth * targetScalePercentage);

                cachedDefaultWidth = Math.Max(120, Math.Min(400, calculatedWidth));
            }
            catch
            {
                cachedDefaultWidth = defaultWidth;
            }

            return cachedDefaultWidth.Value;
        }

        /// <summary>
        /// The saved pane width, clamped, or the default width if none has been saved yet.
        /// </summary>
        private int SavedWidth()
        {
            int width = Properties.Settings.Default.LastWidth;
            return width > 0 ? ClampPaneWidth(width) : EnsureResponsiveDefaultWidth();
        }

        private static int ClampPaneWidth(int width)
        {
            return Math.Max(MinPaneWidth, Math.Min(MaxPaneWidth, width));
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
        /// Keeps the list's highlight on the active sheet.
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

                if (pane.Visible) control.RefreshWorksheets(workbook);
            }
            catch { /* Chart sheets and templates can refuse the refresh */ }
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
        /// Creates the pane on the workbook's window at the saved width, fills the list, and shows it.
        /// </summary>
        private void CreatePane(Excel.Workbook workbook)
        {
            Excel.Window window = FirstWindowOf(workbook);
            if (window == null) return;

            ContractsFileNavigatorControl newControl = new ContractsFileNavigatorControl
            {
                Workbook = workbook
            };

            CustomTaskPane newPane = this.CustomTaskPanes.Add(newControl, "Worksheets", window);
            try
            {
                newPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                newPane.Width = SavedWidth();

                // Width is meaningless when docked top or bottom, so keep the pane on a side. Excel's
                // "NoHorizontal" is the restriction compatible with a left-docked pane, despite its name.
                // Optional: a rejected restriction must never stop the pane from working.
                try { newPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal; }
                catch { }

                newControl.Resize += new EventHandler(NavigatorControl_Resize);
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

            control.RefreshWorksheets(workbook);
            pane.Visible = true;
        }

        /// <summary>
        /// Drops the pane. The resize handler is unhooked first so the removal saves nothing.
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
            try { this.CustomTaskPanes.Remove(oldPane); } catch { /* Already disposed with its window */ }
        }

        /// <summary>
        /// Restarts the save delay on every resize, so a drag is written once when it ends.
        /// </summary>
        private void NavigatorControl_Resize(object sender, EventArgs e)
        {
            if (pane == null || isPaneClosing) return;

            try
            {
                if (!pane.Visible) return;
            }
            catch { return; }

            resizeSaveTimer.Stop();
            resizeSaveTimer.Start();
        }

        /// <summary>
        /// Saves the pane's width once resizing has settled.
        /// </summary>
        private void ResizeSaveTimer_Tick(object sender, EventArgs e)
        {
            resizeSaveTimer.Stop();
            if (pane == null || isPaneClosing) return;

            try
            {
                if (!pane.Visible) return;
                SaveWidth(pane.Width);
            }
            catch { /* Excel is busy */ }
        }

        /// <summary>
        /// Records the width, clamped, unless it is already on file.
        /// </summary>
        private void SaveWidth(int paneWidth)
        {
            int width = ClampPaneWidth(paneWidth);
            if (Properties.Settings.Default.LastWidth == width) return;

            Properties.Settings.Default.LastWidth = width;
            Properties.Settings.Default.Save();
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
