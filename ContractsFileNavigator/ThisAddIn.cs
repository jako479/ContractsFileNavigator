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

        /// <summary>Docked side used when the saved dock position cannot be read.</summary>
        private const Office.MsoCTPDockPosition FallbackDockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;

        /// <summary>Pane width in points used when the saved width cannot be read. Points already scale with DPI.</summary>
        private const int FallbackWidth = 150;

        /// <summary>
        /// Floating pane height in points used when the saved height cannot be read. Only a
        /// floating pane has a height of its own; a docked pane is stretched to the window.
        /// </summary>
        private const int FallbackHeight = 400;

        /// <summary>The settings file .NET keeps per user: the only file settings recovery may delete.</summary>
        private const string UserConfigFileName = "user.config";

        /// <summary>
        /// Registry name of the general-purpose Sheet Navigator add-in. Both add-ins manage the
        /// same kind of pane, so this one stays idle while that one is installed.
        /// </summary>
        private const string SheetNavigatorProgId = "SheetNavigator";

        /// <summary>
        /// How many half-second checks a pane change waits while the file's close is pending. A pane
        /// still shown after that many survived a cancelled close, so the change is the user's.
        /// </summary>
        private const int ClosingRecheckLimit = 4;

        /// <summary>The pane and its list, created on the Contracts window while the file is open.</summary>
        private CustomTaskPane pane;
        private ContractsFileNavigatorControl control;

        /// <summary>Delays recording a resize, dock or float until the user has stopped dragging.</summary>
        private readonly Timer resizeSaveTimer = new Timer { Interval = 500 };

        /// <summary>
        /// Excel raises no event when a sheet tab is dragged to a new position, so the visible pane
        /// re-checks the sheet list once a second. Only a changed list triggers a rebuild.
        /// </summary>
        private readonly Timer refreshTimer = new Timer { Interval = 1000 };

        /// <summary>
        /// Set while the Contracts file is closing, so a layout change raised by the teardown is not
        /// taken for the user's.
        /// </summary>
        private bool isPaneClosing = false;

        /// <summary>
        /// How many times a queued change has waited on <see cref="isPaneClosing"/>; see <see cref="ClosingRecheckLimit"/>.
        /// </summary>
        private int closingRechecks;

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
                    MessageBox.Show(ExcelOwner(), "Contracts File Navigator is not running because Sheet Navigator is installed.\n\n" +
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
                MessageBox.Show(ExcelOwner(), $"Contracts File Navigator failed to initialize components.\n\nError Details: {ex.Message}",
                                "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Carries settings forward after an Office update, since the settings folder is named after the Excel build.
        /// </summary>
        private void UpgradeSettingsIfNeeded()
        {
            try
            {
                if (!ReadSetting(() => Properties.Settings.Default.UpgradeRequired)) return;

                Properties.Settings.Default.Upgrade();
                WriteSetting(() => Properties.Settings.Default.UpgradeRequired = false);

                // On that first run every setting is also written, so the file lists them all for editing;
                // assigning a setting marks it changed, which is what makes Save write it out
                bool enabled = ReadSetting(() => Properties.Settings.Default.Enabled);
                int width = SavedWidth();
                int height = SavedHeight();
                string dockName = DockPositionName(SavedDockPosition());
                WriteSetting(() =>
                {
                    Properties.Settings.Default.Enabled = enabled;
                    Properties.Settings.Default.Width = width;
                    Properties.Settings.Default.Height = height;
                    Properties.Settings.Default.DockPosition = dockName;
                });

                SaveSettings();
            }
            catch (Exception ex)
            {
                // The flag stays set, so the next start tries again; without the log that retry would be invisible
                Diagnostics.Write("Settings upgrade failed: " + ex);
            }
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
                if (!TryRecoverSettings(ex)) throw;
                Properties.Settings.Default.Reload();
                return read();
            }
        }

        /// <summary>
        /// Changes a setting in memory. If .NET reports the settings file corrupt, the file is reset and the write retried.
        /// </summary>
        private static void WriteSetting(Action write)
        {
            try
            {
                write();
            }
            catch (ConfigurationErrorsException ex)
            {
                if (!TryRecoverSettings(ex)) throw;
                Properties.Settings.Default.Reload();
                write();
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
                if (!TryRecoverSettings(ex)) throw;
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// Deletes a corrupt user.config (a crash can leave it truncated, and then every settings call
        /// throws until it is gone). True if the file was deleted; the caller reloads if it needs the defaults.
        /// </summary>
        private static bool TryRecoverSettings(ConfigurationErrorsException ex)
        {
            // Only the add-in's own user.config inside the user's profile, and only on a parse error;
            // anything else (another file, a passing "file in use") is logged and left alone
            string file = ConfigFileNamedBy(ex);
            if (!IsOwnUserConfig(file) || !IsParseError(ex))
            {
                Diagnostics.Write("Settings unavailable, file kept: " + (file ?? "(unknown path)") + " | " + ex);
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
        /// The saved pane width, clamped; a width of zero or less (a garbage value) is the fallback.
        /// </summary>
        private static int SavedWidth()
        {
            return PaneSizeRules.ClampWidth(ReadSetting(() => Properties.Settings.Default.Width), FallbackWidth);
        }

        /// <summary>
        /// The saved floating height, clamped; a height of zero or less (a garbage value) is the fallback.
        /// </summary>
        private static int SavedHeight()
        {
            return PaneSizeRules.ClampHeight(ReadSetting(() => Properties.Settings.Default.Height), FallbackHeight);
        }

        /// <summary>
        /// The saved dock position; anything but "Left", "Right" or "Floating" (a garbage value) is the fallback.
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
            return FallbackDockPosition;
        }

        /// <summary>
        /// The dock position as written in settings: "Left", "Right" or "Floating"; top and bottom throw.
        /// </summary>
        private static string DockPositionName(Office.MsoCTPDockPosition dock)
        {
            // Top and bottom have no saved form and are only reachable if Excel rejected the dock restriction
            switch (dock)
            {
                case Office.MsoCTPDockPosition.msoCTPDockPositionLeft: return "Left";
                case Office.MsoCTPDockPosition.msoCTPDockPositionRight: return "Right";
                case Office.MsoCTPDockPosition.msoCTPDockPositionFloating: return "Floating";
                default: throw new ArgumentOutOfRangeException(nameof(dock), dock, "Only a side-docked or floating pane has a saved form");
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
        /// Records anything the user changed just before the close, then flags the pane as closing so a
        /// layout change raised by the teardown is not taken for the user's.
        /// </summary>
        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            try
            {
                if (pane == null || !IsContractsFile(workbook)) return;

                // A change still waiting for its delay happened before the close began, so it is the user's
                FlushPendingSaves();

                // Excel asks about unsaved changes after this event, so the close may still be cancelled; the
                // flag is cleared when the user works in the file again or when a change outlives the re-checks.
                // The pane itself is replaced when the file is next opened
                isPaneClosing = true;
                closingRechecks = 0;
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

                // A macro with screen updating off may activate every sheet in turn; the list catches up on the next tick
                if (IsPaneWindowActive() && this.Application.ScreenUpdating)
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
        /// cancelled close, and lets the visible pane catch sheet reorders that raise no Excel event.
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
            try
            {
                if (pane != null && !IsAlive()) RemovePane();
                if (pane != null || !IsContractsFile(workbook)) return;
                CreatePane(workbook);
            }
            catch (Exception ex) { Diagnostics.Write("RestorePane failed: " + ex); }
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

            Office.MsoCTPDockPosition dock;
            int width;
            int height;
            try
            {
                // Dock position before size: the pane API expects that order. Only a floating pane takes a height.
                // A value Excel rejects must not stop the pane from being created, or the file could never get one.
                ApplySavedOrDefault("dock position", () => newPane.DockPosition = savedDockPosition, () => newPane.DockPosition = FallbackDockPosition);
                ApplySavedOrDefault("width", () => newPane.Width = savedWidth, () => newPane.Width = FallbackWidth);
                if (newPane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating)
                {
                    ApplySavedOrDefault("height", () => newPane.Height = savedHeight, () => newPane.Height = FallbackHeight);
                }

                // Width is meaningless when docked top or bottom, so keep the pane on a side or floating. Excel's
                // "NoHorizontal" is the restriction compatible with a side-docked pane, despite its name.
                // Optional: a rejected restriction must never stop the pane from working.
                try { newPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal; }
                catch (Exception ex) { Diagnostics.Write("Dock restriction rejected: " + ex); }

                newPane.DockPositionChanged += new EventHandler(Pane_DockPositionChanged);
                newControl.Resize += new EventHandler(Control_Resize);

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
                try { this.CustomTaskPanes.Remove(newPane); } catch (Exception ex) { Diagnostics.Write("Pane removal after failed creation failed: " + ex); }
                throw;
            }

            // The fields are set only once the pane is shown, so a failure part-way leaves no hidden pane
            // behind and the next activation tries again
            pane = newPane;
            control = newControl;
            isPaneClosing = false;
            isFloatingHeightPending = false;
            recordedDockPosition = dock;
            recordedWidth = width;
            recordedHeight = height;
        }

        /// <summary>
        /// Applies a saved pane value; if Excel rejects it the default is tried, and if that is rejected
        /// too the pane keeps Excel's own value.
        /// </summary>
        private static void ApplySavedOrDefault(string property, Action applySaved, Action applyDefault)
        {
            try
            {
                applySaved();
                return;
            }
            catch (Exception ex) { Diagnostics.Write($"Saved {property} rejected, trying the default: {ex.Message}"); }

            try
            {
                applyDefault();
            }
            catch (Exception ex) { Diagnostics.Write($"Default {property} rejected, keeping Excel's: {ex.Message}"); }
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

            try { oldPane.DockPositionChanged -= Pane_DockPositionChanged; } catch (Exception ex) { Diagnostics.Write("Dock unhook failed: " + ex); }
            try { oldControl.Resize -= Control_Resize; } catch (Exception ex) { Diagnostics.Write("Resize unhook failed: " + ex); }
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
        /// True while the pane exists and is shown: the only time a layout event concerns it.
        /// </summary>
        private bool IsShown()
        {
            return pane != null && IsAlive() && pane.Visible;
        }

        /// <summary>
        /// Queues the resize so a drag is written once, when it ends.
        /// </summary>
        private void Control_Resize(object sender, EventArgs e)
        {
            try
            {
                if (!IsShown()) return;

                QueueResizeSave();
            }
            catch (Exception ex) { Diagnostics.Write("Resize failed: " + ex); }
        }

        /// <summary>
        /// Queues a resize, dock or float to be recorded once the delay has passed, restarting the delay.
        /// </summary>
        private void QueueResizeSave()
        {
            resizeSaveTimer.Stop();
            resizeSaveTimer.Start();
        }

        /// <summary>
        /// Records the layout change that has settled, after first giving a pane that just floated its
        /// saved height; a change still waiting on a pending close is queued again.
        /// </summary>
        private void ResizeSaveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                resizeSaveTimer.Stop();
                if (!IsShown()) return;
                if (StillClosing())
                {
                    QueueResizeSave();
                    return;
                }

                if (isFloatingHeightPending) ApplyFloatingHeight();
                RecordPaneChange();
            }
            catch (Exception ex) { Diagnostics.Write("Resize save failed: " + ex); }
        }

        /// <summary>
        /// Records a change still waiting for its delay now, for a close or shutdown that would otherwise drop it.
        /// </summary>
        private void FlushPendingSaves()
        {
            if (!resizeSaveTimer.Enabled) return;
            ResizeSaveTimer_Tick(null, null);
        }

        /// <summary>
        /// True while a queued change must keep waiting because the file's close is still pending;
        /// a pane still shown after the last re-check survived a cancelled close.
        /// </summary>
        private bool StillClosing()
        {
            if (!isPaneClosing) return false;
            if (++closingRechecks < ClosingRecheckLimit) return true;

            isPaneClosing = false;
            return false;
        }

        /// <summary>
        /// Queues the dock or float to be recorded; a pane that just floated is marked so the resize save
        /// gives it the saved height first.
        /// </summary>
        private void Pane_DockPositionChanged(object sender, EventArgs e)
        {
            try
            {
                if (!IsShown()) return;

                // Excel rejects property sets inside this handler and may still resize the pane as the drag
                // ends, so the saved height is applied, and the pane recorded, by the resize save. Docking goes
                // through the same delay, since Excel may change the width at the same time.
                if (pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating) isFloatingHeightPending = true;

                QueueResizeSave();
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
        /// floating, height to the settings. A layout event that changed nothing since the last write writes nothing.
        /// </summary>
        private void RecordPaneChange()
        {
            // A width or height of zero or less is not a change, the saved value stays; a docked pane's
            // height is Excel's, so the saved height stays too
            Office.MsoCTPDockPosition dock = pane.DockPosition;
            int width = PaneSizeRules.ClampWidth(pane.Width, recordedWidth);
            int height = dock == Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                ? PaneSizeRules.ClampHeight(pane.Height, recordedHeight)
                : recordedHeight;
            if (dock == recordedDockPosition && width == recordedWidth && height == recordedHeight) return;

            // Top and bottom are only reachable if Excel rejected the dock restriction; they have no saved form
            if (dock == Office.MsoCTPDockPosition.msoCTPDockPositionTop || dock == Office.MsoCTPDockPosition.msoCTPDockPositionBottom) return;

            string dockName = DockPositionName(dock);
            WriteSetting(() =>
            {
                Properties.Settings.Default.DockPosition = dockName;
                Properties.Settings.Default.Width = width;
                Properties.Settings.Default.Height = height;
            });
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
        /// Records a pending change, stops the timers, unhooks the Excel and pane events and drops the
        /// pane, each step on its own so one failure cannot skip the rest.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // A change still waiting for its delay would otherwise be lost with the timer
            RunLogged("Pending saves", FlushPendingSaves);

            RunLogged("Resize timer stop", () => { resizeSaveTimer.Stop(); resizeSaveTimer.Dispose(); });
            RunLogged("Refresh timer stop", () => { refreshTimer.Stop(); refreshTimer.Dispose(); });

            RunLogged("WorkbookOpen unhook", () => this.Application.WorkbookOpen -= Application_WorkbookOpen);
            RunLogged("WorkbookActivate unhook", () => this.Application.WorkbookActivate -= Application_WorkbookActivate);
            RunLogged("WorkbookAfterSave unhook", () => this.Application.WorkbookAfterSave -= Application_WorkbookAfterSave);
            RunLogged("WorkbookBeforeClose unhook", () => this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose);
            RunLogged("SheetActivate unhook", () => this.Application.SheetActivate -= Application_SheetActivate);

            RunLogged("Pane removal", RemovePane);

            // A repeat count still pending would otherwise be lost with Excel
            Diagnostics.Flush();
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

        /// <summary>
        /// Excel's main window as a message box owner, so the box stays in front of Excel; null if Excel will not say.
        /// </summary>
        private IWin32Window ExcelOwner()
        {
            try
            {
                return new WindowHandle(new IntPtr(this.Application.Hwnd));
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Excel window lookup failed: " + ex);
                return null;
            }
        }

        #region VSTO Generated Code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion

        /// <summary>
        /// Wraps a native window handle for use as a message box owner.
        /// </summary>
        private sealed class WindowHandle : IWin32Window
        {
            public IntPtr Handle { get; }

            /// <summary>
            /// Keeps the handle; nothing is created or owned.
            /// </summary>
            public WindowHandle(IntPtr handle)
            {
                Handle = handle;
            }
        }
    }
}
