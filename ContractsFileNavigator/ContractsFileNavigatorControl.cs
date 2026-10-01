using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ContractsFileNavigator
{
    /// <summary>
    /// The pane's content: a list of the workbook's visible worksheets. Clicking a name activates
    /// that sheet in this pane's own window, so several windows on one workbook stay independent.
    /// </summary>
    public partial class ContractsFileNavigatorControl : UserControl
    {
        /// <summary>
        /// Excel's generic "can't do that now" error, raised for example while a sheet tab name is being typed.
        /// </summary>
        private const int ExcelBusyHResult = unchecked((int)0x800A03EC);

        /// <summary>
        /// True while this control moves the highlight itself, so only the user's selections trigger a jump.
        /// </summary>
        private bool isUpdatingSelection;

        /// <summary>
        /// The Excel window this pane belongs to. Each window has its own active sheet.
        /// </summary>
        internal Excel.Window Window { get; set; }

        /// <summary>
        /// The workbook shown in that window. Falls back to the active workbook.
        /// </summary>
        internal Excel.Workbook Workbook { get; set; }

        /// <summary>
        /// The workbook to list: the one this pane was created for, else Excel's active workbook.
        /// </summary>
        private Excel.Workbook TargetWorkbook
        {
            get { return Workbook ?? Globals.ThisAddIn.Application.ActiveWorkbook; }
        }

        /// <summary>
        /// Builds the list and hooks its mouse and keyboard events.
        /// </summary>
        public ContractsFileNavigatorControl()
        {
            InitializeComponent();

            // A single click selects and jumps; keys are swallowed so the highlight can only move by mouse
            this.worksheetList.SelectedIndexChanged += new EventHandler(WorksheetList_SelectedIndexChanged);
            this.worksheetList.KeyDown += new KeyEventHandler(WorksheetList_KeyDown);

            // Excel raises no event for a sheet rename or reorder, so check as the pointer arrives
            this.worksheetList.MouseEnter += new EventHandler(WorksheetList_MouseEnter);
        }

        /// <summary>
        /// Brings the list in line with the workbook: rebuilds it only if the visible sheet names
        /// changed (added, removed, renamed, reordered, hidden or unhidden), then highlights the active sheet.
        /// </summary>
        public void RefreshWorksheets(Excel.Workbook activeWorkbook)
        {
            if (IsDisposed || activeWorkbook == null) return;

            List<string> names = VisibleSheetNames(activeWorkbook);
            if (!SameAsList(names))
            {
                RunWithoutJumping(() =>
                {
                    this.worksheetList.BeginUpdate();
                    try
                    {
                        this.worksheetList.Items.Clear();
                        foreach (string name in names) this.worksheetList.Items.Add(name);
                    }
                    finally
                    {
                        this.worksheetList.EndUpdate();
                    }
                });
            }

            HighlightActiveSheet();
        }

        /// <summary>
        /// Moves the highlight to this window's active sheet without triggering a jump. A chart sheet
        /// is not listed, so it clears the highlight; a stale one would block a click back to that sheet.
        /// </summary>
        public void HighlightActiveSheet()
        {
            if (IsDisposed) return;

            RunWithoutJumping(() =>
            {
                try
                {
                    object active = Window != null ? Window.ActiveSheet : TargetWorkbook?.ActiveSheet;
                    if (active is Excel.Worksheet currentSheet)
                    {
                        this.worksheetList.SelectedItem = currentSheet.Name;
                    }
                    else
                    {
                        this.worksheetList.ClearSelected();
                    }
                }
                catch (Exception ex) { Diagnostics.Write("Highlight failed: " + ex); }
            });
        }

        /// <summary>
        /// The names of the workbook's visible worksheets, in tab order.
        /// </summary>
        private static List<string> VisibleSheetNames(Excel.Workbook workbook)
        {
            List<string> names = new List<string>();
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible) names.Add(worksheet.Name);
            }
            return names;
        }

        /// <summary>
        /// True if the list already shows exactly these names in this order. The comparison is
        /// case-sensitive so a rename that only changed case still counts.
        /// </summary>
        private bool SameAsList(List<string> names)
        {
            if (this.worksheetList.Items.Count != names.Count) return false;
            for (int i = 0; i < names.Count; i++)
            {
                if (!string.Equals(this.worksheetList.Items[i] as string, names[i], StringComparison.Ordinal)) return false;
            }
            return true;
        }

        /// <summary>
        /// Runs a change to the list with the jump-on-select behavior suspended.
        /// </summary>
        private void RunWithoutJumping(Action action)
        {
            bool wasUpdating = isUpdatingSelection;
            isUpdatingSelection = true;
            try
            {
                action();
            }
            finally
            {
                isUpdatingSelection = wasUpdating;
            }
        }

        /// <summary>
        /// Refreshes if Excel will answer, otherwise just re-highlights the active sheet.
        /// </summary>
        internal void RefreshQuietly()
        {
            try
            {
                Excel.Workbook workbook = TargetWorkbook;
                if (workbook == null) return;

                if (IsExcelEditing(Globals.ThisAddIn.Application))
                {
                    HighlightActiveSheet();
                }
                else
                {
                    RefreshWorksheets(workbook);
                }
            }
            catch (Exception ex) { Diagnostics.Write("Quiet refresh failed: " + ex); }
        }

        /// <summary>
        /// True while Excel is busy or mid-edit (typing in a cell or a sheet tab name).
        /// Most Ribbon commands are disabled then, which is the only reliable signal.
        /// </summary>
        private static bool IsExcelEditing(Excel.Application app)
        {
            if (app.Ready == false || app.Interactive == false) return true;
            return app.CommandBars.GetEnabledMso("FileNewDefault") == false;
        }

        /// <summary>
        /// The workbook's worksheet with this name, or null if there is none (it was renamed or removed).
        /// </summary>
        private static Excel.Worksheet FindSheet(Excel.Workbook workbook, string name)
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (string.Equals(worksheet.Name, name, StringComparison.OrdinalIgnoreCase)) return worksheet;
            }
            return null;
        }

        /// <summary>
        /// Refreshes as the pointer arrives, so the list is current before a click lands.
        /// </summary>
        private void WorksheetList_MouseEnter(object sender, EventArgs e)
        {
            RefreshQuietly();
        }

        /// <summary>
        /// Blocks keyboard navigation in the list; Excel's Ctrl+PgUp/PgDn already covers that.
        /// </summary>
        private void WorksheetList_KeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        /// <summary>
        /// Activates the selected sheet in this pane's window. If the jump cannot happen,
        /// the list is refreshed and the highlight returns to the sheet the window is still on.
        /// </summary>
        private void WorksheetList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUpdatingSelection || IsDisposed) return;

            Excel.Application app = Globals.ThisAddIn.Application;
            bool jumped = false;
            bool? previousScreenUpdating = null;
            string failure = null;

            try
            {
                if (this.worksheetList.SelectedItem == null) return;
                string selectedSheetName = this.worksheetList.SelectedItem.ToString();

                Excel.Workbook workbook = TargetWorkbook;
                if (workbook == null || IsExcelEditing(app)) return;

                // The sheet may have been renamed or removed since the list was filled
                Excel.Worksheet targetSheet = FindSheet(workbook, selectedSheetName);
                if (targetSheet == null) return;

                previousScreenUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;

                // Worksheet.Activate acts on the workbook's active window, so make it this one first
                Window?.Activate();
                targetSheet.Activate();
                jumped = true;
            }
            catch (COMException ex) when (ex.HResult == ExcelBusyHResult)
            {
                // Excel refused because it is mid-edit; expected, so the user is not told
                Diagnostics.Write("Jump refused, Excel is busy: " + ex.Message);
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Jump failed: " + ex);
                failure = ex.Message;
            }
            finally
            {
                // Put Excel's setting back exactly as found
                if (previousScreenUpdating.HasValue)
                {
                    try { app.ScreenUpdating = previousScreenUpdating.Value; }
                    catch (Exception ex) { Diagnostics.Write("ScreenUpdating restore failed: " + ex); }
                }

                if (!jumped) RefreshQuietly();
            }

            // Only after ScreenUpdating is back on, so Excel repaints while the box is up
            if (failure != null)
            {
                MessageBox.Show($"Could not jump to sheet: {failure}", "Navigation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
