using System;
using System.Linq;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Sheaside.Controls;
using Sheaside.Models;
using Sheaside.Services;
using Excel = Microsoft.Office.Interop.Excel;

internal static partial class SmokeTests
{
    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hwnd, int message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern bool GetKeyboardState(byte[] state);
    [DllImport("user32.dll")]
    private static extern bool SetKeyboardState(byte[] state);

    private static void Pointer(ReorderableSheetListBox list, int index, Keys modifiers = Keys.None, bool right = false)
    {
        byte[] previous = new byte[256];
        GetKeyboardState(previous);
        var keyboard = (byte[])previous.Clone();
        keyboard[0x10] = (byte)((modifiers & Keys.Shift) != 0 ? 128 : 0);
        keyboard[0x11] = (byte)((modifiers & Keys.Control) != 0 ? 128 : 0);
        try
        {
            // Changes this test thread's key state, not global input or the user's cursor.
            SetKeyboardState(keyboard);
            Rectangle row = list.GetItemRectangle(index);
            var point = new IntPtr(((row.Top + row.Height / 2) << 16) | 8);
            int flags = (right ? 2 : 1) | ((modifiers & Keys.Control) != 0 ? 8 : 0) | ((modifiers & Keys.Shift) != 0 ? 4 : 0);
            SendMessage(list.Handle, right ? 0x0204 : 0x0201, new IntPtr(flags), point);
            SendMessage(list.Handle, right ? 0x0205 : 0x0202, IntPtr.Zero, point);
            // The hidden test control is not under the real cursor, so WinForms
            // does not synthesize MouseClick. Exercise that managed routing separately.
            if (!right) InvokeControl(list, "OnMouseClick", new MouseEventArgs(MouseButtons.Left, 1, 8, row.Top + row.Height / 2, 0));
            list.ContextMenuStrip.Close();
        }
        finally { SetKeyboardState(previous); }
    }

    private static void TestMultiSelection()
    {
        using (var view = new SheetListControl { Size = new Size(320, 500) })
        using (var otherWindow = new SheetListControl())
        {
            var rows = Enumerable.Range(0, 8).Select(i => new WorksheetItem("s" + i, "Row " + i,
                i == 4 ? WorksheetVisibility.VeryHidden : WorksheetVisibility.Visible, i == 3 ? (int?)255 : null,
                i % 2 == 0 ? SheetKind.Worksheet : SheetKind.Chart)).ToArray();
            view.ShowWorksheets("Book", rows, "s0");
            otherWindow.ShowWorksheets("Book", rows, "s0");
            var list = view.Controls.OfType<ReorderableSheetListBox>().Single();
            list.CreateControl();
            int activations = 0;
            view.WorksheetRequested += (s, e) => activations++;
            Pointer(list, 1);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "s1" }), "plain click selects one");
            Check(activations == 1, "plain click activates only target");
            Pointer(list, 3, Keys.Control);
            Check(view.SelectedSheetIds.OrderBy(i => i).SequenceEqual(new[] { "s1", "s3" }), "Ctrl adds mixed chart row");
            Pointer(list, 1, Keys.Control);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "s3" }), "Ctrl toggles off");
            Pointer(list, 1);
            Pointer(list, 4, Keys.Shift);
            Check(view.SelectedSheetIds.OrderBy(i => i).SequenceEqual(new[] { "s1", "s2", "s3", "s4" }), "Shift range includes VeryHidden");
            Pointer(list, 6, Keys.Control);
            Pointer(list, 7, Keys.Control | Keys.Shift);
            Check(view.SelectedSheetIds.Length == 6 && view.SelectedSheetIds.Contains("s7") && !view.SelectedSheetIds.Contains("s5"), "Ctrl Shift adds range to existing set");
            Check(activations == 2, "modifier selection never requests Excel activation");
            string[] selected = view.SelectedSheetIds.OrderBy(i => i).ToArray();
            view.ShowWorksheets("Book", rows, "s5");
            Check(view.SelectedSheetIds.OrderBy(i => i).SequenceEqual(selected), "external active change preserves selection");
            Check(otherWindow.SelectedSheetIds.SequenceEqual(new[] { "s0" }), "separate pane selection for same book");
            Pointer(list, 3, right: true);
            Check(view.SelectedSheetIds.OrderBy(i => i).SequenceEqual(selected), "right click selected row preserves set");
            typeof(ContextMenuStrip).GetMethod("OnOpening", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .Invoke(list.ContextMenuStrip, new object[] { new System.ComponentModel.CancelEventArgs() });
            Check(list.ContextMenuStrip.Items.Cast<ToolStripItem>().All(item => !item.Enabled), "multi selection disables single-sheet menus");
            Check(!view.Controls.OfType<Button>().Single().Enabled, "multi selection disables add button");
            int starts = 0;
            view.DragStarting += (s, e) => starts++;
            InvokeControl(list, "OnMouseDown", new MouseEventArgs(MouseButtons.Left, 1, 8, 8, 0));
            InvokeControl(list, "OnMouseMove", new MouseEventArgs(MouseButtons.Left, 0, 100, 100, 0));
            Check(starts == 0, "multi selection cannot start reorder");
            Pointer(list, 0, right: true);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "s0" }), "right click unselected row selects just it");
            InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Control | Keys.A));
            Check(view.SelectedSheetIds.Length == 8 && activations == 2, "Ctrl A selects all without activation");
            InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Escape));
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "s5" }), "Escape selects active sheet only");
            Pointer(list, 2);
            Pointer(list, 5, Keys.Control);
            rows[2] = new WorksheetItem("s2", "Renamed", true);
            var reversed = rows.Reverse().ToArray();
            view.ShowWorksheets("Book", reversed, "s0", "s2");
            Check(view.SelectedSheetIds.OrderBy(i => i).SequenceEqual(new[] { "s2", "s5" }), "rename reorder and preferred ID preserve multi selection");
            Check(list.AnchorIndex == 2, "range anchor follows identity after reorder");
            view.ShowWorksheets("Book", reversed.Where(i => i.Id != "s2").ToArray(), "s0");
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "s5" }), "deleted member removed");
            view.ShowWorksheets("Book", reversed.Where(i => i.Id != "s2" && i.Id != "s5").ToArray(), "s0");
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "s0" }), "all selected rows deleted falls back to active");
            Pointer(list, list.SelectedIndex, Keys.Control);
            view.ShowWorksheets("Book", reversed.Where(i => i.Id != "s2" && i.Id != "s5").ToArray(), "s0");
            Check(view.SelectedSheetIds.Length == 0, "Ctrl toggled empty selection survives refresh");
            var many = Enumerable.Range(0, 60).Select(i => new WorksheetItem("m" + i, "Many " + i, true)).ToArray();
            view.ShowWorksheets("Book", many, "m0");
            list.TopIndex = 30;
            int top = list.TopIndex;
            InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Control | Keys.A));
            Check(list.TopIndex == top && view.SelectedSheetIds.Length == 60, "select all does not scroll");
            int invalidations = 0;
            list.Invalidated += (s, e) => invalidations++;
            view.ShowWorksheets("Book", many, "m0");
            Check(invalidations == 0 && list.TopIndex == top, "unchanged refresh does not repaint or scroll");
            InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Escape));
            Check(list.TopIndex == top && view.SelectedSheetIds.SequenceEqual(new[] { "m0" }), "Escape preserves scroll even when active row is offscreen");
        }
    }

    private static void TestExcelSelection()
    {
        var excel = new Excel.Application { Visible = false, DisplayAlerts = false };
        Excel.Workbook book = null;
        Excel.Workbook other = null;
        try
        {
            book = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            var sheet = (Excel.Worksheet)book.Worksheets[1];
            var chart = (Excel.Chart)book.Charts.Add();
            var service = new ExcelWorksheetService(excel);
            int handle = book.Windows[1].Hwnd;
            string id = service.GetWorksheetId(sheet);
            string chartId = service.GetWorksheetId(chart);
            service.RenameWorksheet(service.CaptureTarget(handle, id), "Renamed");
            chart.Name = "ChartRenamed";
            Check(service.GetWorksheetId(sheet) == id && service.GetWorksheetId(chart) == chartId, "sheet and chart IDs survive rename even without CodeName");
            using (var view = new SheetListControl())
            using (var secondView = new SheetListControl())
            {
                view.ShowWorksheets(book.Name, service.GetWorksheets(book), service.GetActiveWorksheetId(book.Windows[1]));
                var list = view.Controls.OfType<ReorderableSheetListBox>().Single();
                view.WorksheetRequested += (s, e) => service.ActivateWorksheet(handle, e.Worksheet.Id);
                var activeBefore = book.Windows[1].ActiveSheet;
                InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Control | Keys.A));
                Check(book.Windows[1].SelectedSheets.Count == 1 && Equals(activeBefore, book.Windows[1].ActiveSheet), "local selection never groups Excel sheets");
                var second = book.NewWindow();
                secondView.ShowWorksheets(book.Name, service.GetWorksheets(book), service.GetActiveWorksheetId(second));
                Check(secondView.SelectedSheetIds.Length == 1 && view.SelectedSheetIds.Length == 2, "same book windows keep independent sets");
                chart.Move(After: sheet);
                view.ShowWorksheets(book.Name, service.GetWorksheets(book), service.GetActiveWorksheetId(second));
                Check(view.SelectedSheetIds.Contains(id) && view.SelectedSheetIds.Contains(chartId), "Excel native reorder preserves IDs and selected set");
                other = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                var sameName = (Excel.Worksheet)other.Worksheets[1];
                sameName.Name = sheet.Name;
                Check(service.GetWorksheetId(sameName) != id, "same name in other book has independent identity");
                second.Activate();
                sheet.Activate();
                for (int i = 1; i <= book.Windows.Count; i++)
                {
                    book.Windows[i].Activate();
                    sheet.Activate();
                }
                chart.Delete();
                view.ShowWorksheets(book.Name, service.GetWorksheets(book), id);
                Check(view.SelectedSheetIds.SequenceEqual(new[] { id }), "Excel deletion prunes selected chart");
                var replacement = (Excel.Chart)book.Charts.Add();
                replacement.Name = "ChartRenamed";
                Check(service.GetWorksheetId(replacement) != chartId, "recreated same name does not inherit deleted selection");
            }
            service.ClearIdentities();
        }
        finally
        {
            if (other != null) other.Close(false);
            if (book != null) book.Close(false);
            excel.Quit();
        }
    }
}
