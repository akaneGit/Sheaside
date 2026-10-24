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
            if (row.Top < 0 || row.Bottom > list.ClientSize.Height)
            {
                list.TopIndex = index;
                row = list.GetItemRectangle(index);
            }
            var point = new IntPtr(((row.Top + row.Height / 2) << 16) | 8);
            int flags = (right ? 2 : 1) | ((modifiers & Keys.Control) != 0 ? 8 : 0) | ((modifiers & Keys.Shift) != 0 ? 4 : 0);
            SendMessage(list.Handle, right ? 0x0204 : 0x0201, new IntPtr(flags), point);
            SendMessage(list.Handle, right ? 0x0205 : 0x0202, IntPtr.Zero, point);
            list.ContextMenuStrip.Close();
        }
        finally { SetKeyboardState(previous); }
    }

    private static void TestMultiSelection()
    {
        var settings = new UserPreferences(); settings.SetShowHidden(true); settings.SetShowVeryHidden(true);
        using (var view = new SheetListControl(settings) { Size = new Size(320, 500) })
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
            Check(!list.ContextMenuStrip.Items[0].Enabled && !list.ContextMenuStrip.Items[3].Enabled && list.ContextMenuStrip.Items[1].Enabled, "multi selection disables add rename and enables common delete");
            Check(!view.AddButton.Enabled, "multi selection disables add button");
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

    private static void TestPointerRegression()
    {
        using (var view = new SheetListControl { Size = new Size(320, 600) })
        {
            var rows = Enumerable.Range(0, 10).Select(i => new WorksheetItem("p" + i, "Row " + i, true)).ToArray();
            view.ShowWorksheets("Book", rows, "p0");
            var list = view.Controls.OfType<ReorderableSheetListBox>().Single();
            list.CreateControl();
            int starts = 0, activations = 0;
            view.DragStarting += (s, e) => starts++; // Cancel=true prevents a blocking OLE loop.
            view.WorksheetRequested += (s, e) => activations++;
            Pointer(list, 0);
            Pointer(list, 8, Keys.Control);
            Pointer(list, 6, Keys.Control | Keys.Shift);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "p0", "p6", "p7", "p8" }), "Ctrl Shift preserves disjoint existing selection");
            Pointer(list, 9, Keys.Control | Keys.Shift);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "p0", "p6", "p7", "p8", "p9" }) && list.AnchorIndex == 8, "successive Ctrl Shift unions retain anchor");
            Pointer(list, 7, Keys.Shift);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "p7", "p8" }), "Shift replaces previous union");
            Pointer(list, 5, Keys.Shift);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "p5", "p6", "p7", "p8" }) && list.AnchorIndex == 8, "successive Shift extends from original anchor");
            Pointer(list, 8, Keys.Control);
            Pointer(list, 9, Keys.Control | Keys.Shift);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "p5", "p6", "p7", "p8", "p9" }), "Ctrl toggled off anchor remains range origin");
            Check(activations == 1, "range gestures never activate Excel");

            Action<int, int, int> drag = (from, to, flags) =>
            {
                Func<int, IntPtr> point = index => new IntPtr(((list.GetItemRectangle(index).Top + list.ItemHeight / 2) << 16) | 8);
                SendMessage(list.Handle, 0x0201, new IntPtr(1 | flags), point(from));
                // Deliberately release modifiers in subsequent messages. The down
                // snapshot must decide reorder eligibility for this entire gesture.
                SendMessage(list.Handle, 0x0200, new IntPtr(1), point(to));
                SendMessage(list.Handle, 0x0200, new IntPtr(1), point(Math.Min(9, to + 1)));
                SendMessage(list.Handle, 0x0202, IntPtr.Zero, point(to));
            };
            string[] before = view.SelectedSheetIds;
            int activeBefore = activations;
            drag(6, 2, 0);
            Check(view.SelectedSheetIds.SequenceEqual(before) && starts == 0, "multi selected-row drag keeps exact set and never reorders");
            Check(activations == activeBefore, "blocked selected-row drag does not activate");
            drag(1, 4, 0);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "p1" }) && starts == 0, "multi unselected-row drag stays single without native range selection");
            Check(activations == activeBefore && !list.IsSelecting && !list.Capture, "blocked drag release clears tracking without activation");
            drag(1, 4, 0);
            Check(starts == 1, "next drag starting single retains reorder");
            Pointer(list, 2);
            drag(3, 6, 8);
            Check(starts == 1 && view.SelectedSheetIds.SequenceEqual(new[] { "p2", "p3" }), "Ctrl drag only applies down toggle");
            Pointer(list, 2);
            drag(4, 7, 4);
            Check(starts == 1 && view.SelectedSheetIds.SequenceEqual(new[] { "p2", "p3", "p4" }), "Shift drag freezes down range");
            drag(6, 8, 12);
            Check(starts == 1 && view.SelectedSheetIds.SequenceEqual(new[] { "p2", "p3", "p4", "p5", "p6" }), "Ctrl Shift drag freezes down union");
            Pointer(list, 3);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "p3" }), "plain release on selected row still collapses to one");
            Pointer(list, 5, Keys.Control);
            var down = new IntPtr(((list.GetItemRectangle(3).Top + list.ItemHeight / 2) << 16) | 8);
            SendMessage(list.Handle, 0x0201, new IntPtr(1), down);
            SendMessage(list.Handle, 0x001F, IntPtr.Zero, IntPtr.Zero);
            SendMessage(list.Handle, 0x0202, IntPtr.Zero, down);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "p3", "p5" }) && !list.IsSelecting && !list.Capture, "cancel mode does not collapse selection");
            Pointer(list, 1);
            int renames = 0;
            view.CommandRequested += (s, e) => { if (e.Command == WorksheetCommand.Rename) renames++; };
            var doublePoint = new IntPtr(((list.GetItemRectangle(1).Top + list.ItemHeight / 2) << 16) | 8);
            SendMessage(list.Handle, 0x0203, new IntPtr(1), doublePoint);
            SendMessage(list.Handle, 0x0202, IntPtr.Zero, doublePoint);
            Check(renames == 1, "native double-click messages still request rename");
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
                // Verify deletion/recreation before creating the second window:
                // Excel 2013 can reject Chart.Delete after multi-window activity.
                sheet.Activate();
                chart.Delete();
                view.ShowWorksheets(book.Name, service.GetWorksheets(book), id);
                Check(view.SelectedSheetIds.SequenceEqual(new[] { id }), "Excel deletion prunes selected chart");
                var replacement = (Excel.Chart)book.Charts.Add();
                replacement.Name = "ChartRenamed";
                Check(service.GetWorksheetId(replacement) != chartId, "recreated same name does not inherit deleted selection");
                chart = replacement;
                chartId = service.GetWorksheetId(chart);
                view.ShowWorksheets(book.Name, service.GetWorksheets(book), service.GetActiveWorksheetId(book.Windows[1]));
                InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Control | Keys.A));
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
