using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Sheaside.Controls;
using Sheaside.Models;
using Sheaside.Services;
using Excel = Microsoft.Office.Interop.Excel;

internal static partial class SmokeTests
{
    private static int assertions;

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            if (!args.Contains("--batch-only"))
            {
                TestNames();
                TestList();
                TestRenameDialog();
                TestReorderRules();
                TestDragControl();
                TestAppearance();
                TestChartList();
                TestMultiSelection();
                TestPointerRegression();
            }
            TestBatch();
            TestReleaseSettings();
            TestFilters();
            TestThemes();
            if (args.Contains("--excel"))
            {
                if (!args.Contains("--batch-only")) { TestExcel(); TestExcelReorder(); TestExcelAppearance(); TestExcelCharts(); TestExcelSelection(); }
                TestExcelBatch();
            }
            Console.WriteLine("PASS: " + assertions + " assertions" + (args.Contains("--excel") ? " (including Excel integration)" : " (managed only)"));
            return 0;
        }
        catch (Exception exception) { Console.Error.WriteLine(exception); return 1; }
    }

    private static void Check(bool condition, string message)
    {
        assertions++;
        if (!condition) throw new Exception("FAIL: " + message);
    }

    private static void Reject(Action action, string message)
    {
        try { action(); }
        catch (InvalidOperationException) { assertions++; return; }
        throw new Exception("FAIL: " + message);
    }

    private static void TestNames()
    {
        string[] existing = { "Other", "グラフ" };
        foreach (string name in new[] { "", " ", new string('a', 32), "a:b", "a\\b", "a/b", "a?b", "a*b", "a[b", "a]b", "a\nb", "'start", "end'", "History", "HISTORY", "other", "グラフ" })
            Check(WorksheetNameRules.Validate(name, existing) != null, "reject name: " + name);
        foreach (string name in new[] { new string('a', 31), "日本語のシート", "O'Brien", "Renamed" })
            Check(WorksheetNameRules.Validate(name, existing) == null, "accept name: " + name);
    }

    private static void TestList()
    {
        using (var view = new SheetListControl())
        {
            var list = view.Controls.OfType<ListBox>().Single();
            var first = new WorksheetItem("Sheet1", "First", true);
            var second = new WorksheetItem("Sheet2", "Second", true);
            var rows = new List<WorksheetItem> { first, second };
            view.ShowWorksheets("Book", rows, "Sheet1");
            Check(list.SelectedIndex == 0, "initial active selection");
            list.ClearSelected();
            list.SelectedIndex = 1;
            view.ShowWorksheets("Book", rows, "Sheet1");
            Check(list.SelectedIndex == 1, "preserve local navigation on timer tick");
            rows[1] = new WorksheetItem("Sheet2", "Renamed", true);
            view.ShowWorksheets("Book", rows, "Sheet1", "Sheet2");
            Check(list.SelectedIndex == 1 && ((WorksheetItem)list.Items[1]).Name == "Renamed", "rename keeps selected identity");
            Check(ReferenceEquals(list.Items[0], first), "unchanged rows retained");
            view.ShowWorksheets("Book", rows, "Sheet2");
            view.ShowWorksheets("Book", rows, "Sheet1");
            Check(list.SelectedIndex == 0, "external active sheet change is synchronized");
            rows.RemoveAt(0);
            view.ShowWorksheets("Book", rows, "Sheet2");
            Check(list.SelectedIndex == 0 && list.Items.Count == 1, "deleted selection falls back to active sheet");
            int activationRequests = 0;
            view.WorksheetRequested += (sender, e) => activationRequests++;
            view.ShowWorksheets("Book", rows, "Sheet2");
            Check(activationRequests == 0, "refresh does not activate Excel sheets");
            WorksheetCommandEventArgs command = null;
            view.CommandRequested += (sender, e) => command = e;
            view.AddButton.PerformClick();
            Check(command != null && command.Command == WorksheetCommand.Add && command.Worksheet.Id == "Sheet2", "add button uses current row");
            typeof(ListBox).GetMethod("OnMouseDoubleClick", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(list, new object[] { new MouseEventArgs(MouseButtons.Left, 2, 4, 4, 0) });
            Check(command.Command == WorksheetCommand.Rename, "double click requests rename");
            var menu = list.ContextMenuStrip;
            typeof(ContextMenuStrip).GetMethod("OnOpening", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(menu, new object[] { new System.ComponentModel.CancelEventArgs() });
            list.SelectedIndex = -1;
            SheetBatchEventArgs batch = null;
            view.BatchCommandRequested += (s, e) => batch = e;
            menu.Items[1].PerformClick();
            Check(batch != null && batch.Command == SheetBatchCommand.Delete && batch.Ids.Single() == "Sheet2", "menu command retains opening target");
            typeof(ContextMenuStrip).GetMethod("OnOpening", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(menu, new object[] { new System.ComponentModel.CancelEventArgs() });
            Check(menu.Items[0].Enabled && !menu.Items[3].Enabled && !menu.Items[4].Enabled && !menu.Items[1].Enabled, "empty space allows only add");
        }
    }

    private static void TestRenameDialog()
    {
        string applied = null;
        using (var dialog = new RenameSheetDialog("Current", name =>
        {
            string error = WorksheetNameRules.Validate(name, new[] { "Duplicate" });
            if (error == null) applied = name;
            return error;
        }))
        {
            var input = dialog.Controls.OfType<TextBox>().Single();
            Check(input.Text == "Current", "rename starts with current name");
            var submit = (Button)dialog.AcceptButton;
            var onClick = typeof(Button).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic);
            input.Text = "Duplicate";
            onClick.Invoke(submit, new object[] { EventArgs.Empty });
            Check(dialog.DialogResult == DialogResult.None && applied == null && input.Text == "Duplicate", "invalid rename keeps dialog/input");
            input.Text = "NewName";
            onClick.Invoke(submit, new object[] { EventArgs.Empty });
            Check(dialog.DialogResult == DialogResult.OK && applied == "NewName", "valid rename accepted");
            Check(((Button)dialog.CancelButton).DialogResult == DialogResult.Cancel, "cancel button available");
        }
    }

    private static void TestExcel()
    {
        Excel.Application excel = null;
        Excel.Workbook book = null;
        Excel.Workbook other = null;
        Excel.Workbook readOnly = null;
        string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "readonly-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            excel = new Excel.Application { Visible = true, DisplayAlerts = false };
            Console.WriteLine("Excel version: " + excel.Version);
            var service = new ExcelWorksheetService(excel);
            book = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            var source = (Excel.Worksheet)book.Worksheets[1];
            source.Name = "Original";
            source.Cells[1, 1].Value2 = "copied value";
            int handle = book.Windows[1].Hwnd;
            var target = service.CaptureTarget(handle, service.GetWorksheetId(source));
            Reject(() => service.EnsureCanDelete(target), "last visible worksheet");
            string addedCode = service.AddWorksheet(target);
            Check(book.Worksheets.Count == 2 && service.GetWorksheetId((Excel.Worksheet)source.Next) == addedCode, "add after selected sheet");
            service.RenameWorksheet(target, "Renamed");
            Check(source.Name == "Renamed", "rename applied");
            service.RenameWorksheet(target, "Renamed");
            Check(source.Name == "Renamed", "unchanged name accepted");
            Reject(() => service.RenameWorksheet(target, ((Excel.Worksheet)source.Next).Name), "duplicate worksheet name");
            var chart = (Excel.Chart)book.Charts.Add(After: book.Sheets[book.Sheets.Count]);
            chart.Name = "ChartName";
            Reject(() => service.RenameWorksheet(target, "chartname"), "chart name conflict");
            source.Activate();
            string copyCode = service.DuplicateWorksheet(target);
            var copy = (Excel.Worksheet)source.Next;
            Check(service.GetWorksheetId(copy) == copyCode && (string)copy.Cells[1, 1].Value2 == "copied value", "copy position and contents");
            other = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            int otherCount = other.Sheets.Count;
            int beforeAdd = book.Sheets.Count;
            service.AddWorksheet(target);
            Check(book.Sheets.Count == beforeAdd + 1 && other.Sheets.Count == otherCount, "captured target survives workbook switch");
            book.Protect(Password: "test", Structure: true);
            Reject(() => service.AddWorksheet(target), "protected add");
            Reject(() => service.RenameWorksheet(target, "Blocked"), "protected rename");
            Reject(() => service.DuplicateWorksheet(target), "protected copy");
            Reject(() => service.DeleteWorksheet(target, source.Name), "protected delete");
            book.Unprotect("test");
            var copyTarget = service.CaptureTarget(handle, copyCode);
            excel.DisplayAlerts = true;
            service.DeleteWorksheet(copyTarget, copy.Name);
            Check(excel.DisplayAlerts, "DisplayAlerts true restored");
            Reject(() => service.RenameWorksheet(copyTarget, "Deleted"), "stale worksheet target rejected");
            excel.DisplayAlerts = false;
            var addedTarget = service.CaptureTarget(handle, addedCode);
            service.DeleteWorksheet(addedTarget, addedTarget.Sheet.Name);
            Check(!excel.DisplayAlerts, "DisplayAlerts false restored");
            int beforeRejectedDelete = book.Sheets.Count;
            Reject(() => service.DeleteWorksheet(target, "Outdated confirmation"), "changed confirmation name rejected");
            Check(book.Sheets.Count == beforeRejectedDelete && !excel.DisplayAlerts, "rejected delete changes neither sheets nor alerts");
            source.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            string hiddenCopyId = service.DuplicateWorksheet(target);
            var hiddenCopy = (Excel.Worksheet)source.Next;
            Check(service.GetWorksheetId(hiddenCopy) == hiddenCopyId && hiddenCopy.Visible == Excel.XlSheetVisibility.xlSheetHidden, "hidden sheet copied after source without unhiding");
            int beforeHiddenDelete = book.Sheets.Count;
            try
            {
                service.DeleteWorksheet(service.CaptureTarget(handle, hiddenCopyId), hiddenCopy.Name);
                Check(book.Sheets.Count == beforeHiddenDelete - 1, "hidden delete succeeded");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Some Excel versions reject this operation. Exercise the native
                // failure path as well as the normal visible-sheet deletion above.
                Check(book.Sheets.Count == beforeHiddenDelete, "native deletion rejection preserves sheet");
                Console.WriteLine("Native Excel rejected hidden-copy deletion; sheet preserved.");
            }
            Check(!excel.DisplayAlerts, "alerts restored after hidden delete, including native failure");
            source.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            var loneSheet = (Excel.Worksheet)other.Worksheets[1];
            var hidden = (Excel.Worksheet)other.Worksheets.Add(After: loneSheet);
            hidden.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            var loneTarget = service.CaptureTarget(other.Windows[1].Hwnd, service.GetWorksheetId(loneSheet));
            Reject(() => service.EnsureCanDelete(loneTarget), "hidden sheets do not satisfy visible-sheet guard");
            other.Charts.Add(After: hidden);
            service.DeleteWorksheet(loneTarget, loneSheet.Name);
            Check(other.Worksheets.Count == 1 && other.Charts.Count == 1, "visible chart permits deleting last visible worksheet");
            var secondWindow = book.NewWindow();
            var secondTarget = service.CaptureTarget(secondWindow.Hwnd, service.GetWorksheetId(source));
            service.RenameWorksheet(secondTarget, "TwoWindows");
            Check(source.Name == "TwoWindows", "second window targets same workbook");
            secondWindow.Close();
            Reject(() => service.AddWorksheet(secondTarget), "closed window rejected");
            book.SaveAs(file, Excel.XlFileFormat.xlOpenXMLWorkbook);
            book.Close(false);
            book = null;
            Reject(() => service.AddWorksheet(target), "closed workbook rejected");
            readOnly = excel.Workbooks.Open(file, ReadOnly: true);
            Reject(() => service.CaptureTarget(readOnly.Windows[1].Hwnd, null), "read-only workbook rejected");
            Console.WriteLine("Excel integration completed.");
        }
        finally
        {
            if (excel != null)
            {
                excel.DisplayAlerts = false;
                if (readOnly != null) readOnly.Close(false);
                if (other != null) other.Close(false);
                if (book != null) book.Close(false);
                excel.Quit();
            }
            if (File.Exists(file)) File.Delete(file);
        }
    }
}
