using System;
using System.Linq;
using System.Windows.Forms;
using Sheaside.Controls;
using Sheaside.Models;
using Sheaside.Services;
using Excel = Microsoft.Office.Interop.Excel;

internal static partial class SmokeTests
{
    private static void TestChartList()
    {
        using (var view = new SheetListControl())
        {
            var chart = new WorksheetItem("chart:c", "Sales", WorksheetVisibility.Visible, 255, SheetKind.Chart);
            var rows = new[] { new WorksheetItem("w", "Data", true), chart };
            view.ShowWorksheets("Book", rows, chart.Id);
            var list = view.Controls.OfType<ListBox>().Single();
            Check(list.SelectedIndex == 1 && chart.DisplayName.Contains("グラフ"), "chart label and active selection");
            view.ShowWorksheets("Book", rows, "w");
            Check(list.SelectedIndex == 0, "chart to worksheet active synchronization");
        }
    }

    private static void TestExcelCharts()
    {
        var excel = new Excel.Application { Visible = false, DisplayAlerts = false };
        Excel.Workbook book = null;
        Excel.Workbook other = null;
        try
        {
            book = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            var data = (Excel.Worksheet)book.Worksheets[1];
            data.Name = "Data";
            var chart = (Excel.Chart)book.Charts.Add(After: data);
            chart.Name = "Sales";
            chart.HasTitle = true;
            chart.ChartTitle.Text = "Chart copy test";
            var tail = (Excel.Worksheet)book.Worksheets.Add(After: chart);
            tail.Name = "Tail";
            // Normalize fixture order explicitly: older Excel can ignore Charts.Add's After.
            chart.Move(After: data);
            tail.Move(After: chart);
            // An embedded chart must not become a tab row.
            var embedded = ((Excel.ChartObjects)data.ChartObjects()).Add(10, 10, 100, 100);
            var service = new ExcelWorksheetService(excel);
            int handle = book.Windows[1].Hwnd;
            Func<string> order = () => string.Join(",", service.GetWorksheets(book).Select(i => i.Name));
            Func<object, ExcelWorksheetService.MoveSession> capture = sheet => service.CaptureMove(handle,
                service.GetWorksheetId(sheet), service.GetWorksheets(book).Select(i => i.Id).ToArray());
            Check(order() == "Data,Sales,Tail", "mixed tab order excludes embedded chart: " + order());
            var chartRow = service.GetWorksheets(book)[1];
            Check(chartRow.Kind == SheetKind.Chart && chartRow.Id.StartsWith("chart:"), "chart metadata and typed identity");
            service.ActivateWorksheet(handle, chartRow.Id);
            Check(Equals(book.Windows[1].ActiveSheet, chart) && service.GetActiveWorksheetId(book.Windows[1]) == chartRow.Id, "chart activation and active identity");
            var target = service.CaptureTarget(handle, chartRow.Id);
            service.RenameWorksheet(target, "Revenue");
            Check(chart.Name == "Revenue", "chart rename");
            Reject(() => service.RenameWorksheet(target, data.Name), "chart name cannot duplicate worksheet");
            var copyId = service.DuplicateWorksheet(target);
            var copy = (Excel.Chart)chart.Next;
            Check(service.GetWorksheetId(copy) == copyId && copy.ChartTitle.Text == "Chart copy test", "chart copy immediately after source retains content");
            var copyTarget = service.CaptureTarget(handle, copyId);
            service.DeleteWorksheet(copyTarget, copy.Name);
            Check(book.Charts.Count == 1 && !excel.DisplayAlerts, "chart delete and alert restoration");
            var addedId = service.AddWorksheet(target);
            Check(chart.Next is Excel.Worksheet && service.GetWorksheetId(chart.Next) == addedId, "add worksheet after chart");
            var added = service.CaptureTarget(handle, addedId);
            service.DeleteWorksheet(added, added.Sheet.Name);
            Check(service.MoveWorksheet(capture(chart), 0) && order() == "Revenue,Data,Tail", "chart move first");
            Check(service.MoveWorksheet(capture(chart), 3) && order() == "Data,Tail,Revenue", "chart move last");
            Check(service.MoveWorksheet(capture(tail), 3) && order() == "Data,Revenue,Tail", "worksheet moves after chart anchor");
            book.Saved = true;
            Check(!service.MoveWorksheet(capture(chart), 1) && !service.MoveWorksheet(capture(chart), 2) && book.Saved, "chart self-drop no mutation");
            service.SetTabColor(target, 0);
            Check(service.GetWorksheets(book)[1].TabColor == 0, "chart black tab color");
            service.SetTabColor(target, null);
            Check(service.GetTabColor(chart) == null, "chart no color");
            service.ChangeVisibility(target, WorksheetVisibility.Visible);
            service.ActivateWorksheet(handle, service.GetWorksheetId(chart));
            Check(chart.Visible == Excel.XlSheetVisibility.xlSheetHidden, "hidden chart click does not reveal");
            Check(service.MoveWorksheet(capture(chart), 0) && chart.Visible == Excel.XlSheetVisibility.xlSheetHidden && order() == "Revenue,Data,Tail", "hidden chart reorder");
            service.ChangeVisibility(target, WorksheetVisibility.Hidden);
            chart.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            Reject(() => service.ChangeVisibility(target, WorksheetVisibility.Hidden), "VeryHidden chart cannot bypass confirmation");
            Check(service.MoveWorksheet(capture(chart), 3) && chart.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden && order() == "Data,Tail,Revenue", "VeryHidden chart reorder");
            service.ChangeVisibility(target, WorksheetVisibility.VeryHidden, chart.Name);
            Check(chart.Visible == Excel.XlSheetVisibility.xlSheetVisible, "VeryHidden chart confirmed reveal");
            var stale = capture(chart);
            book.Protect(Structure: true);
            Reject(() => service.MoveWorksheet(stale, 0), "protected chart move");
            Reject(() => service.SetTabColor(target, 255), "protected chart color");
            Reject(() => service.RenameWorksheet(target, "No"), "protected chart rename");
            Reject(() => service.DuplicateWorksheet(target), "protected chart copy");
            Reject(() => service.DeleteWorksheet(target, chart.Name), "protected chart delete");
            Reject(() => service.ChangeVisibility(target, WorksheetVisibility.Visible), "protected chart visibility");
            book.Unprotect();
            chart.Activate();
            data.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            tail.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Reject(() => service.ChangeVisibility(target, WorksheetVisibility.Visible), "last visible chart cannot hide");
            Reject(() => service.DeleteWorksheet(target, chart.Name), "last visible chart cannot delete");
            data.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            tail.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            var window = book.NewWindow();
            handle = window.Hwnd;
            service.ActivateWorksheet(handle, service.GetWorksheetId(chart));
            Check(Equals(window.ActiveSheet, chart), "chart activation in second window");
            stale = capture(chart);
            other = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Reject(() => service.MoveWorksheet(stale, 0), "chart drag cancels on other workbook");
            var otherChart = (Excel.Chart)other.Charts.Add();
            otherChart.Name = chart.Name;
            service.SetTabColor(target, 255);
            Check(service.GetTabColor(chart) == 255 && service.GetTabColor(otherChart) == null, "chart command stays in captured workbook");
            window.Activate();
            chart.Activate();
            // Excel 2013's hidden automation can reject deleting a chart active
            // in another window. Stabilize both windows; active-chart deletion is
            // already covered by the copy/delete check above in a single window.
            for (int i = 1; i <= book.Windows.Count; i++)
            {
                book.Windows[i].Activate();
                data.Activate();
            }
            service.DeleteWorksheet(target, chart.Name);
            Check(book.Charts.Count == 0 && service.GetWorksheets(book).All(i => i.Kind == SheetKind.Worksheet), "chart disappears after deletion");
            Reject(() => service.SetTabColor(target, 0), "stale deleted chart rejected");
            Console.WriteLine("Chart integration completed.");
        }
        finally
        {
            if (other != null) other.Close(false);
            if (book != null) book.Close(false);
            excel.Quit();
        }
    }
}
