using System;
using System.Linq;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Sheaside.Controls;
using Sheaside.Models;
using Sheaside.Services;
using Excel = Microsoft.Office.Interop.Excel;

internal static partial class SmokeTests
{
    private static void TestBatch()
    {
        var items = Enumerable.Range(0, 4).Select(i => new WorksheetItem("b" + i, "Batch " + i, true)).ToArray();
        int calls = 0;
        var result = SheetBatchRunner.Run(items, () => { }, (index, item) =>
        {
            calls++;
            if (index == 2) throw new COMException("Injected Excel refusal");
            item.Status = index == 0 ? SheetBatchStatus.Succeeded : SheetBatchStatus.Skipped;
        });
        Check(calls == 3 && result.Count(SheetBatchStatus.Succeeded) == 1 && result.Count(SheetBatchStatus.Skipped) == 1 &&
            result.Count(SheetBatchStatus.Failed) == 1 && result.Count(SheetBatchStatus.NotProcessed) == 1, "batch stops on first refusal with complete accounting");
        Check(result.Items[2].Item.Id == "b2" && result.Items[3].Item.Id == "b3" && result.Details.Contains("未処理") && result.Details.Contains("Injected Excel refusal"), "batch report identifies failed and unprocessed names");
        calls = 0;
        result = SheetBatchRunner.Run(items, () => { throw new InvalidOperationException("Protected"); }, (i, item) => calls++);
        Check(calls == 0 && result.Count(SheetBatchStatus.NotProcessed) == 4 && result.Error == "Protected", "preflight fails before any mutation");
        using (var dialog = new BatchDetailsDialog("Delete", "4件", string.Join("\r\n", items.Select(i => i.Name)), true))
        {
            Check(ReferenceEquals(dialog.AcceptButton, dialog.CancelButton), "confirmation defaults to cancel");
            var text = dialog.Controls.OfType<TextBox>().Single();
            Check(text.ReadOnly && text.Multiline && text.Text.Contains("Batch 3"), "confirmation shows every name in scrollable read-only details");
        }
        using (var view = new SheetListControl())
        {
            view.ShowWorksheets("Book", items, "b0");
            var list = view.Controls.OfType<ReorderableSheetListBox>().Single();
            InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Control | Keys.A));
            var menu = list.ContextMenuStrip;
            typeof(ContextMenuStrip).GetMethod("OnOpening", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(menu, new object[] { new CancelEventArgs() });
            var batch = (ToolStripMenuItem)menu.Items[3];
            Check(batch.Enabled && menu.Items.Count == 8, "unified menu has no batch submenu");
            SheetBatchEventArgs request = null;
            view.BatchCommandRequested += (s, e) => request = e;
            list.ClearSelected();
            batch.PerformClick();
            Check(request != null && request.Command == SheetBatchCommand.Delete && request.Ids.SequenceEqual(items.Select(i => i.Id)), "batch menu captures immutable target IDs");
            InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Escape));
            typeof(ContextMenuStrip).GetMethod("OnOpening", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(menu, new object[] { new CancelEventArgs() });
            Check(batch.Enabled && menu.Items[1].Enabled, "single selection uses same delete menu");
        }
    }

    private static void TestExcelBatch()
    {
        var excel = new Excel.Application { Visible = false, DisplayAlerts = false };
        Excel.Workbook book = null, other = null, readOnly = null;
        string file = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "batch-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            book = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            var keep = (Excel.Worksheet)book.Worksheets[1]; keep.Name = "Keep";
            var a = (Excel.Worksheet)book.Worksheets.Add(After: keep); a.Name = "A";
            var b = (Excel.Worksheet)book.Worksheets.Add(After: a); b.Name = "B";
            var chart = (Excel.Chart)book.Charts.Add(); chart.Name = "Chart"; chart.Move(After: b);
            var service = new ExcelWorksheetService(excel);
            var batches = new ExcelSheetBatchService(service);
            int handle = book.Windows[1].Hwnd;
            string[] ids = { service.GetWorksheetId(a), service.GetWorksheetId(b), service.GetWorksheetId(chart) };
            Func<SheetBatchCommand, ExcelSheetBatchService.Plan> capture = command => batches.Capture(handle, ids.Reverse(), command);
            Reject(() => batches.Capture(handle, service.GetWorksheets(book).Select(i => i.Id), SheetBatchCommand.Delete), "batch deletion cannot remove all visible sheets");
            Reject(() => batches.Capture(handle, service.GetWorksheets(book).Select(i => i.Id), SheetBatchCommand.Hide), "batch hiding cannot remove all visible sheets");
            var plan = capture(SheetBatchCommand.Delete);
            var result = batches.Execute(plan, false);
            Check(result.Count(SheetBatchStatus.NotProcessed) == 3 && book.Sheets.Count == 4, "batch deletion requires confirmation");
            b.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            chart.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            plan = capture(SheetBatchCommand.Unhide);
            Check(!plan.NeedsConfirmation && plan.Items.Any(i => i.Visibility == WorksheetVisibility.VeryHidden), "VeryHidden reveal no longer requires confirmation");
            result = batches.Execute(plan, false);
            Check(result.Count(SheetBatchStatus.Succeeded) == 2 && chart.Visible == Excel.XlSheetVisibility.xlSheetVisible, "VeryHidden reveals without confirmation");
            b.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            chart.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            result = batches.Execute(capture(SheetBatchCommand.Hide), true);
            Check(result.Count(SheetBatchStatus.Succeeded) == 2 && result.Count(SheetBatchStatus.Skipped) == 1 && chart.Visible == Excel.XlSheetVisibility.xlSheetHidden, "mixed hide converts VeryHidden and skips Hidden");
            result = batches.Execute(capture(SheetBatchCommand.Unhide), false);
            Check(result.Count(SheetBatchStatus.Succeeded) == 3, "batch reveals Hidden worksheet and chart");
            result = batches.Execute(capture(SheetBatchCommand.Unhide), false);
            Check(result.Count(SheetBatchStatus.Skipped) == 3, "already visible batch skips");
            chart.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            result = batches.Execute(capture(SheetBatchCommand.Unhide), true);
            Check(result.Count(SheetBatchStatus.Succeeded) == 1 && result.Count(SheetBatchStatus.Skipped) == 2, "confirmed VeryHidden batch reveal");
            result = batches.Execute(capture(SheetBatchCommand.VeryHide), false);
            Check(result.Count(SheetBatchStatus.Succeeded) == 3 && chart.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden, "worksheet and chart bulk VeryHidden");
            var guarded = new ExcelSheetBatchService(service, () => false);
            Reject(() => guarded.Capture(handle, ids, SheetBatchCommand.Unhide), "disabled VeryHidden setting rejects hidden targets");
            Reject(() => guarded.Capture(handle, new[] { service.GetWorksheetId(keep) }, SheetBatchCommand.VeryHide), "disabled VeryHidden setting rejects command");
            Reject(() => batches.Capture(handle, new[] { service.GetWorksheetId(keep) }, SheetBatchCommand.VeryHide), "VeryHidden preserves last visible sheet");
            result = batches.Execute(capture(SheetBatchCommand.Unhide), false);
            Check(result.Count(SheetBatchStatus.Succeeded) == 3, "bulk VeryHidden round trip");
            a.Tab.Color = 255;
            result = batches.Execute(capture(SheetBatchCommand.SetColor), false, 255);
            Check(result.Count(SheetBatchStatus.Succeeded) == 2 && result.Count(SheetBatchStatus.Skipped) == 1 && service.GetTabColor(chart) == 255, "batch color and same-color skip including chart");
            result = batches.Execute(capture(SheetBatchCommand.ClearColor), false);
            Check(result.Count(SheetBatchStatus.Succeeded) == 3 && service.GetTabColor(chart) == null, "batch clear color");
            result = batches.Execute(capture(SheetBatchCommand.ClearColor), false);
            Check(result.Count(SheetBatchStatus.Skipped) == 3, "batch clear already unset skips");
            plan = capture(SheetBatchCommand.SetColor);
            b.Name = "Changed";
            result = batches.Execute(plan, false, 123);
            Check(result.Count(SheetBatchStatus.NotProcessed) == 3 && service.GetTabColor(a) == null, "rename during dialog rejects whole batch before mutation");
            b.Name = "B";
            plan = capture(SheetBatchCommand.Delete);
            a.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            result = batches.Execute(plan, true);
            Check(result.Count(SheetBatchStatus.NotProcessed) == 3 && book.Sheets.Count == 4, "visibility changed during delete confirmation rejects all");
            a.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            plan = capture(SheetBatchCommand.SetColor);
            book.Protect(Structure: true);
            result = batches.Execute(plan, false, 123);
            Check(result.Count(SheetBatchStatus.NotProcessed) == 3, "protection applied after capture prevents whole batch");
            Reject(() => capture(SheetBatchCommand.Delete), "protected batch capture rejected");
            book.Unprotect();
            var removed = (Excel.Worksheet)book.Worksheets.Add();
            plan = batches.Capture(handle, new[] { ids[0], service.GetWorksheetId(removed) }, SheetBatchCommand.Delete);
            removed.Delete();
            result = batches.Execute(plan, true);
            Check(result.Count(SheetBatchStatus.NotProcessed) == 2 && book.Sheets.Count == 4, "missing target after confirmation prevents deleting remaining targets");
            b.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            chart.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            plan = capture(SheetBatchCommand.Duplicate);
            Check(plan.Items.Select(i => i.Name).SequenceEqual(new[] { "A", "B", "Chart" }), "batch sources frozen in tab order irrespective of input order");
            excel.ScreenUpdating = true;
            result = batches.Execute(plan, false);
            Check(result.Count(SheetBatchStatus.Succeeded) == 3 && book.Sheets.Count == 7, "batch duplicates visible hidden and chart sources: " + result.Details);
            Check(service.GetWorksheetId(a.Next) == result.Items[0].CreatedId && service.GetWorksheetId(b.Next) == result.Items[1].CreatedId && service.GetWorksheetId(chart.Next) == result.Items[2].CreatedId, "each copy immediately follows its original");
            Check(ExcelSheet.From(b.Next).Visible == Excel.XlSheetVisibility.xlSheetVeryHidden && ExcelSheet.From(chart.Next).Visible == Excel.XlSheetVisibility.xlSheetHidden, "VeryHidden worksheet and hidden chart copy state preserved");
            Check(excel.ScreenUpdating && b.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden && chart.Visible == Excel.XlSheetVisibility.xlSheetHidden, "copy restores source visibility and screen updating");
            string[] copies = result.Items.Select(i => i.CreatedId).ToArray();
            foreach (string id in copies) service.CaptureTarget(handle, id).Sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            keep.Activate();
            excel.DisplayAlerts = true;
            result = batches.Execute(batches.Capture(handle, copies, SheetBatchCommand.Delete), true);
            Check(result.Count(SheetBatchStatus.Succeeded) == 3 && book.Sheets.Count == 4 && excel.DisplayAlerts, "batch deletes only copies and restores alerts=true: " + result.Details);
            excel.DisplayAlerts = false;
            plan = capture(SheetBatchCommand.SetColor);
            other = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            var otherSheet = (Excel.Worksheet)other.Worksheets[1]; otherSheet.Name = a.Name;
            result = batches.Execute(plan, false, 654321);
            Check(result.Count(SheetBatchStatus.Succeeded) == 3 && service.GetTabColor(otherSheet) == null, "window switch cannot redirect batch to other book");
            book.Windows[1].Activate();
            var second = book.NewWindow();
            handle = second.Hwnd;
            result = batches.Execute(capture(SheetBatchCommand.ClearColor), false);
            Check(result.Count(SheetBatchStatus.Succeeded) == 3 && book.Windows[1].SelectedSheets.Count == 1 && second.SelectedSheets.Count == 1, "second-window batch does not group sheets");

            // Deterministically change real Excel protection between runner steps.
            // Hidden automation does not reliably deliver Excel activation events.
            plan = capture(SheetBatchCommand.SetColor);
            int applied = 0;
            try
            {
                result = SheetBatchRunner.Run(plan.Items, () => { }, (index, item) =>
                {
                    var target = service.CaptureTarget(handle, item.Item.Id);
                    service.SetTabColor(target, 123);
                    item.Status = SheetBatchStatus.Succeeded;
                    applied++;
                    if (index == 0) book.Protect(Structure: true);
                });
            }
            finally { book.Unprotect(); }
            Check(applied == 1 && result.Count(SheetBatchStatus.Succeeded) == 1 && result.Count(SheetBatchStatus.Failed) == 1 && result.Count(SheetBatchStatus.NotProcessed) == 1 && service.GetTabColor(b) == null, "Excel protection between steps stops remainder and accounts for partial success");
            Check(!excel.DisplayAlerts, "batch preserves alerts=false");

            plan = capture(SheetBatchCommand.SetColor);
            book.SaveAs(file, Excel.XlFileFormat.xlOpenXMLWorkbook);
            book.Close(false); book = null;
            result = batches.Execute(plan, false, 0);
            Check(result.Count(SheetBatchStatus.NotProcessed) == 3 && service.GetTabColor(otherSheet) == null, "closed captured window cannot redirect batch");
            readOnly = excel.Workbooks.Open(file, ReadOnly: true);
            Reject(() => batches.Capture(readOnly.Windows[1].Hwnd, service.GetWorksheets(readOnly).Select(i => i.Id), SheetBatchCommand.ClearColor), "read-only batch rejected");
            Console.WriteLine("Batch integration completed.");
        }
        finally
        {
            excel.DisplayAlerts = false;
            if (readOnly != null) readOnly.Close(false);
            if (other != null) other.Close(false);
            if (book != null) book.Close(false);
            excel.Quit();
            if (System.IO.File.Exists(file)) System.IO.File.Delete(file);
        }
    }
}
