using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Sheaside.Controls;
using Sheaside.Models;
using Sheaside.Services;
using Excel = Microsoft.Office.Interop.Excel;

internal static partial class SmokeTests
{
    private static void TestReorderRules()
    {
        for (int count = 1; count <= 7; count++)
        for (int source = 0; source < count; source++)
        for (int gap = 0; gap <= count; gap++)
        {
            var original = Enumerable.Range(0, count).ToArray();
            var expected = original.Take(gap).Where(value => value != source).Concat(new[] { source })
                .Concat(original.Skip(gap).Where(value => value != source)).ToArray();
            int destination = WorksheetMoveRules.DestinationIndex(source, gap, count);
            var actual = original.Where(value => value != source).ToList();
            actual.Insert(destination, source);
            Check(actual.SequenceEqual(expected) && (destination == source) == (gap == source || gap == source + 1),
                "move boundary permutation: " + count + "/" + source + "/" + gap);
        }
        Check(WorksheetMoveRules.BoundaryAt(1, 29, 20, 20, 3) == 1, "upper half inserts before");
        Check(WorksheetMoveRules.BoundaryAt(1, 30, 20, 20, 3) == 2, "midpoint inserts after");
        Check(WorksheetMoveRules.BoundaryAt(-1, 100, 0, 0, 3) == 3, "blank area inserts at end");
        Check(WorksheetMoveRules.BoundaryAt(-1, 10, 0, 0, 0) == 0, "empty list boundary");
        foreach (int gap in new[] { -1, 4 })
        {
            try { WorksheetMoveRules.DestinationIndex(0, gap, 3); throw new Exception("Invalid gap accepted"); }
            catch (ArgumentOutOfRangeException) { assertions++; }
        }
    }

    private static void InvokeControl(object control, string method, object args)
    {
        control.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic)
            .Invoke(control, new[] { args });
    }

    private static void TestDragControl()
    {
        using (var view = new SheetListControl { Size = new Size(260, 400) })
        {
            var list = view.Controls.OfType<ReorderableSheetListBox>().Single();
            var a = new WorksheetItem("A", "A", true);
            var b = new WorksheetItem("B", "B", true);
            view.ShowWorksheets("Book", new[] { a, b }, "A");
            list.CreateControl();
            int starts = 0;
            int finishes = 0;
            view.DragStarting += (sender, args) =>
            {
                starts++;
                Check(view.IsDragging && args.WorksheetId == "A" && args.Order.SequenceEqual(new[] { "A", "B" }), "drag snapshot and owner state");
                view.ShowWorksheets("Book", new[] { b }, "B");
                Check(list.Items.Count == 2, "refresh is deferred inside drag");
                // Keep Cancel=true to test threshold/lifecycle without a blocking OS drag loop.
            };
            view.DragCompleted += (sender, args) => { finishes++; Check(!view.IsDragging, "completed event leaves drag mode"); };
            InvokeControl(list, "OnMouseDown", new MouseEventArgs(MouseButtons.Left, 1, 4, 4, 0));
            InvokeControl(list, "OnMouseMove", new MouseEventArgs(MouseButtons.Left, 0, 4, 4, 0));
            Check(starts == 0, "small mouse movement is not a drag");
            InvokeControl(list, "OnMouseMove", new MouseEventArgs(MouseButtons.Left, 0, 100, 40, 0));
            Check(starts == 1 && finishes == 1, "threshold exceeded starts then safely cancels");
            int clicks = 0;
            list.MouseClick += (sender, args) => clicks++;
            InvokeControl(list, "OnMouseClick", new MouseEventArgs(MouseButtons.Left, 1, 4, 4, 0));
            Check(clicks == 0, "drag does not also activate a sheet");
            InvokeControl(list, "OnMouseDown", new MouseEventArgs(MouseButtons.Left, 1, 4, 4, 0));
            InvokeControl(list, "OnMouseUp", new MouseEventArgs(MouseButtons.Left, 1, 4, 4, 0));
            InvokeControl(list, "OnMouseClick", new MouseEventArgs(MouseButtons.Left, 1, 4, 4, 0));
            Check(clicks == 1, "normal click still works after drag");

            // Exercise OLE callbacks with a local token, without moving the user's mouse.
            var type = typeof(ReorderableSheetListBox);
            var token = Guid.NewGuid().ToString("N");
            type.GetProperty("IsDragging").GetSetMethod(true).Invoke(list, new object[] { true });
            type.GetField("cancelled", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(list, false);
            type.GetField("dragToken", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(list, token);
            var data = new DataObject();
            data.SetData("Sheaside.LocalWorksheetMove", false, new string(token.ToCharArray()));
            Rectangle row = list.GetItemRectangle(1);
            Point screen = list.PointToScreen(new Point(8, row.Bottom - 2));
            var over = new DragEventArgs(data, 1, screen.X, screen.Y, DragDropEffects.Move, DragDropEffects.None);
            InvokeControl(list, "OnDragOver", over);
            Check(over.Effect == DragDropEffects.Move, "local drop accepted at lower half");
            int boundary = -1;
            view.DropRequested += (sender, args) => boundary = args.Boundary;
            InvokeControl(list, "OnDragDrop", over);
            Check(boundary == 2 && ReferenceEquals(list.Items[0], a), "drop requests move without optimistic reorder");
            Point blank = list.PointToScreen(new Point(8, row.Bottom + 20));
            var blankOver = new DragEventArgs(data, 1, blank.X, blank.Y, DragDropEffects.Move, DragDropEffects.None);
            InvokeControl(list, "OnDragOver", blankOver);
            Check(blankOver.Effect == DragDropEffects.Move && (int)type.GetField("insertionBoundary", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(list) == 2,
                "blank space shows end insertion marker");
            InvokeControl(list, "OnDragDrop", blankOver);
            Check(boundary == 2, "blank space drop requests end boundary");
            var foreign = new DataObject();
            foreign.SetData("Sheaside.LocalWorksheetMove", false, Guid.NewGuid().ToString("N"));
            var foreignOver = new DragEventArgs(foreign, 1, screen.X, screen.Y, DragDropEffects.Move, DragDropEffects.Move);
            InvokeControl(list, "OnDragEnter", foreignOver);
            Check(foreignOver.Effect == DragDropEffects.None, "another pane's token is rejected");
            Point self = list.PointToScreen(new Point(4, 4));
            var selfOver = new DragEventArgs(data, 1, self.X, self.Y, DragDropEffects.Move, DragDropEffects.Move);
            InvokeControl(list, "OnDragOver", selfOver);
            Check(selfOver.Effect == DragDropEffects.None, "no-op has no drop effect");
            view.CancelDrag();
            var query = new QueryContinueDragEventArgs(1, false, DragAction.Continue);
            InvokeControl(list, "OnQueryContinueDrag", query);
            Check(query.Action == DragAction.Cancel, "external cancellation stops native drag");
            InvokeControl(list, "OnDragOver", over);
            Check(over.Effect == DragDropEffects.None, "cancelled token cannot drop");
            type.GetProperty("IsDragging").GetSetMethod(true).Invoke(list, new object[] { false });
            view.ShowWorksheets("Book", new[] { b, a }, "B", "A");
            Check(list.SelectedIndex == 1 && ((WorksheetItem)list.SelectedItem).Id == "A", "moved sheet stays selected");
            view.ShowWorksheets("Book", new[] { a, b }, "B", "A");
            Check(list.SelectedIndex == 0 && ((WorksheetItem)list.Items[0]).Id == "A", "real order can restore after failed move");
        }
    }

    private static void TestExcelReorder()
    {
        Excel.Application excel = null;
        Excel.Workbook book = null;
        Excel.Workbook other = null;
        try
        {
            excel = new Excel.Application { Visible = true, DisplayAlerts = false };
            book = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            var service = new ExcelWorksheetService(excel);
            var a = (Excel.Worksheet)book.Worksheets[1];
            a.Name = "A";
            foreach (string name in new[] { "B", "C", "D" })
                ((Excel.Worksheet)book.Worksheets.Add(After: book.Sheets[book.Sheets.Count])).Name = name;
            var b = (Excel.Worksheet)book.Worksheets["B"];
            var c = (Excel.Worksheet)book.Worksheets["C"];
            var d = (Excel.Worksheet)book.Worksheets["D"];
            var middleChart = (Excel.Chart)book.Charts.Add(After: b);
            middleChart.Name = "MiddleChart";
            var tailChart = (Excel.Chart)book.Charts.Add(After: d);
            tailChart.Name = "TailChart";
            int handle = book.Windows[1].Hwnd;
            Func<Excel.Worksheet, ExcelWorksheetService.MoveSession> capture = sheet => service.CaptureMove(handle,
                service.GetWorksheetId(sheet), service.GetWorksheets(book).Select(item => item.Id).ToArray());
            Func<string> order = () => string.Join(",", service.GetWorksheets(book).Where(item => item.Kind == SheetKind.Worksheet).Select(item => item.Name));

            Check(service.MoveWorksheet(capture(d), 0) && order() == "D,A,B,C", "move last to first with charts");
            Check(service.MoveWorksheet(capture(d), book.Sheets.Count) && order() == "A,B,C,D", "move first to last with charts");
            Check(service.MoveWorksheet(capture(c), 1) && order() == "A,C,B,D", "move backward before anchor");
            Check(service.MoveWorksheet(capture(c), 3) && order() == "A,B,C,D", "move forward before anchor");
            book.Saved = true;
            Check(!service.MoveWorksheet(capture(c), 2) && !service.MoveWorksheet(capture(c), 3) && book.Saved,
                "both self boundaries avoid Excel mutation");
            Check(excel.Workbooks.Count == 1 && book.Charts.Count == 2 && middleChart.Index < tailChart.Index,
                "no new workbook; chart order preserved");
            b.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            excel.ScreenUpdating = true;
            string activeBeforeHiddenMove = service.GetActiveWorksheetId(book.Windows[1]);
            Check(service.MoveWorksheet(capture(b), 0) && order() == "B,A,C,D" && b.Visible == Excel.XlSheetVisibility.xlSheetHidden,
                "hidden worksheet moves without unhiding");
            Check(excel.ScreenUpdating && service.GetActiveWorksheetId(book.Windows[1]) == activeBeforeHiddenMove, "hidden move restores painting and active sheet");
            b.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            excel.ScreenUpdating = false;
            Check(service.MoveWorksheet(capture(b), book.Sheets.Count) && order() == "A,C,D,B" && b.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden,
                "VeryHidden worksheet moves to end");
            Check(!excel.ScreenUpdating, "already disabled painting remains disabled");
            excel.ScreenUpdating = true;
            bool movedAfterHidden = service.MoveWorksheet(capture(a), book.Sheets.Count);
            Check(movedAfterHidden && order() == "C,D,B,A" && b.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden,
                "move after VeryHidden destination");
            Check(service.MoveWorksheet(capture(a), 0) && order() == "A,C,D,B", "restore order after hidden destination test");
            var stale = capture(a);
            book.Protect(Password: "test", Structure: true);
            Reject(() => service.MoveWorksheet(stale, 4), "structure protected after drag started");
            Check(order() == "A,C,D,B", "protection rejection preserves order");
            book.Unprotect("test");
            stale = capture(a);
            c.Move(Before: a);
            Reject(() => service.MoveWorksheet(stale, 4), "external reorder cancels stale drag");
            Check(order() == "C,A,D,B", "failed stale move leaves actual Excel order intact");
            stale = capture(a);
            var added = (Excel.Worksheet)book.Worksheets.Add(After: d);
            Reject(() => service.MoveWorksheet(stale, 0), "sheet addition cancels stale drag");
            added.Delete();
            Reject(() => service.CaptureMove(handle, service.GetWorksheetId(a), new[] { "wrong snapshot" }), "outdated displayed order refused");
            stale = capture(a);
            var secondWindow = book.NewWindow();
            Reject(() => service.MoveWorksheet(stale, 0), "same workbook window switch cancels move");
            handle = secondWindow.Hwnd;
            stale = capture(a);
            other = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            string before = order();
            Reject(() => service.MoveWorksheet(stale, 0), "other workbook switch cancels move");
            Check(order() == before && other.Worksheets.Count == 1 && excel.Workbooks.Count == 2, "cross-workbook rejection changes neither book");
            var otherSource = (Excel.Worksheet)other.Worksheets[1];
            otherSource.Name = "A";
            ((Excel.Worksheet)other.Worksheets.Add(After: otherSource)).Name = "B";
            var otherSession = service.CaptureMove(other.Windows[1].Hwnd, service.GetWorksheetId(otherSource),
                service.GetWorksheets(other).Select(item => item.Id).ToArray());
            Check(service.MoveWorksheet(otherSession, 2) && string.Join(",", service.GetWorksheets(other).Select(item => item.Name)) == "B,A",
                "reorder in second workbook with overlapping sheet names");
            Check(order() == before && excel.Workbooks.Count == 2, "second-book move leaves first workbook unchanged");
            Check(!excel.DisplayAlerts, "move leaves DisplayAlerts untouched");
            Console.WriteLine("Excel reorder integration completed.");
        }
        finally
        {
            if (excel != null)
            {
                excel.DisplayAlerts = false;
                if (other != null) other.Close(false);
                if (book != null) book.Close(false);
                excel.Quit();
            }
        }
    }
}
