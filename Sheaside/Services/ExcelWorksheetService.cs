using System;
using System.Collections.Generic;
using System.Linq;
using Sheaside.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sheaside.Services
{
    internal sealed class ExcelWorksheetService
    {
        // Short-lived, captured before showing a modal dialog. Never bound to the UI.
        internal sealed class OperationTarget
        {
            public Excel.Window Window;
            public Excel.Workbook Workbook;
            public Excel.Worksheet Worksheet;
            public int WindowHandle;
        }

        internal sealed class MoveSession
        {
            public OperationTarget Target;
            public object[] SheetOrder;
            public Excel.Worksheet[] Worksheets;
            public string[] WorksheetIds;
            public int SourceIndex;
        }
        private readonly Excel.Application application;

        public ExcelWorksheetService(Excel.Application application) { this.application = application; }

        public IList<Excel.Window> GetWindows()
        {
            var result = new List<Excel.Window>();
            var windows = application.Windows;
            for (int i = 1; i <= windows.Count; i++)
            {
                var window = windows[i];
                if (window.Visible) result.Add(window);
            }
            return result;
        }

        public Excel.Window FindWindow(int handle)
        {
            foreach (var window in GetWindows())
                if (window.Hwnd == handle) return window;
            return null;
        }

        public Excel.Workbook GetWorkbook(Excel.Window window)
        {
            // Window.Parent is Application. The active sheet's parent is its workbook,
            // including when the active sheet is a chart sheet.
            var worksheet = window.ActiveSheet as Excel.Worksheet;
            if (worksheet != null) return worksheet.Parent as Excel.Workbook;
            var chart = window.ActiveSheet as Excel.Chart;
            if (chart != null) return chart.Parent as Excel.Workbook;
            return null;
        }

        public IList<WorksheetItem> GetWorksheets(Excel.Workbook workbook)
        {
            var result = new List<WorksheetItem>();
            var worksheets = workbook.Worksheets;
            for (int i = 1; i <= worksheets.Count; i++)
            {
                var sheet = (Excel.Worksheet)worksheets[i];
                result.Add(new WorksheetItem(GetWorksheetId(sheet), sheet.Name, sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible));
            }
            return result;
        }

        public string GetActiveWorksheetId(Excel.Window window)
        {
            var sheet = window.ActiveSheet as Excel.Worksheet;
            return sheet == null ? null : GetWorksheetId(sheet);
        }

        public string GetWorksheetId(Excel.Worksheet sheet)
        {
            // Fresh workbooks can have empty CodeNames until Excel initializes VBA.
            // Names are unique within a workbook, including chart sheets. Prefixes
            // prevent collisions between a real CodeName and the fallback name.
            string codeName = sheet.CodeName;
            return string.IsNullOrEmpty(codeName) ? "name:" + sheet.Name : "code:" + codeName;
        }

        public string ActivateWorksheet(int windowHandle, string codeName)
        {
            var window = FindWindow(windowHandle);
            if (window == null) return "対象のウィンドウは閉じられています。";
            var workbook = GetWorkbook(window);
            if (workbook == null) return "ブックが見つかりません。";

            var worksheets = workbook.Worksheets;
            for (int i = 1; i <= worksheets.Count; i++)
            {
                var sheet = (Excel.Worksheet)worksheets[i];
                if (GetWorksheetId(sheet) != codeName) continue;
                if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    return "非表示のシートには切り替えられません。Excelで再表示してください。";
                window.Activate();
                sheet.Activate();
                return "シート「" + sheet.Name + "」に切り替えました。";
            }
            return "シートが見つかりません。一覧の更新後に選び直してください。";
        }

        public OperationTarget CaptureTarget(int windowHandle, string codeName)
        {
            var window = FindWindow(windowHandle);
            if (window == null) throw new InvalidOperationException("対象のウィンドウは閉じられています。");
            var workbook = GetWorkbook(window);
            if (workbook == null) throw new InvalidOperationException("対象のブックが見つかりません。");
            var target = new OperationTarget { Window = window, Workbook = workbook, WindowHandle = windowHandle };
            if (codeName != null)
            {
                var sheets = workbook.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    var sheet = (Excel.Worksheet)sheets[i];
                    if (GetWorksheetId(sheet) == codeName) { target.Worksheet = sheet; break; }
                }
                if (target.Worksheet == null) throw new InvalidOperationException("対象のシートは既に削除されています。");
            }
            ValidateTarget(target);
            return target;
        }

        private void ValidateTarget(OperationTarget target)
        {
            var currentWindow = FindWindow(target.WindowHandle);
            // Excel can return a different Window COM identity for the same HWND.
            // Check the live handle and workbook identity, not Window.Equals.
            if (currentWindow == null || target.Window.Hwnd != target.WindowHandle ||
                !Equals(GetWorkbook(currentWindow), target.Workbook))
                throw new InvalidOperationException("対象のブックまたはウィンドウが閉じられました。操作をやり直してください。");
            if (target.Workbook.ReadOnly)
                throw new InvalidOperationException("読み取り専用のブックではシートを変更できません。書き込み可能なブックでお試しください。");
            if (target.Workbook.ProtectStructure)
                throw new InvalidOperationException("ブックの構成が保護されています。Excelで保護を解除してからお試しください。");
            if (target.Workbook.MultiUserEditing)
                throw new InvalidOperationException("従来の共有ブックではこの操作はできません。共有を解除してからお試しください。");
            if (target.Worksheet != null)
            {
                var sheets = target.Workbook.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                    if (Equals(sheets[i], target.Worksheet)) return;
                throw new InvalidOperationException("対象のシートは既に削除されています。操作をやり直してください。");
            }
        }

        public string AddWorksheet(OperationTarget target)
        {
            ValidateTarget(target);
            target.Window.Activate();
            ValidateTarget(target);
            var sheets = target.Workbook.Sheets;
            object after = target.Worksheet ?? sheets[sheets.Count];
            var added = (Excel.Worksheet)target.Workbook.Worksheets.Add(After: after);
            return GetWorksheetId(added);
        }

        public void RenameWorksheet(OperationTarget target, string newName)
        {
            ValidateTarget(target);
            if (target.Worksheet == null) throw new InvalidOperationException("シートを選択してください。");
            var names = new List<string>();
            var sheets = target.Workbook.Sheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                dynamic sheet = sheets[i];
                if (!Equals((object)sheet, target.Worksheet)) names.Add((string)sheet.Name);
            }
            string error = WorksheetNameRules.Validate(newName, names);
            if (error != null) throw new InvalidOperationException(error);
            if (target.Worksheet.Name != newName) target.Worksheet.Name = newName;
        }

        public string DuplicateWorksheet(OperationTarget target)
        {
            ValidateTarget(target);
            if (target.Worksheet == null) throw new InvalidOperationException("シートを選択してください。");
            target.Window.Activate();
            ValidateTarget(target);
            target.Worksheet.Copy(After: target.Worksheet);
            var copy = (Excel.Worksheet)target.Worksheet.Next;
            return GetWorksheetId(copy);
        }

        public void EnsureCanDelete(OperationTarget target)
        {
            ValidateTarget(target);
            if (target.Worksheet == null) throw new InvalidOperationException("シートを選択してください。");
            if (target.Worksheet.Visible != Excel.XlSheetVisibility.xlSheetVisible) return;
            int visibleCount = 0;
            var sheets = target.Workbook.Sheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                dynamic sheet = sheets[i];
                if ((int)sheet.Visible == (int)Excel.XlSheetVisibility.xlSheetVisible) visibleCount++;
            }
            if (visibleCount <= 1) throw new InvalidOperationException("最後の表示シートは削除できません。");
        }

        public void DeleteWorksheet(OperationTarget target, string confirmedName)
        {
            // Recheck after the confirmation dialog: the workbook may have changed.
            EnsureCanDelete(target);
            if (target.Worksheet.Name != confirmedName)
                throw new InvalidOperationException("確認中にシート名が変わりました。対象を確認して、もう一度削除してください。");
            bool previousAlerts = application.DisplayAlerts;
            try
            {
                application.DisplayAlerts = false;
                target.Worksheet.Delete();
            }
            finally { application.DisplayAlerts = previousAlerts; }
        }

        public MoveSession CaptureMove(int windowHandle, string worksheetId, IList<string> displayedOrder)
        {
            var target = CaptureTarget(windowHandle, worksheetId);
            if (target.Worksheet == null) throw new InvalidOperationException("移動するシートを選択してください。");
            var allSheets = target.Workbook.Sheets;
            var order = new object[allSheets.Count];
            for (int i = 0; i < order.Length; i++) order[i] = allSheets[i + 1];
            var worksheetObjects = order.OfType<Excel.Worksheet>().ToArray();
            var ids = worksheetObjects.Select(GetWorksheetId).ToArray();
            if (!ids.SequenceEqual(displayedOrder))
                throw new InvalidOperationException("シート一覧が変更されました。更新後にもう一度ドラッグしてください。");
            var session = new MoveSession
            {
                Target = target, SheetOrder = order, Worksheets = worksheetObjects, WorksheetIds = ids,
                SourceIndex = Array.IndexOf(ids, worksheetId)
            };
            ValidateMove(session);
            return session;
        }

        public void ValidateMove(MoveSession session)
        {
            ValidateTarget(session.Target);
            var activeWindow = application.ActiveWindow;
            if (activeWindow == null || activeWindow.Hwnd != session.Target.WindowHandle ||
                !Equals(GetWorkbook(activeWindow), session.Target.Workbook))
                throw new InvalidOperationException("ブックまたはウィンドウが切り替わったため、並び替えを中止しました。");
            var sheets = session.Target.Workbook.Sheets;
            if (sheets.Count != session.SheetOrder.Length)
                throw new InvalidOperationException("ドラッグ中にシート構成が変更されました。もう一度操作してください。");
            for (int i = 0; i < sheets.Count; i++)
                if (!Equals(sheets[i + 1], session.SheetOrder[i]))
                    throw new InvalidOperationException("ドラッグ中にシート順が変更されました。もう一度操作してください。");
            if (!session.Worksheets.Select(GetWorksheetId).SequenceEqual(session.WorksheetIds))
                throw new InvalidOperationException("ドラッグ中にシートの識別情報が変わりました。もう一度操作してください。");
        }

        public bool MoveWorksheet(MoveSession session, int boundary)
        {
            ValidateMove(session);
            int destination = WorksheetMoveRules.DestinationIndex(session.SourceIndex, boundary, session.Worksheets.Length);
            if (destination == session.SourceIndex) return false;
            var source = session.Target.Worksheet;
            var anchor = boundary < session.Worksheets.Length ? session.Worksheets[boundary] : session.Worksheets.Last();
            if (!Equals(source.Parent, session.Target.Workbook) || !Equals(anchor.Parent, session.Target.Workbook))
                throw new InvalidOperationException("異なるブックへの移動はできません。");
            // Exactly one argument is mandatory; neither would create a NEW workbook.
            // Chart sheets are not anchors. Preserve their order relative to each other.
            var visibility = source.Visible;
            bool hidden = visibility != Excel.XlSheetVisibility.xlSheetVisible;
            bool screenUpdating = application.ScreenUpdating;
            object previouslyActiveSheet = hidden ? session.Target.Window.ActiveSheet : null;
            try
            {
                // Excel can reject moving a VeryHidden sheet directly. Do not expose
                // its contents: suspend painting and restore visibility in all cases.
                if (hidden)
                {
                    application.ScreenUpdating = false;
                    source.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    ValidateMove(session);
                }
                if (boundary < session.Worksheets.Length) source.Move(Before: anchor);
                else source.Move(After: anchor);
            }
            finally
            {
                if (hidden)
                {
                    try
                    {
                        source.Visible = visibility;
                        var currentWindow = application.ActiveWindow;
                        if (currentWindow != null && currentWindow.Hwnd == session.Target.WindowHandle)
                        {
                            var previousWorksheet = previouslyActiveSheet as Excel.Worksheet;
                            var previousChart = previouslyActiveSheet as Excel.Chart;
                            if (previousWorksheet != null) previousWorksheet.Activate();
                            else if (previousChart != null) previousChart.Activate();
                        }
                    }
                    finally { application.ScreenUpdating = screenUpdating; }
                }
            }
            return true;
        }

        // Do not FinalReleaseComObject on Excel/VSTO's shared RCWs. In particular,
        // releasing an event argument or Application can disconnect other add-ins.
        // No workbook/worksheet RCW is stored in the view or a long-lived field.
    }
}
