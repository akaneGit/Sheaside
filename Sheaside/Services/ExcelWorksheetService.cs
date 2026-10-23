using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
            public ExcelSheet Sheet;
            public int WindowHandle;
        }

        internal sealed class MoveSession
        {
            public OperationTarget Target;
            public object[] SheetOrder;
            public ExcelSheet[] Sheets;
            public string[] WorksheetIds;
            public int SourceIndex;
        }
        private readonly Excel.Application application;
        // Keep identities stable even before Excel initializes CodeName. References
        // live only as long as the open workbook; deleted rows are pruned on refresh.
        private readonly Dictionary<Excel.Workbook, Dictionary<object, string>> identities =
            new Dictionary<Excel.Workbook, Dictionary<object, string>>();

        public void ClearIdentities() { identities.Clear(); }

        public ExcelWorksheetService(Excel.Application application) { this.application = application; }

        public IList<Excel.Window> GetWindows()
        {
            var openBooks = new HashSet<Excel.Workbook>();
            var books = application.Workbooks;
            for (int i = 1; i <= books.Count; i++) openBooks.Add(books[i]);
            foreach (var book in identities.Keys.Where(book => !openBooks.Contains(book)).ToArray()) identities.Remove(book);
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
            var liveSheets = new HashSet<object>();
            var worksheets = workbook.Sheets;
            for (int i = 1; i <= worksheets.Count; i++)
            {
                var sheet = ExcelSheet.From(worksheets[i]);
                liveSheets.Add(sheet.Native);
                result.Add(new WorksheetItem(GetWorksheetId(sheet), sheet.Name, (WorksheetVisibility)(int)sheet.Visible, GetTabColor(sheet), sheet.Kind));
            }
            Dictionary<object, string> registry;
            if (identities.TryGetValue(workbook, out registry))
                foreach (object sheet in registry.Keys.Where(sheet => !liveSheets.Contains(sheet)).ToArray()) registry.Remove(sheet);
            return result;
        }

        public string GetActiveWorksheetId(Excel.Window window)
        {
            var sheet = window.ActiveSheet;
            return sheet == null ? null : GetWorksheetId(sheet);
        }

        public string GetWorksheetId(object nativeSheet)
        {
            var sheet = nativeSheet as ExcelSheet ?? ExcelSheet.From(nativeSheet);
            var workbook = (Excel.Workbook)sheet.Parent;
            Dictionary<object, string> registry;
            if (!identities.TryGetValue(workbook, out registry))
                identities.Add(workbook, registry = new Dictionary<object, string>());
            string id;
            if (!registry.TryGetValue(sheet.Native, out id))
                registry.Add(sheet.Native, id = (sheet.Kind == SheetKind.Chart ? "chart:" : "sheet:") + Guid.NewGuid().ToString("N"));
            return id;
        }

        public string ActivateWorksheet(int windowHandle, string codeName)
        {
            var window = FindWindow(windowHandle);
            if (window == null) return "対象のウィンドウは閉じられています。";
            var workbook = GetWorkbook(window);
            if (workbook == null) return "ブックが見つかりません。";

            var worksheets = workbook.Sheets;
            for (int i = 1; i <= worksheets.Count; i++)
            {
                var sheet = ExcelSheet.From(worksheets[i]);
                if (GetWorksheetId(sheet) != codeName) continue;
                if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    return sheet.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden
                        ? "VeryHiddenのシートです。右クリックの「再表示」から表示できます。"
                        : "非表示のシートです。右クリックの「再表示」から表示できます。";
                window.Activate();
                sheet.Activate();
                return null;
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
                var sheets = workbook.Sheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    var sheet = ExcelSheet.From(sheets[i]);
                    if (GetWorksheetId(sheet) == codeName) { target.Sheet = sheet; break; }
                }
                if (target.Sheet == null) throw new InvalidOperationException("対象のシートは既に削除されています。");
            }
            ValidateTarget(target);
            return target;
        }

        internal void ValidateTarget(OperationTarget target)
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
            if (target.Sheet != null)
            {
                var sheets = target.Workbook.Sheets;
                for (int i = 1; i <= sheets.Count; i++)
                    if (Equals(sheets[i], target.Sheet.Native)) return;
                throw new InvalidOperationException("対象のシートは既に削除されています。操作をやり直してください。");
            }
        }

        public string AddWorksheet(OperationTarget target)
        {
            ValidateTarget(target);
            target.Window.Activate();
            ValidateTarget(target);
            var sheets = target.Workbook.Sheets;
            object after = target.Sheet?.Native ?? sheets[sheets.Count];
            var added = (Excel.Worksheet)target.Workbook.Worksheets.Add(After: after);
            return GetWorksheetId(added);
        }

        public void RenameWorksheet(OperationTarget target, string newName)
        {
            ValidateTarget(target);
            if (target.Sheet == null) throw new InvalidOperationException("シートを選択してください。");
            var names = new List<string>();
            var sheets = target.Workbook.Sheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                dynamic sheet = sheets[i];
                if (!Equals((object)sheet, target.Sheet.Native)) names.Add((string)sheet.Name);
            }
            string error = WorksheetNameRules.Validate(newName, names);
            if (error != null) throw new InvalidOperationException(error);
            if (target.Sheet.Name != newName) target.Sheet.Name = newName;
        }

        public string DuplicateWorksheet(OperationTarget target)
        {
            ValidateTarget(target);
            if (target.Sheet == null) throw new InvalidOperationException("シートを選択してください。");
            target.Window.Activate();
            ValidateTarget(target);
            var visibility = target.Sheet.Visible;
            bool temporarilyShow = visibility == Excel.XlSheetVisibility.xlSheetVeryHidden;
            bool screenUpdating = application.ScreenUpdating;
            object previousActive = target.Window.ActiveSheet;
            try
            {
                if (temporarilyShow)
                {
                    application.ScreenUpdating = false;
                    target.Sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                }
                return CopyWorksheet(target, visibility);
            }
            finally
            {
                if (temporarilyShow)
                {
                    try
                    {
                        target.Sheet.Visible = visibility;
                        var activeWindow = application.ActiveWindow;
                        if (activeWindow != null && activeWindow.Hwnd == target.WindowHandle)
                        {
                            var worksheet = previousActive as Excel.Worksheet;
                            var chart = previousActive as Excel.Chart;
                            if (worksheet != null) worksheet.Activate();
                            else if (chart != null) chart.Activate();
                        }
                    }
                    finally { application.ScreenUpdating = screenUpdating; }
                }
            }
        }

        private string CopyWorksheet(OperationTarget target, Excel.XlSheetVisibility visibility)
        {
            ValidateTarget(target);
            var before = new HashSet<string>(GetWorksheets(target.Workbook).Select(i => i.Id));
            string sourceId = GetWorksheetId(target.Sheet);
            target.Sheet.CopyAfter();
            var after = GetWorksheets(target.Workbook).ToArray();
            var added = after.Where(i => !before.Contains(i.Id)).ToArray();
            if (added.Length != 1 || added[0].Kind != target.Sheet.Kind)
                throw new InvalidOperationException("複製先を一意に確認できません。実際の一覧を確認してください。");
            string copyId = added[0].Id;
            var copy = CaptureTarget(target.WindowHandle, copyId);
            if (!Equals(copy.Workbook, target.Workbook)) throw new InvalidOperationException("複製先のブックが変わりました。");
            int sourceIndex = Array.FindIndex(after, i => i.Id == sourceId);
            int copyIndex = Array.FindIndex(after, i => i.Id == copyId);
            if (sourceIndex < 0) throw new InvalidOperationException("複製中に元シートが削除されました。");
            // Excel can ignore an After anchor when it is hidden. Do not infer
            // identity from Next; normalize the identified copy's actual tab order.
            if (copyIndex != sourceIndex + 1)
                MoveWorksheet(CaptureMove(target.WindowHandle, copyId, after.Select(i => i.Id).ToArray()), sourceIndex + 1);
            ValidateTarget(copy);
            if (copy.Sheet.Visible != visibility) copy.Sheet.Visible = visibility;
            return copyId;
        }

        public int? GetTabColor(object nativeSheet)
        {
            var sheet = nativeSheet as ExcelSheet ?? ExcelSheet.From(nativeSheet);
            var tab = sheet.Tab;
            return Convert.ToInt32(tab.ColorIndex) == (int)Excel.XlColorIndex.xlColorIndexNone
                ? (int?)null : Convert.ToInt32(tab.Color);
        }

        public void SetTabColor(OperationTarget target, int? color)
        {
            ValidateTarget(target);
            if (target.Sheet == null) throw new InvalidOperationException("シートを選択してください。");
            if (color.HasValue) target.Sheet.Tab.Color = color.Value;
            else target.Sheet.Tab.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
        }

        public int[] GetThemeColors(Excel.Workbook workbook)
        {
            try
            {
                dynamic theme = ((dynamic)workbook).Theme;
                var result = new int[10];
                for (int i = 0; i < result.Length; i++) result[i] = Convert.ToInt32(theme.ThemeColorScheme.Colors(i + 1).RGB);
                return result;
            }
            catch (Exception e) when (e is COMException || e is InvalidComObjectException || e is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            { return null; }
        }
        public bool MatchesTabColor(ExcelSheet sheet, TabColorChoice choice)
        {
            if (!choice.Theme.HasValue)
            {
                if (GetTabColor(sheet) != choice.Rgb) return false;
                if (!choice.Rgb.HasValue) return true;
                try { int theme = Convert.ToInt32(sheet.Tab.ThemeColor); return theme < 1 || theme > 12; }
                catch (COMException) { return true; }
            }
            try { return Convert.ToInt32(sheet.Tab.ThemeColor) == choice.Theme.Value && Math.Abs(sheet.Tab.TintAndShade - choice.Tint) < .0001; }
            catch (COMException) { return false; } // RGB tabs have no ThemeColor.
        }
        public void SetTabColorChoice(OperationTarget target, TabColorChoice choice)
        {
            ValidateTarget(target);
            if (target.Sheet == null) throw new InvalidOperationException("シートを選択してください。");
            if (!choice.Theme.HasValue) { SetTabColor(target, choice.Rgb); return; }
            target.Sheet.Tab.ThemeColor = (Excel.XlThemeColor)choice.Theme.Value;
            target.Sheet.Tab.TintAndShade = choice.Tint;
        }

        public void ChangeVisibility(OperationTarget target, WorksheetVisibility expected, string confirmedName = null)
        {
            ValidateTarget(target);
            if (target.Sheet == null) throw new InvalidOperationException("シートを選択してください。");
            if ((int)target.Sheet.Visible != (int)expected)
                throw new InvalidOperationException("表示状態が変わりました。一覧を確認して操作をやり直してください。");
            if (expected == WorksheetVisibility.VeryHidden && target.Sheet.Name != confirmedName)
                throw new InvalidOperationException("VeryHiddenの再表示には、対象シート名を確認してください。");
            if (expected == WorksheetVisibility.Visible && CountVisibleSheets(target.Workbook) <= 1)
                throw new InvalidOperationException("最後の表示シートは非表示にできません。");
            target.Sheet.Visible = expected == WorksheetVisibility.Visible
                ? Excel.XlSheetVisibility.xlSheetHidden : Excel.XlSheetVisibility.xlSheetVisible;
        }

        public void HideWorksheet(OperationTarget target)
        { SetVisibility(target, WorksheetVisibility.Hidden); }

        public void SetVisibility(OperationTarget target, WorksheetVisibility visibility)
        {
            ValidateTarget(target);
            if (target.Sheet == null) throw new InvalidOperationException("シートを選択してください。");
            if (!Enum.IsDefined(typeof(WorksheetVisibility), visibility)) throw new InvalidOperationException("不正な表示状態です。");
            if (visibility != WorksheetVisibility.Visible && target.Sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible && CountVisibleSheets(target.Workbook) <= 1)
                throw new InvalidOperationException("最後の表示シートは非表示にできません。");
            target.Sheet.Visible = (Excel.XlSheetVisibility)visibility;
        }

        private static int CountVisibleSheets(Excel.Workbook workbook)
        {
            int count = 0;
            var sheets = workbook.Sheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                dynamic sheet = sheets[i];
                if ((int)sheet.Visible == (int)Excel.XlSheetVisibility.xlSheetVisible) count++;
            }
            return count;
        }

        public void EnsureCanDelete(OperationTarget target)
        {
            ValidateTarget(target);
            if (target.Sheet == null) throw new InvalidOperationException("シートを選択してください。");
            if (target.Sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible) return;
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
            if (target.Sheet.Name != confirmedName)
                throw new InvalidOperationException("確認中にシート名が変わりました。対象を確認して、もう一度削除してください。");
            bool previousAlerts = application.DisplayAlerts;
            try
            {
                application.DisplayAlerts = false;
                target.Sheet.Delete();
            }
            finally { application.DisplayAlerts = previousAlerts; }
        }

        public MoveSession CaptureMove(int windowHandle, string worksheetId, IList<string> displayedOrder)
        {
            var target = CaptureTarget(windowHandle, worksheetId);
            if (target.Sheet == null) throw new InvalidOperationException("移動するシートを選択してください。");
            var allSheets = target.Workbook.Sheets;
            var order = new object[allSheets.Count];
            for (int i = 0; i < order.Length; i++) order[i] = allSheets[i + 1];
            var worksheetObjects = order.Select(ExcelSheet.From).ToArray();
            var ids = worksheetObjects.Select(GetWorksheetId).ToArray();
            if (!ids.SequenceEqual(displayedOrder))
                throw new InvalidOperationException("シート一覧が変更されました。更新後にもう一度ドラッグしてください。");
            var session = new MoveSession
            {
                Target = target, SheetOrder = order, Sheets = worksheetObjects, WorksheetIds = ids,
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
            if (!session.Sheets.Select(GetWorksheetId).SequenceEqual(session.WorksheetIds))
                throw new InvalidOperationException("ドラッグ中にシートの識別情報が変わりました。もう一度操作してください。");
        }

        public bool MoveWorksheet(MoveSession session, int boundary)
        {
            ValidateMove(session);
            int destination = WorksheetMoveRules.DestinationIndex(session.SourceIndex, boundary, session.Sheets.Length);
            if (destination == session.SourceIndex) return false;
            var source = session.Target.Sheet;
            var anchor = boundary < session.Sheets.Length ? session.Sheets[boundary] : session.Sheets.Last();
            if (!Equals(source.Parent, session.Target.Workbook) || !Equals(anchor.Parent, session.Target.Workbook))
                throw new InvalidOperationException("異なるブックへの移動はできません。");
            // Exactly one argument is mandatory; neither would create a NEW workbook.
            // Anchors follow the complete Excel tab order, including chart sheets.
            var visibility = source.Visible;
            bool hidden = visibility != Excel.XlSheetVisibility.xlSheetVisible;
            var anchorVisibility = anchor.Visible;
            bool anchorHidden = anchorVisibility != Excel.XlSheetVisibility.xlSheetVisible;
            bool suspendPainting = hidden || anchorHidden;
            bool screenUpdating = application.ScreenUpdating;
            object previouslyActiveSheet = suspendPainting ? session.Target.Window.ActiveSheet : null;
            try
            {
                // Excel can reject moving a VeryHidden sheet directly. Do not expose
                // its contents: suspend painting and restore visibility in all cases.
                if (suspendPainting)
                {
                    application.ScreenUpdating = false;
                    if (hidden) source.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    // Excel may silently insert BEFORE a hidden After anchor.
                    if (anchorHidden) anchor.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    ValidateMove(session);
                }
                source.Move(anchor, boundary < session.Sheets.Length);
                var expected = session.SheetOrder.ToList();
                expected.RemoveAt(session.SourceIndex);
                expected.Insert(destination, source.Native);
                var actual = session.Target.Workbook.Sheets;
                for (int i = 0; i < expected.Count; i++)
                    if (!Equals(actual[i + 1], expected[i]))
                        throw new InvalidOperationException("Excelが指定と異なる順序に移動しました。一覧の実際の順序を確認してください。");
            }
            finally
            {
                if (suspendPainting)
                {
                    try
                    {
                        try { if (hidden) source.Visible = visibility; }
                        finally { if (anchorHidden) anchor.Visible = anchorVisibility; }
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
        // The view stores IDs only. The identity registry retains RCWs while their
        // workbook is open, pruning deleted sheets and clearing on shutdown.
    }
}
