using System.Collections.Generic;
using Sheaside.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sheaside.Services
{
    internal sealed class ExcelWorksheetService
    {
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
                result.Add(new WorksheetItem(sheet.CodeName, sheet.Name, sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible));
            }
            return result;
        }

        public string GetActiveCodeName(Excel.Window window)
        {
            var sheet = window.ActiveSheet as Excel.Worksheet;
            return sheet == null ? null : sheet.CodeName;
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
                if (sheet.CodeName != codeName) continue;
                if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    return "非表示のシートには切り替えられません。Excelで再表示してください。";
                window.Activate();
                sheet.Activate();
                return "シート「" + sheet.Name + "」に切り替えました。";
            }
            return "シートが見つかりません。一覧の更新後に選び直してください。";
        }

        // Do not FinalReleaseComObject on Excel/VSTO's shared RCWs. In particular,
        // releasing an event argument or Application can disconnect other add-ins.
        // No workbook/worksheet RCW is stored in the view or a long-lived field.
    }
}
