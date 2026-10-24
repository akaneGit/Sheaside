using System;
using Sheaside.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sheaside.Services
{
    // Short-lived typed adapter. Never pass this COM-owning object to a control.
    internal sealed class ExcelSheet
    {
        private readonly Excel.Worksheet worksheet;
        private readonly Excel.Chart chart;
        private ExcelSheet(object sheet)
        {
            worksheet = sheet as Excel.Worksheet;
            chart = sheet as Excel.Chart;
            Native = sheet;
        }
        public static ExcelSheet From(object sheet)
        {
            if (sheet is Excel.Worksheet || sheet is Excel.Chart) return new ExcelSheet(sheet);
            throw new InvalidOperationException("このシート種別には対応していません。");
        }
        public object Native { get; private set; }
        public SheetKind Kind { get { return worksheet != null ? SheetKind.Worksheet : SheetKind.Chart; } }
        public string Name
        {
            get { return worksheet != null ? worksheet.Name : chart.Name; }
            set { if (worksheet != null) worksheet.Name = value; else chart.Name = value; }
        }
        public string CodeName { get { return worksheet != null ? worksheet.CodeName : chart.CodeName; } }
        public bool IsProtected { get { return worksheet != null
            ? worksheet.ProtectContents || worksheet.ProtectDrawingObjects || worksheet.ProtectScenarios
            : chart.ProtectContents || chart.ProtectDrawingObjects || chart.ProtectData || chart.ProtectFormatting || chart.ProtectSelection; } }
        public object Parent { get { return worksheet != null ? worksheet.Parent : chart.Parent; } }
        public object Next { get { return worksheet != null ? worksheet.Next : chart.Next; } }
        public Excel.Tab Tab { get { return worksheet != null ? worksheet.Tab : chart.Tab; } }
        public Excel.XlSheetVisibility Visible
        {
            get { return worksheet != null ? worksheet.Visible : chart.Visible; }
            set { if (worksheet != null) worksheet.Visible = value; else chart.Visible = value; }
        }
        public void Activate() { if (worksheet != null) worksheet.Activate(); else chart.Activate(); }
        public void Delete() { if (worksheet != null) worksheet.Delete(); else chart.Delete(); }
        public void CopyAfter()
        {
            if (worksheet != null) worksheet.Copy(After: Native);
            else chart.Copy(After: Native);
        }
        public void Move(ExcelSheet anchor, bool before)
        {
            if (worksheet != null)
            {
                if (before) worksheet.Move(Before: anchor.Native); else worksheet.Move(After: anchor.Native);
            }
            else
            {
                if (before) chart.Move(Before: anchor.Native); else chart.Move(After: anchor.Native);
            }
        }
    }
}
