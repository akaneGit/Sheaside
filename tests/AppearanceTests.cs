using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.ComponentModel;
using System.Windows.Forms;
using Sheaside.Controls;
using Sheaside.Models;
using Sheaside.Services;
using Excel = Microsoft.Office.Interop.Excel;

internal static partial class SmokeTests
{
    private static void TestAppearance()
    {
        foreach (SheetKind kind in Enum.GetValues(typeof(SheetKind)))
            foreach (WorksheetVisibility state in Enum.GetValues(typeof(WorksheetVisibility)))
            {
                var row = new WorksheetItem("icon", "名前", state, null, kind);
                Check(ReorderableSheetListBox.KindIcon(row) == (kind == SheetKind.Chart ? LineIcon.Chart : LineIcon.Worksheet), "sheet kind icon " + kind + state);
                Check(ReorderableSheetListBox.StateIcon(row) == (state == WorksheetVisibility.Visible ? LineIcon.None : state == WorksheetVisibility.Hidden ? LineIcon.Hidden : LineIcon.Locked), "state icon " + kind + state);
                Check(row.AccessibleDescription.Contains(state.ToString()) && row.AccessibleDescription.Contains(kind == SheetKind.Chart ? "グラフ" : "ワーク"), "accessible row description " + kind + state);
            }
        foreach (int size in new[] { 16, 32, 256 })
            using (var logo = RibbonArtwork.Create(size))
                Check(logo.Width == size && logo.Height == size && logo.GetPixel(0, 0).A == 0 && logo.GetPixel(size / 2, size / 2).A > 0, "brand icon size/transparency " + size);
        var palette = TabColorChoice.Palette(TabColorChoice.DefaultTheme(), true);
        Check(palette.Count == 60 && palette.All(c => c.Theme.HasValue), "theme palette has sixty linked colors");
        Check(palette[0].Theme == 2 && palette[1].Theme == 1 && palette[4].Theme == 5, "Excel palette theme order");
        Check(palette[14].Tint == .8 && palette[54].Tint == -.5, "theme tint shades");
        Check(TabColorChoice.Palette(null, true).All(c => !c.Theme.HasValue), "unavailable theme falls back to fixed RGB");
        Reject(() => TabColorChoice.Themed(0, 0, 0), "invalid theme rejected");
        Reject(() => TabColorChoice.Themed(5, double.NaN, 0), "invalid tint rejected");
        using (var dialog = new ColorPaletteDialog(null))
            Check(dialog.Selection == null, "opening palette does not change color");
        using (var dialog = new ColorPaletteDialog(TabColorChoice.DefaultTheme()))
        {
            var grid = dialog.Controls.OfType<FlowLayoutPanel>().Single().Controls.OfType<TableLayoutPanel>().Single(c => c.Name == "ThemeGrid");
            Check(grid.Controls.Count == 60, "palette grid exposes all theme shades");
            typeof(Button).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(grid.GetControlFromPosition(4, 2), new object[] { EventArgs.Empty });
            Check(dialog.Selection.Theme == 5 && dialog.Selection.Tint == .6 && dialog.DialogResult == DialogResult.OK, "palette button selects linked theme tint");
        }
        foreach (float scale in new[] { 1F, 1.25F, 1.5F })
        using (var dialog = new ColorPaletteDialog(null))
        {
            dialog.Show();
            dialog.Scale(new SizeF(scale, scale));
            dialog.PerformLayout(); Application.DoEvents();
            var layout = dialog.Controls.OfType<FlowLayoutPanel>().Single();
            var standard = layout.Controls.OfType<TableLayoutPanel>().Single(c => c.Name == "StandardGrid");
            Check(standard.Controls.Count == 10 && standard.Controls.Cast<Control>().All(c => c.Width > 10 && c.Height > 10 && standard.ClientRectangle.Contains(c.Bounds)), "standard colors visible with fallback at scale " + scale);
            Check(!layout.VerticalScroll.Visible && !layout.HorizontalScroll.Visible && layout.Controls.Cast<Control>().All(c => layout.ClientRectangle.Contains(c.Bounds)), "complete palette fits without scroll at scale " + scale);
            ((Button)standard.GetControlFromPosition(1, 0)).PerformClick();
            Check(dialog.Selection.Rgb == ColorTranslator.ToOle(Color.Red) && !dialog.Selection.Theme.HasValue, "standard red selection independent of theme");
        }
        var settings = new UserPreferences(); settings.SetShowHidden(true); settings.SetShowVeryHidden(true);
        using (var view = new SheetListControl(settings) { Size = new Size(260, 400) })
        {
            var list = view.Controls.OfType<ListBox>().Single();
            var rows = Enumerable.Range(0, 40).Select(i => new WorksheetItem("s" + i, "Sheet " + i, true)).ToArray();
            view.ShowWorksheets("Book", rows, "s0");
            list.ClearSelected();
            list.SelectedIndex = 20;
            list.TopIndex = 18;
            int top = list.TopIndex;
            rows[20] = new WorksheetItem("s20", "Sheet 20", WorksheetVisibility.VeryHidden, 0);
            view.ShowWorksheets("Book", rows, "s0");
            Check(list.SelectedIndex == 20 && list.TopIndex == top, "appearance preserves selection and scroll");
            Check(((WorksheetItem)list.Items[20]).DisplayName.Contains("VeryHidden"), "VeryHidden label");
            Check(((WorksheetItem)list.Items[20]).TabColor == 0, "black is not no color");
            var menu = list.ContextMenuStrip;
            typeof(ContextMenuStrip).GetMethod("OnOpening", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(menu, new object[] { new CancelEventArgs() });
            Check(menu.Items[6].Enabled && menu.Items[7].Enabled && !menu.Items[8].Enabled, "VeryHidden can hide or reveal via unified menu");
            rows[20] = new WorksheetItem("s20", "Sheet 20", WorksheetVisibility.Hidden, null);
            view.ShowWorksheets("Book", rows, "s0");
            Check(((WorksheetItem)list.Items[20]).DisplayName.Contains("非表示"), "Hidden label distinct");
            typeof(ContextMenuStrip).GetMethod("OnOpening", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(menu, new object[] { new CancelEventArgs() });
            Check(!menu.Items[6].Enabled && menu.Items[7].Enabled && menu.Items[8].Enabled, "Hidden can reveal or become VeryHidden");
            using (var bitmap = new Bitmap(260, 400))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                typeof(ReorderableSheetListBox).GetMethod("OnDrawItem", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(list, new object[] { new DrawItemEventArgs(graphics, list.Font, new Rectangle(0, 0, 260, list.ItemHeight), 20, DrawItemState.Selected) });
                Check(bitmap.GetPixel(7, 7).ToArgb() == SystemColors.Window.ToArgb(), "unset swatch uses system theme");
            }
        }
    }

    private static void TestExcelAppearance()
    {
        var excel = new Excel.Application { Visible = true, DisplayAlerts = false };
        Excel.Workbook book = null;
        Excel.Workbook other = null;
        Excel.Workbook readOnly = null;
        string file = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appearance-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            book = excel.Workbooks.Add();
            var service = new ExcelWorksheetService(excel);
            var first = (Excel.Worksheet)book.Worksheets[1];
            var second = (Excel.Worksheet)book.Worksheets.Add(After: first);
            int handle = book.Windows[1].Hwnd;
            var target = service.CaptureTarget(handle, service.GetWorksheetId(second));
            foreach (var color in new[] { Color.Black, Color.Red, Color.FromArgb(12, 45, 210) })
            {
                int ole = ColorTranslator.ToOle(color);
                service.SetTabColor(target, ole);
                Check(service.GetWorksheets(book).Single(i => i.Id == service.GetWorksheetId(second)).TabColor == ole, "tab color roundtrip " + color);
            }
            service.SetTabColor(target, null);
            Check(service.GetTabColor(second) == null, "reset color");
            second.Tab.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
            Check(service.GetTabColor(second) == Convert.ToInt32(second.Tab.Color), "theme resolves actual color");
            service.ChangeVisibility(target, WorksheetVisibility.Visible);
            Check(second.Visible == Excel.XlSheetVisibility.xlSheetHidden, "hide");
            Check(service.ActivateWorksheet(handle, service.GetWorksheetId(second)) == null && !Equals(book.Windows[1].ActiveSheet, second), "hidden selection is silent and preserves active sheet");
            Check(second.Visible == Excel.XlSheetVisibility.xlSheetHidden, "click never reveals hidden");
            service.ChangeVisibility(target, WorksheetVisibility.Hidden);
            Check(second.Visible == Excel.XlSheetVisibility.xlSheetVisible, "unhide");
            second.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            Check(service.GetWorksheets(book).Single(i => i.Id == service.GetWorksheetId(second)).Visibility == WorksheetVisibility.VeryHidden, "read VeryHidden");
            Reject(() => service.ChangeVisibility(target, WorksheetVisibility.Hidden), "stale normal reveal cannot bypass confirmation");
            Reject(() => service.ChangeVisibility(target, WorksheetVisibility.VeryHidden), "VeryHidden requires confirmation name");
            Reject(() => service.ChangeVisibility(target, WorksheetVisibility.VeryHidden, "wrong"), "stale confirmation rejected");
            Check(service.ActivateWorksheet(handle, service.GetWorksheetId(second)) == null && !Equals(book.Windows[1].ActiveSheet, second), "hidden selection is silent and preserves active sheet");
            Check(second.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden, "click never reveals VeryHidden");
            service.ChangeVisibility(target, WorksheetVisibility.VeryHidden, second.Name);
            Check(second.Visible == Excel.XlSheetVisibility.xlSheetVisible, "confirmed VeryHidden reveal");
            var secondWindow = book.NewWindow();
            var windowTarget = service.CaptureTarget(secondWindow.Hwnd, service.GetWorksheetId(second));
            other = excel.Workbooks.Add();
            var otherSheet = (Excel.Worksheet)other.Worksheets[1];
            otherSheet.Name = second.Name;
            service.SetTabColor(windowTarget, 255);
            service.ChangeVisibility(windowTarget, WorksheetVisibility.Visible);
            Check(service.GetTabColor(second) == 255 && service.GetTabColor(otherSheet) == null, "color stays in captured workbook across window switch");
            Check(second.Visible == Excel.XlSheetVisibility.xlSheetHidden && otherSheet.Visible == Excel.XlSheetVisibility.xlSheetVisible, "visibility stays in captured workbook");
            service.ChangeVisibility(windowTarget, WorksheetVisibility.Hidden);
            secondWindow.Activate();
            book.Protect(Structure: true);
            Reject(() => service.SetTabColor(target, 255), "protected color rejected after capture");
            Reject(() => service.ChangeVisibility(target, WorksheetVisibility.Visible), "protected visibility rejected after capture");
            book.Unprotect();
            for (int i = 1; i <= book.Worksheets.Count; i++)
            {
                var sheet = (Excel.Worksheet)book.Worksheets[i];
                if (!Equals(sheet, first)) sheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            }
            var firstTarget = service.CaptureTarget(handle, service.GetWorksheetId(first));
            Reject(() => service.ChangeVisibility(firstTarget, WorksheetVisibility.Visible), "last visible cannot hide");
            var chart = (Excel.Chart)book.Charts.Add();
            service.ChangeVisibility(firstTarget, WorksheetVisibility.Visible);
            Check(first.Visible == Excel.XlSheetVisibility.xlSheetHidden && chart.Visible == Excel.XlSheetVisibility.xlSheetVisible, "visible chart counts");
            Check(!excel.DisplayAlerts, "appearance preserves alerts");
            book.SaveAs(file, Excel.XlFileFormat.xlOpenXMLWorkbook);
            book.Close(false);
            book = null;
            Reject(() => service.SetTabColor(windowTarget, 0), "closed captured window rejects color");
            readOnly = excel.Workbooks.Open(file, ReadOnly: true);
            Reject(() => service.CaptureTarget(readOnly.Windows[1].Hwnd, service.GetWorksheetId((Excel.Worksheet)readOnly.Worksheets[1])), "read-only appearance target rejected");
        }
        finally
        {
            if (book != null) book.Close(false);
            if (readOnly != null) readOnly.Close(false);
            if (other != null) other.Close(false);
            excel.Quit();
            if (System.IO.File.Exists(file)) System.IO.File.Delete(file);
        }
    }
}
