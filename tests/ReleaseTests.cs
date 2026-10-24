using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Sheaside.Controls;
using Sheaside.Models;
using Sheaside.Services;

internal static partial class SmokeTests
{
    private static void TestReleaseSettings()
    {
        string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings-" + Guid.NewGuid().ToString("N") + ".xml");
        try
        {
            var settings = new UserPreferences(path);
            Check(!settings.ShowHidden && !settings.ShowVeryHidden, "both visibility switches default off");
            using (var first = new SheetListControl(settings))
            using (var second = new SheetListControl(settings))
            {
                var mixed = new[] { new WorksheetItem("v", "Visible", true), new WorksheetItem("h", "Hidden", WorksheetVisibility.Hidden, null), new WorksheetItem("vh", "Secret", WorksheetVisibility.VeryHidden, null, SheetKind.Chart) };
                first.ShowWorksheets("Book", mixed, "v"); second.ShowWorksheets("Book", mixed, "v");
                settings.SetShowHidden(true);
                Check(first.Controls.OfType<ListBox>().Single().Items.Count == 2 && second.Controls.OfType<ListBox>().Single().Items.Count == 2, "Hidden switch reaches both window views");
                settings.SetShowVeryHidden(true); settings.SetShowHidden(false);
                Check(first.Controls.OfType<ListBox>().Single().Items.Cast<WorksheetItem>().Select(i => i.Id).SequenceEqual(new[] { "v", "vh" }), "VeryHidden independent of Hidden");
                Check(new UserPreferences(path).ShowVeryHidden && !new UserPreferences(path).ShowHidden, "independent visibility options persist");
                settings.SetShowVeryHidden(false);
            }
            Check(!settings.ShowVeryHidden, "VeryHidden default off");
            using (var view = new SheetListControl(settings))
            {
                var rows = new[] { new WorksheetItem("v", "Visible", true), new WorksheetItem("h", "Secret", WorksheetVisibility.VeryHidden, null, SheetKind.Chart) };
                view.ShowWorksheets("Book", rows, "v");
                var list = view.Controls.OfType<ReorderableSheetListBox>().Single();
                Check(list.Items.Count == 1 && view.IsFiltered && list.ReorderEnabled, "VeryHidden omitted without disabling reorder");
                settings.SetShowVeryHidden(true);
                Check(list.Items.Count == 2 && !view.IsFiltered, "setting reveals cached VeryHidden rows");
                InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Control | Keys.A));
                settings.SetShowVeryHidden(false);
                Check(view.SelectedSheetIds.SequenceEqual(new[] { "v" }), "excluded VeryHidden removed from selected targets");
                settings.SetShowVeryHidden(true);
                Check(view.SelectedSheetIds.SequenceEqual(new[] { "v" }), "revealed rows do not silently regain selection");
            }
            Check(new UserPreferences(path).ShowVeryHidden, "settings persist across instances");
            settings.SetShowHidden(true);
            var restored = new UserPreferences(path);
            Check(restored.ShowHidden && restored.ShowVeryHidden, "both display settings persist");
            int events = 0; restored.Changed += (s, e) => events++;
            restored.SetShowHidden(true);
            Check(events == 0, "unchanged settings do not trigger redraw");
            File.WriteAllText(path, "<Sheaside theme=\"obsolete\" dockLeft=\"true\" startupVisible=\"true\" />");
            Check(new UserPreferences(path).LoadError == null && !new UserPreferences(path).ShowHidden, "obsolete settings ignored safely");
            var failing = new UserPreferences(Path.Combine(path, "child.xml"));
            try { failing.SetShowVeryHidden(true); throw new Exception("Expected write failure"); }
            catch (IOException) { Check(!failing.ShowVeryHidden, "failed save keeps in-memory settings unchanged"); }
            File.WriteAllText(path, "invalid xml");
            var corrupt = new UserPreferences(path);
            Check(!corrupt.ShowVeryHidden && corrupt.LoadError != null, "corrupt settings fall back safely");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    private static void TestThemes()
    {
        using (var button = new ModernButton { Text = "追加", Icon = LineIcon.Add, Size = new System.Drawing.Size(120, 34), BackColor = System.Drawing.Color.White, ForeColor = System.Drawing.Color.Black })
        {
            InvokeControl(button, "OnMouseEnter", EventArgs.Empty);
            Check(button.IsHovered, "modern button hover state");
            InvokeControl(button, "OnMouseLeave", EventArgs.Empty);
            Check(!button.IsHovered && button.GetPreferredSize(System.Drawing.Size.Empty).Height >= 32, "modern button spacing and hover reset");
            foreach (int dpi in new[] { 96, 120, 144 })
                using (var bitmap = new System.Drawing.Bitmap(180, 60))
                {
                    bitmap.SetResolution(dpi, dpi);
                    using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                        foreach (LineIcon icon in Enum.GetValues(typeof(LineIcon))) ModernButton.DrawIcon(graphics, icon, new System.Drawing.Rectangle(5, 5, dpi / 6, dpi / 6), System.Drawing.Color.Black);
                    Check(bitmap.GetPixel(5 + dpi / 12, 5 + dpi / 12).A > 0, "line icons render at DPI " + dpi);
                }
        }
        Check(OfficeThemeReader.Decode(16, 0) == OfficeTheme.Colorful && OfficeThemeReader.Decode(16, 5) == OfficeTheme.White, "Office light theme mapping");
        Check(OfficeThemeReader.Decode(16, 3) == OfficeTheme.DarkGray && OfficeThemeReader.Decode(16, 4) == OfficeTheme.Black, "Office dark theme mapping");
        Check(OfficeThemeReader.Decode(15, 2) == OfficeTheme.DarkGray && OfficeThemeReader.Decode(15, 1) == OfficeTheme.LightGray, "Office 2013 theme mapping");
        Check(OfficeThemeReader.Decode(16, null) == OfficeTheme.System && OfficeThemeReader.Decode(16, 99) == OfficeTheme.System, "missing or unknown Office theme safely falls back");
        Check(OfficeThemeReader.Decode(16, 7, true) == OfficeTheme.Black && OfficeThemeReader.Decode(16, 7, false) == OfficeTheme.White, "Office system mode follows Windows preference");
        Check(OfficeThemeReader.Read("invalid") == OfficeTheme.System, "invalid Office version safely falls back");
        var settings = new UserPreferences();
        using (var first = new SheetListControl(settings))
        using (var second = new SheetListControl(settings))
        {
            var rows = new[] { new WorksheetItem("a", "Sheet", true) };
            first.ShowWorksheets("Book", rows, "a"); second.ShowWorksheets("Book", rows, "a");
            var list = first.Controls.OfType<ReorderableSheetListBox>().Single();
            var other = second.Controls.OfType<ReorderableSheetListBox>().Single();
            first.SetOfficeTheme(OfficeTheme.Black); second.SetOfficeTheme(OfficeTheme.Black);
            Check(list.BackColor == other.BackColor && (SystemInformation.HighContrast || list.BackColor.R < 50), "dark theme reaches all panes and respects high contrast");
            Check(first.SelectedSheetIds.Single() == "a", "theme change preserves selection");
            first.SetOfficeTheme(OfficeTheme.White);
            Check(SystemInformation.HighContrast || list.BackColor == System.Drawing.Color.White, "light theme uses light list background");
            first.SetOfficeTheme(OfficeTheme.System);
            Check(list.BackColor == System.Drawing.SystemColors.Window && list.ForeColor == System.Drawing.SystemColors.WindowText, "system theme restores system colors");
            first.Size = new System.Drawing.Size(220, 400); first.PerformLayout();
            Check(list.Width > 150 && list.Height > 80, "compact pane keeps usable list area");
            first.Size = new System.Drawing.Size(500, 700); first.PerformLayout();
            Check(list.Width > 400 && list.Height > 300, "list expands with pane size");
            Check(first.AddButton is ModernButton && first.DeleteButton is ModernButton && first.AddButton.Height == first.DeleteButton.Height, "action buttons use unified renderer and equal height");
            first.Size = new System.Drawing.Size(300, 520);
            settings.SetShowHidden(true); settings.SetShowVeryHidden(true);
            first.ShowWorksheets("UI preview", new[] { new WorksheetItem("a", "売上", true), new WorksheetItem("b", "月別売上データ_2026年度_集計用の非表示シート", WorksheetVisibility.Hidden, 255, SheetKind.Worksheet, true), new WorksheetItem("c", "グラフ", WorksheetVisibility.Visible, null, SheetKind.Chart), new WorksheetItem("d", "内部集計", WorksheetVisibility.VeryHidden, 0, SheetKind.Worksheet, true) }, "c");
            list.ClearSelected(); list.SetSelected(0, true);
            first.PerformLayout();
            using (var bitmap = new System.Drawing.Bitmap(first.Width, first.Height))
            {
                first.DrawToBitmap(bitmap, first.ClientRectangle);
                bitmap.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rc5-pane.png"));
            }
        }
    }

    private static void TestFilters()
    {
        var settings = new UserPreferences();
        using (var view = new SheetListControl(settings))
        {
            var rows = new[] { new WorksheetItem("v", "Sales 日本", true), new WorksheetItem("h", "Sales hidden", WorksheetVisibility.Hidden, null), new WorksheetItem("c", "Sales chart", WorksheetVisibility.Visible, null, SheetKind.Chart), new WorksheetItem("vh", "Sales secret", WorksheetVisibility.VeryHidden, null) };
            view.ShowWorksheets("Book", rows, "v");
            var list = view.Controls.OfType<ReorderableSheetListBox>().Single();
            Check(list.Items.Count == 2 && !view.ClearSearchButton.Visible, "visible worksheets and charts shown by default");
            view.SetSearch("SALES");
            Check(list.Items.Count == 2 && view.ClearSearchButton.Visible, "case insensitive realtime search");
            settings.SetShowHidden(true); settings.SetShowVeryHidden(true);
            Check(list.Items.Count == 4, "display switches combine with search");
            InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Control | Keys.A));
            view.SetSearch("日本");
            Check(list.Items.Count == 1 && view.SelectedSheetIds.SequenceEqual(new[] { "v" }), "search prunes excluded selections");
            view.Size = new System.Drawing.Size(260, 400);
            list.CreateControl();
            int starts = 0;
            view.DragStarting += (s, e) =>
            {
                starts++;
                Check(e.SelectedIds.SequenceEqual(new[] { "v" }) && e.Order.SequenceEqual(new[] { "v" }), "filtered drag carries only displayed selected IDs");
            };
            InvokeControl(list, "OnMouseDown", new MouseEventArgs(MouseButtons.Left, 1, 8, 8, 0));
            InvokeControl(list, "OnMouseMove", new MouseEventArgs(MouseButtons.Left, 0, 100, 40, 0));
            Check(starts == 1 && list.ReorderEnabled, "search permits drag without restoring excluded targets");
            view.ClearSearchButton.PerformClick();
            Check(list.Items.Count == 4 && !view.ClearSearchButton.Visible && settings.ShowHidden && settings.ShowVeryHidden, "search reset preserves ribbon visibility switches");
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "v" }), "search reset does not resurrect selections");
            view.SetSearch("missing");
            Check(list.Items.Count == 0 && !view.DeleteButton.Enabled, "empty results have no destructive targets");
            rows[0] = new WorksheetItem("v", "missing now", true);
            view.ShowWorksheets("Book", rows, "v");
            Check(list.Items.Count == 1, "external rename refreshes search");
            view.SetSearch("Sales"); settings.SetShowVeryHidden(false); settings.SetShowHidden(false);
            Check(list.Items.Count == 1 && ((WorksheetItem)list.Items[0]).Kind == SheetKind.Chart, "excluded hidden sheets cannot match search");
            view.Dispose(); settings.SetShowHidden(true);
            Check(view.IsDisposed, "disposed panel detaches settings events");
        }
    }
}
