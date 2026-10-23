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
            Check(!settings.ShowVeryHidden, "VeryHidden default off");
            using (var view = new SheetListControl(settings))
            {
                var rows = new[] { new WorksheetItem("v", "Visible", true), new WorksheetItem("h", "Secret", WorksheetVisibility.VeryHidden, null, SheetKind.Chart) };
                view.ShowWorksheets("Book", rows, "v");
                var list = view.Controls.OfType<ReorderableSheetListBox>().Single();
                Check(list.Items.Count == 1 && view.IsFiltered, "VeryHidden omitted and reorder disabled");
                settings.SetShowVeryHidden(true);
                Check(list.Items.Count == 2 && !view.IsFiltered, "setting reveals cached VeryHidden rows");
                InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Control | Keys.A));
                settings.SetShowVeryHidden(false);
                Check(view.SelectedSheetIds.SequenceEqual(new[] { "v" }), "excluded VeryHidden removed from selected targets");
                settings.SetShowVeryHidden(true);
                Check(view.SelectedSheetIds.SequenceEqual(new[] { "v" }), "revealed rows do not silently regain selection");
            }
            Check(new UserPreferences(path).ShowVeryHidden, "settings persist across instances");
            settings.SetOptions(true, false);
            var restored = new UserPreferences(path);
            Check(!restored.StartupVisible && restored.ShowVeryHidden, "all settings persist together");
            int events = 0; restored.Changed += (s, e) => events++;
            restored.SetOptions(true, false);
            Check(events == 0, "unchanged settings do not trigger redraw");
            File.WriteAllText(path, "<Sheaside theme=\"invalid-obsolete\" dockLeft=\"true\" startupVisible=\"false\" showVeryHidden=\"true\" />");
            Check(new UserPreferences(path).LoadError == null && !new UserPreferences(path).StartupVisible, "obsolete settings ignored" );
            using (var dialog = new PreferencesDialog(restored))
                Check(!dialog.StartupVisible && dialog.ShowVeryHidden, "settings dialog initializes saved values");
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
        }
    }

    private static void TestFilters()
    {
        var settings = new UserPreferences(); settings.SetShowVeryHidden(true);
        using (var view = new SheetListControl(settings))
        {
            var rows = new[] {
                new WorksheetItem("a", "Sales 日本", true),
                new WorksheetItem("b", "sales chart", WorksheetVisibility.Hidden, null, SheetKind.Chart),
                new WorksheetItem("c", "Other", true),
                new WorksheetItem("d", "sales secret", WorksheetVisibility.VeryHidden, null) };
            view.ShowWorksheets("Book", rows, "a");
            var list = view.Controls.OfType<ReorderableSheetListBox>().Single();
            InvokeControl(list, "OnKeyDown", new KeyEventArgs(Keys.Control | Keys.A));
            view.SetFilter("SALES", null, null);
            Check(view.SelectedSheetIds.SequenceEqual(new[] { "a", "b", "d" }) && view.IsFiltered, "search preserves tab order and prunes excluded selection");
            view.SetFilter("", WorksheetVisibility.Hidden, SheetKind.Chart);
            Check(list.Items.Count == 1 && ((WorksheetItem)list.Items[0]).Id == "b", "combined visibility and chart filters");
            view.SetFilter("日本", null, SheetKind.Worksheet);
            Check(list.Items.Count == 1 && ((WorksheetItem)list.Items[0]).Id == "a", "Japanese substring search");
            int drags = 0;
            view.DragStarting += (s, e) => drags++;
            InvokeControl(list, "OnMouseDown", new MouseEventArgs(MouseButtons.Left, 1, 8, 8, 0));
            InvokeControl(list, "OnMouseMove", new MouseEventArgs(MouseButtons.Left, 0, 100, 100, 0));
            Check(drags == 0 && !list.IsDragging, "filtered single row cannot initiate a drag");
            InvokeControl(list, "OnMouseUp", new MouseEventArgs(MouseButtons.Left, 1, 100, 100, 0));
            view.SetFilter("absent", null, null);
            Check(list.Items.Count == 0 && view.SelectedSheetIds.Length == 0, "empty search has no operation targets");
            view.SetFilter("", null, null);
            Check(list.Items.Count == 4 && view.SelectedSheetIds.Length == 0 && !view.IsFiltered, "clear filters does not resurrect selection");
            view.SetFilter("new", null, null);
            rows[2] = new WorksheetItem("c", "New title", true);
            view.ShowWorksheets("Book", rows, "a");
            Check(list.Items.Count == 1 && ((WorksheetItem)list.Items[0]).Id == "c", "external rename updates filtered snapshot");
            settings.SetShowVeryHidden(false);
            view.SetFilter("", WorksheetVisibility.VeryHidden, null);
            Check(list.Items.Count == 0, "VeryHidden filter cannot bypass opt-in");
            view.SetFilter("sales", null, null);
            Check(!view.HasFacetFilter && view.IsFiltered, "text search is independent of facet indicator");
            view.Size = new System.Drawing.Size(190, 300);
            view.ToggleFilterPopup();
            Check(view.FilterPopup.Right <= view.ClientSize.Width && view.FilterPopup.Bottom <= view.ClientSize.Height, "popup stays inside narrow pane");
            var options = view.FilterPopup.Controls.OfType<FlowLayoutPanel>().Single().Controls.OfType<Button>().ToArray();
            Check(options.All(b => !b.Text.Contains("VeryHidden")), "VeryHidden not offered when setting off");
            options.Single(b => b.Text.Contains("非表示")).PerformClick();
            Check(view.HasFacetFilter && list.Items.Count == 1, "popup applies state while preserving query");
            view.ToggleFilterPopup();
            Check(view.ClearFilterButton.Visible && !view.FilterPopup.Controls.OfType<FlowLayoutPanel>().Single().Controls.OfType<Button>().Any(b => b.Text == "フィルター解除"), "only external facet reset exists");
            view.HandlePopupPointer(view.PointToScreen(new System.Drawing.Point(-10, -10)));
            Check(!view.FilterPopup.Visible && view.HasFacetFilter, "outside click closes without resetting facets");
            view.ToggleFilterPopup(); view.ToggleFilterPopup();
            Check(!view.FilterPopup.Visible, "filter toggle closes popup");
            view.ToggleFilterPopup();
            var escape = Message.Create(IntPtr.Zero, 0x100, (IntPtr)Keys.Escape, IntPtr.Zero);
            typeof(SheetListControl).GetMethod("ProcessCmdKey", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).Invoke(view, new object[] { escape, Keys.Escape });
            Check(!view.FilterPopup.Visible && view.HasFacetFilter, "escape closes popup and retains facets");
            view.ClearFilterButton.PerformClick();
            Check(!view.ClearFilterButton.Visible, "reset hidden when facets cleared");
            Check(!view.HasFacetFilter && list.Items.Count == 2 && view.IsFiltered, "one-click facet reset keeps search");
            settings.SetShowVeryHidden(true);
            view.SetFilter("", WorksheetVisibility.VeryHidden, null);
            view.ToggleFilterPopup();
            settings.SetShowVeryHidden(false);
            Check(!view.FilterPopup.Visible && !view.HasFacetFilter && !list.Items.Cast<WorksheetItem>().Any(i => i.Visibility == WorksheetVisibility.VeryHidden), "setting off dismisses popup and removes obsolete VeryHidden facet");
        }
        using (var disposedView = new SheetListControl(settings))
        {
            disposedView.Dispose();
            settings.SetOptions(false, false);
            Check(disposedView.IsDisposed, "disposed view detaches shared settings event");
        }
    }
}
