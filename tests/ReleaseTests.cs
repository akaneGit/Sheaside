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
            File.WriteAllText(path, "invalid xml");
            var corrupt = new UserPreferences(path);
            Check(!corrupt.ShowVeryHidden && corrupt.LoadError != null, "corrupt settings fall back safely");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
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
        }
    }
}
