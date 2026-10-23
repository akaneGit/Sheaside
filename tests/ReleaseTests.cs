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
}
