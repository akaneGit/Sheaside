using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sheaside
{
    [ComVisible(true)]
    public sealed class SheasideRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            using (var stream = typeof(SheasideRibbon).Assembly.GetManifestResourceStream("Sheaside.SheasideRibbon.xml"))
            using (var reader = new StreamReader(stream)) return reader.ReadToEnd();
        }

        public void OnLoad(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
            Invalidate();
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            var manager = Globals.ThisAddIn.PaneManager;
            return manager != null && manager.IsVisible(control.Context as Excel.Window);
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.PaneManager != null && control.Context is Excel.Window;
        }

        public void OnToggle(Office.IRibbonControl control, bool pressed)
        {
            Globals.ThisAddIn.PaneManager?.SetVisible(control.Context as Excel.Window, pressed);
            Invalidate();
        }

        internal void Invalidate()
        {
            try { ribbon?.Invalidate(); }
            catch (COMException) { /* Office may already be closing its ribbon. */ }
        }
        public bool GetDisplayPressed(Office.IRibbonControl control)
        {
            var settings = Globals.ThisAddIn.PaneManager?.Preferences;
            return settings != null && (control.Id == "SheasideHidden" ? settings.ShowHidden : settings.ShowVeryHidden);
        }
        public void OnDisplayToggle(Office.IRibbonControl control, bool pressed)
        {
            var settings = Globals.ThisAddIn.PaneManager?.Preferences;
            if (settings == null) return;
            try
            {
                if (control.Id == "SheasideHidden") settings.SetShowHidden(pressed);
                else if (control.Id == "SheasideVeryHidden") settings.SetShowVeryHidden(pressed);
            }
            catch (System.Exception e) when (e is IOException || e is System.UnauthorizedAccessException)
            { System.Windows.Forms.MessageBox.Show("表示設定を保存できませんでした。\r\n" + e.Message, "Sheaside"); }
            finally { Invalidate(); }
        }
    }
}
