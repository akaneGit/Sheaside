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
            try { ribbon?.InvalidateControl("SheasideToggle"); }
            catch (COMException) { /* Office may already be closing its ribbon. */ }
        }
    }
}
