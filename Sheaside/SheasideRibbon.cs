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
        private readonly object brandImage = CreateBrandImage();
        private static object CreateBrandImage()
        {
            using (var image = Controls.RibbonArtwork.Create(32)) return PictureConverter.Convert(image);
        }
        public object GetBrandImage(Office.IRibbonControl control) { return brandImage; }
        public object GetStateImage(Office.IRibbonControl control)
        {
            var theme = Services.OfficeTheme.System;
            try { theme = Services.OfficeThemeReader.Read(Globals.ThisAddIn.Application.Version); }
            catch (COMException) { }
            using (var image = new System.Drawing.Bitmap(16, 16))
            {
                using (var graphics = System.Drawing.Graphics.FromImage(image))
                    Controls.ModernButton.DrawIcon(graphics, control.Id == "SheasideHidden" ? Controls.LineIcon.Hidden : Controls.LineIcon.Locked,
                        new System.Drawing.Rectangle(0, 0, 16, 16), Controls.UiAppearance.Foreground(theme));
                return PictureConverter.Convert(image);
            }
        }
        private sealed class PictureConverter : System.Windows.Forms.AxHost
        {
            private PictureConverter() : base("") { }
            internal static object Convert(System.Drawing.Image image) { return GetIPictureDispFromPicture(image); }
        }

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
