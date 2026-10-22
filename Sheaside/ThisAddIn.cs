using System;
using Sheaside.Services;
using Office = Microsoft.Office.Core;

namespace Sheaside
{
    public partial class ThisAddIn
    {
        private SheasideRibbon ribbon;
        internal SheetPaneManager PaneManager { get; private set; }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new SheasideRibbon();
            return ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            PaneManager = new SheetPaneManager(Application, CustomTaskPanes);
            PaneManager.StateChanged += OnPaneStateChanged;
            PaneManager.Start();
            ribbon?.Invalidate();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (PaneManager == null) return;
            PaneManager.StateChanged -= OnPaneStateChanged;
            PaneManager.Dispose();
            PaneManager = null;
        }

        private void OnPaneStateChanged(object sender, EventArgs e) { ribbon?.Invalidate(); }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
