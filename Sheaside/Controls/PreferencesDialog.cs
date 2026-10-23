using System.Drawing;
using System.Windows.Forms;
using Sheaside.Services;

namespace Sheaside.Controls
{
    internal sealed class PreferencesDialog : Form
    {
        private readonly CheckBox veryHidden, startup;
        public bool ShowVeryHidden { get { return veryHidden.Checked; } }
        public bool StartupVisible { get { return startup.Checked; } }

        public PreferencesDialog(UserPreferences settings, OfficeTheme theme = OfficeTheme.System)
        {
            Text = "Sheaside — 設定";
            Font = new Font("Yu Gothic UI", 9F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(460, 280);
            MinimumSize = new Size(476, 0);
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = MinimizeBox = false;
            ShowInTaskbar = false;
            Padding = new Padding(16);
            var layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, MinimumSize = new Size(428, 0), ColumnCount = 1, RowCount = 5 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 5; i++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            veryHidden = new CheckBox { Text = "VeryHiddenシートを一覧に表示する", Checked = settings.ShowVeryHidden, Dock = DockStyle.Top, AutoSize = true, Margin = new Padding(0, 0, 0, 6) };
            layout.Controls.Add(veryHidden, 0, 0);
            layout.Controls.Add(Description("Excelの通常の「再表示」には出ないシートです。ONにすると一覧で表示し、表示状態や色を管理できます。"), 0, 1);
            startup = new CheckBox { Text = "新しいウィンドウでパネルを表示する", Checked = settings.StartupVisible, Dock = DockStyle.Top, AutoSize = true, Margin = new Padding(0, 0, 0, 6) };
            layout.Controls.Add(startup, 0, 2);
            layout.Controls.Add(Description("Excelの［表示］→［新しいウィンドウを開く］で作成する別ウィンドウや、新しく開いたブックでの初期表示です。現在のパネル表示は変更しません。"), 0, 3);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, WrapContents = false, Margin = new Padding(0, 4, 0, 0) };
            var cancel = new ModernButton { Text = "キャンセル", AutoSize = true, DialogResult = DialogResult.Cancel };
            var save = new ModernButton { Text = "保存", AutoSize = true, DialogResult = DialogResult.OK };
            buttons.Controls.Add(cancel); buttons.Controls.Add(save);
            layout.Controls.Add(buttons, 0, 4);
            AcceptButton = save; CancelButton = cancel;
            Controls.Add(layout);
            UiAppearance.Apply(this, theme);
        }
        private static Label Description(string text)
        {
            return new Label { Text = text, AutoSize = true, MaximumSize = new Size(404, 0), Margin = new Padding(24, 0, 0, 18) };
        }
    }
}
