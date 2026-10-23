using System.Drawing;
using System.Windows.Forms;
using Sheaside.Services;

namespace Sheaside.Controls
{
    internal sealed class PreferencesDialog : Form
    {
        private readonly CheckBox veryHidden, startup;
        private readonly ComboBox theme, position;
        public bool ShowVeryHidden { get { return veryHidden.Checked; } }
        public bool StartupVisible { get { return startup.Checked; } }
        public PaneTheme Theme { get { return (PaneTheme)theme.SelectedIndex; } }
        public bool DockLeft { get { return position.SelectedIndex == 1; } }

        public PreferencesDialog(UserPreferences settings)
        {
            Text = "Sheaside — 設定";
            Font = new Font("Yu Gothic UI", 9F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(430, 260);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = MinimizeBox = false;
            ShowInTaskbar = false;
            Padding = new Padding(12);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 5 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            veryHidden = new CheckBox { Text = "VeryHiddenシートを一覧に表示する", Checked = settings.ShowVeryHidden, Dock = DockStyle.Fill, AutoSize = true };
            layout.Controls.Add(veryHidden, 0, 0); layout.SetColumnSpan(veryHidden, 2);
            startup = new CheckBox { Text = "新しいウィンドウでパネルを表示する", Checked = settings.StartupVisible, Dock = DockStyle.Fill, AutoSize = true };
            layout.Controls.Add(startup, 0, 1); layout.SetColumnSpan(startup, 2);
            theme = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, AccessibleName = "テーマ" };
            theme.Items.AddRange(new object[] { "システム", "ライト", "ダーク" }); theme.SelectedIndex = (int)settings.Theme;
            position = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, AccessibleName = "パネル位置" };
            position.Items.AddRange(new object[] { "右", "左" }); position.SelectedIndex = settings.DockLeft ? 1 : 0;
            layout.Controls.Add(new Label { Text = "テーマ", AutoSize = true }, 0, 2); layout.Controls.Add(theme, 1, 2);
            layout.Controls.Add(new Label { Text = "パネル位置", AutoSize = true }, 0, 3); layout.Controls.Add(position, 1, 3);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true };
            var cancel = new Button { Text = "キャンセル", AutoSize = true, DialogResult = DialogResult.Cancel };
            var save = new Button { Text = "保存", AutoSize = true, DialogResult = DialogResult.OK };
            buttons.Controls.Add(cancel); buttons.Controls.Add(save);
            layout.Controls.Add(buttons, 0, 4); layout.SetColumnSpan(buttons, 2);
            AcceptButton = save; CancelButton = cancel;
            Controls.Add(layout);
        }
    }
}
