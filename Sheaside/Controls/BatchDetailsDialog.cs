using System.Drawing;
using System.Windows.Forms;

namespace Sheaside.Controls
{
    internal sealed class BatchDetailsDialog : Form
    {
        public BatchDetailsDialog(string title, string message, string details, bool confirmation)
        {
            Text = "Sheaside — " + title;
            Font = new Font("Yu Gothic UI", 9F);
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(420, 300);
            Size = new Size(560, 450);
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            Padding = new Padding(12);
            var label = new Label { Text = message, Dock = DockStyle.Top, Height = 66, AutoEllipsis = true };
            var names = new TextBox { Text = details, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
                WordWrap = false, Dock = DockStyle.Fill, AccessibleName = confirmation ? "操作対象一覧" : "処理結果一覧" };
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft };
            var cancel = new Button { Text = confirmation ? "キャンセル" : "閉じる", DialogResult = DialogResult.Cancel, AutoSize = true };
            buttons.Controls.Add(cancel);
            if (confirmation)
                buttons.Controls.Add(new Button { Text = "実行する", DialogResult = DialogResult.OK, AutoSize = true });
            CancelButton = cancel;
            AcceptButton = cancel; // Enter must not accidentally confirm a destructive batch.
            Controls.Add(names);
            Controls.Add(label);
            Controls.Add(buttons);
        }
    }
}
