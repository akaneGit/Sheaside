using System;
using System.Drawing;
using System.Windows.Forms;

namespace Sheaside.Controls
{
    internal sealed class RenameSheetDialog : Form
    {
        public RenameSheetDialog(string currentName, Func<string, string> tryRename)
        {
            Text = "シートの名前変更";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(420, 175);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            var label = new Label { Text = "新しいシート名", AutoSize = true, Location = new Point(12, 12) };
            var nameBox = new TextBox { Text = currentName, Location = new Point(12, 36), Width = 396, AccessibleName = "新しいシート名" };
            var errorLabel = new Label { Location = new Point(12, 65), Size = new Size(396, 64), ForeColor = Color.Firebrick };
            var ok = new Button { Text = "変更", Location = new Point(246, 136), Width = 78 };
            var cancel = new Button { Text = "キャンセル", Location = new Point(330, 136), Width = 78, DialogResult = DialogResult.Cancel };
            ok.Click += (sender, args) =>
            {
                // The controller applies the rename and returns a user-facing error.
                // Keep the dialog/input open on validation or Excel errors.
                var error = tryRename(nameBox.Text);
                if (error == null) DialogResult = DialogResult.OK;
                else { errorLabel.Text = error; nameBox.Focus(); }
            };
            Controls.AddRange(new Control[] { label, nameBox, errorLabel, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;
            Shown += (sender, args) => { nameBox.Focus(); nameBox.SelectAll(); };
        }
    }
}
