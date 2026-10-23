using System.Drawing;
using System.Windows.Forms;
using Sheaside.Models;

namespace Sheaside.Controls
{
    internal sealed class ColorPaletteDialog : Form
    {
        public TabColorChoice Selection { get; private set; }
        public ColorPaletteDialog(int[] colors)
        {
            Text = "Sheaside — タブの色"; Font = new Font("Yu Gothic UI", 9F);
            AutoScaleMode = AutoScaleMode.Font; ClientSize = new Size(320, 330);
            StartPosition = FormStartPosition.CenterParent; FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false; ShowInTaskbar = false; Padding = new Padding(10);
            var layout = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
            layout.Controls.Add(new Label { Text = colors == null ? "テーマの色（取得不可：代替RGB）" : "テーマの色", AutoSize = true });
            var grid = new TableLayoutPanel { Width = 290, Height = 150, ColumnCount = 10, RowCount = 6 };
            for (int c = 0; c < 10; c++) grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10));
            for (int r = 0; r < 6; r++) grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100F/6));
            var palette = TabColorChoice.Palette(colors, colors != null);
            for (int i = 0; i < palette.Count; i++) grid.Controls.Add(ColorButton(palette[i]), i % 10, i / 10);
            layout.Controls.Add(grid);
            layout.Controls.Add(new Label { Text = "標準の色", AutoSize = true });
            var standard = new FlowLayoutPanel { Width = 290, Height = 28, WrapContents = false };
            foreach (Color color in new[] { Color.DarkRed, Color.Red, Color.Orange, Color.Yellow, Color.LightGreen, Color.Green, Color.Cyan, Color.Blue, Color.DarkBlue, Color.Purple })
            {
                var button = ColorButton(TabColorChoice.Fixed(ColorTranslator.ToOle(color), color.Name));
                button.Size = new Size(26, 24); standard.Controls.Add(button);
            }
            layout.Controls.Add(standard);
            var none = new Button { Text = "色なし", Width = 290 };
            none.Click += (s, e) => Choose(TabColorChoice.Fixed(null)); layout.Controls.Add(none);
            var other = new Button { Text = "その他の色…", Width = 290 };
            other.Click += (s, e) =>
            {
                using (var dialog = new ColorDialog { FullOpen = true })
                    if (dialog.ShowDialog(this) == DialogResult.OK) Choose(TabColorChoice.Fixed(ColorTranslator.ToOle(dialog.Color)));
            };
            layout.Controls.Add(other);
            var cancel = new Button { Text = "キャンセル", DialogResult = DialogResult.Cancel, Width = 290 };
            layout.Controls.Add(cancel); CancelButton = cancel;
            Controls.Add(layout);
        }
        private Button ColorButton(TabColorChoice choice)
        {
            var button = new Button { BackColor = ColorTranslator.FromOle(choice.Rgb.Value), Dock = DockStyle.Fill,
                Margin = new Padding(1), FlatStyle = FlatStyle.Flat, AccessibleName = choice.Name, UseVisualStyleBackColor = false };
            button.Click += (s, e) => Choose(choice);
            return button;
        }
        private void Choose(TabColorChoice choice) { Selection = choice; DialogResult = DialogResult.OK; Close(); }
    }
}
