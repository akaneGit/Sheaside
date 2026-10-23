using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Sheaside.Controls
{
    internal enum LineIcon { None, Add, Delete, Search, Filter, Settings }

    // Original geometric artwork, drawn in a 16-unit coordinate system; no external font/assets.
    internal sealed class ModernButton : Button
    {
        internal LineIcon Icon { get; set; }
        private bool hovered, selected;
        internal bool Selected { get { return selected; } set { selected = value; Invalidate(); } }
        internal bool IsHovered { get { return hovered; } }
        public ModernButton()
        {
            FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0;
            UseVisualStyleBackColor = false;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Margin = new Padding(3);
        }
        protected override void OnMouseEnter(EventArgs e) { hovered = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hovered = false; Invalidate(); base.OnMouseLeave(e); }
        public override Size GetPreferredSize(Size proposedSize)
        {
            var text = TextRenderer.MeasureText(Text, Font);
            return new Size(text.Width + (Icon == LineIcon.None ? 24 : 48), Math.Max(Font.Height + 16, 32));
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            if (Width < 3 || Height < 3) return;
            var g = e.Graphics;
            float scale = g.DpiX / 96F;
            bool dark = BackColor.GetBrightness() < .5;
            Color accent = SystemInformation.HighContrast ? SystemColors.Highlight : dark ? Color.FromArgb(126, 211, 164) : Color.FromArgb(16, 110, 62);
            Color foreground = !Enabled ? SystemColors.GrayText : Selected ? (SystemInformation.HighContrast ? SystemColors.HighlightText : accent) : ForeColor;
            Color fill = Selected ? (SystemInformation.HighContrast ? SystemColors.Highlight : dark ? Color.FromArgb(39, 65, 48) : Color.FromArgb(226, 242, 232)) :
                hovered && Enabled ? (dark ? ControlPaint.Light(BackColor, .15F) : ControlPaint.Dark(BackColor, .05F)) : BackColor;
            g.Clear(BackColor); g.SmoothingMode = SmoothingMode.AntiAlias;
            float radius = Math.Min(4 * scale, Math.Min(Width, Height) / 2F);
            using (var path = new GraphicsPath())
            using (var brush = new SolidBrush(fill))
            using (var pen = new Pen(Selected ? accent : ControlPaint.Dark(BackColor, .12F)))
            {
                path.AddArc(1, 1, radius, radius, 180, 90);
                path.AddArc(Width - radius - 2, 1, radius, radius, 270, 90);
                path.AddArc(Width - radius - 2, Height - radius - 2, radius, radius, 0, 90);
                path.AddArc(1, Height - radius - 2, radius, radius, 90, 90); path.CloseFigure();
                g.FillPath(brush, path);
                if (Selected || hovered || Focused) g.DrawPath(pen, path);
            }
            int iconSize = (int)Math.Round(16 * scale), gap = (int)Math.Round(6 * scale);
            int textWidth = string.IsNullOrEmpty(Text) ? 0 : TextRenderer.MeasureText(g, Text, Font).Width;
            int contentWidth = Math.Min(Width - 12, textWidth + (Icon == LineIcon.None ? 0 : iconSize + (textWidth > 0 ? gap : 0)));
            int left = Math.Max(4, (Width - contentWidth) / 2);
            if (Icon != LineIcon.None)
            {
                DrawIcon(g, Icon, new Rectangle(left, (Height - iconSize) / 2, iconSize, iconSize), foreground);
                left += iconSize + gap;
            }
            if (textWidth > 0)
                TextRenderer.DrawText(g, Text, Font, new Rectangle(left, 0, Math.Max(1, Width - left - 4), Height), foreground,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
            if (Focused && ShowFocusCues) ControlPaint.DrawFocusRectangle(g, Rectangle.Inflate(ClientRectangle, -4, -4), foreground, fill);
        }
        internal static void DrawIcon(Graphics g, LineIcon icon, Rectangle bounds, Color color)
        {
            var state = g.Save();
            try
            {
                g.TranslateTransform(bounds.X, bounds.Y); g.ScaleTransform(bounds.Width / 16F, bounds.Height / 16F);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var pen = new Pen(color, 1.3F) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                {
                    switch (icon)
                    {
                        case LineIcon.Add: g.DrawLine(pen, 8, 2, 8, 14); g.DrawLine(pen, 2, 8, 14, 8); break;
                        case LineIcon.Delete:
                            g.DrawLine(pen, 2, 4, 14, 4); g.DrawLines(pen, new[] { new Point(4, 4), new Point(5, 14), new Point(11, 14), new Point(12, 4) });
                            g.DrawRectangle(pen, 6, 1, 4, 3); g.DrawLine(pen, 7, 6, 7, 11); g.DrawLine(pen, 9, 6, 9, 11); break;
                        case LineIcon.Search: g.DrawEllipse(pen, 2, 2, 8, 8); g.DrawLine(pen, 9, 9, 14, 14); break;
                        case LineIcon.Filter: g.DrawPolygon(pen, new[] { new Point(1, 2), new Point(15, 2), new Point(10, 8), new Point(10, 14), new Point(6, 12), new Point(6, 8) }); break;
                        case LineIcon.Settings:
                            g.DrawEllipse(pen, 4, 4, 8, 8); g.DrawEllipse(pen, 6, 6, 4, 4);
                            for (int i = 0; i < 8; i++) { double a = i * Math.PI / 4; g.DrawLine(pen, (float)(8 + 4 * Math.Cos(a)), (float)(8 + 4 * Math.Sin(a)), (float)(8 + 6.5 * Math.Cos(a)), (float)(8 + 6.5 * Math.Sin(a))); } break;
                    }
                }
            }
            finally { g.Restore(state); }
        }
    }
}
