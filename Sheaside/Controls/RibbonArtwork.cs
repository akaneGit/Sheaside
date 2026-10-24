using System.Drawing;
using System.Drawing.Drawing2D;

namespace Sheaside.Controls
{
    // Original vector interpretation of the supplied stacked sheets / wave reference.
    internal static class RibbonArtwork
    {
        internal static Bitmap Create(int size)
        {
            var image = new Bitmap(size, size);
            using (var g = Graphics.FromImage(image))
            using (var green = new SolidBrush(Color.FromArgb(0, 82, 52)))
            using (var teal = new SolidBrush(Color.FromArgb(0, 174, 145)))
            using (var white = new SolidBrush(Color.White))
            using (var outline = new Pen(Color.FromArgb(0, 82, 52), 1.3F))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias; g.ScaleTransform(size / 32F, size / 32F);
                using (var back = Sheet(4, 3, 21, 21)) g.FillPath(green, back);
                if (size > 16) using (var middle = Sheet(7, 7, 21, 19)) { g.FillPath(teal, middle); g.DrawPath(Pens.White, middle); }
                using (var front = Sheet(11, 11, 19, 17)) { g.FillPath(white, front); g.DrawPath(outline, front); }
                using (var wave = new GraphicsPath())
                {
                    wave.AddBezier(1, 25, 8, 13, 17, 31, 23, 23);
                    wave.AddBezier(23, 23, 26, 18, 21, 17, 20, 21);
                    wave.AddBezier(20, 21, 23, 14, 33, 21, 30, 28);
                    wave.AddBezier(30, 28, 21, 32, 13, 28, 9, 26);
                    wave.AddBezier(9, 26, 6, 24, 3, 24, 1, 25); wave.CloseFigure();
                    using (var border = new Pen(Color.White, size <= 16 ? 1.2F : 1.5F)) { g.FillPath(teal, wave); g.DrawPath(border, wave); }
                }
            }
            return image;
        }
        private static GraphicsPath Sheet(float x, float y, float width, float height)
        {
            var path = new GraphicsPath();
            path.AddPolygon(new[] { new PointF(x,y+2), new PointF(x+2,y), new PointF(x+width-7,y), new PointF(x+width-5,y+3), new PointF(x+width-2,y+3), new PointF(x+width,y+5), new PointF(x+width,y+height), new PointF(x,y+height) });
            return path;
        }
    }
}
