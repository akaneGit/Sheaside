using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Sheaside.Models;
using Sheaside.Services;

namespace Sheaside.Controls
{
    internal sealed class ReorderableSheetListBox : ListBox
    {
        private const string DragFormat = "Sheaside.LocalWorksheetMove";
        private readonly Timer scrollTimer;
        private Point mouseDownPoint;
        private WorksheetItem mouseDownItem;
        private string dragToken;
        private int sourceIndex;
        private int insertionBoundary = -1;
        private bool cancelled;
        private bool consumedClick;

        public ReorderableSheetListBox()
        {
            AllowDrop = true;
            scrollTimer = new Timer { Interval = 120 };
            scrollTimer.Tick += ScrollDuringDrag;
        }

        public bool IsDragging { get; private set; }
        public event EventHandler<WorksheetDragStartingEventArgs> DragStarting;
        public event EventHandler<WorksheetDropEventArgs> DropRequested;
        public event EventHandler DragCompleted;

        public void CancelDrag()
        {
            cancelled = true;
            mouseDownItem = null;
            SetInsertionBoundary(-1);
            scrollTimer.Stop();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            consumedClick = false;
            mouseDownItem = null;
            if (e.Button == MouseButtons.Left && e.Clicks == 1)
            {
                int index = IndexFromPoint(e.Location);
                if (index != NoMatches) mouseDownItem = Items[index] as WorksheetItem;
                mouseDownPoint = e.Location;
            }
            base.OnMouseDown(e);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (e.Button != MouseButtons.Left || mouseDownItem == null || IsDragging) return;
            Size threshold = SystemInformation.DragSize;
            var deadZone = new Rectangle(mouseDownPoint.X - threshold.Width / 2,
                mouseDownPoint.Y - threshold.Height / 2, threshold.Width, threshold.Height);
            if (deadZone.Contains(e.Location)) return;
            sourceIndex = Items.IndexOf(mouseDownItem);
            if (sourceIndex < 0) { mouseDownItem = null; return; }

            consumedClick = true;
            IsDragging = true;
            cancelled = false;
            try
            {
                var request = new WorksheetDragStartingEventArgs(mouseDownItem.Id,
                    Items.Cast<WorksheetItem>().Select(item => item.Id).ToArray());
                DragStarting?.Invoke(this, request);
                if (request.Cancel || cancelled) return;
                // A string survives OLE's data marshaling; object identity does not.
                // Only the originating control owns this per-drag token.
                dragToken = Guid.NewGuid().ToString("N");
                var data = new DataObject();
                data.SetData(DragFormat, false, dragToken);
                scrollTimer.Start();
                DoDragDrop(data, DragDropEffects.Move);
            }
            finally
            {
                scrollTimer.Stop();
                SetInsertionBoundary(-1);
                dragToken = null;
                mouseDownItem = null;
                IsDragging = false;
                DragCompleted?.Invoke(this, EventArgs.Empty);
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            mouseDownItem = null;
            base.OnMouseUp(e);
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            if (!consumedClick) base.OnMouseClick(e);
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            if (!consumedClick) base.OnMouseDoubleClick(e);
        }

        protected override void OnQueryContinueDrag(QueryContinueDragEventArgs e)
        {
            if (cancelled || e.EscapePressed || (e.KeyState & 2) != 0) e.Action = DragAction.Cancel;
            base.OnQueryContinueDrag(e);
        }

        private bool IsLocalMove(IDataObject data)
        {
            return IsDragging && !cancelled && dragToken != null && data != null &&
                data.GetDataPresent(DragFormat, false) && string.Equals(data.GetData(DragFormat, false) as string, dragToken, StringComparison.Ordinal);
        }

        private int BoundaryAt(Point point)
        {
            if (!ClientRectangle.Contains(point)) return -1;
            int hit = IndexFromPoint(point);
            Rectangle row = hit < 0 ? Rectangle.Empty : GetItemRectangle(hit);
            return WorksheetMoveRules.BoundaryAt(hit, point.Y, row.Top, row.Height, Items.Count);
        }

        private bool ChangesOrder(int boundary)
        {
            return boundary >= 0 && WorksheetMoveRules.DestinationIndex(sourceIndex, boundary, Items.Count) != sourceIndex;
        }

        protected override void OnDragEnter(DragEventArgs e) { UpdateDrag(e); base.OnDragEnter(e); }
        protected override void OnDragOver(DragEventArgs e) { UpdateDrag(e); base.OnDragOver(e); }

        private void UpdateDrag(DragEventArgs e)
        {
            int boundary = IsLocalMove(e.Data) ? BoundaryAt(PointToClient(new Point(e.X, e.Y))) : -1;
            bool allowed = ChangesOrder(boundary) && (e.AllowedEffect & DragDropEffects.Move) != 0;
            SetInsertionBoundary(allowed ? boundary : -1);
            e.Effect = allowed ? DragDropEffects.Move : DragDropEffects.None;
        }

        protected override void OnDragLeave(EventArgs e)
        {
            SetInsertionBoundary(-1);
            base.OnDragLeave(e);
        }

        protected override void OnDragDrop(DragEventArgs e)
        {
            scrollTimer.Stop();
            UpdateDrag(e);
            int boundary = insertionBoundary;
            SetInsertionBoundary(-1);
            if (boundary >= 0) DropRequested?.Invoke(this, new WorksheetDropEventArgs(boundary));
            base.OnDragDrop(e);
        }

        private void ScrollDuringDrag(object sender, EventArgs e)
        {
            if (!IsDragging || cancelled || Items.Count == 0) return;
            Point point = PointToClient(MousePosition);
            if (!ClientRectangle.Contains(point)) return;
            int margin = Math.Min(ItemHeight, ClientSize.Height / 3);
            int next = TopIndex;
            if (point.Y < margin) next = Math.Max(0, next - 1);
            else if (point.Y >= ClientSize.Height - margin) next = Math.Min(Items.Count - 1, next + 1);
            if (next != TopIndex) { TopIndex = next; Invalidate(); }
            int boundary = BoundaryAt(point);
            SetInsertionBoundary(ChangesOrder(boundary) ? boundary : -1);
        }

        private int LineY(int boundary)
        {
            if (boundary < 0 || Items.Count == 0) return -1;
            int y = boundary < Items.Count ? GetItemRectangle(boundary).Top : GetItemRectangle(Items.Count - 1).Bottom;
            return y < 0 || y > ClientSize.Height ? -1 : Math.Max(1, Math.Min(ClientSize.Height - 2, y));
        }

        private void SetInsertionBoundary(int boundary)
        {
            if (insertionBoundary == boundary) return;
            int oldY = LineY(insertionBoundary);
            insertionBoundary = boundary;
            int newY = LineY(boundary);
            if (oldY >= 0) Invalidate(new Rectangle(0, oldY - 4, ClientSize.Width, 9));
            if (newY >= 0) Invalidate(new Rectangle(0, newY - 4, ClientSize.Width, 9));
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            // ListBox is a native control: paint the marker after its WM_PAINT,
            // including when the insertion point is in the empty client area.
            if (m.Msg != 0x000F || insertionBoundary < 0) return;
            int y = LineY(insertionBoundary);
            if (y < 0) return;
            using (var graphics = Graphics.FromHwnd(Handle))
            using (var pen = new Pen(SystemColors.Highlight, 2))
            {
                graphics.DrawLine(pen, 2, y, ClientSize.Width - 3, y);
                graphics.DrawLine(pen, 2, y - 3, 2, y + 3);
                graphics.DrawLine(pen, ClientSize.Width - 3, y - 3, ClientSize.Width - 3, y + 3);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { CancelDrag(); scrollTimer.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
