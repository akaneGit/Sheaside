using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
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
        private string activeWorksheetId;
        private bool blockDragGesture;
        private bool dispatchingPointer;
        public Keys PointerModifiers { get; private set; }
        public bool IsSelecting { get; private set; }

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);

        public int AnchorIndex
        {
            get { return IsHandleCreated ? SendMessage(Handle, 0x019D, IntPtr.Zero, IntPtr.Zero).ToInt32() : SelectedIndex; }
            set { if (value >= 0 && value < Items.Count) SendMessage(Handle, 0x019C, new IntPtr(value), IntPtr.Zero); }
        }

        public ReorderableSheetListBox()
        {
            AllowDrop = true;
            SelectionMode = SelectionMode.MultiExtended;
            DrawMode = DrawMode.OwnerDrawFixed;
            ItemHeight = Font.Height + 10;
            scrollTimer = new Timer { Interval = 120 };
            scrollTimer.Tick += ScrollDuringDrag;
        }

        public bool IsDragging { get; private set; }
        public void SetActiveWorksheet(string id)
        {
            if (activeWorksheetId == id) return;
            string previous = activeWorksheetId;
            activeWorksheetId = id;
            for (int i = 0; i < Items.Count; i++)
                if (((WorksheetItem)Items[i]).Id == id || ((WorksheetItem)Items[i]).Id == previous)
                    Invalidate(GetItemRectangle(i));
        }

        public void UpdateTextExtent()
        {
            HorizontalExtent = Items.Cast<WorksheetItem>().Select(item =>
                TextRenderer.MeasureText("● " + item.DisplayName, Font).Width + 28).DefaultIfEmpty(0).Max();
        }

        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            ItemHeight = Font.Height + 10;
            UpdateTextExtent();
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= Items.Count) return;
            var item = (WorksheetItem)Items[e.Index];
            e.DrawBackground();
            bool selected = (e.State & DrawItemState.Selected) != 0;
            Color text = selected ? SystemColors.HighlightText : SystemColors.WindowText;
            var swatch = new Rectangle(e.Bounds.Left + 5, e.Bounds.Top + 5, 10, Math.Max(5, e.Bounds.Height - 10));
            using (var brush = new SolidBrush(item.TabColor.HasValue ? ColorTranslator.FromOle(item.TabColor.Value) : SystemColors.Window))
                e.Graphics.FillRectangle(brush, swatch);
            using (var pen = new Pen(selected ? SystemColors.HighlightText : SystemColors.WindowText))
                e.Graphics.DrawRectangle(pen, swatch);
            string caption = (item.Id == activeWorksheetId ? "● " : "　 ") + item.DisplayName;
            var bounds = new Rectangle(e.Bounds.Left + 22, e.Bounds.Top, Math.Max(0, e.Bounds.Width - 22), e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, caption, Font, bounds, text,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);
            e.DrawFocusRectangle();
        }
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
            if (!dispatchingPointer)
            {
                PointerModifiers = ModifierKeys;
                blockDragGesture = SelectedItems.Count > 1 || (PointerModifiers & (Keys.Control | Keys.Shift)) != 0;
            }
            if (e.Button == MouseButtons.Right)
            {
                int hit = IndexFromPoint(e.Location);
                if (hit < 0 || !GetSelected(hit))
                {
                    ClearSelected();
                    if (hit >= 0) { SetSelected(hit, true); AnchorIndex = hit; }
                }
            }
            if (e.Button == MouseButtons.Left && e.Clicks == 1 && !blockDragGesture)
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
            if (e.Button != MouseButtons.Left || mouseDownItem == null || IsDragging || SelectedItems.Count > 1) return;
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
            if (!consumedClick && (PointerModifiers & (Keys.Control | Keys.Shift)) == 0) base.OnMouseClick(e);
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            if (!consumedClick && SelectedItems.Count == 1 && (PointerModifiers & (Keys.Control | Keys.Shift)) == 0) base.OnMouseDoubleClick(e);
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
            if (m.Msg == 0x0202 || m.Msg == 0x0215) IsSelecting = false;
            if (m.Msg == 0x0204) // Native right-down must not discard a multi-selection.
            {
                Focus();
                int packed = m.LParam.ToInt32();
                OnMouseDown(new MouseEventArgs(MouseButtons.Right, 1, (short)packed, (short)(packed >> 16), 0));
                return;
            }
            if (m.Msg == 0x0201 || m.Msg == 0x0203)
            {
                IsSelecting = true;
                PointerModifiers = ModifierKeys;
                blockDragGesture = SelectedItems.Count > 1 || (PointerModifiers & (Keys.Control | Keys.Shift)) != 0;
                dispatchingPointer = true;
                try { base.WndProc(ref m); }
                finally { dispatchingPointer = false; }
                return;
            }
            base.WndProc(ref m);
            // ListBox is a native control: paint the marker after its WM_PAINT,
            // including when the insertion point is in the empty client area.
            if (m.Msg != 0x000F || insertionBoundary < 0) return;
            int y = LineY(insertionBoundary);
            if (y < 0) return;
            using (var graphics = Graphics.FromHwnd(Handle))
            using (var pen = new Pen(SystemColors.Highlight, 2))
            using (var outline = new Pen(SystemColors.Window, 4))
            {
                graphics.DrawLine(outline, 2, y, ClientSize.Width - 3, y);
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
