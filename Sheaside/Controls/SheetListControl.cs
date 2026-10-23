using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Sheaside.Models;

namespace Sheaside.Controls
{
    internal sealed class SheetListControl : UserControl
    {
        private readonly Label workbookLabel;
        private readonly ReorderableSheetListBox sheetList;
        private readonly Label statusLabel;
        private readonly Label selectionLabel;
        private readonly Button addButton;
        private readonly SheetSelectionState selection = new SheetSelectionState();
        private readonly ContextMenuStrip sheetMenu;
        private WorksheetItem menuTarget;
        private string[] batchMenuTargets = new string[0];
        private bool updating;

        public SheetListControl()
        {
            AutoScaleMode = AutoScaleMode.Font;
            Font = new Font("Yu Gothic UI", 9F);
            Padding = new Padding(8);
            workbookLabel = new Label { Dock = DockStyle.Top, Height = 32, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
            statusLabel = new Label { Dock = DockStyle.Bottom, Height = 64, AutoEllipsis = true };
            selectionLabel = new Label { Dock = DockStyle.Bottom, Height = 26, AutoEllipsis = true };
            sheetList = new ReorderableSheetListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                HorizontalScrollbar = true,
                DisplayMember = "DisplayName",
                AccessibleName = "シート一覧"
            };
            // MouseClick also handles clicking an already selected row (e.g. after a failed activation).
            sheetList.MouseClick += OnSheetClicked;
            sheetList.KeyDown += OnSheetKeyDown;
            sheetList.MouseDoubleClick += (sender, args) =>
            {
                int index = sheetList.IndexFromPoint(args.Location);
                if (args.Button == MouseButtons.Left && index != ListBox.NoMatches && sheetList.SelectedItems.Count == 1)
                    RequestCommand(WorksheetCommand.Rename, (WorksheetItem)sheetList.Items[index]);
            };
            sheetMenu = new ContextMenuStrip();
            sheetMenu.Items.Add("シートの追加", null, (sender, args) => RequestCommand(WorksheetCommand.Add, menuTarget));
            sheetMenu.Items.Add("名前の変更", null, (sender, args) => RequestCommand(WorksheetCommand.Rename, menuTarget));
            sheetMenu.Items.Add("シートの複製", null, (sender, args) => RequestCommand(WorksheetCommand.Duplicate, menuTarget));
            sheetMenu.Items.Add("シートの削除", null, (sender, args) => RequestCommand(WorksheetCommand.Delete, menuTarget));
            var hide = sheetMenu.Items.Add("非表示にする", null, (sender, args) => RequestCommand(WorksheetCommand.Hide, menuTarget));
            var unhide = sheetMenu.Items.Add("再表示", null, (sender, args) => RequestCommand(WorksheetCommand.Unhide, menuTarget));
            var reveal = sheetMenu.Items.Add("VeryHiddenを再表示…", null, (sender, args) => RequestCommand(WorksheetCommand.RevealVeryHidden, menuTarget));
            var colors = new ToolStripMenuItem("タブの色");
            colors.DropDownItems.Add("色を選択…", null, (sender, args) => RequestCommand(WorksheetCommand.SetTabColor, menuTarget));
            colors.DropDownItems.Add("色なし", null, (sender, args) => RequestCommand(WorksheetCommand.ClearTabColor, menuTarget));
            sheetMenu.Items.Add(colors);
            var batchMenu = new ToolStripMenuItem("選択シートの一括操作");
            foreach (var command in new[] { SheetBatchCommand.Delete, SheetBatchCommand.Hide, SheetBatchCommand.Unhide,
                SheetBatchCommand.SetColor, SheetBatchCommand.ClearColor, SheetBatchCommand.Duplicate })
            {
                var captured = command;
                batchMenu.DropDownItems.Add(BatchCommandName(command), null, (sender, args) =>
                {
                    if (!updating && !IsDragging && batchMenuTargets.Length > 1)
                        BatchCommandRequested?.Invoke(this, new SheetBatchEventArgs(captured, batchMenuTargets));
                });
            }
            sheetMenu.Items.Add(batchMenu);
            sheetMenu.Opening += (sender, args) =>
            {
                menuTarget = sheetList.SelectedItem as WorksheetItem;
                bool single = sheetList.SelectedItems.Count == 1;
                sheetMenu.Items[0].Enabled = sheetList.SelectedItems.Count <= 1;
                for (int i = 1; i < sheetMenu.Items.Count; i++) sheetMenu.Items[i].Enabled = single;
                hide.Enabled = single && menuTarget?.Visibility == WorksheetVisibility.Visible;
                unhide.Enabled = single && menuTarget?.Visibility == WorksheetVisibility.Hidden;
                reveal.Enabled = single && menuTarget?.Visibility == WorksheetVisibility.VeryHidden;
                batchMenuTargets = SelectedSheetIds;
                batchMenu.Visible = batchMenuTargets.Length > 1;
                batchMenu.Enabled = batchMenuTargets.Length > 1;
                batchMenu.Text = "選択した " + batchMenuTargets.Length + "件を一括操作";
            };
            sheetList.MouseDown += (sender, args) =>
            {
                if (args.Button == MouseButtons.Left)
                    selection.PreserveSelection = (sheetList.PointerModifiers & (Keys.Control | Keys.Shift)) != 0;
                CaptureSelection();
            };
            sheetList.ContextMenuStrip = sheetMenu;
            addButton = new Button { Text = "+ シートを追加", Dock = DockStyle.Top, Height = 32, AccessibleName = "シートを追加" };
            addButton.Click += (sender, args) => RequestCommand(WorksheetCommand.Add, sheetList.SelectedItem as WorksheetItem);
            sheetList.SelectedIndexChanged += (sender, args) => { if (!updating) CaptureSelection(); };
            Controls.Add(sheetList);
            Controls.Add(selectionLabel);
            Controls.Add(statusLabel);
            Controls.Add(addButton);
            Controls.Add(workbookLabel);
        }

        public event EventHandler<WorksheetRequestedEventArgs> WorksheetRequested;
        public event EventHandler<WorksheetCommandEventArgs> CommandRequested;
        public event EventHandler<SheetBatchEventArgs> BatchCommandRequested;
        internal static string BatchCommandName(SheetBatchCommand command)
        {
            switch (command)
            {
                case SheetBatchCommand.Delete: return "一括削除…";
                case SheetBatchCommand.Hide: return "一括非表示…";
                case SheetBatchCommand.Unhide: return "一括再表示";
                case SheetBatchCommand.SetColor: return "一括タブ色変更…";
                case SheetBatchCommand.ClearColor: return "一括で色なし";
                default: return "一括複製";
            }
        }
        public bool IsMenuOpen { get { return sheetMenu.Visible; } }
        public bool IsDragging { get { return sheetList.IsDragging; } }
        public event EventHandler<WorksheetDragStartingEventArgs> DragStarting
        {
            add { sheetList.DragStarting += value; }
            remove { sheetList.DragStarting -= value; }
        }
        public event EventHandler<WorksheetDropEventArgs> DropRequested
        {
            add { sheetList.DropRequested += value; }
            remove { sheetList.DropRequested -= value; }
        }
        public event EventHandler DragCompleted
        {
            add { sheetList.DragCompleted += value; }
            remove { sheetList.DragCompleted -= value; }
        }
        public void CancelDrag() { sheetList.CancelDrag(); }
        public string[] SelectedSheetIds
        {
            get
            {
                var ids = new HashSet<string>(selection.SelectedIds);
                return sheetList.Items.Cast<WorksheetItem>().Where(item => ids.Contains(item.Id)).Select(item => item.Id).ToArray();
            }
        }

        private void CaptureSelection()
        {
            if (updating) return;
            selection.Capture(sheetList.SelectedItems.Cast<WorksheetItem>().Select(item => item.Id));
            int anchor = sheetList.AnchorIndex;
            if (anchor >= 0 && anchor < sheetList.Items.Count) selection.AnchorId = ((WorksheetItem)sheetList.Items[anchor]).Id;
            UpdateSelectionSummary();
        }

        private void UpdateSelectionSummary()
        {
            int count = sheetList.SelectedItems.Count;
            selectionLabel.Text = count + "件選択中" + (count > 1 ? "（右クリックで一括操作）" : "");
            addButton.Enabled = count <= 1;
        }

        private void ApplySelection()
        {
            var ids = new HashSet<string>(selection.SelectedIds);
            int anchor = -1;
            for (int i = 0; i < sheetList.Items.Count; i++)
            {
                var item = (WorksheetItem)sheetList.Items[i];
                bool wanted = ids.Contains(item.Id);
                if (sheetList.GetSelected(i) != wanted) sheetList.SetSelected(i, wanted);
                if (item.Id == selection.AnchorId) anchor = i;
            }
            sheetList.AnchorIndex = anchor;
            UpdateSelectionSummary();
        }

        public void ShowWorksheets(string workbookName, IList<WorksheetItem> worksheets, string activeId, string preferredId = null)
        {
            if (IsDragging || sheetList.IsSelecting) return;
            workbookLabel.Text = workbookName;
            CaptureSelection();
            selection.Reconcile(worksheets.Select(item => item.Id), activeId, preferredId);

            bool changed = sheetList.Items.Count != worksheets.Count;
            for (int i = 0; !changed && i < worksheets.Count; i++)
                changed = !SameItem((WorksheetItem)sheetList.Items[i], worksheets[i]);
            sheetList.SetActiveWorksheet(activeId);
            updating = true;
            if (changed) sheetList.BeginUpdate();
            try
            {
                int topIndex = sheetList.TopIndex;
                for (int i = 0; i < worksheets.Count; i++)
                {
                    if (i >= sheetList.Items.Count) sheetList.Items.Add(worksheets[i]);
                    else if (!SameItem((WorksheetItem)sheetList.Items[i], worksheets[i])) sheetList.Items[i] = worksheets[i];
                }
                while (sheetList.Items.Count > worksheets.Count) sheetList.Items.RemoveAt(sheetList.Items.Count - 1);
                if (changed) sheetList.UpdateTextExtent();

                ApplySelection();
                if (sheetList.Items.Count > 0)
                    sheetList.TopIndex = Math.Min(Math.Max(0, topIndex), sheetList.Items.Count - 1);
                if (changed || statusLabel.Text.Length == 0)
                    statusLabel.Text = worksheets.Count == 0 ? "シートがありません。" : "クリック、または選択して Enter で切り替えます。";
            }
            finally
            {
                if (changed) sheetList.EndUpdate();
                updating = false;
            }
        }

        public void ShowStatus(string message) { statusLabel.Text = message; }

        private static bool SameItem(WorksheetItem left, WorksheetItem right)
        {
            return left.Id == right.Id && left.Name == right.Name && left.Kind == right.Kind && left.Visibility == right.Visibility && left.TabColor == right.TabColor;
        }

        private void RequestCommand(WorksheetCommand command, WorksheetItem worksheet)
        {
            if (!updating && !IsDragging && sheetList.SelectedItems.Count <= 1 && (command == WorksheetCommand.Add || worksheet != null))
                CommandRequested?.Invoke(this, new WorksheetCommandEventArgs(command, worksheet));
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) sheetMenu.Dispose();
            base.Dispose(disposing);
        }

        private void OnSheetClicked(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && sheetList.IndexFromPoint(e.Location) != ListBox.NoMatches)
            {
                selection.PreserveSelection = false;
                RequestWorksheet();
            }
        }

        private void OnSheetKeyDown(object sender, KeyEventArgs e)
        {
            if ((e.Control && e.KeyCode == Keys.A) || e.KeyCode == Keys.Escape)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                if (IsDragging) { CancelDrag(); return; }
                if (e.KeyCode == Keys.Escape) selection.ResetToActive();
                else
                {
                    selection.Capture(sheetList.Items.Cast<WorksheetItem>().Select(item => item.Id));
                    selection.PreserveSelection = true;
                }
                int top = sheetList.TopIndex;
                updating = true;
                try { ApplySelection(); if (sheetList.Items.Count > 0) sheetList.TopIndex = top; }
                finally { updating = false; }
                return;
            }
            if (e.Control || e.Shift) selection.PreserveSelection = true;
            if (e.KeyCode != Keys.Enter) return;
            e.Handled = true;
            e.SuppressKeyPress = true;
            RequestWorksheet();
        }

        private void RequestWorksheet()
        {
            var worksheet = sheetList.SelectedItem as WorksheetItem;
            if (!updating && !IsDragging && sheetList.SelectedItems.Count == 1 && worksheet != null)
                WorksheetRequested?.Invoke(this, new WorksheetRequestedEventArgs(worksheet));
        }
    }
}
