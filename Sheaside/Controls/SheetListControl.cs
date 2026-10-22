using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Sheaside.Models;

namespace Sheaside.Controls
{
    internal sealed class SheetListControl : UserControl
    {
        private readonly Label workbookLabel;
        private readonly ReorderableSheetListBox sheetList;
        private readonly Label statusLabel;
        private readonly ContextMenuStrip sheetMenu;
        private WorksheetItem menuTarget;
        private string lastActiveId;
        private bool hasSnapshot;
        private bool updating;

        public SheetListControl()
        {
            AutoScaleMode = AutoScaleMode.Font;
            Font = new Font("Yu Gothic UI", 9F);
            Padding = new Padding(8);
            workbookLabel = new Label { Dock = DockStyle.Top, Height = 32, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
            statusLabel = new Label { Dock = DockStyle.Bottom, Height = 64, AutoEllipsis = true };
            sheetList = new ReorderableSheetListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                HorizontalScrollbar = true,
                DisplayMember = "DisplayName",
                AccessibleName = "ワークシート一覧"
            };
            // MouseClick also handles clicking an already selected row (e.g. after a failed activation).
            sheetList.MouseClick += OnSheetClicked;
            sheetList.KeyDown += OnSheetKeyDown;
            sheetList.MouseDoubleClick += (sender, args) =>
            {
                int index = sheetList.IndexFromPoint(args.Location);
                if (args.Button == MouseButtons.Left && index != ListBox.NoMatches)
                    RequestCommand(WorksheetCommand.Rename, (WorksheetItem)sheetList.Items[index]);
            };
            sheetMenu = new ContextMenuStrip();
            sheetMenu.Items.Add("シートの追加", null, (sender, args) => RequestCommand(WorksheetCommand.Add, menuTarget));
            sheetMenu.Items.Add("名前の変更", null, (sender, args) => RequestCommand(WorksheetCommand.Rename, menuTarget));
            sheetMenu.Items.Add("シートの複製", null, (sender, args) => RequestCommand(WorksheetCommand.Duplicate, menuTarget));
            sheetMenu.Items.Add("シートの削除", null, (sender, args) => RequestCommand(WorksheetCommand.Delete, menuTarget));
            sheetMenu.Opening += (sender, args) =>
            {
                menuTarget = sheetList.SelectedItem as WorksheetItem;
                for (int i = 1; i < sheetMenu.Items.Count; i++) sheetMenu.Items[i].Enabled = menuTarget != null;
            };
            sheetList.MouseDown += (sender, args) =>
            {
                if (args.Button == MouseButtons.Right)
                    sheetList.SelectedIndex = sheetList.IndexFromPoint(args.Location);
            };
            sheetList.ContextMenuStrip = sheetMenu;
            var addButton = new Button { Text = "+ シートを追加", Dock = DockStyle.Top, Height = 32, AccessibleName = "シートを追加" };
            addButton.Click += (sender, args) => RequestCommand(WorksheetCommand.Add, sheetList.SelectedItem as WorksheetItem);
            Controls.Add(sheetList);
            Controls.Add(statusLabel);
            Controls.Add(addButton);
            Controls.Add(workbookLabel);
        }

        public event EventHandler<WorksheetRequestedEventArgs> WorksheetRequested;
        public event EventHandler<WorksheetCommandEventArgs> CommandRequested;
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

        public void ShowWorksheets(string workbookName, IList<WorksheetItem> worksheets, string activeId, string preferredId = null)
        {
            if (IsDragging) return;
            workbookLabel.Text = workbookName;
            var previousSelection = sheetList.SelectedItem as WorksheetItem;
            string selection = preferredId ?? (!hasSnapshot || lastActiveId != activeId
                ? activeId : previousSelection?.Id);
            int selectedIndex = -1;
            for (int i = 0; i < worksheets.Count; i++)
                if (worksheets[i].Id == selection) selectedIndex = i;
            if (selectedIndex < 0)
                for (int i = 0; i < worksheets.Count; i++)
                    if (worksheets[i].Id == activeId) selectedIndex = i;

            bool changed = sheetList.Items.Count != worksheets.Count;
            for (int i = 0; !changed && i < worksheets.Count; i++)
                changed = !SameItem((WorksheetItem)sheetList.Items[i], worksheets[i]);
            lastActiveId = activeId;
            hasSnapshot = true;
            if (!changed && sheetList.SelectedIndex == selectedIndex) return;
            updating = true;
            sheetList.BeginUpdate();
            try
            {
                int topIndex = sheetList.TopIndex;
                for (int i = 0; i < worksheets.Count; i++)
                {
                    if (i >= sheetList.Items.Count) sheetList.Items.Add(worksheets[i]);
                    else if (!SameItem((WorksheetItem)sheetList.Items[i], worksheets[i])) sheetList.Items[i] = worksheets[i];
                }
                while (sheetList.Items.Count > worksheets.Count) sheetList.Items.RemoveAt(sheetList.Items.Count - 1);

                sheetList.SelectedIndex = selectedIndex;
                if (changed && sheetList.Items.Count > 0)
                    sheetList.TopIndex = Math.Min(Math.Max(0, topIndex), sheetList.Items.Count - 1);
                if (changed || statusLabel.Text.Length == 0)
                    statusLabel.Text = worksheets.Count == 0 ? "ワークシートがありません。" : "クリック、または選択して Enter で切り替えます。";
            }
            finally
            {
                sheetList.EndUpdate();
                updating = false;
            }
        }

        public void ShowStatus(string message) { statusLabel.Text = message; }

        private static bool SameItem(WorksheetItem left, WorksheetItem right)
        {
            return left.Id == right.Id && left.Name == right.Name && left.IsVisible == right.IsVisible;
        }

        private void RequestCommand(WorksheetCommand command, WorksheetItem worksheet)
        {
            if (!updating && !IsDragging && (command == WorksheetCommand.Add || worksheet != null))
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
                RequestWorksheet();
        }

        private void OnSheetKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;
            e.Handled = true;
            e.SuppressKeyPress = true;
            RequestWorksheet();
        }

        private void RequestWorksheet()
        {
            var worksheet = sheetList.SelectedItem as WorksheetItem;
            if (!updating && !IsDragging && worksheet != null)
                WorksheetRequested?.Invoke(this, new WorksheetRequestedEventArgs(worksheet));
        }
    }
}
