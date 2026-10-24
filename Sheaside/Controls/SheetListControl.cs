using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Sheaside.Models;
using Sheaside.Services;

namespace Sheaside.Controls
{
    internal sealed class SheetListControl : UserControl
    {
        private readonly Label workbookLabel;
        private readonly ReorderableSheetListBox sheetList;
        private readonly Label selectionLabel;
        private readonly Button addButton;
        private readonly Button deleteButton;
        internal Button AddButton { get { return addButton; } }
        internal Button DeleteButton { get { return deleteButton; } }
        private readonly SheetSelectionState selection = new SheetSelectionState();
        private readonly ContextMenuStrip sheetMenu;
        private WorksheetItem menuTarget;
        private string[] batchMenuTargets = new string[0];
        private bool updating;
        private readonly UserPreferences preferences;
        private IList<WorksheetItem> snapshot = new WorksheetItem[0];
        private string snapshotBook, snapshotActive;
        private readonly TextBox searchBox;
        internal Button ClearSearchButton { get; private set; }
        public bool IsFiltered { get; private set; }

        public SheetListControl() : this(new UserPreferences()) { }
        public SheetListControl(UserPreferences preferences)
        {
            this.preferences = preferences;
            AutoScaleMode = AutoScaleMode.Font;
            Size = new Size(260, 500);
            Font = new Font("Yu Gothic UI", 9F);
            Padding = new Padding(8);
            workbookLabel = new Label { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
            selectionLabel = new Label { Dock = DockStyle.Bottom, Height = 26, AutoEllipsis = true };
            sheetList = new ReorderableSheetListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                HorizontalScrollbar = true,
                DisplayMember = "AccessibleDescription",
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
            sheetMenu.Items.Add("シートの複製", null, (sender, args) => RequestBatch(SheetBatchCommand.Duplicate));
            sheetMenu.Items.Add("シートの削除…", null, (sender, args) => RequestBatch(SheetBatchCommand.Delete));
            var hide = sheetMenu.Items.Add("非表示にする", null, (sender, args) => RequestBatch(SheetBatchCommand.Hide));
            var unhide = sheetMenu.Items.Add("再表示", null, (sender, args) => RequestBatch(SheetBatchCommand.Unhide));
            var reveal = sheetMenu.Items.Add("VeryHiddenにする", null, (sender, args) => RequestBatch(SheetBatchCommand.VeryHide));
            var colors = new ToolStripMenuItem("タブの色");
            colors.DropDownItems.Add("色を選択…", null, (sender, args) => RequestBatch(SheetBatchCommand.SetColor));
            colors.DropDownItems.Add("色なし", null, (sender, args) => RequestBatch(SheetBatchCommand.ClearColor));
            sheetMenu.Items.Add(colors);
            var add = sheetMenu.Items[0]; var rename = sheetMenu.Items[1];
            var duplicate = sheetMenu.Items[2]; var delete = sheetMenu.Items[3];
            sheetMenu.Items.Clear();
            sheetMenu.Items.AddRange(new ToolStripItem[] { add, delete, new ToolStripSeparator(), rename, duplicate,
                new ToolStripSeparator(), hide, unhide, reveal, new ToolStripSeparator(), colors });
            sheetMenu.Opening += (sender, args) =>
            {
                menuTarget = sheetList.SelectedItem as WorksheetItem;
                bool single = sheetList.SelectedItems.Count == 1;
                sheetMenu.Items[0].Enabled = sheetList.SelectedItems.Count <= 1;
                for (int i = 1; i < sheetMenu.Items.Count; i++) sheetMenu.Items[i].Enabled = sheetList.SelectedItems.Count > 0;
                rename.Enabled = single;
                var selected = sheetList.SelectedItems.Cast<WorksheetItem>().ToArray();
                hide.Enabled = selected.Any(i => i.Visibility != WorksheetVisibility.Hidden);
                unhide.Enabled = selected.Any(i => !i.IsVisible);
                reveal.Visible = preferences.ShowVeryHidden;
                reveal.Enabled = preferences.ShowVeryHidden && selected.Any(i => i.Visibility != WorksheetVisibility.VeryHidden);
                batchMenuTargets = SelectedSheetIds;
            };
            sheetList.MouseDown += (sender, args) =>
            {
                if (args.Button == MouseButtons.Left)
                    selection.PreserveSelection = (sheetList.PointerModifiers & (Keys.Control | Keys.Shift)) != 0;
                CaptureSelection();
            };
            sheetList.ContextMenuStrip = sheetMenu;
            addButton = new ModernButton { Icon = LineIcon.Add, Text = "追加", Dock = DockStyle.Fill, AccessibleName = "シートを追加" };
            addButton.Click += (sender, args) => RequestCommand(WorksheetCommand.Add, sheetList.SelectedItem as WorksheetItem);
            deleteButton = new ModernButton { Icon = LineIcon.Delete, Text = "削除", Dock = DockStyle.Fill, AccessibleName = "選択シートを削除", Enabled = false };
            deleteButton.Click += (s, e) =>
            {
                batchMenuTargets = SelectedSheetIds;
                RequestBatch(SheetBatchCommand.Delete);
            };
            var actions = new TableLayoutPanel { Dock = DockStyle.Top, Height = 40, ColumnCount = 2, RowCount = 1 };
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            actions.Controls.Add(addButton, 0, 0); actions.Controls.Add(deleteButton, 1, 0);
            sheetList.SelectedIndexChanged += (sender, args) => { if (!updating) CaptureSelection(); };
            Controls.Add(sheetList);
            Controls.Add(selectionLabel);
            Controls.Add(actions);
            var bookRow = new Panel { Dock = DockStyle.Top, Height = 26 };
            bookRow.Controls.Add(workbookLabel);
            searchBox = new TextBox { Dock = DockStyle.Fill, AccessibleName = "シート名を部分一致検索" };
            ClearSearchButton = new ModernButton { Text = "解除", Dock = DockStyle.Right, Width = 48, Visible = false, AccessibleName = "検索を解除" };
            ClearSearchButton.Click += (s, e) => SetSearch("");
            var searchRow = new Panel { Dock = DockStyle.Top, Height = 32, Padding = new Padding(0, 3, 0, 3) };
            var searchIcon = new Panel { Dock = DockStyle.Left, Width = 26, AccessibleName = "検索" };
            searchIcon.Paint += (s, e) => { int size = (int)Math.Round(16 * e.Graphics.DpiX / 96F); ModernButton.DrawIcon(e.Graphics, LineIcon.Search, new Rectangle(2, (searchIcon.Height - size) / 2, size, size), searchIcon.ForeColor); };
            searchRow.Controls.Add(searchBox); searchRow.Controls.Add(searchIcon); searchRow.Controls.Add(ClearSearchButton);
            searchBox.TextChanged += (s, e) => RefreshSearch();
            Controls.Add(searchRow); Controls.Add(bookRow);
            Controls.Add(new Label { Text = "Sheaside", Dock = DockStyle.Top, Height = 28, TextAlign = ContentAlignment.MiddleLeft });
            preferences.Changed += OnPreferencesChanged;
            ApplyTheme();
        }
        public void SetSearch(string query) { searchBox.Text = query ?? ""; }
        private void RefreshSearch()
        {
            ClearSearchButton.Visible = searchBox.Text.Length > 0;
            CancelDrag(); sheetMenu.Close(); batchMenuTargets = new string[0];
            ShowWorksheets(snapshotBook, snapshot, snapshotActive);
        }
        private void OnPreferencesChanged(object sender, EventArgs e) { RefreshSearch(); }
        private OfficeTheme officeTheme;
        internal void SetOfficeTheme(OfficeTheme theme)
        {
            if (officeTheme == theme) return;
            officeTheme = theme;
            ApplyTheme();
        }
        private void ApplyTheme()
        {
            UiAppearance.Apply(this, officeTheme);
            sheetList.BackColor = UiAppearance.Window(officeTheme);
            sheetList.ForeColor = UiAppearance.SystemTheme(officeTheme) ? SystemColors.WindowText : ForeColor;
            sheetMenu.BackColor = BackColor; sheetMenu.ForeColor = ForeColor;
            sheetList.Invalidate();
        }
        protected override void OnSystemColorsChanged(EventArgs e)
        {
            base.OnSystemColorsChanged(e);
            if (preferences != null && sheetList != null && sheetMenu != null) ApplyTheme();
        }
        private void RequestBatch(SheetBatchCommand command)
        {
            if (!updating && !IsDragging && batchMenuTargets.Length > 0)
                BatchCommandRequested?.Invoke(this, new SheetBatchEventArgs(command, batchMenuTargets));
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
                case SheetBatchCommand.VeryHide: return "VeryHiddenにする";
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
            selectionLabel.Text = count + "件選択中 / " + sheetList.Items.Count + "件表示";
            addButton.Enabled = count <= 1;
            deleteButton.Enabled = count > 0;
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
            snapshot = worksheets.ToArray(); snapshotBook = workbookName; snapshotActive = activeId;
            worksheets = worksheets.Where(i => preferences.ShowVeryHidden || i.Visibility != WorksheetVisibility.VeryHidden).ToArray();
            worksheets = worksheets.Where(i => preferences.ShowHidden || i.Visibility != WorksheetVisibility.Hidden).ToArray();
            string query = searchBox.Text;
            worksheets = worksheets.Where(i => i.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0).ToArray();
            IsFiltered = query.Length > 0 || worksheets.Count != snapshot.Count;
            sheetList.ReorderEnabled = true;
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
            }
            finally
            {
                if (changed) sheetList.EndUpdate();
                updating = false;
            }
        }

        public void ShowStatus(string message)
        { if (!string.IsNullOrEmpty(message)) MessageBox.Show(this, message, "Sheaside", MessageBoxButtons.OK, MessageBoxIcon.Information); }

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
            if (disposing) { preferences.Changed -= OnPreferencesChanged; sheetMenu.Dispose(); }
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
