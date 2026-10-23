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
        private int visibilityFilter, kindFilter;
        private readonly Button filterButton;
        private readonly Label filterSummary;
        private readonly Panel filterPopup;
        internal bool HasFacetFilter { get { return visibilityFilter != 0 || kindFilter != 0; } }
        internal Panel FilterPopup { get { return filterPopup; } }
        private bool changingFilter;
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
                if (filterPopup != null) filterPopup.Visible = false;
                if (args.Button == MouseButtons.Left)
                    selection.PreserveSelection = (sheetList.PointerModifiers & (Keys.Control | Keys.Shift)) != 0;
                CaptureSelection();
            };
            sheetList.ContextMenuStrip = sheetMenu;
            addButton = new Button { Text = "＋ 追加", Dock = DockStyle.Fill, AccessibleName = "シートを追加" };
            addButton.Click += (sender, args) => RequestCommand(WorksheetCommand.Add, sheetList.SelectedItem as WorksheetItem);
            deleteButton = new Button { Text = "削除", Dock = DockStyle.Fill, AccessibleName = "選択シートを削除", Enabled = false };
            deleteButton.Click += (s, e) =>
            {
                batchMenuTargets = SelectedSheetIds;
                RequestBatch(SheetBatchCommand.Delete);
            };
            var actions = new TableLayoutPanel { Dock = DockStyle.Top, Height = 34, ColumnCount = 2, RowCount = 1 };
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            actions.Controls.Add(addButton, 0, 0); actions.Controls.Add(deleteButton, 1, 0);
            sheetList.SelectedIndexChanged += (sender, args) => { if (!updating) CaptureSelection(); };
            Controls.Add(sheetList);
            Controls.Add(selectionLabel);
            Controls.Add(actions);
            var header = new Panel { Dock = DockStyle.Top, Height = 32 };
            var settingsButton = new Button { Text = "設定", Dock = DockStyle.Right, Width = 54, AccessibleName = "Sheasideの設定" };
            header.Controls.Add(workbookLabel); header.Controls.Add(settingsButton);
            settingsButton.Click += (s, e) =>
            {
                using (var dialog = new PreferencesDialog(preferences))
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    try { preferences.SetOptions(dialog.ShowVeryHidden, dialog.StartupVisible); }
                    catch (Exception ex) when (ex is System.IO.IOException || ex is UnauthorizedAccessException)
                    { MessageBox.Show(this, "設定を保存できませんでした。\r\n" + ex.Message, "Sheaside"); }
                }
            };
            preferences.Changed += OnPreferencesChanged;
            var filters = new Panel { Dock = DockStyle.Top, Height = 58 };
            var searchRow = new Panel { Dock = DockStyle.Top, Height = 28 };
            searchBox = new TextBox { Dock = DockStyle.Fill, AccessibleName = "シート名を部分一致検索" };
            filterButton = new Button { Dock = DockStyle.Right, Width = 30, AccessibleName = "状態と種類のフィルター" };
            filterButton.Paint += (s, e) =>
            {
                using (var pen = new Pen(HasFacetFilter ? Color.ForestGreen : filterButton.ForeColor, 2))
                    e.Graphics.DrawPolygon(pen, new[] { new Point(7, 7), new Point(22, 7), new Point(17, 13), new Point(17, 21), new Point(12, 18), new Point(12, 13) });
            };
            filterButton.Click += (s, e) => ToggleFilterPopup();
            searchRow.Controls.Add(searchBox); searchRow.Controls.Add(filterButton);
            filterSummary = new Label { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
            var clear = new Button { Text = "解除", Dock = DockStyle.Right, Width = 48, AccessibleName = "状態・種類フィルターを解除" };
            clear.Click += (s, e) => SetFilter(searchBox.Text, null, null);
            var summaryRow = new Panel { Dock = DockStyle.Fill };
            summaryRow.Controls.Add(filterSummary); summaryRow.Controls.Add(clear);
            filters.Controls.Add(summaryRow); filters.Controls.Add(searchRow);
            searchBox.TextChanged += (s, e) => RefreshFilter();
            Controls.Add(filters);
            filterPopup = new Panel { Visible = false, BorderStyle = BorderStyle.FixedSingle, AutoScroll = true };
            Controls.Add(filterPopup);
            SizeChanged += (s, e) => PositionFilterPopup();
            Controls.Add(header);
            ApplyTheme();
        }

        public void SetFilter(string query, WorksheetVisibility? visibility, SheetKind? kind)
        {
            changingFilter = true;
            try
            {
                searchBox.Text = query ?? "";
                visibilityFilter = !visibility.HasValue ? 0 : visibility == WorksheetVisibility.Visible ? 1 : visibility == WorksheetVisibility.Hidden ? 2 : 3;
                kindFilter = !kind.HasValue ? 0 : kind == SheetKind.Worksheet ? 1 : 2;
            }
            finally { changingFilter = false; }
            RefreshFilter();
        }
        internal void ToggleFilterPopup()
        {
            if (filterPopup.Visible) { filterPopup.Visible = false; return; }
            foreach (Control child in filterPopup.Controls.Cast<Control>().ToArray()) child.Dispose();
            var contents = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            contents.Controls.Add(new Label { Text = "状態", AutoSize = true });
            string[] states = { "すべて", "表示", "非表示", "VeryHidden" };
            for (int i = 0; i < (preferences.ShowVeryHidden ? 4 : 3); i++)
            {
                int value = i;
                var button = new Button { Text = (visibilityFilter == i ? "✓ " : "　") + states[i], Width = 170, Height = 26 };
                button.Click += (s, e) => { visibilityFilter = value; filterPopup.Visible = false; RefreshFilter(); };
                contents.Controls.Add(button);
            }
            contents.Controls.Add(new Label { Text = "種類", AutoSize = true });
            string[] kinds = { "すべて", "ワークシート", "グラフシート" };
            for (int i = 0; i < 3; i++)
            {
                int value = i;
                var button = new Button { Text = (kindFilter == i ? "✓ " : "　") + kinds[i], Width = 170, Height = 26 };
                button.Click += (s, e) => { kindFilter = value; filterPopup.Visible = false; RefreshFilter(); };
                contents.Controls.Add(button);
            }
            var clear = new Button { Text = "フィルター解除", Width = 170, Height = 28 };
            clear.Click += (s, e) => { filterPopup.Visible = false; SetFilter(searchBox.Text, null, null); };
            contents.Controls.Add(clear);
            filterPopup.Controls.Add(contents);
            ApplyTheme();
            PositionFilterPopup(); filterPopup.Visible = true; filterPopup.BringToFront();
        }
        private void PositionFilterPopup()
        {
            if (filterPopup == null || filterButton == null) return;
            int top = Math.Min(PointToClient(filterButton.PointToScreen(new Point(0, filterButton.Height))).Y, Math.Max(0, ClientSize.Height - 80));
            filterPopup.SetBounds(Math.Max(0, ClientSize.Width - 204), top, Math.Max(1, Math.Min(200, ClientSize.Width)), Math.Max(1, Math.Min(320, ClientSize.Height - top)));
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape && filterPopup.Visible) { filterPopup.Visible = false; return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        private void RefreshFilter()
        {
            if (changingFilter) return;
            filterSummary.Text = (visibilityFilter == 0 ? "状態:すべて" : new[] { "", "表示", "非表示", "VeryHidden" }[visibilityFilter]) + " / " + (kindFilter == 0 ? "種類:すべて" : kindFilter == 1 ? "ワークシート" : "グラフシート");
            filterButton.Invalidate();
            CancelDrag();
            sheetMenu.Close(); batchMenuTargets = new string[0];
            ShowWorksheets(snapshotBook, snapshot, snapshotActive);
        }

        private void OnPreferencesChanged(object sender, EventArgs e)
        {
            ApplyTheme();
            if (!preferences.ShowVeryHidden && visibilityFilter == 3) visibilityFilter = 0;
            filterPopup.Visible = false;
            RefreshFilter();
        }
        private OfficeTheme officeTheme;
        internal void SetOfficeTheme(OfficeTheme theme)
        {
            if (officeTheme == theme) return;
            officeTheme = theme;
            ApplyTheme();
        }
        private void ApplyTheme()
        {
            bool system = officeTheme == OfficeTheme.System || SystemInformation.HighContrast;
            bool dark = !system && (officeTheme == OfficeTheme.Black || officeTheme == OfficeTheme.DarkGray);
            Color background = system ? SystemColors.Control : officeTheme == OfficeTheme.DarkGray ? Color.FromArgb(68, 68, 68) : dark ? Color.FromArgb(32, 32, 32) : officeTheme == OfficeTheme.LightGray ? Color.Gainsboro : Color.WhiteSmoke;
            Color foreground = system ? SystemColors.ControlText : dark ? Color.WhiteSmoke : Color.FromArgb(24, 24, 24);
            ApplyColors(this, background, foreground);
            sheetList.BackColor = system ? SystemColors.Window : officeTheme == OfficeTheme.DarkGray ? Color.FromArgb(58, 58, 58) : dark ? Color.FromArgb(24, 24, 24) : Color.White;
            sheetList.ForeColor = system ? SystemColors.WindowText : foreground;
            sheetMenu.BackColor = background; sheetMenu.ForeColor = foreground;
            sheetList.Invalidate();
        }
        private static void ApplyColors(Control parent, Color background, Color foreground)
        {
            parent.BackColor = background; parent.ForeColor = foreground;
            foreach (Control child in parent.Controls) ApplyColors(child, background, foreground);
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
        public bool IsMenuOpen { get { return sheetMenu.Visible || filterPopup.Visible; } }
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
            string query = searchBox.Text;
            int state = visibilityFilter, kind = kindFilter;
            worksheets = worksheets.Where(i => i.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 &&
                (state == 0 || (state == 1 && i.IsVisible) || (state == 2 && i.Visibility == WorksheetVisibility.Hidden) || (state == 3 && i.Visibility == WorksheetVisibility.VeryHidden)) &&
                (kind == 0 || (kind == 1 && i.Kind == SheetKind.Worksheet) || (kind == 2 && i.Kind == SheetKind.Chart))).ToArray();
            IsFiltered = query.Length > 0 || state != 0 || kind != 0 || worksheets.Count != snapshot.Count;
            sheetList.ReorderEnabled = !IsFiltered;
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
