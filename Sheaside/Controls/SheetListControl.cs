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
        private readonly ListBox sheetList;
        private readonly Label statusLabel;
        private bool updating;

        public SheetListControl()
        {
            AutoScaleMode = AutoScaleMode.Font;
            Font = new Font("Yu Gothic UI", 9F);
            Padding = new Padding(8);
            workbookLabel = new Label { Dock = DockStyle.Top, Height = 32, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
            statusLabel = new Label { Dock = DockStyle.Bottom, Height = 64, AutoEllipsis = true };
            sheetList = new ListBox
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
            Controls.Add(sheetList);
            Controls.Add(statusLabel);
            Controls.Add(workbookLabel);
        }

        public event EventHandler<WorksheetRequestedEventArgs> WorksheetRequested;

        public void ShowWorksheets(string workbookName, IList<WorksheetItem> worksheets, string activeCodeName)
        {
            updating = true;
            sheetList.BeginUpdate();
            try
            {
                workbookLabel.Text = workbookName;
                bool changed = sheetList.Items.Count != worksheets.Count;
                for (int i = 0; !changed && i < worksheets.Count; i++)
                {
                    var previous = (WorksheetItem)sheetList.Items[i];
                    var current = worksheets[i];
                    changed = previous.CodeName != current.CodeName || previous.Name != current.Name || previous.IsVisible != current.IsVisible;
                }

                int topIndex = sheetList.TopIndex;
                if (changed)
                {
                    sheetList.Items.Clear();
                    foreach (var worksheet in worksheets) sheetList.Items.Add(worksheet);
                }

                int selectedIndex = -1;
                for (int i = 0; i < worksheets.Count; i++)
                    if (worksheets[i].CodeName == activeCodeName) selectedIndex = i;
                // Keep keyboard navigation usable while the list has focus.
                if (changed || !sheetList.Focused) sheetList.SelectedIndex = selectedIndex;
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
            if (!updating && worksheet != null)
                WorksheetRequested?.Invoke(this, new WorksheetRequestedEventArgs(worksheet));
        }
    }
}
