using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Sheaside.Controls;
using Sheaside.Models;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sheaside.Services
{
    internal sealed class SheetPaneManager : IDisposable
    {
        private sealed class PaneEntry
        {
            public CustomTaskPane Pane;
            public SheetListControl View;
            public EventHandler<WorksheetRequestedEventArgs> WorksheetHandler;
            public EventHandler<WorksheetCommandEventArgs> CommandHandler;
            public string PreferredSelection;
        }

        private readonly Excel.Application application;
        private readonly CustomTaskPaneCollection taskPanes;
        private readonly ExcelWorksheetService worksheets;
        private readonly Dictionary<int, PaneEntry> entries = new Dictionary<int, PaneEntry>();
        private readonly Timer refreshTimer;
        private bool refreshing;
        private bool disposed;
        private bool operating;
        private int? activeHandle;

        public SheetPaneManager(Excel.Application application, CustomTaskPaneCollection taskPanes)
        {
            this.application = application;
            this.taskPanes = taskPanes;
            worksheets = new ExcelWorksheetService(application);
            // Excel has no WindowClosed or WorksheetRenamed event. A UI-thread timer
            // reconciles actual windows after a close/save dialog (including Cancel).
            refreshTimer = new Timer { Interval = 1000 };
            refreshTimer.Tick += OnRefresh;
        }

        public event EventHandler StateChanged;

        public void Start()
        {
            application.WindowActivate += OnWindowActivate;
            application.WorkbookActivate += OnWorkbookActivate;
            application.SheetActivate += OnSheetActivate;
            refreshTimer.Start();
            Refresh();
        }

        public bool IsVisible(Excel.Window window)
        {
            try
            {
                PaneEntry entry;
                return !disposed && window != null && entries.TryGetValue(window.Hwnd, out entry) && entry.Pane.Visible;
            }
            catch (COMException) { return false; }
        }

        public void SetVisible(Excel.Window window, bool visible)
        {
            if (disposed || window == null) return;
            try
            {
                Refresh();
                PaneEntry entry;
                if (entries.TryGetValue(window.Hwnd, out entry)) entry.Pane.Visible = visible;
            }
            catch (COMException exception) { Trace.TraceWarning("Sheaside: {0}", exception); }
        }

        private void OnRefresh(object sender, EventArgs e) { Refresh(); }
        private void OnWindowActivate(Excel.Workbook workbook, Excel.Window window) { Refresh(); }
        private void OnWorkbookActivate(Excel.Workbook workbook) { Refresh(); }
        private void OnSheetActivate(object sheet) { Refresh(); }

        private void Refresh()
        {
            if (disposed || refreshing || operating) return;
            refreshing = true;
            try
            {
                var windows = worksheets.GetWindows();
                // Build a complete snapshot before removing anything. Failed COM reads
                // during cell editing must not be mistaken for a closed window.
                var liveHandles = new HashSet<int>(windows.Select(window => window.Hwnd));
                foreach (int handle in entries.Keys.Where(handle => !liveHandles.Contains(handle)).ToArray())
                    RemovePane(handle);

                foreach (var window in windows)
                {
                    int handle = window.Hwnd;
                    PaneEntry entry;
                    if (!entries.TryGetValue(handle, out entry)) entry = CreatePane(window, handle);
                    if (!entry.Pane.Visible || entry.View.IsMenuOpen) continue;
                    var workbook = worksheets.GetWorkbook(window);
                    if (workbook != null)
                    {
                        entry.View.ShowWorksheets(workbook.Name, worksheets.GetWorksheets(workbook), worksheets.GetActiveWorksheetId(window), entry.PreferredSelection);
                        entry.PreferredSelection = null;
                    }
                }

                var activeWindow = application.ActiveWindow;
                int? currentHandle = activeWindow == null ? (int?)null : activeWindow.Hwnd;
                if (activeHandle != currentHandle)
                {
                    activeHandle = currentHandle;
                    StateChanged?.Invoke(this, EventArgs.Empty);
                }
            }
            catch (COMException exception)
            {
                // Excel may reject calls while editing a cell or displaying a dialog.
                // Keep the last good list and retry on the next UI-thread tick.
                Trace.TraceWarning("Sheaside refresh: {0}", exception.Message);
            }
            finally { refreshing = false; }
        }

        private PaneEntry CreatePane(Excel.Window window, int handle)
        {
            var view = new SheetListControl();
            CustomTaskPane pane = null;
            try
            {
                pane = taskPanes.Add(view, "Sheaside", window);
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                pane.Width = 260;
                pane.Visible = true;
                var entry = new PaneEntry { Pane = pane, View = view };
                entry.WorksheetHandler = (sender, args) => ActivateWorksheet(handle, entry, args.Worksheet);
                view.WorksheetRequested += entry.WorksheetHandler;
                entry.CommandHandler = (sender, args) => ExecuteCommand(handle, entry, args);
                view.CommandRequested += entry.CommandHandler;
                pane.VisibleChanged += OnVisibleChanged;
                entries.Add(handle, entry);
                StateChanged?.Invoke(this, EventArgs.Empty);
                return entry;
            }
            catch
            {
                if (pane != null) taskPanes.Remove(pane);
                view.Dispose();
                throw;
            }
        }

        private void ActivateWorksheet(int handle, PaneEntry entry, WorksheetItem worksheet)
        {
            if (disposed || operating) return;
            try { entry.View.ShowStatus(worksheets.ActivateWorksheet(handle, worksheet.Id)); }
            catch (COMException exception)
            {
                Trace.TraceWarning("Sheaside activation: {0}", exception);
                entry.View.ShowStatus("シートを切り替えられません。セルの編集やダイアログを閉じて、もう一度お試しください。");
            }
            Refresh();
        }

        private void ExecuteCommand(int handle, PaneEntry entry, WorksheetCommandEventArgs args)
        {
            if (disposed || operating) return;
            operating = true;
            string status = null;
            try
            {
                var target = worksheets.CaptureTarget(handle, args.Worksheet?.Id);
                switch (args.Command)
                {
                    case WorksheetCommand.Add:
                        entry.PreferredSelection = worksheets.AddWorksheet(target);
                        status = "シートを追加しました。";
                        break;
                    case WorksheetCommand.Rename:
                        using (var dialog = new RenameSheetDialog(target.Worksheet.Name, newName =>
                        {
                            try { worksheets.RenameWorksheet(target, newName); return null; }
                            catch (Exception exception) when (IsOperationError(exception)) { return OperationErrorMessage(exception); }
                        }))
                        {
                            if (dialog.ShowDialog(entry.View) == DialogResult.OK)
                            {
                                entry.PreferredSelection = worksheets.GetWorksheetId(target.Worksheet);
                                status = "シート名を変更しました。";
                            }
                        }
                        break;
                    case WorksheetCommand.Duplicate:
                        entry.PreferredSelection = worksheets.DuplicateWorksheet(target);
                        status = "元シートの直後に複製しました。";
                        break;
                    case WorksheetCommand.Delete:
                        worksheets.EnsureCanDelete(target);
                        string name = target.Worksheet.Name;
                        if (MessageBox.Show(entry.View, "シート「" + name + "」を削除しますか？\nこの操作は元に戻せません。",
                            "Sheaside — シートの削除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                        {
                            worksheets.DeleteWorksheet(target, name);
                            status = "シート「" + name + "」を削除しました。";
                        }
                        break;
                }
            }
            catch (Exception exception) when (IsOperationError(exception))
            {
                status = OperationErrorMessage(exception);
                MessageBox.Show(entry.View, status, "Sheaside", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                operating = false;
                Refresh();
                if (!entry.View.IsDisposed && status != null) entry.View.ShowStatus(status);
            }
        }

        private static bool IsOperationError(Exception exception)
        {
            return exception is COMException || exception is InvalidOperationException || exception is InvalidComObjectException;
        }

        private static string OperationErrorMessage(Exception exception)
        {
            Trace.TraceWarning("Sheaside operation: {0}", exception);
            return exception is InvalidOperationException ? exception.Message
                : "Excelで操作を完了できませんでした。セルの編集やダイアログを閉じ、一覧の状態を確認してからお試しください。";
        }

        private void OnVisibleChanged(object sender, EventArgs e)
        {
            StateChanged?.Invoke(this, EventArgs.Empty);
            Refresh();
        }

        private void RemovePane(int handle)
        {
            var entry = entries[handle];
            entry.Pane.VisibleChanged -= OnVisibleChanged;
            try { taskPanes.Remove(entry.Pane); }
            catch
            {
                entry.Pane.VisibleChanged += OnVisibleChanged;
                throw;
            }
            entry.View.WorksheetRequested -= entry.WorksheetHandler;
            entry.View.CommandRequested -= entry.CommandHandler;
            entry.View.Dispose();
            entries.Remove(handle);
            StateChanged?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            refreshTimer.Stop();
            refreshTimer.Tick -= OnRefresh;
            refreshTimer.Dispose();
            application.WindowActivate -= OnWindowActivate;
            application.WorkbookActivate -= OnWorkbookActivate;
            application.SheetActivate -= OnSheetActivate;
            foreach (var entry in entries.Values)
            {
                entry.Pane.VisibleChanged -= OnVisibleChanged;
                entry.View.WorksheetRequested -= entry.WorksheetHandler;
                entry.View.CommandRequested -= entry.CommandHandler;
            }
            entries.Clear();
            // VSTO owns task pane/control disposal at shutdown. Remove must not be
            // called here: the collection may already have been disposed by VSTO.
        }
    }
}
