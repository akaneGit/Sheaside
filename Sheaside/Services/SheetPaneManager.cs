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
            public EventHandler<WorksheetDragStartingEventArgs> DragStartingHandler;
            public EventHandler<WorksheetDropEventArgs> DropHandler;
            public EventHandler DragCompletedHandler;
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
        private PaneEntry draggingEntry;
        private ExcelWorksheetService.MoveSession moveSession;
        private bool dragCancelled;
        private string dragStatus;

        public SheetPaneManager(Excel.Application application, CustomTaskPaneCollection taskPanes)
        {
            this.application = application;
            this.taskPanes = taskPanes;
            worksheets = new ExcelWorksheetService(application);
            // Events handle activation immediately; a conditional UI-thread timer
            // covers changes without dedicated events (visibility, tab color, close).
            refreshTimer = new Timer { Interval = 3000 };
            refreshTimer.Tick += OnRefresh;
        }

        public event EventHandler StateChanged;

        public void Start()
        {
            application.WindowActivate += OnWindowActivate;
            application.WorkbookActivate += OnWorkbookActivate;
            application.SheetActivate += OnSheetActivate;
            application.WindowDeactivate += OnWindowDeactivate;
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

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private void OnRefresh(object sender, EventArgs e)
        {
            if (disposed || operating) return;
            if (draggingEntry != null) { Refresh(); return; }
            try
            {
                // No visibility/tab-color events exist. Reconcile only while Excel
                // is foreground, ready, and its active window's pane is visible.
                var window = application.ActiveWindow;
                PaneEntry entry;
                if (window == null || GetForegroundWindow() != new IntPtr(window.Hwnd) || !application.Ready ||
                    !entries.TryGetValue(window.Hwnd, out entry) || !entry.Pane.Visible || entry.View.IsMenuOpen) return;
                Refresh(window.Hwnd);
            }
            catch (Exception exception) when (IsOperationError(exception)) { Trace.TraceWarning("Sheaside timer: {0}", exception.Message); }
        }
        private void OnWindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            if (!operating) CancelCurrentDrag("ウィンドウが切り替わったため、並び替えを中止しました。");
            Refresh();
        }
        private void OnWindowDeactivate(Excel.Workbook workbook, Excel.Window window)
        {
            if (!operating) CancelCurrentDrag("ウィンドウが切り替わったため、並び替えを中止しました。");
        }
        private void OnWorkbookActivate(Excel.Workbook workbook)
        {
            if (!operating) CancelCurrentDrag("ブックが切り替わったため、並び替えを中止しました。");
            Refresh();
        }
        private void OnSheetActivate(object sheet) { Refresh(); }

        private void Refresh(int? onlyHandle = null)
        {
            if (disposed || refreshing || operating) return;
            if (draggingEntry != null)
            {
                if (!dragCancelled)
                {
                    try { worksheets.ValidateMove(moveSession); }
                    catch (Exception exception) when (IsOperationError(exception)) { CancelCurrentDrag(OperationErrorMessage(exception)); }
                }
                return; // Never dispose/rebind a native control inside its OLE drag loop.
            }
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
                    if (onlyHandle.HasValue && onlyHandle.Value != handle) continue;
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
            catch (Exception exception) when (IsOperationError(exception))
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
                entry.DragStartingHandler = (sender, args) => StartDrag(handle, entry, args);
                entry.DropHandler = (sender, args) => DropSheet(entry, args);
                entry.DragCompletedHandler = (sender, args) => FinishDrag(entry);
                view.DragStarting += entry.DragStartingHandler;
                view.DropRequested += entry.DropHandler;
                view.DragCompleted += entry.DragCompletedHandler;
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
            if (disposed || operating || draggingEntry != null) return;
            try { entry.View.ShowStatus(worksheets.ActivateWorksheet(handle, worksheet.Id)); }
            catch (Exception exception) when (IsOperationError(exception))
            {
                Trace.TraceWarning("Sheaside activation: {0}", exception);
                entry.View.ShowStatus("シートを切り替えられません。セルの編集やダイアログを閉じて、もう一度お試しください。");
            }
            Refresh();
        }

        private void ExecuteCommand(int handle, PaneEntry entry, WorksheetCommandEventArgs args)
        {
            if (disposed || operating || draggingEntry != null) return;
            operating = true;
            string status = null;
            try
            {
                var target = worksheets.CaptureTarget(handle, args.Worksheet?.Id);
                if (args.Worksheet != null) entry.PreferredSelection = args.Worksheet.Id;
                switch (args.Command)
                {
                    case WorksheetCommand.Hide:
                        worksheets.ChangeVisibility(target, WorksheetVisibility.Visible);
                        status = "シートを非表示にしました。";
                        break;
                    case WorksheetCommand.Unhide:
                        worksheets.ChangeVisibility(target, WorksheetVisibility.Hidden);
                        status = "シートを再表示しました。";
                        break;
                    case WorksheetCommand.RevealVeryHidden:
                        string revealName = target.Sheet.Name;
                        if (MessageBox.Show(entry.View, "VeryHiddenのシート「" + revealName + "」を再表示しますか？",
                            "Sheaside — VeryHiddenの再表示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                        {
                            worksheets.ChangeVisibility(target, WorksheetVisibility.VeryHidden, revealName);
                            status = "VeryHiddenのシートを再表示しました。";
                        }
                        break;
                    case WorksheetCommand.SetTabColor:
                        using (var dialog = new ColorDialog { FullOpen = true })
                        {
                            int? color = worksheets.GetTabColor(target.Sheet);
                            dialog.Color = color.HasValue ? System.Drawing.ColorTranslator.FromOle(color.Value) : System.Drawing.SystemColors.Window;
                            if (dialog.ShowDialog(entry.View) == DialogResult.OK)
                            {
                                worksheets.SetTabColor(target, System.Drawing.ColorTranslator.ToOle(dialog.Color));
                                status = "タブの色を変更しました。";
                            }
                        }
                        break;
                    case WorksheetCommand.ClearTabColor:
                        worksheets.SetTabColor(target, null);
                        status = "タブの色を「色なし」にしました。";
                        break;
                    case WorksheetCommand.Add:
                        entry.PreferredSelection = worksheets.AddWorksheet(target);
                        status = "シートを追加しました。";
                        break;
                    case WorksheetCommand.Rename:
                        using (var dialog = new RenameSheetDialog(target.Sheet.Name, newName =>
                        {
                            try { worksheets.RenameWorksheet(target, newName); return null; }
                            catch (Exception exception) when (IsOperationError(exception)) { return OperationErrorMessage(exception); }
                        }))
                        {
                            if (dialog.ShowDialog(entry.View) == DialogResult.OK)
                            {
                                entry.PreferredSelection = worksheets.GetWorksheetId(target.Sheet);
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
                        string name = target.Sheet.Name;
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

        private void StartDrag(int handle, PaneEntry entry, WorksheetDragStartingEventArgs args)
        {
            if (disposed || operating || draggingEntry != null) return;
            try
            {
                moveSession = worksheets.CaptureMove(handle, args.WorksheetId, args.Order);
                draggingEntry = entry;
                dragCancelled = false;
                dragStatus = null;
                entry.PreferredSelection = args.WorksheetId;
                args.Cancel = false;
            }
            catch (Exception exception) when (IsOperationError(exception))
            {
                moveSession = null;
                entry.View.ShowStatus(OperationErrorMessage(exception));
            }
        }

        private void CancelCurrentDrag(string message)
        {
            if (draggingEntry == null || dragCancelled) return;
            dragCancelled = true;
            dragStatus = message;
            draggingEntry.View.CancelDrag();
        }

        private void DropSheet(PaneEntry entry, WorksheetDropEventArgs args)
        {
            if (disposed || operating || draggingEntry != entry || dragCancelled || moveSession == null) return;
            operating = true;
            try
            {
                bool moved = worksheets.MoveWorksheet(moveSession, args.Boundary);
                entry.PreferredSelection = worksheets.GetWorksheetId(moveSession.Target.Sheet);
                dragStatus = moved ? "シートを並び替えました。" : null;
            }
            catch (Exception exception) when (IsOperationError(exception))
            {
                dragStatus = OperationErrorMessage(exception);
            }
            finally { operating = false; }
            // FinishDrag reloads Excel's real order after DoDragDrop returns,
            // including when Move failed or an Excel event changed the workbook.
        }

        private void FinishDrag(PaneEntry entry)
        {
            if (draggingEntry != null && draggingEntry != entry) return;
            string status = dragStatus;
            draggingEntry = null;
            moveSession = null; // Release our short-lived COM references, not shared RCWs.
            dragStatus = null;
            dragCancelled = false;
            Refresh();
            if (!entry.View.IsDisposed && status != null) entry.View.ShowStatus(status);
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
            if (!operating) CancelCurrentDrag("パネルの表示状態が変わったため、並び替えを中止しました。");
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
            DetachDragHandlers(entry);
            entry.View.Dispose();
            entries.Remove(handle);
            StateChanged?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            CancelCurrentDrag("並び替えを中止しました。");
            moveSession = null;
            refreshTimer.Stop();
            refreshTimer.Tick -= OnRefresh;
            refreshTimer.Dispose();
            application.WindowActivate -= OnWindowActivate;
            application.WorkbookActivate -= OnWorkbookActivate;
            application.SheetActivate -= OnSheetActivate;
            application.WindowDeactivate -= OnWindowDeactivate;
            foreach (var entry in entries.Values)
            {
                entry.Pane.VisibleChanged -= OnVisibleChanged;
                entry.View.WorksheetRequested -= entry.WorksheetHandler;
                entry.View.CommandRequested -= entry.CommandHandler;
                DetachDragHandlers(entry);
            }
            entries.Clear();
            // VSTO owns task pane/control disposal at shutdown. Remove must not be
            // called here: the collection may already have been disposed by VSTO.
        }

        private static void DetachDragHandlers(PaneEntry entry)
        {
            entry.View.DragStarting -= entry.DragStartingHandler;
            entry.View.DropRequested -= entry.DropHandler;
            entry.View.DragCompleted -= entry.DragCompletedHandler;
        }
    }
}
