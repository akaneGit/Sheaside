using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Sheaside.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sheaside.Services
{
    internal sealed class ExcelSheetBatchService
    {
        internal sealed class Plan
        {
            internal Plan(ExcelWorksheetService.OperationTarget context, SheetBatchCommand command, WorksheetItem[] items)
            { Context = context; Command = command; Items = Array.AsReadOnly(items); }
            internal ExcelWorksheetService.OperationTarget Context { get; private set; }
            public SheetBatchCommand Command { get; private set; }
            public ReadOnlyCollection<WorksheetItem> Items { get; private set; }
            public bool NeedsConfirmation { get { return Command == SheetBatchCommand.Delete; } }
        }

        private readonly ExcelWorksheetService sheets;
        private readonly Func<bool> allowVeryHidden;
        private bool executing;
        public ExcelSheetBatchService(ExcelWorksheetService sheets, Func<bool> allowVeryHidden = null)
        { this.sheets = sheets; this.allowVeryHidden = allowVeryHidden ?? (() => true); }

        public Plan Capture(int handle, IEnumerable<string> ids, SheetBatchCommand command)
        {
            if (!Enum.IsDefined(typeof(SheetBatchCommand), command)) throw new InvalidOperationException("未対応の一括操作です。");
            var context = sheets.CaptureTarget(handle, null);
            var requested = new HashSet<string>(ids);
            var items = sheets.GetWorksheets(context.Workbook).Where(i => requested.Contains(i.Id)).ToArray();
            if (items.Length == 0 || items.Length != requested.Count)
                throw new InvalidOperationException("選択対象が変わりました。一覧を更新して選び直してください。");
            var plan = new Plan(context, command, items); // Frozen source order; never enumerate copies as sources.
            ValidateRemaining(plan, 0);
            return plan;
        }

        private ExcelWorksheetService.OperationTarget Resolve(Plan plan, WorksheetItem item)
        {
            sheets.ValidateTarget(plan.Context);
            var target = sheets.CaptureTarget(plan.Context.WindowHandle, item.Id);
            if (!Equals(target.Workbook, plan.Context.Workbook) || !Equals(target.Sheet.Parent, plan.Context.Workbook))
                throw new InvalidOperationException("対象ブックが変わったため中断しました。");
            if (target.Sheet.Name != item.Name || (int)target.Sheet.Visible != (int)item.Visibility || target.Sheet.Kind != item.Kind)
                throw new InvalidOperationException("対象「" + item.Name + "」の名前または表示状態が変わりました。確認し直してください。");
            return target;
        }

        private void ValidateRemaining(Plan plan, int start)
        {
            sheets.ValidateTarget(plan.Context);
            if (!allowVeryHidden() && (plan.Command == SheetBatchCommand.VeryHide || plan.Items.Any(i => i.Visibility == WorksheetVisibility.VeryHidden)))
                throw new InvalidOperationException("VeryHiddenの管理には設定を有効にしてください。");
            for (int i = start; i < plan.Items.Count; i++) Resolve(plan, plan.Items[i]);
            if (plan.Command == SheetBatchCommand.Delete || plan.Command == SheetBatchCommand.Hide || plan.Command == SheetBatchCommand.VeryHide)
            {
                var pending = new HashSet<string>(plan.Items.Skip(start).Select(i => i.Id));
                var current = sheets.GetWorksheets(plan.Context.Workbook);
                if (!current.Any(i => i.IsVisible && !pending.Contains(i.Id)))
                    throw new InvalidOperationException("操作後に最低1枚の表示シートを残してください。");
            }
        }

        public SheetBatchResult Execute(Plan plan, bool confirmed, int? color = null)
        {
            if (executing) throw new InvalidOperationException("一括処理は既に実行中です。");
            executing = true;
            try
            {
                return SheetBatchRunner.Run(plan.Items, () =>
                {
                    if (plan.NeedsConfirmation && !confirmed) throw new InvalidOperationException("対象の確認が必要です。変更していません。");
                    if (plan.Command == SheetBatchCommand.SetColor && !color.HasValue) throw new InvalidOperationException("色を選択してください。");
                    ValidateRemaining(plan, 0); // All selected IDs re-resolved AFTER every dialog, before any mutation.
                }, (index, outcome) =>
                {
                    ValidateRemaining(plan, index);
                    var target = Resolve(plan, outcome.Item); // Resolve again immediately before the individual operation.
                    bool skipped = false;
                    switch (plan.Command)
                    {
                        case SheetBatchCommand.Delete:
                            sheets.DeleteWorksheet(target, outcome.Item.Name);
                            if (sheets.GetWorksheets(plan.Context.Workbook).Any(i => i.Id == outcome.Item.Id))
                                throw new InvalidOperationException("Excelが削除を完了しませんでした。");
                            break;
                        case SheetBatchCommand.Hide:
                        case SheetBatchCommand.VeryHide:
                            var visibility = plan.Command == SheetBatchCommand.VeryHide ? WorksheetVisibility.VeryHidden : WorksheetVisibility.Hidden;
                            skipped = outcome.Item.Visibility == visibility;
                            if (!skipped) sheets.SetVisibility(target, visibility);
                            if ((int)target.Sheet.Visible != (int)visibility)
                                throw new InvalidOperationException("非表示への変更を確認できませんでした。");
                            break;
                        case SheetBatchCommand.Unhide:
                            skipped = outcome.Item.IsVisible;
                            if (!skipped) sheets.ChangeVisibility(target, outcome.Item.Visibility, outcome.Item.Name);
                            if (target.Sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                                throw new InvalidOperationException("再表示への変更を確認できませんでした。");
                            break;
                        case SheetBatchCommand.SetColor:
                        case SheetBatchCommand.ClearColor:
                            int? desired = plan.Command == SheetBatchCommand.ClearColor ? null : color;
                            skipped = sheets.GetTabColor(target.Sheet) == desired;
                            if (!skipped) sheets.SetTabColor(target, desired);
                            if (sheets.GetTabColor(target.Sheet) != desired)
                                throw new InvalidOperationException("タブ色の変更を確認できませんでした。");
                            break;
                        case SheetBatchCommand.Duplicate:
                            var beforeIds = new HashSet<string>(sheets.GetWorksheets(plan.Context.Workbook).Select(i => i.Id));
                            outcome.CreatedId = sheets.DuplicateWorksheet(target);
                            var copy = sheets.CaptureTarget(plan.Context.WindowHandle, outcome.CreatedId);
                            var currentOrder = sheets.GetWorksheets(plan.Context.Workbook).ToArray();
                            var added = currentOrder.Where(i => !beforeIds.Contains(i.Id)).ToArray();
                            int originalIndex = Array.FindIndex(currentOrder, i => i.Id == outcome.Item.Id);
                            if (!Equals(copy.Workbook, plan.Context.Workbook) || added.Length != 1 ||
                                added[0].Id != outcome.CreatedId || copy.Sheet.Kind != outcome.Item.Kind ||
                                originalIndex < 0 || originalIndex + 1 >= currentOrder.Length || currentOrder[originalIndex + 1].Id != outcome.CreatedId)
                                throw new InvalidOperationException("複製先を確認できませんでした。");
                            outcome.Detail = "複製先：" + copy.Sheet.Name;
                            break;
                    }
                    outcome.Status = skipped ? SheetBatchStatus.Skipped : SheetBatchStatus.Succeeded;
                    if (skipped) outcome.Detail = "既に指定の状態です。";
                });
            }
            finally { executing = false; }
        }
    }
}
