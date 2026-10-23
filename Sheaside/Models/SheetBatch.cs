using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Sheaside.Models
{
    internal enum SheetBatchCommand { Delete, Hide, Unhide, SetColor, ClearColor, Duplicate, VeryHide }
    internal enum SheetBatchStatus { NotProcessed, Succeeded, Skipped, Failed }

    internal sealed class SheetBatchEventArgs : EventArgs
    {
        public SheetBatchEventArgs(SheetBatchCommand command, IEnumerable<string> ids)
        { Command = command; Ids = Array.AsReadOnly(ids.ToArray()); }
        public SheetBatchCommand Command { get; private set; }
        public ReadOnlyCollection<string> Ids { get; private set; }
    }

    internal sealed class SheetBatchItemResult
    {
        public SheetBatchItemResult(WorksheetItem item) { Item = item; }
        public WorksheetItem Item { get; private set; }
        public SheetBatchStatus Status { get; internal set; }
        public string Detail { get; internal set; }
        public string CreatedId { get; internal set; }
    }

    internal sealed class SheetBatchResult
    {
        public SheetBatchResult(IEnumerable<WorksheetItem> items)
        { Items = Array.AsReadOnly(items.Select(i => new SheetBatchItemResult(i)).ToArray()); }
        public ReadOnlyCollection<SheetBatchItemResult> Items { get; private set; }
        public string Error { get; internal set; }
        public string Summary { get { return "成功 " + Count(SheetBatchStatus.Succeeded) + "件 / 失敗 " + Count(SheetBatchStatus.Failed)
            + "件 / スキップ " + Count(SheetBatchStatus.Skipped) + "件 / 未処理 " + Count(SheetBatchStatus.NotProcessed) + "件"; } }
        public int Count(SheetBatchStatus status) { return Items.Count(i => i.Status == status); }
        public string Details
        {
            get { return (Error == null ? "" : Error + "\r\n\r\n") + string.Join("\r\n", Items.Select(i =>
                StatusName(i.Status) + "：" + i.Item.DisplayName + (string.IsNullOrEmpty(i.Detail) ? "" : " — " + i.Detail))); }
        }
        private static string StatusName(SheetBatchStatus status)
        {
            switch (status)
            {
                case SheetBatchStatus.Succeeded: return "成功";
                case SheetBatchStatus.Skipped: return "スキップ";
                case SheetBatchStatus.Failed: return "失敗（実状態を要確認）";
                default: return "未処理";
            }
        }
    }
}
