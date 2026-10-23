using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Sheaside.Models;

namespace Sheaside.Services
{
    // Shared stop-on-first-error execution and complete accounting, independent of COM/UI.
    internal static class SheetBatchRunner
    {
        public static SheetBatchResult Run(IEnumerable<WorksheetItem> items, Action preflight, Action<int, SheetBatchItemResult> apply)
        {
            var result = new SheetBatchResult(items);
            try { preflight(); }
            catch (Exception e) when (IsOperationError(e)) { result.Error = e.Message; return result; }
            for (int index = 0; index < result.Items.Count; index++)
            {
                var item = result.Items[index];
                try { apply(index, item); }
                catch (Exception e) when (IsOperationError(e))
                {
                    item.Status = SheetBatchStatus.Failed;
                    item.Detail = e.Message;
                    result.Error = "処理を中断しました。成功した変更は元に戻していません。失敗対象もExcelの実状態を確認してください。";
                    break;
                }
            }
            return result;
        }
        private static bool IsOperationError(Exception e)
        { return e is COMException || e is InvalidOperationException || e is InvalidComObjectException; }
    }
}
