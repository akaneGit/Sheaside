using System;
using System.Collections.Generic;
using System.Linq;

namespace Sheaside.Services
{
    internal static class WorksheetMoveRules
    {
        public static string[] GroupOrder(IList<string> full, IList<string> displayed, IList<string> selected, int boundary)
        {
            if (full == null || displayed == null || selected == null || displayed.Count == 0 || selected.Count == 0 ||
                boundary < 0 || boundary > displayed.Count || full.Any(string.IsNullOrEmpty) ||
                full.Distinct().Count() != full.Count || displayed.Distinct().Count() != displayed.Count ||
                selected.Distinct().Count() != selected.Count || selected.Any(id => !displayed.Contains(id)) ||
                !full.Where(displayed.Contains).SequenceEqual(displayed))
                throw new InvalidOperationException("移動対象または一覧が変更されています。更新後にもう一度操作してください。");
            // A filtered trailing gap means immediately after its last displayed tab.
            int gap = boundary < displayed.Count ? full.IndexOf(displayed[boundary]) : full.IndexOf(displayed.Last()) + 1;
            var moving = new HashSet<string>(selected);
            int insertion = full.Take(gap).Count(id => !moving.Contains(id));
            var remaining = full.Where(id => !moving.Contains(id)).ToArray();
            return remaining.Take(insertion).Concat(full.Where(moving.Contains)).Concat(remaining.Skip(insertion)).ToArray();
        }

        public static int DestinationIndex(int sourceIndex, int boundary, int count)
        {
            if (count < 1 || sourceIndex < 0 || sourceIndex >= count || boundary < 0 || boundary > count)
                throw new ArgumentOutOfRangeException(nameof(boundary));
            return boundary > sourceIndex ? boundary - 1 : boundary;
        }

        public static int BoundaryAt(int hitIndex, int mouseY, int rowTop, int rowHeight, int count)
        {
            if (hitIndex < 0) return count; // Blank space below the last row.
            if (hitIndex >= count || rowHeight <= 0) throw new ArgumentOutOfRangeException(nameof(hitIndex));
            return hitIndex + (mouseY - rowTop >= rowHeight / 2 ? 1 : 0);
        }
    }
}
