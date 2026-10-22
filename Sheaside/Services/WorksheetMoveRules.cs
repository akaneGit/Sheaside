using System;

namespace Sheaside.Services
{
    internal static class WorksheetMoveRules
    {
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
