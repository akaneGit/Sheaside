using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Sheaside.Models
{
    internal sealed class WorksheetDragStartingEventArgs : CancelEventArgs
    {
        public WorksheetDragStartingEventArgs(string worksheetId, IList<string> order) : base(true)
        {
            WorksheetId = worksheetId;
            Order = new List<string>(order).AsReadOnly();
        }

        public string WorksheetId { get; private set; }
        public IList<string> Order { get; private set; }
    }

    internal sealed class WorksheetDropEventArgs : EventArgs
    {
        public WorksheetDropEventArgs(int boundary) { Boundary = boundary; }
        // Gap in the original, unmodified list: 0..Count (not a final row index).
        public int Boundary { get; private set; }
    }
}
