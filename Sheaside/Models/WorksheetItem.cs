using System;

namespace Sheaside.Models
{
    // Only managed values cross the boundary between Excel and the view.
    internal sealed class WorksheetItem
    {
        public WorksheetItem(string worksheetId, string name, bool isVisible)
        {
            Id = worksheetId;
            Name = name;
            IsVisible = isVisible;
        }

        public string Id { get; private set; }
        public string Name { get; private set; }
        public bool IsVisible { get; private set; }
        public string DisplayName { get { return IsVisible ? Name : Name + "（非表示）"; } }
    }

    internal sealed class WorksheetRequestedEventArgs : EventArgs
    {
        public WorksheetRequestedEventArgs(WorksheetItem worksheet) { Worksheet = worksheet; }
        public WorksheetItem Worksheet { get; private set; }
    }

    internal enum WorksheetCommand { Add, Rename, Duplicate, Delete }

    internal sealed class WorksheetCommandEventArgs : EventArgs
    {
        public WorksheetCommandEventArgs(WorksheetCommand command, WorksheetItem worksheet)
        {
            Command = command;
            Worksheet = worksheet;
        }

        public WorksheetCommand Command { get; private set; }
        public WorksheetItem Worksheet { get; private set; }
    }
}
