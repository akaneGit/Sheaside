using System;

namespace Sheaside.Models
{
    // Only managed values cross the boundary between Excel and the view.
    internal sealed class WorksheetItem
    {
        public WorksheetItem(string codeName, string name, bool isVisible)
        {
            CodeName = codeName;
            Name = name;
            IsVisible = isVisible;
        }

        public string CodeName { get; private set; }
        public string Name { get; private set; }
        public bool IsVisible { get; private set; }
        public string DisplayName { get { return IsVisible ? Name : Name + "（非表示）"; } }
    }

    internal sealed class WorksheetRequestedEventArgs : EventArgs
    {
        public WorksheetRequestedEventArgs(WorksheetItem worksheet) { Worksheet = worksheet; }
        public WorksheetItem Worksheet { get; private set; }
    }
}
