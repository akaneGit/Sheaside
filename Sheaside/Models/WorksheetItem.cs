using System;

namespace Sheaside.Models
{
    internal enum SheetKind { Worksheet, Chart }
    internal enum WorksheetVisibility { Visible = -1, Hidden = 0, VeryHidden = 2 }
    // Only managed values cross the boundary between Excel and the view.
    internal sealed class WorksheetItem
    {
        public WorksheetItem(string worksheetId, string name, bool isVisible)
            : this(worksheetId, name, isVisible ? WorksheetVisibility.Visible : WorksheetVisibility.Hidden, null) { }

        public WorksheetItem(string worksheetId, string name, WorksheetVisibility visibility, int? tabColor, SheetKind kind = SheetKind.Worksheet, bool isProtected = false)
        {
            Id = worksheetId;
            Name = name;
            Visibility = visibility;
            TabColor = tabColor;
            Kind = kind;
            IsProtected = isProtected;
        }

        public string Id { get; private set; }
        public string Name { get; private set; }
        public SheetKind Kind { get; private set; }
        public bool IsProtected { get; private set; }
        public WorksheetVisibility Visibility { get; private set; }
        // OLE RGB value; null is different from black (0).
        public int? TabColor { get; private set; }
        public bool IsVisible { get { return Visibility == WorksheetVisibility.Visible; } }
        public string AccessibleDescription { get { return Name + " — " + (Kind == SheetKind.Chart ? "グラフシート" : "ワークシート") + " / " +
            (IsVisible ? "表示（Visible）" : Visibility == WorksheetVisibility.Hidden ? "非表示（Hidden）" : "強制非表示（VeryHidden）") + (IsProtected ? " / シート保護あり" : ""); } }
        public string DisplayName { get { return (Kind == SheetKind.Chart ? "［グラフ］" : "") + Name + (IsVisible ? "" : (Visibility == WorksheetVisibility.VeryHidden ? "（VeryHidden）" : "（非表示）")); } }
    }

    internal sealed class WorksheetRequestedEventArgs : EventArgs
    {
        public WorksheetRequestedEventArgs(WorksheetItem worksheet) { Worksheet = worksheet; }
        public WorksheetItem Worksheet { get; private set; }
    }

    internal enum WorksheetCommand { Add, Rename, Duplicate, Delete, Hide, Unhide, RevealVeryHidden, SetTabColor, ClearTabColor }

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
