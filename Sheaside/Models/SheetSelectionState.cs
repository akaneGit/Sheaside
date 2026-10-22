using System.Collections.Generic;
using System.Linq;

namespace Sheaside.Models
{
    // One instance per task-pane/window. No Excel group selection or COM references.
    internal sealed class SheetSelectionState
    {
        private readonly HashSet<string> selected = new HashSet<string>();
        private bool initialized;
        public string ActiveId { get; private set; }
        public string AnchorId { get; set; }
        public bool PreserveSelection { get; set; }
        public string[] SelectedIds { get { return selected.ToArray(); } }

        public void Capture(IEnumerable<string> ids)
        {
            selected.Clear();
            selected.UnionWith(ids);
            if (selected.Count > 1) PreserveSelection = true;
        }

        public void Reconcile(IEnumerable<string> existingIds, string activeId, string preferredId)
        {
            var existing = new HashSet<string>(existingIds);
            bool hadSelection = selected.Count > 0;
            selected.IntersectWith(existing);
            if (!initialized || (!PreserveSelection && (ActiveId != activeId || preferredId != null)) ||
                (hadSelection && selected.Count == 0))
            {
                selected.Clear();
                string id = preferredId != null && existing.Contains(preferredId) ? preferredId : activeId;
                if (id != null && existing.Contains(id)) selected.Add(id);
            }
            if (AnchorId == null || !existing.Contains(AnchorId)) AnchorId = selected.FirstOrDefault();
            ActiveId = activeId;
            initialized = true;
        }

        public void ResetToActive()
        {
            selected.Clear();
            if (ActiveId != null) selected.Add(ActiveId);
            AnchorId = ActiveId;
            PreserveSelection = false;
        }
    }
}
