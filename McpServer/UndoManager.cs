using Kompas6API5;
using Kompas6Constants3D;
using KompasAPI7;
using McpKompas.Tools;

namespace McpKompas;

/// <summary>
/// Tracks operations performed via MCP tools and supports undoing them.
/// For 2D: stores object references that can be deleted with ksDeleteObj.
/// For 3D: stores entity references that can be deleted from the part's feature tree.
/// </summary>
public static class UndoManager
{
    public enum OpDomain { Drawing2D, Part3D }

    public record Operation(
        int Index,
        OpDomain Domain,
        string Description,
        /// <summary>
        /// For 2D: the reference returned by ksLineSeg, ksCircle, etc.
        /// For 3D: not used directly — we find the last entity by type.
        /// </summary>
        int ObjRef,
        /// <summary>
        /// For 3D operations: the Obj3dType of the created entity.
        /// </summary>
        Obj3dType? EntityType
    );

    private static readonly List<Operation> _history = new();
    private static int _nextIndex = 1;
    private static readonly object _lock = new();

    /// <summary>
    /// Record a 2D drawing operation.
    /// </summary>
    public static void Record2D(string description, int objRef)
    {
        lock (_lock)
        {
            _history.Add(new Operation(_nextIndex++, OpDomain.Drawing2D, description, objRef, null));
        }
    }

    /// <summary>
    /// Record a 2D drawing operation with multiple object references (e.g. rectangle = 4 lines).
    /// Each ref becomes its own history entry sharing the same description.
    /// </summary>
    public static void Record2DMulti(string description, params int[] objRefs)
    {
        lock (_lock)
        {
            int idx = _nextIndex++;
            foreach (var r in objRefs)
                _history.Add(new Operation(idx, OpDomain.Drawing2D, description, r, null));
        }
    }

    /// <summary>
    /// Record a 3D feature operation (extrude, cut, fillet, etc.).
    /// </summary>
    public static void Record3D(string description, Obj3dType entityType)
    {
        lock (_lock)
        {
            _history.Add(new Operation(_nextIndex++, OpDomain.Part3D, description, 0, entityType));
        }
    }

    /// <summary>
    /// Get the operation history (newest first).
    /// </summary>
    public static List<Operation> GetHistory(int limit = 20)
    {
        lock (_lock)
        {
            return _history
                .GroupBy(o => o.Index)
                .Select(g => g.First())
                .OrderByDescending(o => o.Index)
                .Take(limit)
                .ToList();
        }
    }

    /// <summary>
    /// Undo the last N logical operations.
    /// Returns a description of what was undone or an error message.
    /// </summary>
    public static string Undo(int count = 1)
    {
        lock (_lock)
        {
            if (_history.Count == 0)
                return "Nothing to undo.";

            var undone = new List<string>();

            for (int i = 0; i < count; i++)
            {
                if (_history.Count == 0) break;

                // Find the highest index (last logical operation)
                int lastIdx = _history.Max(o => o.Index);
                var ops = _history.Where(o => o.Index == lastIdx).ToList();

                if (ops.Count == 0) break;

                var first = ops[0];
                string result;

                if (first.Domain == OpDomain.Drawing2D)
                    result = Undo2D(ops);
                else
                    result = Undo3D(first);

                undone.Add($"#{lastIdx} {first.Description}: {result}");
                _history.RemoveAll(o => o.Index == lastIdx);
            }

            return undone.Count > 0
                ? "Undone:\n" + string.Join("\n", undone)
                : "Nothing to undo.";
        }
    }

    /// <summary>
    /// Clear all history (e.g. when document is closed).
    /// </summary>
    public static void Clear()
    {
        lock (_lock)
        {
            _history.Clear();
            _nextIndex = 1;
        }
    }

    /// <summary>
    /// Forget a 2D object reference (e.g. after it was manually deleted).
    /// Prevents a later undo from trying to delete an already-gone object.
    /// </summary>
    public static void Forget(int objRef)
    {
        if (objRef == 0) return;
        lock (_lock)
        {
            _history.RemoveAll(o => o.Domain == OpDomain.Drawing2D && o.ObjRef == objRef);
        }
    }

    // ------------------------------------------------------------------

    private static string Undo2D(List<Operation> ops)
    {
        // We need the active 2D doc to delete objects
        var k = KompasConnector.Instance;
        if (k == null) return "not connected";

        ksDocument2D? doc;

        // Snapshot sketch edit state once to avoid races
        var sketchState = SketchEditState.Current;
        if (sketchState.HasValue)
            doc = sketchState.Value.edit;
        else
            doc = (ksDocument2D?)k.ActiveDocument2D();

        if (doc == null)
            return "no active 2D document";

        int deleted = 0;
        foreach (var op in ops)
        {
            if (op.ObjRef != 0 && doc.ksDeleteObj(op.ObjRef) != 0)
                deleted++;
        }

        return $"deleted {deleted}/{ops.Count} object(s)";
    }

    // Kompas TransferInterface "iCycle" constant: 2 = active document context.
    private const int CYC_ACTIVE_DOCUMENT = 2;

    private static string Undo3D(Operation op)
    {
        var k = KompasConnector.Instance;
        if (k == null) return "not connected";

        var doc = (ksDocument3D?)k.ActiveDocument3D();
        if (doc == null) return "no active 3D document";

        var part = (ksPart?)doc.GetPart((short)Part_Type.pTop_Part);
        if (part == null) return "no top-level part";

        if (op.EntityType == null)
            return "unknown entity type";

        // Find the last entity of this type
        var coll = (ksEntityCollection?)part.EntityCollection((short)op.EntityType.Value);
        if (coll == null || coll.GetCount() == 0)
            return "no entity found";

        var entity = (ksEntity)coll.GetByIndex(coll.GetCount() - 1);

        // Transfer to API7 and delete via IFeature7. Delete() returns bool — must check.
        try
        {
            var feature = (IFeature7?)k.TransferInterface(entity, CYC_ACTIVE_DOCUMENT, 0);
            if (feature == null)
                return "delete failed — could not transfer to API7";
            return feature.Delete()
                ? "deleted"
                : "delete failed — Kompas rejected feature.Delete()";
        }
        catch (Exception ex)
        {
            return $"delete failed — {ex.Message}";
        }
    }
}
