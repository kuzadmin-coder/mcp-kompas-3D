using System.ComponentModel;
using System.Text;
using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;
using ModelContextProtocol.Server;

namespace McpKompas.Tools;

[McpServerToolType]
public sealed class Drawing3DTools
{
    // ---------------------------------------------------------------
    //  Helpers
    // ---------------------------------------------------------------

    /// <summary>
    /// Returns the active 3D part context. On failure returns null and sets
    /// <paramref name="error"/> to an "ERROR: ..." string.
    /// </summary>
    private static (KompasObject k, ksDocument3D doc, ksPart part)? RequirePart(out string error)
    {
        error = "";
        var k = KompasConnector.Instance;
        if (k == null)
        {
            error = "ERROR: Not connected to Kompas. Call kompas_connect first.";
            return null;
        }
        var doc = (ksDocument3D?)k.ActiveDocument3D();
        if (doc == null)
        {
            error = "ERROR: No active 3D document. Create a part first with kompas_create_3d_part.";
            return null;
        }
        var part = (ksPart?)doc.GetPart((short)Part_Type.pTop_Part);
        if (part == null)
        {
            error = "ERROR: Could not get the top-level part from the active 3D document.";
            return null;
        }
        return (k, doc, part);
    }

    /// <summary>
    /// Returns the last entity of the given type, or null if none exist.
    /// </summary>
    private static ksEntity? GetLastEntity(ksPart part, Obj3dType type)
    {
        ksEntityCollection? coll = (ksEntityCollection?)part.EntityCollection((short)type);
        if (coll == null || coll.GetCount() == 0)
            return null;
        return (ksEntity)coll.GetByIndex(coll.GetCount() - 1);
    }

    // ---------------------------------------------------------------
    //  Sketch operations
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_create_sketch")]
    [Description("Create a sketch on one of the standard coordinate planes in the active 3D part.")]
    public static string CreateSketch(
        [Description("Plane: 'XOY', 'XOZ', or 'YOZ'.")]
        string plane = "XOY",
        [Description("Rotation angle of the sketch in degrees.")]
        double angle = 0.0)
    {
        var ctx = RequirePart(out var err);
        if (ctx == null) return err;
        var (_, _, part) = ctx.Value;

        Obj3dType planeType = plane.ToUpperInvariant() switch
        {
            "XOZ" => Obj3dType.o3d_planeXOZ,
            "YOZ" => Obj3dType.o3d_planeYOZ,
            _     => Obj3dType.o3d_planeXOY,
        };

        ksEntity? entitySketch = (ksEntity?)part.NewEntity((short)Obj3dType.o3d_sketch);
        if (entitySketch == null)
            return "ERROR: Could not create sketch entity.";

        ksSketchDefinition? sketchDef = (ksSketchDefinition?)entitySketch.GetDefinition();
        if (sketchDef == null)
            return "ERROR: Could not get SketchDefinition.";

        ksEntity? basePlane = (ksEntity?)part.GetDefaultEntity((short)planeType);
        if (basePlane == null)
            return $"ERROR: Could not get plane {plane}.";

        sketchDef.SetPlane(basePlane);
        sketchDef.angle = angle;
        if (!entitySketch.Create())
            return $"ERROR: Kompas rejected the sketch on plane {plane}.";

        UndoManager.Record3D($"Sketch on {plane} angle={angle}°", Obj3dType.o3d_sketch);
        return $"Sketch created on plane {plane}, angle={angle}°. " +
               "Call kompas_sketch_begin_edit to start drawing inside it.";
    }

    [McpServerTool(Name = "kompas_sketch_begin_edit")]
    [Description("Enter edit mode for the last created sketch. Use kompas_draw_* tools to draw 2D geometry inside it. Call kompas_sketch_end_edit when done.")]
    public static string SketchBeginEdit()
    {
        var ctx = RequirePart(out var err);
        if (ctx == null) return err;
        var (_, _, part) = ctx.Value;

        var sketch = GetLastEntity(part, Obj3dType.o3d_sketch);
        if (sketch == null)
            return "ERROR: No sketch found. Create a sketch first with kompas_create_sketch.";

        ksSketchDefinition? sketchDef = (ksSketchDefinition?)sketch.GetDefinition();
        if (sketchDef == null)
            return "ERROR: Could not get SketchDefinition.";

        ksDocument2D? sketchEdit = (ksDocument2D?)sketchDef.BeginEdit();
        if (sketchEdit == null)
            return "ERROR: Could not begin sketch editing.";

        SketchEditState.Current = (sketch, sketchDef, sketchEdit);
        return "Sketch edit mode started. Draw geometry with 2D drawing tools. " +
               "Call kompas_sketch_end_edit when done.";
    }

    [McpServerTool(Name = "kompas_sketch_end_edit")]
    [Description("Finish editing the current sketch and commit geometry to the 3D model.")]
    public static string SketchEndEdit()
    {
        var state = SketchEditState.Current;
        if (state == null)
            return "ERROR: No sketch is currently being edited.";

        state.Value.def.EndEdit();
        state.Value.sketch.Update();
        SketchEditState.Current = null;

        return "Sketch editing finished and geometry committed.";
    }

    // ---------------------------------------------------------------
    //  3D operations
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_extrude")]
    [Description("Apply a Boss-Extrusion operation to the last created sketch in the active 3D part.")]
    public static string Extrude(
        [Description("Extrusion depth in mm (forward direction).")]
        double depth,
        [Description("Direction: 'normal' (forward), 'reverse' (backward), 'both' (symmetric).")]
        string direction = "normal",
        [Description("Depth in mm for the reverse direction when direction='both'. 0 = same as depth.")]
        double depth_reverse = 0.0,
        [Description("Thin-wall extrusion: specify wall thickness > 0 to enable.")]
        double thin_wall = 0.0)
    {
        var ctx = RequirePart(out var err);
        if (ctx == null) return err;
        var (_, _, part) = ctx.Value;
        var sketch = GetLastEntity(part, Obj3dType.o3d_sketch);
        if (sketch == null)
            return "ERROR: No sketch found. Create a sketch first with kompas_create_sketch.";

        ksEntity? extrEntity = (ksEntity?)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
        if (extrEntity == null)
            return "ERROR: Could not create extrusion entity.";

        ksBossExtrusionDefinition? extrDef = (ksBossExtrusionDefinition?)extrEntity.GetDefinition();
        if (extrDef == null)
            return "ERROR: Could not get BossExtrusionDefinition.";

        Direction_Type dirType = direction.ToLowerInvariant() switch
        {
            "reverse" => Direction_Type.dtReverse,
            "both"    => Direction_Type.dtBoth,
            _         => Direction_Type.dtNormal,
        };

        extrDef.directionType = (short)dirType;

        // ksBossExtrusionDefinition.SetSideParam(bool normal, ...): true configures side 1
        // (the "normal"/forward side), false configures side 2 (the reverse side).
        // The depth must be set on the side that actually extrudes — otherwise Kompas
        // silently rejects the operation (depth=0 on the active side ⇒ no feature).
        if (dirType == Direction_Type.dtBoth)
        {
            double rev = depth_reverse > 0 ? depth_reverse : depth;
            extrDef.SetSideParam(true,  (short)End_Type.etBlind, depth, 0, false);
            extrDef.SetSideParam(false, (short)End_Type.etBlind, rev,   0, false);
        }
        else
        {
            // dtNormal → side 1 (true); dtReverse → side 2 (false)
            bool normalSide = dirType == Direction_Type.dtNormal;
            extrDef.SetSideParam(normalSide, (short)End_Type.etBlind, depth, 0, false);
        }

        if (thin_wall > 0)
            extrDef.SetThinParam(true, (short)Direction_Type.dtNormal, thin_wall, thin_wall);

        extrDef.SetSketch(sketch);
        bool ok = extrEntity.Create();
        if (!ok)
            return "ERROR: Kompas rejected the extrusion. Common causes: empty sketch, " +
                   "sketch has no closed contour, sketch contains only construction geometry, " +
                   "or the extrusion would produce zero volume. Inspect the sketch and try again.";

        UndoManager.Record3D($"Extrusion depth={depth}mm dir={direction}", Obj3dType.o3d_bossExtrusion);
        return $"Extrusion created: depth={depth}mm, direction={direction}" +
               (thin_wall > 0 ? $", thin wall={thin_wall}mm" : "") + ".";
    }

    [McpServerTool(Name = "kompas_cut_extrude")]
    [Description("Apply a Cut-Extrusion (pocket/hole) to the last created sketch in the active 3D part.")]
    public static string CutExtrude(
        [Description("Cut depth in mm.")]
        double depth,
        [Description("Direction: 'normal', 'reverse', or 'both'.")]
        string direction = "normal",
        [Description("Depth in mm for the reverse direction when direction='both'. 0 = same as depth.")]
        double depth_reverse = 0.0)
    {
        var ctx = RequirePart(out var err);
        if (ctx == null) return err;
        var (_, _, part) = ctx.Value;
        var sketch = GetLastEntity(part, Obj3dType.o3d_sketch);
        if (sketch == null)
            return "ERROR: No sketch found. Create a sketch first with kompas_create_sketch.";

        ksEntity? cutEntity = (ksEntity?)part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
        if (cutEntity == null)
            return "ERROR: Could not create cut extrusion entity.";

        ksCutExtrusionDefinition? cutDef = (ksCutExtrusionDefinition?)cutEntity.GetDefinition();
        if (cutDef == null)
            return "ERROR: Could not get CutExtrusionDefinition.";

        Direction_Type dirType = direction.ToLowerInvariant() switch
        {
            "reverse" => Direction_Type.dtReverse,
            "both"    => Direction_Type.dtBoth,
            _         => Direction_Type.dtNormal,
        };

        cutDef.directionType = (short)dirType;

        // Same side-selection rule as Extrude: depth must be set on the side that cuts.
        if (dirType == Direction_Type.dtBoth)
        {
            double rev = depth_reverse > 0 ? depth_reverse : depth;
            cutDef.SetSideParam(true,  (short)End_Type.etBlind, depth, 0, false);
            cutDef.SetSideParam(false, (short)End_Type.etBlind, rev,   0, false);
        }
        else
        {
            bool normalSide = dirType == Direction_Type.dtNormal;
            cutDef.SetSideParam(normalSide, (short)End_Type.etBlind, depth, 0, false);
        }

        cutDef.SetSketch(sketch);
        bool ok = cutEntity.Create();
        if (!ok)
            return "ERROR: Kompas rejected the cut extrusion. Common causes: empty sketch, " +
                   "sketch outside the existing body, or the cut would not intersect any material.";

        UndoManager.Record3D($"Cut extrusion depth={depth}mm dir={direction}", Obj3dType.o3d_cutExtrusion);
        return $"Cut extrusion created: depth={depth}mm, direction={direction}.";
    }

    [McpServerTool(Name = "kompas_revolve")]
    [Description("Apply a Rotation (revolve) operation to the last sketch. The sketch must contain an axis line (style=3).")]
    public static string Revolve(
        [Description("Angle of revolution in degrees (360 for full rotation).")]
        double angle = 360.0,
        [Description("Direction: 'normal', 'reverse', or 'both'.")]
        string direction = "normal",
        [Description("Angle in degrees for the reverse direction when direction='both'. 0 = same as angle.")]
        double angle_reverse = 0.0,
        [Description("Thin-wall revolution: wall thickness > 0 to enable.")]
        double thin_wall = 0.0)
    {
        var ctx = RequirePart(out var err);
        if (ctx == null) return err;
        var (_, _, part) = ctx.Value;
        var sketch = GetLastEntity(part, Obj3dType.o3d_sketch);
        if (sketch == null)
            return "ERROR: No sketch found. Create a sketch first with kompas_create_sketch.";

        ksEntity? rotEntity = (ksEntity?)part.NewEntity((short)Obj3dType.o3d_bossRotated);
        if (rotEntity == null)
            return "ERROR: Could not create rotation entity.";

        ksBossRotatedDefinition? rotDef = (ksBossRotatedDefinition?)rotEntity.GetDefinition();
        if (rotDef == null)
            return "ERROR: Could not get BossRotatedDefinition.";

        Direction_Type dirType = direction.ToLowerInvariant() switch
        {
            "reverse" => Direction_Type.dtReverse,
            "both"    => Direction_Type.dtBoth,
            _         => Direction_Type.dtNormal,
        };

        rotDef.directionType = (short)dirType;

        // Same side-selection rule as Extrude: angleNormal drives side 1, angleReverse
        // drives side 2. Setting only angleNormal when directionType=dtReverse leaves
        // the active side at angle=0, and Kompas silently rejects the operation.
        ksRotatedParam? rotProp = (ksRotatedParam?)rotDef.RotatedParam();
        if (rotProp != null)
        {
            rotProp.direction = (short)dirType;
            if (dirType == Direction_Type.dtBoth)
            {
                rotProp.angleNormal = angle;
                rotProp.angleReverse = angle_reverse > 0 ? angle_reverse : angle;
            }
            else if (dirType == Direction_Type.dtReverse)
            {
                rotProp.angleReverse = angle;
            }
            else
            {
                rotProp.angleNormal = angle;
            }
        }

        if (thin_wall > 0)
            rotDef.SetThinParam(true, (short)Direction_Type.dtNormal, thin_wall, thin_wall);

        rotDef.SetSketch(sketch);
        bool ok = rotEntity.Create();
        if (!ok)
            return "ERROR: Kompas rejected the revolution. Common causes: sketch has no axis " +
                   "line (style=3), the contour crosses the axis, or the contour is not closed.";

        UndoManager.Record3D($"Revolution angle={angle}° dir={direction}", Obj3dType.o3d_bossRotated);
        return $"Revolution created: angle={angle}°, direction={direction}" +
               (thin_wall > 0 ? $", thin wall={thin_wall}mm" : "") + ".";
    }

    [McpServerTool(Name = "kompas_fillet")]
    [Description("Apply a fillet (rounded edge) to edges near a specified 3D point. " +
                 "Use x=0,y=0,z=0 to select all edges of the part.")]
    public static string Fillet(
        [Description("Fillet radius in mm.")]
        double radius,
        [Description("X coordinate of a point near the edges to fillet (mm). Use 0 for all edges.")]
        double x = 0.0,
        [Description("Y coordinate of a point near the edges to fillet (mm). Use 0 for all edges.")]
        double y = 0.0,
        [Description("Z coordinate of a point near the edges to fillet (mm). Use 0 for all edges.")]
        double z = 0.0)
    {
        var ctx = RequirePart(out var err);
        if (ctx == null) return err;
        var (_, _, part) = ctx.Value;

        ksEntity? filletEntity = (ksEntity?)part.NewEntity((short)Obj3dType.o3d_fillet);
        if (filletEntity == null)
            return "ERROR: Could not create fillet entity.";

        ksFilletDefinition? filletDef = (ksFilletDefinition?)filletEntity.GetDefinition();
        if (filletDef == null)
            return "ERROR: Could not get FilletDefinition.";

        filletDef.radius = radius;
        filletDef.tangent = true;

        ksEntityCollection? edgeArr = (ksEntityCollection?)filletDef.array();
        if (edgeArr == null)
            return "ERROR: Could not get fillet edge collection.";

        // Sentinel: (0,0,0) explicitly means "all edges". Don't call SelectByPoint —
        // if origin happens to lie on an edge it would silently pick a single edge,
        // contradicting the documented behavior.
        bool selectAll = x == 0.0 && y == 0.0 && z == 0.0;
        int addedEdges = 0;

        if (selectAll)
        {
            ksEntityCollection? allEdges = (ksEntityCollection?)part.EntityCollection((short)Obj3dType.o3d_edge);
            if (allEdges != null)
            {
                for (int i = 0; i < allEdges.GetCount(); i++)
                {
                    edgeArr.Add(allEdges.GetByIndex(i));
                    addedEdges++;
                }
            }
        }
        else
        {
            ksEntityCollection? edges = (ksEntityCollection?)part.EntityCollection((short)Obj3dType.o3d_edge);
            if (edges != null && edges.SelectByPoint(x, y, z))
            {
                int count = edges.GetCount();
                for (int i = 0; i < count; i++)
                {
                    edgeArr.Add(edges.GetByIndex(i));
                    addedEdges++;
                }
            }
        }

        if (addedEdges == 0)
            return $"ERROR: No edges to fillet (point=({x},{y},{z})). " +
                   "Use x=0,y=0,z=0 only when the part actually has edges, or pick a point near a real edge.";

        if (!filletEntity.Create())
            return "ERROR: Kompas rejected the fillet. Common causes: no edges selected, " +
                   "radius larger than the smallest adjacent face, or tangency conflict.";
        UndoManager.Record3D($"Fillet radius={radius}mm", Obj3dType.o3d_fillet);
        return $"Fillet created with radius={radius}mm.";
    }

    [McpServerTool(Name = "kompas_chamfer")]
    [Description("Apply a chamfer (beveled edge) to faces near a specified 3D point.")]
    public static string Chamfer(
        [Description("Chamfer distance in mm.")]
        double distance,
        [Description("X coordinate of a point near the faces to chamfer (mm). Use 0 for all faces.")]
        double x = 0.0,
        [Description("Y coordinate of a point near the faces to chamfer (mm). Use 0 for all faces.")]
        double y = 0.0,
        [Description("Z coordinate of a point near the faces to chamfer (mm). Use 0 for all faces.")]
        double z = 0.0)
    {
        var ctx = RequirePart(out var err);
        if (ctx == null) return err;
        var (_, _, part) = ctx.Value;

        ksEntity? chamfEntity = (ksEntity?)part.NewEntity((short)Obj3dType.o3d_chamfer);
        if (chamfEntity == null)
            return "ERROR: Could not create chamfer entity.";

        ksChamferDefinition? chamfDef = (ksChamferDefinition?)chamfEntity.GetDefinition();
        if (chamfDef == null)
            return "ERROR: Could not get ChamferDefinition.";

        chamfDef.tangent = false;
        chamfDef.SetChamferParam(true, distance, distance);

        ksEntityCollection? faceArr = (ksEntityCollection?)chamfDef.array();
        if (faceArr == null)
            return "ERROR: Could not get chamfer face collection.";

        // Same (0,0,0) = "all faces" sentinel as Fillet — see comment there.
        bool selectAll = x == 0.0 && y == 0.0 && z == 0.0;
        int addedFaces = 0;

        if (selectAll)
        {
            ksEntityCollection? allFaces = (ksEntityCollection?)part.EntityCollection((short)Obj3dType.o3d_face);
            if (allFaces != null)
            {
                for (int i = 0; i < allFaces.GetCount(); i++)
                {
                    faceArr.Add(allFaces.GetByIndex(i));
                    addedFaces++;
                }
            }
        }
        else
        {
            ksEntityCollection? faces = (ksEntityCollection?)part.EntityCollection((short)Obj3dType.o3d_face);
            if (faces != null && faces.SelectByPoint(x, y, z))
            {
                int count = faces.GetCount();
                for (int i = 0; i < count; i++)
                {
                    faceArr.Add(faces.GetByIndex(i));
                    addedFaces++;
                }
            }
        }

        if (addedFaces == 0)
            return $"ERROR: No faces to chamfer (point=({x},{y},{z})). " +
                   "Use x=0,y=0,z=0 only when the part actually has faces, or pick a point near a real face.";

        if (!chamfEntity.Create())
            return "ERROR: Kompas rejected the chamfer. Common causes: no faces selected " +
                   "or distance larger than the adjacent geometry permits.";
        UndoManager.Record3D($"Chamfer distance={distance}mm", Obj3dType.o3d_chamfer);
        return $"Chamfer created with distance={distance}mm.";
    }

    // ---------------------------------------------------------------
    //  Part info
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_get_part_info")]
    [Description("Get information about the active 3D part: file, sketches, and feature counts.")]
    public static string GetPartInfo()
    {
        var ctx = RequirePart(out var err);
        if (ctx == null) return err;
        var (_, doc, part) = ctx.Value;
        var sb = new StringBuilder();

        sb.AppendLine($"File: {doc.fileName}");
        sb.AppendLine($"Author: {doc.author}");
        sb.AppendLine($"Comment: {doc.comment}");

        var countOf = (Obj3dType t) =>
        {
            var c = (ksEntityCollection?)part.EntityCollection((short)t);
            return c?.GetCount() ?? 0;
        };

        sb.AppendLine($"Sketches: {countOf(Obj3dType.o3d_sketch)}");
        sb.AppendLine($"Extrusions: {countOf(Obj3dType.o3d_bossExtrusion)}");
        sb.AppendLine($"Cut extrusions: {countOf(Obj3dType.o3d_cutExtrusion)}");
        sb.AppendLine($"Revolutions: {countOf(Obj3dType.o3d_bossRotated)}");
        sb.AppendLine($"Fillets: {countOf(Obj3dType.o3d_fillet)}");
        sb.AppendLine($"Chamfers: {countOf(Obj3dType.o3d_chamfer)}");
        sb.AppendLine($"Faces: {countOf(Obj3dType.o3d_face)}");
        sb.AppendLine($"Edges: {countOf(Obj3dType.o3d_edge)}");

        return sb.ToString().Trim();
    }
}

/// <summary>
/// Holds the currently open sketch edit session so that 2D draw tools
/// can redirect their output into the sketch editor. Thread-safe —
/// matters when running in HTTP mode where tools may be called concurrently.
/// </summary>
internal static class SketchEditState
{
    private static (ksEntity sketch, ksSketchDefinition def, ksDocument2D edit)? _current;
    private static readonly object _lock = new();

    public static (ksEntity sketch, ksSketchDefinition def, ksDocument2D edit)? Current
    {
        get { lock (_lock) { return _current; } }
        set { lock (_lock) { _current = value; } }
    }
}
