using System.ComponentModel;
using System.Text;
using Kompas6API5;
using Kompas6Constants;
using KAPITypes;
using ModelContextProtocol.Server;

namespace McpKompas.Tools;

[McpServerToolType]
public sealed class Drawing2DTools
{
    // ---------------------------------------------------------------
    //  Helpers
    // ---------------------------------------------------------------

    /// <summary>
    /// Returns the active 2D document, checking first for an open sketch edit session.
    /// On failure returns null and sets <paramref name="error"/> to an "ERROR: ..." string
    /// that the caller can return directly to the MCP client.
    /// </summary>
    private static ksDocument2D? RequireDoc2D(out string error)
    {
        error = "";
        var sketchState = SketchEditState.Current;
        if (sketchState.HasValue)
            return sketchState.Value.edit;

        var k = KompasConnector.Instance;
        if (k == null)
        {
            error = "ERROR: Not connected to Kompas. Call kompas_connect first.";
            return null;
        }
        var doc = (ksDocument2D?)k.ActiveDocument2D();
        if (doc == null)
        {
            error = "ERROR: No active 2D document. Create or open a drawing/fragment first, " +
                    "or start a sketch session with kompas_sketch_begin_edit.";
            return null;
        }
        return doc;
    }

    // ---------------------------------------------------------------
    //  Layers and Views
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_set_layer")]
    [Description("Set the active drawing layer. Standard layers: 0=base, 1=auxiliary, 2=dimensioning, 3=center lines, 4=hatching, 5=invisible.")]
    public static string SetLayer(
        [Description("Layer number (0–255).")]
        int layer)
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        doc.ksLayer(layer);
        return $"Layer set to {layer}.";
    }

    [McpServerTool(Name = "kompas_create_view")]
    [Description("Create a new user view in the active 2D document.")]
    public static string CreateView(
        [Description("View number to assign (2 and above; 0 and 1 are reserved).")]
        int number,
        [Description("X coordinate of view origin (mm).")]
        double x,
        [Description("Y coordinate of view origin (mm).")]
        double y,
        [Description("View scale, e.g. 1.0 for 1:1, 0.5 for 1:2.")]
        double scale = 1.0,
        [Description("View rotation angle in degrees.")]
        double angle = 0.0,
        [Description("Human-readable name for the view.")]
        string name = "")
    {
        if (number < 2)
            return "ERROR: View number must be 2 or greater (0 and 1 are reserved by Kompas).";

        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;

        ksViewParam? par = (ksViewParam?)k.GetParamStruct((short)StructType2DEnum.ko_ViewParam);
        if (par == null)
            return "ERROR: Could not get ViewParam struct.";

        par.Init();
        par.x = x;
        par.y = y;
        par.scale_ = scale;
        par.angle = angle;
        par.state = ldefin2d.stACTIVE;
        par.name = name;
        int n = number;
        // ksCreateSheetView returns the new view's reference, or 0 on failure
        int viewRef = doc.ksCreateSheetView(par, ref n);
        if (viewRef == 0)
            return $"ERROR: Kompas could not create view {number}. " +
                   "The number may already be taken or the document may not support views (fragments don't).";

        return $"View {number} created at ({x},{y}), scale={scale}, angle={angle}.";
    }

    // ---------------------------------------------------------------
    //  Basic geometry
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_draw_line_segment")]
    [Description("Draw a line segment in the active 2D document or current sketch.")]
    public static string DrawLineSegment(
        [Description("Start X (mm).")]
        double x1,
        [Description("Start Y (mm).")]
        double y1,
        [Description("End X (mm).")]
        double x2,
        [Description("End Y (mm).")]
        double y2,
        [Description("Line style: 1=solid, 2=dashed, 3=chain, 4=chain thick, 5=dashed dotted, 6=invisible.")]
        int style = 1)
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        int ref_ = doc.ksLineSeg(x1, y1, x2, y2, style);
        if (ref_ != 0)
        {
            UndoManager.Record2D($"Line ({x1},{y1})-({x2},{y2})", ref_);
            KompasConnector.RefreshView();
            return $"Line drawn from ({x1},{y1}) to ({x2},{y2}), ref={ref_}.";
        }
        return "ERROR: Failed to draw line segment.";
    }

    [McpServerTool(Name = "kompas_draw_circle")]
    [Description("Draw a circle in the active 2D document or current sketch.")]
    public static string DrawCircle(
        [Description("Center X (mm).")]
        double cx,
        [Description("Center Y (mm).")]
        double cy,
        [Description("Radius (mm).")]
        double radius,
        [Description("Line style: 1=solid, 2=dashed, 3=chain, 4=chain thick, 6=invisible.")]
        int style = 1)
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        int ref_ = doc.ksCircle(cx, cy, radius, style);
        if (ref_ != 0)
        {
            UndoManager.Record2D($"Circle ({cx},{cy}) r={radius}", ref_);
            KompasConnector.RefreshView();
            return $"Circle drawn at ({cx},{cy}), r={radius}, ref={ref_}.";
        }
        return "ERROR: Failed to draw circle.";
    }

    [McpServerTool(Name = "kompas_draw_arc")]
    [Description("Draw an arc by center, radius, start angle and end angle.")]
    public static string DrawArc(
        [Description("Center X (mm).")]
        double cx,
        [Description("Center Y (mm).")]
        double cy,
        [Description("Radius (mm).")]
        double radius,
        [Description("Start angle in degrees (counter-clockwise from positive X axis).")]
        double angle1,
        [Description("End angle in degrees (counter-clockwise from positive X axis).")]
        double angle2,
        [Description("Direction: 1=counter-clockwise, -1=clockwise.")]
        int direction = 1,
        [Description("Line style: 1=solid, 2=dashed, 3=chain.")]
        int style = 1)
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        // ksArcByAngle direction param is short
        int ref_ = doc.ksArcByAngle(cx, cy, radius, angle1, angle2, (short)direction, style);
        if (ref_ != 0)
        {
            UndoManager.Record2D($"Arc ({cx},{cy}) r={radius} {angle1}°-{angle2}°", ref_);
            KompasConnector.RefreshView();
            return $"Arc drawn at ({cx},{cy}), r={radius}, {angle1}°-{angle2}°, ref={ref_}.";
        }
        return "ERROR: Failed to draw arc.";
    }

    [McpServerTool(Name = "kompas_draw_rectangle")]
    [Description("Draw an axis-aligned rectangle as four line segments.")]
    public static string DrawRectangle(
        [Description("Left edge X (mm).")]
        double x,
        [Description("Bottom edge Y (mm).")]
        double y,
        [Description("Width (mm).")]
        double width,
        [Description("Height (mm).")]
        double height,
        [Description("Line style: 1=solid, 2=dashed.")]
        int style = 1)
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        int r1 = doc.ksLineSeg(x,          y,          x + width, y,          style);
        int r2 = doc.ksLineSeg(x + width,  y,          x + width, y + height, style);
        int r3 = doc.ksLineSeg(x + width,  y + height, x,         y + height, style);
        int r4 = doc.ksLineSeg(x,          y + height, x,         y,          style);
        UndoManager.Record2DMulti($"Rectangle ({x},{y}) {width}x{height}", r1, r2, r3, r4);
        KompasConnector.RefreshView();
        return $"Rectangle drawn at ({x},{y}), {width}x{height}.";
    }

    [McpServerTool(Name = "kompas_draw_infinite_line")]
    [Description("Draw a construction (infinite) line through a point at a given angle. Useful for center lines and axes.")]
    public static string DrawInfiniteLine(
        [Description("X coordinate of the point the line passes through (mm).")]
        double x,
        [Description("Y coordinate of the point the line passes through (mm).")]
        double y,
        [Description("Angle in degrees (0=horizontal, 90=vertical).")]
        double angle)
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        // ksLine takes (x, y, angle) — no style parameter
        int ref_ = doc.ksLine(x, y, angle);
        if (ref_ != 0)
        {
            UndoManager.Record2D($"Infinite line ({x},{y}) {angle}°", ref_);
            KompasConnector.RefreshView();
            return $"Infinite line through ({x},{y}) at {angle}°, ref={ref_}.";
        }
        return "ERROR: Failed to draw infinite line.";
    }

    // ---------------------------------------------------------------
    //  Text and annotations
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_add_text")]
    [Description("Add a simple text annotation at the given position.")]
    public static string AddText(
        [Description("X coordinate (mm).")]
        double x,
        [Description("Y coordinate (mm).")]
        double y,
        [Description("The text string to insert.")]
        string text,
        [Description("Font height in mm (e.g. 3.5, 5.0, 7.0).")]
        double height = 3.5,
        [Description("Rotation angle in degrees.")]
        double angle = 0.0)
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        // ksText(x, y, ang, hStr, ksuStr, bitVector, s)
        // ksuStr = character spacing ratio (1.0 = normal), bitVector = 0 (no special flags)
        int ref_ = doc.ksText(x, y, angle, height, 1.0, 0, text);
        if (ref_ != 0)
        {
            UndoManager.Record2D($"Text \"{text}\"", ref_);
            KompasConnector.RefreshView();
            return $"Text added at ({x},{y}): \"{text}\", ref={ref_}.";
        }
        return "ERROR: Failed to add text.";
    }

    [McpServerTool(Name = "kompas_add_linear_dimension")]
    [Description("Add a linear dimension between two points in the active 2D document.")]
    public static string AddLinearDimension(
        [Description("First point X (mm).")]
        double x1,
        [Description("First point Y (mm).")]
        double y1,
        [Description("Second point X (mm).")]
        double x2,
        [Description("Second point Y (mm).")]
        double y2,
        [Description("Dimension line offset X from midpoint (mm) — moves the text label.")]
        double offset_x = 0.0,
        [Description("Dimension line offset Y from midpoint (mm) — moves the text label.")]
        double offset_y = 10.0,
        [Description("Projection style: 0=parallel to line, 1=horizontal, 2=vertical.")]
        int projection = 0)
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;

        ksLDimParam? dimPar = (ksLDimParam?)k.GetParamStruct((short)StructType2DEnum.ko_LDimParam);
        if (dimPar == null)
            return "ERROR: Could not get LDimParam struct.";

        ksLDimSourceParam? srcPar = (ksLDimSourceParam?)k.GetParamStruct((short)StructType2DEnum.ko_LDimSource);
        if (srcPar == null)
            return "ERROR: Could not get LDimSourceParam struct.";

        srcPar.Init();
        srcPar.x1 = x1;
        srcPar.y1 = y1;
        srcPar.x2 = x2;
        srcPar.y2 = y2;
        srcPar.dx = offset_x;
        srcPar.dy = offset_y;
        srcPar.ps = (short)projection;

        dimPar.SetSPar(srcPar);
        int ref_ = doc.ksLinDimension(dimPar);
        if (ref_ != 0)
        {
            UndoManager.Record2D($"Linear dimension ({x1},{y1})-({x2},{y2})", ref_);
            KompasConnector.RefreshView();
            return $"Linear dimension added, ref={ref_}.";
        }
        return "ERROR: Failed to add linear dimension.";
    }

    [McpServerTool(Name = "kompas_add_radial_dimension")]
    [Description("Add a radial dimension to a circle or arc.")]
    public static string AddRadialDimension(
        [Description("Center X of circle/arc (mm).")]
        double cx,
        [Description("Center Y of circle/arc (mm).")]
        double cy,
        [Description("Radius (mm).")]
        double radius,
        [Description("Angle at which the dimension arrow points (degrees from positive X axis).")]
        double angle = 45.0)
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;

        ksRDimParam? dimPar = (ksRDimParam?)k.GetParamStruct((short)StructType2DEnum.ko_RDimParam);
        if (dimPar == null)
            return "ERROR: Could not get RDimParam struct.";

        ksRDimSourceParam? srcPar = (ksRDimSourceParam?)k.GetParamStruct((short)StructType2DEnum.ko_RDimSource);
        if (srcPar == null)
            return "ERROR: Could not get RDimSourceParam struct.";

        srcPar.xc = cx;
        srcPar.yc = cy;
        srcPar.rad = radius;

        dimPar.SetSPar(srcPar);
        int ref_ = doc.ksRadDimension(dimPar);
        if (ref_ != 0)
        {
            UndoManager.Record2D($"Radial dimension r={radius}", ref_);
            KompasConnector.RefreshView();
            return $"Radial dimension r={radius} added, ref={ref_}.";
        }
        return "ERROR: Failed to add radial dimension.";
    }

    // ---------------------------------------------------------------
    //  Hatching
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_add_hatch")]
    [Description("Add a hatching region bounded by a closed rectangle (defined by x, y, width, height).")]
    public static string AddHatch(
        [Description("Left edge X of bounding rectangle (mm).")]
        double x,
        [Description("Bottom edge Y of bounding rectangle (mm).")]
        double y,
        [Description("Width of bounding rectangle (mm).")]
        double width,
        [Description("Height of bounding rectangle (mm).")]
        double height,
        [Description("Hatch step (line spacing) in mm.")]
        double step = 2.5,
        [Description("Hatch angle in degrees.")]
        double angle = 45.0,
        [Description("Line style of hatch boundary: 1=solid, 2=dashed.")]
        int style = 1)
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;

        // First draw the visible bounding contour (these remain as standalone objects).
        // ksHatch's internal boundary is invisible — without these lines the hatch has no outline.
        int b1 = doc.ksLineSeg(x,         y,          x + width, y,          style);
        int b2 = doc.ksLineSeg(x + width, y,          x + width, y + height, style);
        int b3 = doc.ksLineSeg(x + width, y + height, x,         y + height, style);
        int b4 = doc.ksLineSeg(x,         y + height, x,         y,          style);

        // Start hatch object (returns 1 on success)
        if (doc.ksHatch(style, angle, step, 0, 0, 0) != 1)
            return "ERROR: ksHatch returned failure.";

        // Re-draw boundary as the hatch contour (invisible — used only to bound the fill)
        doc.ksLineSeg(x,         y,          x + width, y,          style);
        doc.ksLineSeg(x + width, y,          x + width, y + height, style);
        doc.ksLineSeg(x + width, y + height, x,         y + height, style);
        doc.ksLineSeg(x,         y + height, x,         y,          style);

        int ref_ = doc.ksEndObj();
        if (ref_ != 0)
        {
            // Record all 5 refs under one logical operation so undo removes hatch + outline together
            UndoManager.Record2DMulti($"Hatch ({x},{y}) {width}x{height}", b1, b2, b3, b4, ref_);
            KompasConnector.RefreshView();
            return $"Hatch added at ({x},{y}), {width}x{height}, step={step}, angle={angle}°, ref={ref_}.";
        }
        return "ERROR: Failed to create hatch object.";
    }

    // ---------------------------------------------------------------
    //  Groups and transforms
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_begin_group")]
    [Description("Begin recording drawing objects into a group. Objects drawn after this call belong to the group until kompas_end_group.")]
    public static string BeginGroup()
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        int gr = doc.ksNewGroup((short)0);
        if (gr != 0)
            return $"Group started, ref={gr}.";
        return "ERROR: Failed to start group.";
    }

    [McpServerTool(Name = "kompas_end_group")]
    [Description("Finish recording objects into the current group.")]
    public static string EndGroup()
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        int gr = doc.ksEndGroup();
        if (gr != 0)
        {
            UndoManager.Record2D("Group", gr);
            return $"Group closed, ref={gr}.";
        }
        return "ERROR: No active group.";
    }

    [McpServerTool(Name = "kompas_rotate_group")]
    [Description("Rotate a group reference (from kompas_end_group) around a point by a given angle.")]
    public static string RotateGroup(
        [Description("Group reference returned by kompas_end_group.")]
        int group_ref,
        [Description("Center of rotation X (mm).")]
        double cx,
        [Description("Center of rotation Y (mm).")]
        double cy,
        [Description("Rotation angle in degrees (positive = counter-clockwise).")]
        double angle)
    {
        if (group_ref == 0) return "ERROR: group_ref must be non-zero.";
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        // ksRotateObj returns 1 on success, 0 if the ref is not a valid group/object
        if (doc.ksRotateObj(group_ref, cx, cy, angle) == 0)
            return $"ERROR: Could not rotate group {group_ref}. The reference may be stale or not a group.";
        KompasConnector.RefreshView();
        return $"Group {group_ref} rotated {angle}° around ({cx},{cy}).";
    }

    [McpServerTool(Name = "kompas_move_object")]
    [Description("Translate a drawing object or group by (dx, dy).")]
    public static string MoveObject(
        [Description("Object or group reference.")]
        int object_ref,
        [Description("Translation in X (mm).")]
        double dx,
        [Description("Translation in Y (mm).")]
        double dy)
    {
        if (object_ref == 0) return "ERROR: object_ref must be non-zero.";
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        if (doc.ksMoveObj(object_ref, dx, dy) == 0)
            return $"ERROR: Could not move object {object_ref}. The reference may be stale.";
        KompasConnector.RefreshView();
        return $"Object {object_ref} moved by ({dx},{dy}).";
    }

    // ---------------------------------------------------------------
    //  Splines (Bezier & NURBS)
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_draw_bezier")]
    [Description("Draw a Bezier curve through control points in the active 2D document or current sketch.")]
    public static string DrawBezier(
        [Description("Array of X coordinates of control points (mm), e.g. [0, 20, 50, 70, 100].")]
        double[] xs,
        [Description("Array of Y coordinates of control points (mm), same length as xs.")]
        double[] ys,
        [Description("Whether the curve is closed (connects last point to first).")]
        bool closed = false,
        [Description("Line style: 1=solid, 2=dashed, 3=chain.")]
        int style = 1)
    {
        if (xs.Length != ys.Length)
            return "ERROR: xs and ys arrays must have the same length.";
        if (xs.Length < 2)
            return "ERROR: Need at least 2 control points.";

        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        doc.ksBezier((short)(closed ? 1 : 0), style);
        for (int i = 0; i < xs.Length; i++)
            doc.ksPoint(xs[i], ys[i], 0);
        int ref_ = doc.ksEndObj();

        if (ref_ != 0)
        {
            UndoManager.Record2D($"Bezier {xs.Length} pts", ref_);
            KompasConnector.RefreshView();
            return $"Bezier curve drawn with {xs.Length} points, closed={closed}, ref={ref_}.";
        }
        return "ERROR: Failed to draw Bezier curve.";
    }

    [McpServerTool(Name = "kompas_draw_nurbs")]
    [Description("Draw a NURBS curve through weighted control points in the active 2D document or current sketch.")]
    public static string DrawNurbs(
        [Description("Array of X coordinates of control points (mm).")]
        double[] xs,
        [Description("Array of Y coordinates of control points (mm), same length as xs.")]
        double[] ys,
        [Description("Array of weights for each control point (same length as xs). Use 1.0 for uniform weighting.")]
        double[]? weights = null,
        [Description("Degree of the NURBS curve (e.g. 3 for cubic).")]
        int degree = 3,
        [Description("Whether the curve is closed (periodic).")]
        bool closed = false,
        [Description("Line style: 1=solid, 2=dashed, 3=chain.")]
        int style = 1)
    {
        if (xs.Length != ys.Length)
            return "ERROR: xs and ys arrays must have the same length.";
        if (xs.Length < 2)
            return "ERROR: Need at least 2 control points.";
        if (weights != null && weights.Length != xs.Length)
            return "ERROR: weights array must have the same length as xs/ys.";

        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;

        doc.ksNurbs((short)degree, closed, style);

        ksNurbsPointParam? par = (ksNurbsPointParam?)k.GetParamStruct(
            (short)StructType2DEnum.ko_NurbsPointParam);
        if (par == null)
            return "ERROR: Could not get NurbsPointParam struct.";

        for (int i = 0; i < xs.Length; i++)
        {
            par.Init();
            par.x = xs[i];
            par.y = ys[i];
            par.weight = weights != null ? weights[i] : 1.0;
            doc.ksNurbsPoint(par);
        }

        int ref_ = doc.ksEndObj();

        if (ref_ != 0)
        {
            UndoManager.Record2D($"NURBS {xs.Length} pts deg={degree}", ref_);
            KompasConnector.RefreshView();
            return $"NURBS curve drawn with {xs.Length} points, degree={degree}, closed={closed}, ref={ref_}.";
        }
        return "ERROR: Failed to draw NURBS curve.";
    }

    // ---------------------------------------------------------------
    //  Utilities
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_delete_object")]
    [Description("Delete a drawing object by its reference.")]
    public static string DeleteObject(
        [Description("Object reference to delete.")]
        int object_ref)
    {
        var doc = RequireDoc2D(out var err);
        if (doc == null) return err;
        int res = doc.ksDeleteObj(object_ref);
        if (res != 0)
        {
            UndoManager.Forget(object_ref);
            return $"Object {object_ref} deleted.";
        }
        return $"ERROR: Could not delete object {object_ref}.";
    }
}
