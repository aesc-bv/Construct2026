using System;
using System.Linq;
using System.Drawing;
using System.Globalization;
using SpaceClaim.Api.V251;
using SpaceClaim.Api.V251.Geometry;
using SpaceClaim.Api.V251.Modeler;
using Point = SpaceClaim.Api.V251.Geometry.Point;
using Vector = SpaceClaim.Api.V251.Geometry.Vector;

namespace AESCConstruct25.FrameGenerator.Utilities
{
    public static class JointCurveHelper
    {
        /// <summary>
        /// Draws an extended center‐line in world‐space by transforming the local segment
        /// through its parent component’s Placement matrix.
        /// </summary>
        public static void DrawExtendedCenterLine(
            Component component,
            CurveSegment localSegment,
            string name = "CenterLine",
            double scale = 10
        )
        {
            // 1) local→world transform
            Matrix toWorld = component.Placement;

            // 2) endpoints in world
            Point w0 = toWorld * localSegment.StartPoint;
            Point w1 = toWorld * localSegment.EndPoint;

            // 3) extension in world units
            Vector dir = w1 - w0;
            Direction unit = dir.Direction;
            Vector uv = unit.ToVector();
            Point p0 = w0 - uv * (0.005 * scale);
            Point p1 = w1 + uv * (0.005 * scale);

            // 4) draw in component’s part
            CreateDebugCurve(component.Template, p0, p1, Color.Blue, name);
        }

        static double GetCustomDimension(Component comp, string key, double defaultWidth)
        {
            if (comp.Template.CustomProperties.TryGetValue(key, out var cp)
             && double.TryParse(cp.Value.ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var mm)
             && mm > 0)
                return mm * 0.001;
            return defaultWidth;
        }

        /// <summary>
        /// From a local‐space segment and its neighbour, compute the two offset lines in world‐space.
        /// </summary>
        public static (Line inner, Line outer, Vector perp) GetOffsetEdges(
            Component component,
            CurveSegment localSeg,
            Component otherComponent,
            double width,    // from Construct_w (in metres)
            double offsetX   // from custom property "offsetX"
        )
        {
            const double tol = 1e-6;

            var part = component.Template;
            var oldDebugCurves = part.Curves
                .OfType<DesignCurve>()
                .Where(dc => dc.Name == "DEBUG_InnerOffset" || dc.Name == "DEBUG_OuterOffset")
                .ToList();
            foreach (var dc in oldDebugCurves)
                dc.Delete();

            // 1) Lift this segment to world:
            var placementA = component.Placement;
            Point wA0 = placementA * localSeg.StartPoint;
            Point wA1 = placementA * localSeg.EndPoint;
            Vector dA = (wA1 - wA0).Direction.ToVector();

            // 2) And lift the neighbour’s segment:
            var otherDC = otherComponent.Template
                               .Curves
                               .OfType<DesignCurve>()
                               .FirstOrDefault()?.Shape as CurveSegment
                           ?? throw new ArgumentException("otherComponent must have a curve");
            var placementB = otherComponent.Placement;
            Point wB0 = placementB * otherDC.StartPoint;
            Point wB1 = placementB * otherDC.EndPoint;
            Vector dB = (wB1 - wB0).Direction.ToVector();

            // 3) Compute the shared‐plane normal:
            Vector planeN = Vector.Cross(dA, dB);
            if (planeN.Magnitude < tol)
            {
                // fallback if nearly colinear
                planeN = Vector.Cross(dA, Vector.Create(0, 1, 0));
                if (planeN.Magnitude < tol)
                    planeN = Vector.Cross(dA, Vector.Create(1, 0, 0));
            }
            planeN = planeN.Direction.ToVector();

            // 4) Find the shared endpoint ("corner") in world:
            var segA_w = CurveSegment.Create(wA0, wA1);
            var segB_w = CurveSegment.Create(wB0, wB1);
            Point corner = FindSharedPoint(segA_w, segB_w);

            // 5) Identify the "other" ends of each segment (the ray from corner):
            Point wAother = ((corner - wA0).Magnitude < tol) ? wA1 : wA0;
            Point wBother = ((corner - wB0).Magnitude < tol) ? wB1 : wB0;

            // 6) Build the interior bisector in world:
            Vector vA = (wAother - corner).Direction.ToVector();
            Vector vB = (wBother - corner).Direction.ToVector();
            Vector bis = vA + vB;
            if (bis.Magnitude < tol)
            {
                // fallback if nearly a straight line
                bis = Vector.Cross(planeN, dA).Direction.ToVector();
            }
            else
            {
                bis = bis.Direction.ToVector();
            }

            // 7) Compute the two possible perp-directions to A in the plane,
            //    then pick the one whose dot with the bisector is positive (interior):
            Vector p1 = Vector.Cross(planeN, dA).Direction.ToVector();
            Vector perpWorld = Vector.Dot(p1, bis) >= 0 ? p1 : -p1;

            // 8) Convert that perp back into component‐local:
            Point Wc = corner;
            Point Wc2 = Point.Create(corner.X + perpWorld.X,
                                     corner.Y + perpWorld.Y,
                                     corner.Z + perpWorld.Z);
            var invA = placementA.Inverse;
            Point Lc = invA * Wc;
            Point Lc2 = invA * Wc2;
            Vector perpLocal = (Lc2 - Lc).Direction.ToVector();

            // 9) Choose the “base” local endpoint (the one farther from corner):
            Point sL = localSeg.StartPoint;
            Point eL = localSeg.EndPoint;
            double ds = ((placementA * sL) - corner).Magnitude;
            double de = ((placementA * eL) - corner).Magnitude;
            Point baseLocal = ds > de + tol ? sL : eL;

            // 10) Compute total shift = half the profile‐width + offsetX
            //double halfWidth = width * 0.5;
            //double shiftInner = halfWidth + offsetX;
            //double shiftOuter = halfWidth - offsetX;

            // 11) build the two local‐space offset lines
            //var dirLocal = (eL - sL).Direction;
            // interior‐side line (red in your debug)
            Vector localX = Vector.Create(1, 0, 0);
            double flip = Vector.Dot(perpLocal, localX) >= 0
                          ? +1.0
                          : -1.0;
            double effectiveOffset = offsetX * flip;

            double angleDeg = 0;
            if (component.Template.CustomProperties
                         .TryGetValue("RotationAngle", out var rotProp)
             && double.TryParse(rotProp.Value.ToString(),
                                NumberStyles.Float,
                                CultureInfo.InvariantCulture,
                                out var a))
            {
                // ensure positive
                angleDeg = (a % 360 + 360) % 360;
            }

            // 2) Default to the “w” dimension
            double chosenWidth = width;
            string chosenDimension = "w";

            // 3) If rotated near 90° or near 270°, swap to h/b/D
            const double tolcheck = 1e-3;
            if (Math.Abs(angleDeg - 90) < tolcheck || Math.Abs(angleDeg - 270) < tolcheck)
            {
                // try h → b → D
                double? h = GetCustomDimension(component, "Construct_h", width);
                double? b = GetCustomDimension(component, "Construct_b", width);
                double? D = GetCustomDimension(component, "Construct_D", width);

                if (h.HasValue)
                {
                    chosenWidth = h.Value;
                    chosenDimension = "h";
                }
                else if (b.HasValue)
                {
                    chosenWidth = b.Value;
                    chosenDimension = "b";
                }
                else if (D.HasValue)
                {
                    chosenWidth = D.Value;
                    chosenDimension = "D";
                }
            }

            // 3) Log it:
            Logger.Log(
                $"GetOffsetEdges: using dimension '{chosenDimension}' = {chosenWidth:F4} m " +
                $"because RotationAngle = {angleDeg:F1}°"
            );

            // 11’) build two shifts that still sum to exactly width:
            double halfWidth = chosenWidth * 0.5;
            double shiftInner = halfWidth + effectiveOffset;
            double shiftOuter = halfWidth - effectiveOffset;

            // 12’) rebuild the two lines as before:
            var dirLocal = (eL - sL).Direction;
            Line innerLocal = Line.Create(baseLocal + perpLocal * shiftInner, dirLocal);
            Line outerLocal = Line.Create(baseLocal - perpLocal * shiftOuter, dirLocal);

            //12) debug‐draw(unchanged)
            double dbgLen = 0.1;
            {
                var v = innerLocal.Direction.ToVector() * (dbgLen / 2);
                CreateDebugCurve(part,
                    innerLocal.Origin - v,
                    innerLocal.Origin + v,
                    Color.Red,
                    "DEBUG_InnerOffset");
            }
            {
                var v = outerLocal.Direction.ToVector() * (dbgLen / 2);
                CreateDebugCurve(part,
                    outerLocal.Origin - v,
                    outerLocal.Origin + v,
                    Color.Green,
                    "DEBUG_OuterOffset");
            }

            return (innerLocal, outerLocal, perpLocal);
        }

        //public static Point? IntersectLines(
        //    Component componentA,
        //    Line localLineA,
        //    Component componentB,
        //    Line localLineB
        //)
        //{
        //    // -- STEP 1: lift localLineA into worldA --
        //    // take two local‐space points on lineA: its Origin, and Origin + Direction
        //    Point localA_O = localLineA.Origin;
        //    Vector localA_dirVec = localLineA.Direction.ToVector();
        //    Point localA_P = localA_O + localA_dirVec;

        //    // transform both to world via componentA.Placement
        //    Point worldA_O = componentA.Placement * localA_O;
        //    Point worldA_P = componentA.Placement * localA_P;

        //    // recompute worldA direction
        //    Vector worldA_dir = (worldA_P - worldA_O).Direction.ToVector();
        //    Line worldLineA = Line.Create(worldA_O, worldA_dir.Direction);

        //    // -- STEP 2: lift localLineB into worldB --
        //    Point localB_O = localLineB.Origin;
        //    Vector localB_dirVec = localLineB.Direction.ToVector();
        //    Point localB_P = localB_O + localB_dirVec;

        //    Point worldB_O = componentB.Placement * localB_O;
        //    Point worldB_P = componentB.Placement * localB_P;

        //    Vector worldB_dir = (worldB_P - worldB_O).Direction.ToVector();
        //    Line worldLineB = Line.Create(worldB_O, worldB_dir.Direction);

        //    // -- STEP 3: compute intersection of worldLineA and worldLineB in world coords --
        //    Point p = worldLineA.Origin;
        //    Vector d1 = worldLineA.Direction.ToVector();
        //    Point q = worldLineB.Origin;
        //    Vector d2 = worldLineB.Direction.ToVector();
        //    Vector r = q - p;

        //    double a = Vector.Dot(d1, d1);
        //    double b = Vector.Dot(d1, d2);
        //    double c = Vector.Dot(d2, d2);
        //    double d = Vector.Dot(d1, r);
        //    double e = Vector.Dot(d2, r);

        //    double denom = a * c - b * b;
        //    //Logger.Log($"IntersectLines: denom={denom:F6}, a={a:F6}, b={b:F6}, c={c:F6}");
        //    if (Math.Abs(denom) < 1e-10)
        //        return null;

        //    double s = (d * c - b * e) / denom;
        //    return p + d1 * s;
        //}

        //public static (Line inner, Line outer, Vector perp) GetOffsetEdges(
        //    Component component,
        //    CurveSegment localSeg,
        //    Component otherComponent,
        //    double width,    // full profile width
        //    double offsetX   // unused
        //)
        //{
        //    //var oldOri = Application.UserOptions.WorldOrientation;
        //    //Application.UserOptions.WorldOrientation = WorldOrientation.UpIsY;
        //    //try
        //    //{
        //        const double tol = 1e-6;
        //        const double tol90 = 1e-3;

        //        // 0) log inputs
        //        Logger.Log($"GetOffsetEdges START: comp='{component.Name}', other='{otherComponent.Name}', width={width:F4}, offsetX={offsetX:F4}");

        //        var part = component.Template;

        //        // clear old debug curves
        //        foreach (var dc in part.Curves
        //            .OfType<DesignCurve>()
        //            .Where(dc => dc.Name == "DEBUG_InnerOffset" || dc.Name == "DEBUG_OuterOffset")
        //            .ToList())
        //            dc.Delete();

        //        // 1) read optional offsetY (for future use)
        //        double offsetY = 0;
        //        if (part.CustomProperties.TryGetValue("offsetY", out var oy) &&
        //            double.TryParse(oy.Value.ToString().Replace(',', '.'),
        //                            NumberStyles.Any,
        //                            CultureInfo.InvariantCulture,
        //                            out offsetY))
        //        {
        //            Logger.Log($"  offsetY = {offsetY:F4}");
        //        }

        //        // 2) read RotationAngle
        //        double angleDeg = 0;
        //        if (part.CustomProperties.TryGetValue("RotationAngle", out var rp)
        //         && double.TryParse(rp.Value.ToString().Replace(',', '.'),
        //                            NumberStyles.Any,
        //                            CultureInfo.InvariantCulture,
        //                            out var a))
        //        {
        //            angleDeg = (a % 360 + 360) % 360;
        //        }
        //        Logger.Log($"  RotationAngle = {angleDeg:F1}°");

        //        // 3) lift this segment to world and compute dA
        //        var placementA = component.Placement;
        //        Point wA0 = placementA * localSeg.StartPoint;
        //        Point wA1 = placementA * localSeg.EndPoint;
        //        Vector dA = (wA1 - wA0).Direction.ToVector();
        //        Logger.Log(
        //            $"  wA0 = ({wA0.X:F4},{wA0.Y:F4},{wA0.Z:F4}), " +
        //            $"wA1 = ({wA1.X:F4},{wA1.Y:F4},{wA1.Z:F4}), " +
        //            $"dA = ({dA.X:F4},{dA.Y:F4},{dA.Z:F4})"
        //        );

        //        // 4) lift the other component’s segment and compute dB
        //        var otherDC = otherComponent.Template
        //                           .Curves
        //                           .OfType<DesignCurve>()
        //                           .FirstOrDefault()?.Shape as CurveSegment
        //                       ?? throw new ArgumentException("otherComponent must have a curve");
        //        Point wB0 = otherComponent.Placement * otherDC.StartPoint;
        //        Point wB1 = otherComponent.Placement * otherDC.EndPoint;
        //        Vector dB = (wB1 - wB0).Direction.ToVector();
        //        Logger.Log(
        //            $"  wB0 = ({wB0.X:F4},{wB0.Y:F4},{wB0.Z:F4}), " +
        //            $"wB1 = ({wB1.X:F4},{wB1.Y:F4},{wB1.Z:F4}), " +
        //            $"dB = ({dB.X:F4},{dB.Y:F4},{dB.Z:F4})"
        //        );

        //        // 5) compute shared‐plane normal
        //        Vector planeN = Vector.Cross(dA, dB);
        //        if (planeN.Magnitude < tol)
        //        {
        //            planeN = Vector.Cross(dA, Vector.Create(0, 1, 0));
        //            if (planeN.Magnitude < tol)
        //                planeN = Vector.Cross(dA, Vector.Create(1, 0, 0));
        //        }
        //        planeN = planeN.Direction.ToVector();
        //        if (planeN.Y < 0) planeN = -planeN;
        //        Logger.Log($"  planeN = ({planeN.X:F4},{planeN.Y:F4},{planeN.Z:F4})");

        //        // 6) find shared “corner” world‐point
        //        var segA_w = CurveSegment.Create(wA0, wA1);
        //        var segB_w = CurveSegment.Create(wB0, wB1);
        //        Point corner = FindSharedPoint(segA_w, segB_w);
        //        Logger.Log($"  corner = ({corner.X:F4},{corner.Y:F4},{corner.Z:F4})");

        //        // 7) interior‐bisector
        //        Point wAother = ((corner - wA0).Magnitude < tol) ? wA1 : wA0;
        //        Point wBother = ((corner - wB0).Magnitude < tol) ? wB1 : wB0;
        //        Vector vA = (wAother - corner).Direction.ToVector();
        //        Vector vB = (wBother - corner).Direction.ToVector();
        //        Vector bis = (vA + vB).Magnitude < tol
        //                   ? Vector.Cross(planeN, dA).Direction.ToVector()
        //                   : (vA + vB).Direction.ToVector();
        //        Logger.Log($"  bisector = ({bis.X:F4},{bis.Y:F4},{bis.Z:F4})");

        //        // 8) pick true perp in plane
        //        Vector p1 = Vector.Cross(planeN, dA).Direction.ToVector();
        //        Vector perpWorld = Vector.Dot(p1, bis) >= 0 ? p1 : -p1;
        //        Logger.Log($"  raw perp candidate p1 = ({p1.X:F4},{p1.Y:F4},{p1.Z:F4})");
        //        Logger.Log($"  chosen perpWorld = ({perpWorld.X:F4},{perpWorld.Y:F4},{perpWorld.Z:F4})");

        //        // 9) map perp into local
        //        var invA = placementA.Inverse;
        //        Point Lc = invA * corner;
        //        Point Lc2 = invA * Point.Create(
        //            corner.X + perpWorld.X,
        //            corner.Y + perpWorld.Y,
        //            corner.Z + perpWorld.Z
        //        );
        //        Vector perpLocal = (Lc2 - Lc).Direction.ToVector();
        //        Logger.Log($"  perpLocal = ({perpLocal.X:F4},{perpLocal.Y:F4},{perpLocal.Z:F4})");

        //        // 10) pick half‐dim & local axis by rotation
        //        double halfDim;
        //        Vector axisLocal;
        //        if (Math.Abs(angleDeg % 180) < tol90)
        //        {
        //            halfDim = width * 0.5;
        //            axisLocal = Vector.Create(1, 0, 0);
        //        }
        //        else
        //        {
        //            double? h = GetCustomDimension(component, "Construct_h", width);
        //            double? b = GetCustomDimension(component, "Construct_b", width);
        //            double? D = GetCustomDimension(component, "Construct_D", width);
        //            double chosenHeight = h ?? b ?? D ?? width;
        //            halfDim = chosenHeight * 0.5;
        //            axisLocal = Vector.Create(0, 1, 0);
        //        }
        //        Logger.Log($"  axisLocal = ({axisLocal.X:F4},{axisLocal.Y:F4},{axisLocal.Z:F4}), halfDim = {halfDim:F4}");

        //        // 11) direction test and shift
        //        double flip = Vector.Dot(perpLocal, axisLocal) >= 0 ? +1.0 : -1.0;
        //        Vector shiftLocal = axisLocal * (halfDim * flip);
        //        Logger.Log($"  flip = {flip:F1}, shiftLocal = ({shiftLocal.X:F4},{shiftLocal.Y:F4},{shiftLocal.Z:F4})");

        //        // 12) build the two local‐space offset lines
        //        var dirLocal = (localSeg.EndPoint - localSeg.StartPoint).Direction;
        //        var innerOrigin = Point.Create(shiftLocal.X, shiftLocal.Y, shiftLocal.Z);
        //        var outerOrigin = Point.Create(-shiftLocal.X, -shiftLocal.Y, -shiftLocal.Z);
        //        var innerLocal = Line.Create(innerOrigin, dirLocal);
        //        var outerLocal = Line.Create(outerOrigin, dirLocal);
        //        Logger.Log(
        //            $"  innerOrigin = ({innerOrigin.X:F4},{innerOrigin.Y:F4},{innerOrigin.Z:F4}), " +
        //            $"innerDir = ({innerLocal.Direction.X:F4},{innerLocal.Direction.Y:F4},{innerLocal.Direction.Z:F4})"
        //        );
        //        Logger.Log(
        //            $"  outerOrigin = ({outerOrigin.X:F4},{outerOrigin.Y:F4},{outerOrigin.Z:F4}), " +
        //            $"outerDir = ({outerLocal.Direction.X:F4},{outerLocal.Direction.Y:F4},{outerLocal.Direction.Z:F4})"
        //        );

        //        double dbgLen = 0.1;
        //        {
        //            var v = innerLocal.Direction.ToVector() * (dbgLen / 2);
        //            CreateDebugCurve(part,
        //                innerLocal.Origin - v,
        //                innerLocal.Origin + v,
        //                Color.Red,
        //                "DEBUG_InnerOffset");
        //        }
        //        {
        //            var v = outerLocal.Direction.ToVector() * (dbgLen / 2);
        //            CreateDebugCurve(part,
        //                outerLocal.Origin - v,
        //                outerLocal.Origin + v,
        //                Color.Green,
        //                "DEBUG_OuterOffset");
        //        }

        //        Logger.Log("GetOffsetEdges END");
        //        return (innerLocal, outerLocal, axisLocal);
        //    //}
        //    //finally
        //    //{
        //    //    Application.UserOptions.WorldOrientation = oldOri;
        //    //}
        //}


        public static Point? IntersectLines(
            Component componentA,
            Line localLineA,
            Component componentB,
            Line localLineB
        )
        {
            // Lift A into world
            Point A0 = componentA.Placement * localLineA.Origin;
            Point A1 = componentA.Placement * (localLineA.Origin + localLineA.Direction.ToVector());
            Vector v = (A1 - A0).Direction.ToVector();
            Line worldA = Line.Create(A0, v.Direction);

            // Lift B into world
            Point B0 = componentB.Placement * localLineB.Origin;
            Point B1 = componentB.Placement * (localLineB.Origin + localLineB.Direction.ToVector());
            Vector w = (B1 - B0).Direction.ToVector();
            Line worldB = Line.Create(B0, w.Direction);

            // Compute intersection in world
            Point p = worldA.Origin;
            Vector d1 = worldA.Direction.ToVector();
            Point q = worldB.Origin;
            Vector d2 = worldB.Direction.ToVector();
            Vector r = q - p;

            double a = Vector.Dot(d1, d1);
            double b = Vector.Dot(d1, d2);
            double c = Vector.Dot(d2, d2);
            double d = Vector.Dot(d1, r);
            double e = Vector.Dot(d2, r);
            double denom = a * c - b * b;
            //Logger.Log($"IntersectLines: denom={denom:F6}");
            if (Math.Abs(denom) < 1e-10)
                return null;

            double s = (d * c - b * e) / denom;
            return p + d1 * s;
        }


        public static Point? IntersectTLines(
            Component componentA,
            CurveSegment segA,
            Component componentB,
            CurveSegment segB
        )
        {
            const double tol = 1e-8;

            // 1) Lift A’s segment into world
            Point wA0 = componentA.Placement * segA.StartPoint;
            Point wA1 = componentA.Placement * segA.EndPoint;

            // 2) Lift B’s segment into world
            Point wB0 = componentB.Placement * segB.StartPoint;
            Point wB1 = componentB.Placement * segB.EndPoint;

            // 3) Direction unit-vectors
            Vector d1 = (wA1 - wA0).Direction.ToVector();
            Vector d2 = (wB1 - wB0).Direction.ToVector();

            // 4) Build the 2×2 system for closest points on two infinite lines
            Point p = wA0;
            Point q = wB0;
            Vector r = q - p;
            double a = Vector.Dot(d1, d1);
            double b = Vector.Dot(d1, d2);
            double c = Vector.Dot(d2, d2);
            double d = Vector.Dot(d1, r);
            double e = Vector.Dot(d2, r);
            double denom = a * c - b * b;
            if (Math.Abs(denom) < tol)
                return null; // nearly parallel

            // 5) Solve for parameters s (on A) and t (on B)
            double s = (d * c - b * e) / denom;
            double t = (a * e - b * d) / denom;

            // 6) Convert to normalized segment parameters u,v in [0,1]
            double lenA = (wA1 - wA0).Magnitude;
            double lenB = (wB1 - wB0).Magnitude;
            if (lenA < tol || lenB < tol)
                return null;
            double u = s / lenA;
            double v = t / lenB;

            // 7) Reject if the closest point lies outside the finite segments
            if (u < -tol || u > 1 + tol || v < -tol || v > 1 + tol)
                return null;

            // 8) Return the intersection point along A
            return p + d1 * s;
        }


        public static void CreateDebugCurve(
            Part part,
            Point start,
            Point end,
            Color color,
            string name
        )
        {
            var seg = CurveSegment.Create(start, end);
            var dc = DesignCurve.Create(part, seg);
            dc.SetColor(null, color);
            dc.SetVisibility(null, true);
            dc.Name = name;
        }

        public static Point FindSharedPoint(CurveSegment a, CurveSegment b)
        {
            foreach (var pA in new[] { a.StartPoint, a.EndPoint })
                foreach (var pB in new[] { b.StartPoint, b.EndPoint })
                    if ((pA - pB).Magnitude < 1e-6)
                        return pA;
            return a.StartPoint;
        }

        public static double GetProfileWidth(Component comp)
        {
            if (comp.Template.CustomProperties
                    .TryGetValue("Construct_w", out var wp)
                && double.TryParse(wp.Value.ToString(),
                                   NumberStyles.Any,
                                   CultureInfo.InvariantCulture,
                                   out var w_mm))
            {
                // convert from millimetres to metres for joint logic
                return w_mm * 0.001;
            }

            var fullDb = comp.Template.Bodies
                           .FirstOrDefault(b => b.Name == "ExtrudedProfile")
                       ?? comp.Template.Bodies.FirstOrDefault(b => b.Shape != null);

            if (fullDb != null)
            {
                var bb = fullDb.Shape.GetBoundingBox(Matrix.Identity, tight: true);
                var seg = comp.Template.Curves
                                .OfType<DesignCurve>()
                                .FirstOrDefault()?.Shape as CurveSegment;
                if (seg != null)
                {
                    var d = (seg.EndPoint - seg.StartPoint).Direction.ToVector();
                    var flat = Vector.Create(d.X, 0, d.Z).Direction.ToVector();
                    var perp = Vector.Cross(Vector.Create(0, 1, 0), flat).Direction.ToVector();

                    var corners = Box.Create(new[] { bb.MinCorner, bb.MaxCorner }).Corners;
                    var center = bb.Center;
                    double maxDist = corners
                        .Select(c => Math.Abs(Vector.Dot(perp, c - center)))
                        .Max();
                    return maxDist * 2.0;
                }
            }

            return 0.0;
        }

        public static double GetOffset(Component comp, string key)
        {
            if (!comp.Template.CustomProperties.TryGetValue(key, out var prop))
                return 0.0;
            var s = prop.Value.ToString().Replace(',', '.');
            return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)
                 ? v
                 : 0.0;
        }

        public static (double forward, double backward) PickDirection(
            Plane plane,
            Component component,
            bool connectedAtStart,
            bool connectedAtEnd,
            double longLen,
            double shortLen
        )
        {
            var extr = component.Template
                         .Bodies
                         .FirstOrDefault(b => b.Name == "ExtrudedProfile")
                         ?.Shape;
            if (extr == null)
                return (longLen, shortLen);

            var bb = extr.GetBoundingBox(Matrix.Identity, tight: true);
            var toComp = bb.Center - plane.Frame.Origin;
            bool onPos = Vector.Dot(toComp, plane.Frame.DirZ.ToVector()) > 0;

            return onPos
                ? (shortLen, longLen)
                : (longLen, shortLen);
        }

        // inside JointCurveHelper.cs

        /// <summary>
        /// For a T-joint: picks forward/backward extrusion lengths by comparing
        /// the world-space vector from the T-point to the attached end of B
        /// against the cutterPlane's Z axis (which is A’s world direction).
        /// </summary>
        public static (double forward, double backward) PickTJointDirectionWorld(
            Plane cutterPlane,
            Component componentB,
            Point pTworld,
            bool connectedAtStart,
            bool connectedAtEnd,
            double attachLen,
            double freeLen
        )
        {
            // 1) Get raw B curve
            var rawB = componentB.Template
                           .Curves
                           .OfType<DesignCurve>()
                           .FirstOrDefault()?.Shape as CurveSegment;
            if (rawB == null)
                return (freeLen, attachLen);

            // 2) World‐space point of the free end (not attached to A)
            Point worldFree = connectedAtStart
                ? componentB.Placement * rawB.EndPoint
                : componentB.Placement * rawB.StartPoint;

            // 3) Vector from T‐point toward the FREE end
            Vector toFree = (worldFree - pTworld).Direction.ToVector();

            // 4) Cutter plane forward‐axis (+Z) in world
            Vector planeZ = cutterPlane.Frame.DirZ.ToVector();

            // 5) Dot gives signed distance: 
            //    >0 ⇒ free-end lies in +Z half-space
            double dot = Vector.Dot(toFree, planeZ);
            //Logger.Log($"PickTJointDirectionWorld: dot(toFree,planeZ) = {dot:F4}");

            // 6) If dot > 0, then forward (+Z) points toward free end ⇒ forward=freeLen
            //    else forward=attachLen
            if (dot > 0)
                return (forward: freeLen, backward: attachLen);
            else
                return (forward: attachLen, backward: freeLen);
        }

    }


    public static class DirectionExtensions
    {
        public static Vector ToVector(this Direction d)
            => Vector.Create(d.X, d.Y, d.Z);
    }
}
