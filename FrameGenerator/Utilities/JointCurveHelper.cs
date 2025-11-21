using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;

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
                }
                else if (b.HasValue)
                {
                    chosenWidth = b.Value;
                }
                else if (D.HasValue)
                {
                    chosenWidth = D.Value;
                }
            }

            // 3) Log it:
            // Logger.Log(
            //    $"GetOffsetEdges: using dimension '{chosenDimension}' = {chosenWidth:F4} m " +
            //    $"because RotationAngle = {angleDeg:F1}°"
            //);

            // 11’) build two shifts that still sum to exactly width:
            double halfWidth = chosenWidth * 0.5;
            double shiftInner = halfWidth + effectiveOffset;
            double shiftOuter = halfWidth - effectiveOffset;

            // 12’) rebuild the two lines as before:
            var dirLocal = (eL - sL).Direction;
            Line innerLocal = Line.Create(baseLocal + perpLocal * shiftInner, dirLocal);
            Line outerLocal = Line.Create(baseLocal - perpLocal * shiftOuter, dirLocal);

            //12) debug‐draw(unchanged)
            //double dbgLen = 0.1;
            //{
            //    var v = innerLocal.Direction.ToVector() * (dbgLen / 2);
            //    CreateDebugCurve(part,
            //        innerLocal.Origin - v,
            //        innerLocal.Origin + v,
            //        Color.Red,
            //        "DEBUG_InnerOffset");
            //}
            //{
            //    var v = outerLocal.Direction.ToVector() * (dbgLen / 2);
            //    CreateDebugCurve(part,
            //        outerLocal.Origin - v,
            //        outerLocal.Origin + v,
            //        Color.Green,
            //        "DEBUG_OuterOffset");
            //}

            return (innerLocal, outerLocal, perpLocal);
        }

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
            // Helper to read a numeric custom property (mm → m)
            static bool TryGetMmAsMeters(Component c, string key, out double meters)
            {
                meters = 0.0;
                if (c.Template.CustomProperties.TryGetValue(key, out var prop)
                    && double.TryParse(prop.Value.ToString().Replace(',', '.'),
                                       NumberStyles.Any, CultureInfo.InvariantCulture, out var mm)
                    && mm > 0)
                {
                    meters = mm * 0.001;
                    return true;
                }
                return false;
            }

            // Read rotation (default 0)
            double angleDeg = 0.0;
            if (comp.Template.CustomProperties.TryGetValue("RotationAngle", out var rotProp))
                double.TryParse(rotProp.Value.ToString().Replace(',', '.'),
                                NumberStyles.Float, CultureInfo.InvariantCulture, out angleDeg);

            // Normalize to [0,360)
            angleDeg = (angleDeg % 360 + 360) % 360;

            // Treat “exactly 90 or 270” with a tiny epsilon to handle float text like 89.9999999
            const double eps = 1e-6;
            bool useH = Math.Abs(angleDeg - 90.0) < eps || Math.Abs(angleDeg - 270.0) < eps;

            // Prefer Construct_h when rotated 90/270, else Construct_w
            if (useH)
            {
                if (TryGetMmAsMeters(comp, "Construct_h", out var h_m))
                    return h_m;

                if (TryGetMmAsMeters(comp, "Construct_a", out var a_m))
                    return a_m;

                // fallback if h missing: use w
                if (TryGetMmAsMeters(comp, "Construct_w", out var w_m))
                    return w_m;

                return 0.0;
            }
            else
            {
                if (TryGetMmAsMeters(comp, "Construct_w", out var w_m))
                    return w_m;

                return 0.0;
            }
        }

        //public static double GetProfileWidth(Component comp)
        //{


        //    if (comp.Template.CustomProperties
        //            .TryGetValue("Construct_w", out var wp)
        //        && double.TryParse(wp.Value.ToString(),
        //                           NumberStyles.Any,
        //                           CultureInfo.InvariantCulture,
        //                           out var w_mm))
        //    {
        //        // convert from millimetres to metres for joint logic
        //        return w_mm * 0.001;
        //    }

        //    var fullDb = comp.Template.Bodies
        //                   .FirstOrDefault(b => b.Name == "ExtrudedProfile")
        //               ?? comp.Template.Bodies.FirstOrDefault(b => b.Shape != null);

        //    if (fullDb != null)
        //    {
        //        var bb = fullDb.Shape.GetBoundingBox(Matrix.Identity, tight: true);
        //        var seg = comp.Template.Curves
        //                        .OfType<DesignCurve>()
        //                        .FirstOrDefault()?.Shape as CurveSegment;
        //        if (seg != null)
        //        {
        //            var d = (seg.EndPoint - seg.StartPoint).Direction.ToVector();
        //            var flat = Vector.Create(d.X, 0, d.Z).Direction.ToVector();
        //            var perp = Vector.Cross(Vector.Create(0, 1, 0), flat).Direction.ToVector();

        //            var corners = Box.Create(new[] { bb.MinCorner, bb.MaxCorner }).Corners;
        //            var center = bb.Center;
        //            double maxDist = corners
        //                .Select(c => Math.Abs(Vector.Dot(perp, c - center)))
        //                .Max();
        //            return maxDist * 2.0;
        //        }
        //    }

        //    return 0.0;
        //}
        //public static double GetProfileWidth(Component comp)
        //{
        //    const double tol = 1e-9;

        //    // 1) Find a solid body to measure (prefer the fresh ExtrudedProfile)
        //    var db = comp.Template.Bodies
        //                 .FirstOrDefault(b => b.Name == "ExtrudedProfile" && b.Shape != null)
        //          ?? comp.Template.Bodies.FirstOrDefault(b => b.Shape != null);
        //    if (db?.Shape == null)
        //        return 0.0;

        //    // 2) Get the construction segment in PART-local space
        //    var seg = comp.Template.Curves.OfType<DesignCurve>()
        //                 .FirstOrDefault()?.Shape as CurveSegment;
        //    if (seg == null)
        //        return 0.0;

        //    // 3) Build the oriented frame in WORLD:
        //    //    ẑ = along the member (construction line)
        //    //    ŷ = world-up projected to ⟂ ẑ (fallback to world-X if needed)
        //    //    x̂ = ẑ × ŷ
        //    Vector z = (seg.EndPoint - seg.StartPoint).Direction.ToVector();
        //    if (z.Magnitude < tol) return 0.0;

        //    Vector up = Vector.Create(0, 1, 0);
        //    up = up - Vector.Dot(up, z) * z;                  // reject z component
        //    if (up.Magnitude < tol)
        //    {
        //        up = Vector.Create(1, 0, 0);
        //        up = up - Vector.Dot(up, z) * z;
        //        if (up.Magnitude < tol) return 0.0;           // degenerate edge case
        //    }
        //    up = up.Direction.ToVector();

        //    Vector x = Vector.Cross(z, up);
        //    if (x.Magnitude < tol) return 0.0;
        //    x = x.Direction.ToVector();

        //    Vector y = Vector.Cross(z, x).Direction.ToVector();

        //    // Origin of the frame doesn’t affect extents; use world origin.
        //    var frameWorld = Frame.Create(Point.Origin, x.Direction, y.Direction);

        //    // 4) Compose transform for the OBB query:
        //    //    Part-local → World (comp.Placement) → Frame-local (inverse of frame mapping)
        //    Matrix worldFromFrame = Matrix.CreateMapping(frameWorld);      // frame-local → world
        //    Matrix toFrameLocal = worldFromFrame.Inverse * comp.Placement; // part-local → frame-local

        //    // 5) Tight AABB in our oriented frame; X-extent is the width
        //    var bb = db.Shape.GetBoundingBox(toFrameLocal, tight: true);
        //    var size = bb.MaxCorner - bb.MinCorner;

        //    // Result is in model units (meters in your codebase).
        //    return Math.Abs(size.X / 10000.0);
        //}

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
