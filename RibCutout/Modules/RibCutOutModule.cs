using AESCConstruct25.FrameGenerator.Utilities; // Logger
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Point = SpaceClaim.Api.V242.Geometry.Point;

namespace AESCConstruct25.RibCutout.Modules
{
    internal static class RibCutOutModule
    {
        // -----------------------------
        // Config - flip this to enable/disable scene debug
        // -----------------------------
        static class Config
        {
            public static bool EnableDebugViz = true;       // turn on/off all debug geometry
            public static double AxisLen = 0.01;            // ~10 mm triad, in meters
            public static Color CutterGhostColor = Color.FromArgb(230, 60, 60);
            public static Color WeldGhostColor = Color.FromArgb(150, 70, 200);
            public static Color EdgeColor = Color.Orange;
            public static Color MidPlaneColor = Color.Cyan;
            public static Color AxisXColor = Color.Red;
            public static Color AxisYColor = Color.LimeGreen;
            public static Color AxisZColor = Color.Blue;
            public static Color DirLineColor = Color.Cyan;
        }

        // Will be set at start of ProcessPairs so helpers can add scene objects
        static Part s_part;

        // -----------------------------
        // Formatting helpers
        // -----------------------------
        static string F(double d) => d.ToString("0.########", CultureInfo.InvariantCulture);
        static string P(Point p) => $"P({F(p.X)},{F(p.Y)},{F(p.Z)})";
        static string D(Direction d) { var v = d.ToVector(); return $"D({F(v.X)},{F(v.Y)},{F(v.Z)})"; }
        static string B(Box b) => $"Box[min={P(b.MinCorner)}, max={P(b.MaxCorner)}]";
        static string FrameInfo(Frame f)
        {
            try { return $"Frame[O={P(f.Origin)}, X={D(f.DirX)}, Y={D(f.DirY)}, Z={D(f.DirZ)}]"; }
            catch { return "Frame[unavailable]"; }
        }

        // -----------------------------
        // Core helpers (with logging)
        // -----------------------------
        static Point CenterOf(Box box)
        {
            try
            {
                var v = 0.5 * (box.MinCorner.Vector + box.MaxCorner.Vector);
                var p = Point.Create(v.X, v.Y, v.Z);
                Logger.Log($"CenterOf -> {P(p)}; {B(box)}");
                return p;
            }
            catch (Exception ex) { Logger.Log($"CenterOf ERROR: {ex.Message}"); throw; }
        }

        static Point ToWorld(Frame f, double lx, double ly, double lz)
        {
            try
            {
                var ox = f.Origin.X; var oy = f.Origin.Y; var oz = f.Origin.Z;
                var vx = f.DirX.ToVector(); var vy = f.DirY.ToVector(); var vz = f.DirZ.ToVector();
                var p = Point.Create(
                    ox + lx * vx.X + ly * vy.X + lz * vz.X,
                    oy + lx * vx.Y + ly * vy.Y + lz * vz.Y,
                    oz + lx * vx.Z + ly * vy.Z + lz * vz.Z
                );
                Logger.Log($"ToWorld local({F(lx)},{F(ly)},{F(lz)}) -> {P(p)} in {FrameInfo(f)}");
                return p;
            }
            catch (Exception ex) { Logger.Log($"ToWorld ERROR: {ex.Message}"); throw; }
        }

        // -----------------------------
        // Owner/frame lookup & vector helpers (new)
        // -----------------------------

        // Find DesignBody that owns a Body.Shape (by reference)
        static DesignBody TryFindOwner(Body target)
        {
            try
            {
                if (s_part == null || target == null) return null;
                return s_part.Bodies.FirstOrDefault(db => object.ReferenceEquals(db.Shape, target));
            }
            catch { return null; }
        }

        // Read a body's frame without returning null (Try-pattern)
        static bool TryGetOwnerFrame(DesignBody owner, out Frame frame)
        {
            frame = default;
            if (owner == null) return false;

            try
            {
                var t = owner.GetType();

                // properties commonly seen across SpaceClaim builds
                var prop = t.GetProperty("Frame", BindingFlags.Public | BindingFlags.Instance)
                           ?? t.GetProperty("CoordinateSystem", BindingFlags.Public | BindingFlags.Instance)
                           ?? t.GetProperty("PlacementFrame", BindingFlags.Public | BindingFlags.Instance);

                if (prop != null)
                {
                    var val = prop.GetValue(owner, null);
                    if (val is Frame fr)
                    {
                        frame = fr;
                        return true;
                    }
                }

                // methods, depending on build
                var m = t.GetMethod("GetCoordinateSystem", BindingFlags.Public | BindingFlags.Instance)
                      ?? t.GetMethod("GetFrame", BindingFlags.Public | BindingFlags.Instance);

                if (m != null)
                {
                    var res = m.Invoke(owner, null);
                    if (res is Frame fr2)
                    {
                        frame = fr2;
                        return true;
                    }
                }
            }
            catch
            {
                // ignore and return false
            }

            return false;
        }

        // Small vector helpers
        static Vector VNorm(Vector v)
        {
            var m = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            return (m > 1e-12) ? Vector.Create(v.X / m, v.Y / m, v.Z / m) : v;
        }
        static Vector VCross(Vector a, Vector b)
        {
            return Vector.Create(
                a.Y * b.Z - a.Z * b.Y,
                a.Z * b.X - a.X * b.Z,
                a.X * b.Y - a.Y * b.X
            );
        }

        // Build a right-handed frame with Z = zDir, and X/Y aligned to the owner's rotation (projected to be ⟂ Z)
        static Frame MakeAlignedFrame(Point origin, Direction zDir, DesignBody ownerOrNull, string logTag)
        {
            var z = VNorm(zDir.ToVector());

            // choose owner axes if available, otherwise world axes
            Vector xRef = Vector.Create(1, 0, 0);
            Vector yRef = Vector.Create(0, 1, 0);

            if (TryGetOwnerFrame(ownerOrNull, out var of))
            {
                Logger.Log($"{logTag}: owner frame -> {FrameInfo(of)} (DesignBody='{ownerOrNull?.Name}')");
                xRef = of.DirX.ToVector();
                yRef = of.DirY.ToVector();
            }
            else
            {
                Logger.Log($"{logTag}: owner frame not available, using world XY as reference.");
            }

            // Pick the reference axis most perpendicular to z (to avoid degeneracy)
            var dotX = Math.Abs(Vector.Dot(xRef, z));
            var dotY = Math.Abs(Vector.Dot(yRef, z));
            var x0 = (dotX < dotY) ? xRef : yRef;

            // Project onto plane ⟂ z and normalize
            var xProj = x0 - Vector.Dot(x0, z) * z;
            var x = VNorm(xProj);

            // If degenerate (parallel), pick a safe fallback
            if (Math.Sqrt(x.X * x.X + x.Y * x.Y + x.Z * x.Z) < 1e-9)
                x = VNorm(Math.Abs(z.X) < 0.9 ? Vector.Create(1, 0, 0) : Vector.Create(0, 1, 0));

            // y completes a right-handed basis
            var y = VNorm(VCross(z, x));

            var fx = Direction.Create(x.X, x.Y, x.Z);
            var fy = Direction.Create(y.X, y.Y, y.Z);
            var local = Frame.Create(origin, fx, fy);

            return local;
        }

        static void LogOwnerRotation(string label, Body body)
        {
            try
            {
                var owner = TryFindOwner(body);
                if (owner != null && TryGetOwnerFrame(owner, out var of))
                    Logger.Log($"Owner rotation -> {FrameInfo(of)} (DesignBody='{owner.Name}')");
                else
                    Logger.Log("Owner frame not available.");
            }
            catch (Exception ex) { Logger.Log($"{label}: rotation log ERROR: {ex.Message}"); }
        }

        // -----------------------------
        // Scene Debug helpers (safe no-ops if s_part is null or viz disabled)
        // -----------------------------
        static void AddGhostBody(Body src, string name, Color color)
        {
            if (!Config.EnableDebugViz || s_part == null || src == null) return;
            try { var db = DesignBody.Create(s_part, name, src.Copy()); db.SetColor(null, color); }
            catch (Exception ex) { Logger.Log($"DebugViz AddGhostBody ERROR: {ex.Message}"); }
        }

        static void AddPolyline(IReadOnlyList<Point> pts, string name, Color color)
        {
            if (!Config.EnableDebugViz || s_part == null || pts == null || pts.Count < 2) return;
            try
            {
                for (int i = 0; i < pts.Count - 1; i++)
                {
                    var seg = CurveSegment.Create(pts[i], pts[i + 1]);
                    var dc = DesignCurve.Create(s_part, seg);
                    dc.Name = name;
                    dc.SetColor(null, color);
                }
            }
            catch (Exception ex) { Logger.Log($"DebugViz AddPolyline ERROR: {ex.Message}"); }
        }

        static void AddAxes(Frame f, string name, double len)
        {
            if (!Config.EnableDebugViz) return;
            try
            {
                var o = f.Origin;
                AddPolyline(new[] { o, o + f.DirX.ToVector() * len }, $"{name}/X", Config.AxisXColor);
                AddPolyline(new[] { o, o + f.DirY.ToVector() * len }, $"{name}/Y", Config.AxisYColor);
                AddPolyline(new[] { o, o + f.DirZ.ToVector() * len }, $"{name}/Z", Config.AxisZColor);
            }
            catch (Exception ex) { Logger.Log($"DebugViz Axes ERROR: {ex.Message}"); }
        }

        static void AddBoxEdges(Frame local, double x0, double y0, double z0, double x1, double y1, double z1, string name)
        {
            if (!Config.EnableDebugViz) return;
            try
            {
                // 8 corners in local -> world
                var p000 = ToWorld(local, x0, y0, z0);
                var p001 = ToWorld(local, x0, y0, z1);
                var p010 = ToWorld(local, x0, y1, z0);
                var p011 = ToWorld(local, x0, y1, z1);
                var p100 = ToWorld(local, x1, y0, z0);
                var p101 = ToWorld(local, x1, y0, z1);
                var p110 = ToWorld(local, x1, y1, z0);
                var p111 = ToWorld(local, x1, y1, z1);

                var c = Config.EdgeColor;

                // edges
                AddPolyline(new[] { p000, p100, p110, p010, p000 }, $"{name}/edge-xy(z0)", c);
                AddPolyline(new[] { p001, p101, p111, p011, p001 }, $"{name}/edge-xy(z1)", c);
                AddPolyline(new[] { p000, p001 }, $"{name}/edge-z0", c);
                AddPolyline(new[] { p100, p101 }, $"{name}/edge-z1", c);
                AddPolyline(new[] { p110, p111 }, $"{name}/edge-z2", c);
                AddPolyline(new[] { p010, p011 }, $"{name}/edge-z3", c);

                // also log for debugging
                Logger.Log($"{name}/corners: " +
                           $"p000={P(p000)}, p001={P(p001)}, p010={P(p010)}, p011={P(p011)}, " +
                           $"p100={P(p100)}, p101={P(p101)}, p110={P(p110)}, p111={P(p111)}");
            }
            catch (Exception ex) { Logger.Log($"DebugViz BoxEdges ERROR: {ex.Message}"); }
        }

        static void AddMidPlaneRect(Frame local, double x0, double x1, double y0, double y1, string name)
        {
            if (!Config.EnableDebugViz) return;
            try
            {
                // draw rectangle on local z=0 (split plane)
                var a = ToWorld(local, x0, y0, 0);
                var b = ToWorld(local, x1, y0, 0);
                var c = ToWorld(local, x1, y1, 0);
                var d = ToWorld(local, x0, y1, 0);
                AddPolyline(new[] { a, b, c, d, a }, $"{name}/midZ", Config.MidPlaneColor);
                Logger.Log($"{name}/midZ: a={P(a)} b={P(b)} c={P(c)} d={P(d)}");
            }
            catch (Exception ex) { Logger.Log($"DebugViz MidPlane ERROR: {ex.Message}"); }
        }

        // -----------------------------
        // Solid constructors
        // -----------------------------

        // Rectangular prism cutter in a local frame where Z = "forward" (cut direction)
        static Body MakeRectPrismCutter(Frame local, double x0, double y0, double z0, double x1, double y1, double z1)
        {
            try
            {
                Logger.Log($"MakeRectPrismCutter: local={FrameInfo(local)}; x0={F(x0)}, y0={F(y0)}, z0={F(z0)}, x1={F(x1)}, y1={F(y1)}, z1={F(z1)}");

                // Base rectangle on plane z=z0 in local frame
                var p0 = ToWorld(local, x0, y0, z0);
                var p1 = ToWorld(local, x1, y0, z0);
                var p2 = ToWorld(local, x1, y1, z0);
                var p3 = ToWorld(local, x0, y1, z0);

                var baseOrigin = ToWorld(local, 0, 0, z0);
                var baseFrame = Frame.Create(baseOrigin, local.DirX, local.DirY); // normal=local.DirZ
                var plane = Plane.Create(baseFrame);

                // sanity check distances along normal
                var n = local.DirZ.ToVector();
                double d0 = Vector.Dot((p0 - baseOrigin), n);
                double d1 = Vector.Dot((p1 - baseOrigin), n);
                double d2 = Vector.Dot((p2 - baseOrigin), n);
                double d3 = Vector.Dot((p3 - baseOrigin), n);
                Logger.Log($"MakeRectPrismCutter: plane-z0 checks d0={F(d0)} d1={F(d1)} d2={F(d2)} d3={F(d3)}");

                var segs = new[]
                {
                    CurveSegment.Create(p0, p1),
                    CurveSegment.Create(p1, p2),
                    CurveSegment.Create(p2, p3),
                    CurveSegment.Create(p3, p0),
                };

                var profile = new Profile(plane, segs);

                double height = Math.Max(0, z1 - z0);
                var body = Body.ExtrudeProfile(profile, height);

                Logger.Log($"MakeRectPrismCutter: height={F(height)}; result volume={F(body.Volume)}");
                return body;
            }
            catch (Exception ex) { Logger.Log($"MakeRectPrismCutter ERROR: {ex.Message}"); throw; }
        }

        // Cylinder extruded along a chosen LOCAL axis ('X','Y','Z'), base at the corresponding min value
        static Body MakeAxisCylinder(Frame local, char axis, double cx, double cy, double cz, double radius, double height, int sides = 24)
        {
            try
            {
                radius = Math.Max(0, radius);
                height = Math.Max(0, height);

                // Pick a profile plane whose normal = requested axis
                Frame circleFrame;
                switch (char.ToUpperInvariant(axis))
                {
                    case 'X': circleFrame = Frame.Create(ToWorld(local, cx, cy, cz), local.DirY, local.DirZ); break; // normal X
                    case 'Y': circleFrame = Frame.Create(ToWorld(local, cx, cy, cz), local.DirZ, local.DirX); break; // normal Y
                    default: circleFrame = Frame.Create(ToWorld(local, cx, cy, cz), local.DirX, local.DirY); break; // normal Z
                }

                var plane = Plane.Create(circleFrame);
                var circle = Circle.Create(circleFrame, radius);
                var loop = new[] { CurveSegment.Create(circle) };
                var prof = new Profile(plane, loop);

                var body = Body.ExtrudeProfile(prof, height);
                Logger.Log($"MakeAxisCylinder(axis={axis}): centerLocal=({F(cx)},{F(cy)},{F(cz)}), r={F(radius)}, h={F(height)}; vol={F(body.Volume)}");
                return body;
            }
            catch (Exception ex) { Logger.Log($"MakeAxisCylinder ERROR: {ex.Message}"); throw; }
        }

        // Try to get endpoints of an edge, across API variants.
        static bool TryGetEdgeEndpoints(Edge e, out Point p0, out Point p1)
        {
            p0 = default; p1 = default;
            try
            {
                // common pattern: StartVertex/EndVertex -> Position
                var t = e.GetType();
                var svProp = t.GetProperty("StartVertex") ?? t.GetProperty("Start");
                var evProp = t.GetProperty("EndVertex") ?? t.GetProperty("End");
                if (svProp != null && evProp != null)
                {
                    var sv = svProp.GetValue(e, null);
                    var ev = evProp.GetValue(e, null);
                    if (sv != null && ev != null)
                    {
                        var vt = sv.GetType();
                        var posProp = vt.GetProperty("Position") ?? vt.GetProperty("Point");
                        if (posProp != null)
                        {
                            p0 = (Point)posProp.GetValue(sv, null);
                            p1 = (Point)posProp.GetValue(ev, null);
                            return true;
                        }
                    }
                }

                // sometimes edges expose StartPoint / EndPoint directly
                var sp = t.GetProperty("StartPoint") ?? t.GetProperty("PointStart");
                var ep = t.GetProperty("EndPoint") ?? t.GetProperty("PointEnd");
                if (sp != null && ep != null)
                {
                    p0 = (Point)sp.GetValue(e, null);
                    p1 = (Point)ep.GetValue(e, null);
                    return true;
                }
            }
            catch { /* ignore */ }
            return false;
        }

        // Build a local frame whose X is along a given vector (used for the probe tube)
        static Frame MakeFrameWithX(Point origin, Direction xDir)
        {
            var x = VNorm(xDir.ToVector());
            // pick a safe ref that is not parallel to X
            var refZ = (Math.Abs(x.Z) < 0.9) ? Vector.Create(0, 0, 1) : Vector.Create(0, 1, 0);
            var z = VNorm(VCross(x, refZ));          // z ⟂ x
            if (Math.Sqrt(z.X * z.X + z.Y * z.Y + z.Z * z.Z) < 1e-9)
                z = VNorm(VCross(x, Vector.Create(0, 1, 0)));
            var y = VNorm(VCross(z, x));             // right-handed
            return Frame.Create(origin,
                Direction.Create(x.X, x.Y, x.Z),
                Direction.Create(y.X, y.Y, y.Z));
        }

        // Probe whether an edge (as a thin cylinder) intersects 'other' body
        static bool EdgeIntersectsBody(Point a, Point b, Body other, out double score)
        {
            score = 0;
            try
            {
                var v = b - a;
                var len = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
                if (len < 1e-6) return false;

                var mid = Point.Create((a.X + b.X) * 0.5, (a.Y + b.Y) * 0.5, (a.Z + b.Z) * 0.5);
                var xDir = Direction.Create(v.X, v.Y, v.Z);
                var f = MakeFrameWithX(mid, xDir);

                // build a very thin tube centered on the edge, along local X
                var tube = MakeAxisCylinder(f, 'X', 0, 0, 0, 0.0005, len, 16);

                var test = other.Copy();
                test.Intersect(new[] { tube });
                score = test.Volume;              // larger = stronger intersection
                return score > 1e-12;
            }
            catch { return false; }
        }

        // Find a good Y direction from 'source' that intersects 'other'.
        // Returns true and sets yDir if found.
        static bool TryPickIntersectingY(Body source, Body other, out Direction yDir, string logTag)
        {
            yDir = default;
            try
            {
                // Enumerate edges via reflection-friendly path
                var edgesProp = source.GetType().GetProperty("Edges", BindingFlags.Public | BindingFlags.Instance);
                var edgesObj = edgesProp?.GetValue(source, null) as System.Collections.IEnumerable;
                if (edgesObj == null)
                {
                    Logger.Log($"{logTag}: source body has no Edges enumerable.");
                    return false;
                }

                double bestScore = 0;
                Vector bestVec = default;
                foreach (var eo in edgesObj)
                {
                    if (eo is Edge e && TryGetEdgeEndpoints(e, out var p0, out var p1))
                    {
                        if (EdgeIntersectsBody(p0, p1, other, out var s))
                        {
                            var v = p1 - p0;
                            if (s > bestScore && Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z) > 1e-6)
                            {
                                bestScore = s;
                                bestVec = VNorm(v);
                            }
                        }
                    }
                }

                if (bestScore > 0)
                {
                    yDir = Direction.Create(bestVec.X, bestVec.Y, bestVec.Z);
                    Logger.Log($"{logTag}: picked Y from intersecting edge (score={F(bestScore)}) -> {D(yDir)}");
                    return true;
                }

                Logger.Log($"{logTag}: no intersecting edge found.");
                return false;
            }
            catch (Exception ex)
            {
                Logger.Log($"{logTag}: TryPickIntersectingY ERROR: {ex.Message}");
                return false;
            }
        }

        // -----------------------------
        // Main entry
        // -----------------------------
        public static void ProcessPairs(
            Document doc,
            List<(Body A, Body B)> pairs,
            bool perpendicularCut,
            double toleranceMM,
            bool applyMiddleTolerance,
            bool reverseDirection,
            bool addWeldRound,
            double weldRadiusMM
        )
        {
            var opId = Guid.NewGuid().ToString("N").Substring(0, 8);

            try
            {
                var part = doc?.MainPart;
                s_part = part;
                if (part == null)
                {
                    Logger.Log($"[{opId}] ERROR: Document.MainPart is null. Aborting.");
                    return;
                }

                // Units: SpaceClaim is meters; UI is mm
                double halfTol = Math.Max(0.0, toleranceMM) / 2000.0;   // tolerance/2 in meters
                double weldR = Math.Max(0.0, weldRadiusMM) / 1000.0;    // radius in meters

                if (pairs == null || pairs.Count == 0)
                {
                    Logger.Log($"[{opId}] No pairs supplied. Nothing to do.");
                    return;
                }

                for (int i = 0; i < pairs.Count; i++)
                {
                    var (origA, origB) = pairs[i];

                    try
                    {
                        if (origA == null || origB == null)
                        {
                            Logger.Log($"[{opId}] Pair[{i}] ERROR: One or both bodies are null. Skipping.");
                            continue;
                        }

                        // Log original body "rotation" (owner frame if available)
                        LogOwnerRotation($"[{opId}] Pair[{i}] A", origA);
                        LogOwnerRotation($"[{opId}] Pair[{i}] B", origB);

                        var bbA = origA.GetBoundingBox(Matrix.Identity, true);
                        var bbB = origB.GetBoundingBox(Matrix.Identity, true);

                        // Overlap (in current world)
                        var overlap = origA.Copy();
                        overlap.Intersect(new[] { origB.Copy() });

                        if (overlap.Volume <= 0)
                        {
                            Logger.Log($"[{opId}] Pair[{i}] No overlap. Skipping.");
                            continue;
                        }

                        var centerA = CenterOf(origA.GetBoundingBox(Matrix.Identity, true));
                        var centerB = CenterOf(origB.GetBoundingBox(Matrix.Identity, true));
                        var centerOverlap = CenterOf(overlap.GetBoundingBox(Matrix.Identity, true));

                        // Local function that builds the cutter for a given target (no subtraction here).
                        Body BuildHalfLapCutter(
                            Body target,
                            Body other,              // the *other* body in the pair
                            DesignBody alignOwner,   // resolved owner for 'target'
                            Point targetCenter,
                            string debugPrefix
                        )
                        {
                            try
                            {
                                // ---------------------------------------------------------------------
                                // PERPENDICULAR CUT PATH: force cut normal to the target body's X.
                                // ---------------------------------------------------------------------
                                if (perpendicularCut)
                                {
                                    // Direction of cut (forward): from target toward overlap, optionally reversed
                                    var dir = (centerOverlap - targetCenter).Direction;
                                    var dirOrig = D(dir);
                                    if (reverseDirection)
                                    {
                                        var v = dir.ToVector();
                                        dir = Direction.Create(-v.X, -v.Y, -v.Z);
                                    }

                                    // Split overlap by plane through center, normal 'dir'
                                    var splitPlane = Plane.Create(Frame.Create(centerOverlap, dir));
                                    var overlapForTarget = overlap.Copy();
                                    overlapForTarget.Split(splitPlane, null);
                                    var pieces = overlapForTarget.SeparatePieces();

                                    if (pieces.Count != 2)
                                    {
                                        Logger.Log($"[{opId}] Pair[{i}] {debugPrefix}: Unexpected piece count ({pieces.Count}). Aborting this half.");
                                        return null;
                                    }

                                    // Choose half closer to target along -dir
                                    Body chosen = null;
                                    double bestDot = double.NegativeInfinity;
                                    foreach (var piece in pieces)
                                    {
                                        var pc = CenterOf(piece.GetBoundingBox(Matrix.Identity, true));
                                        var v = pc - centerOverlap;
                                        double score = -Vector.Dot(v, dir.ToVector());
                                        if (score > bestDot) { bestDot = score; chosen = piece; }
                                    }
                                    if (chosen == null)
                                    {
                                        Logger.Log($"[{opId}] Pair[{i}] {debugPrefix}: No chosen piece. Aborting this half.");
                                        return null;
                                    }

                                    // Local frame: Z = cut direction, X/Y aligned with the target body's rotation
                                    Direction yHint;
                                    var gotY = TryPickIntersectingY(target, other, out yHint, $"[{opId}] Pair[{i}] {debugPrefix}");

                                    Frame local;
                                    if (gotY)
                                    {
                                        var z = VNorm(dir.ToVector());
                                        var y0 = VNorm(yHint.ToVector());

                                        // make y perpendicular to z (project off any small component)
                                        var yProj = y0 - Vector.Dot(y0, z) * z;
                                        var y = VNorm(yProj);
                                        if (Math.Sqrt(y.X * y.X + y.Y * y.Y + y.Z * y.Z) < 1e-9)
                                        {
                                            var owner = alignOwner;
                                            if (owner != null && TryGetOwnerFrame(owner, out var of2))
                                            {
                                                y0 = of2.DirY.ToVector();
                                                yProj = y0 - Vector.Dot(y0, z) * z;
                                                y = VNorm(yProj);
                                            }
                                            else
                                            {
                                                y = (Math.Abs(z.X) < 0.9) ? Vector.Create(1, 0, 0) : Vector.Create(0, 1, 0);
                                                y = VNorm(y - Vector.Dot(y, z) * z);
                                            }
                                        }

                                        var x = VNorm(VCross(y, z));       // X = Y × Z
                                        y = VNorm(VCross(z, x));           // re-orthonormalize Y

                                        local = Frame.Create(
                                            centerOverlap,
                                            Direction.Create(x.X, x.Y, x.Z),
                                            Direction.Create(y.X, y.Y, y.Z)
                                        );
                                    }
                                    else
                                    {
                                        // fallback: owner-aligned XY with Z≈dir
                                        local = MakeAlignedFrame(
                                            centerOverlap,
                                            dir,
                                            alignOwner,
                                            $"[{opId}] Pair[{i}] {debugPrefix}"
                                        );
                                    }

                                    // Correct world->local mapping
                                    Matrix localToWorld, worldToLocal;
                                    try { localToWorld = Matrix.CreateMapping(local); worldToLocal = localToWorld.Inverse; }
                                    catch { localToWorld = Matrix.CreateMapping(local); worldToLocal = localToWorld.Inverse; }

                                    // Oriented bbox in local coords (do not mutate geometry)
                                    var bbLocal = chosen.GetBoundingBox(worldToLocal, true);
                                    var min = bbLocal.MinCorner; var max = bbLocal.MaxCorner;

                                    // Spans (informational for weld logic)
                                    double sx = max.X - min.X, sy = max.Y - min.Y, sz = max.Z - min.Z;

                                    // For welds only (tolerance is now independent)
                                    char perpAxis = (sy <= sx) ? 'Y' : 'X';
                                    char sideAxis = (perpAxis == 'X') ? 'Y' : 'X';

                                    // Start extents (no tol)
                                    double x0a = min.X, x1a = max.X;
                                    double y0a = min.Y, y1a = max.Y; // keep your current "huge Y" behavior
                                    double z0a = min.Z, z1a = max.Z;

                                    // TOLERANCE (updated): X±halfTol, Z+halfTol only if checkbox
                                    y0a -= halfTol * 2.0;
                                    y1a += halfTol * 2.0;

                                    double fwdExtraA = applyMiddleTolerance ? halfTol : 0.0;
                                    if (!reverseDirection)
                                    {
                                        z1a += fwdExtraA;  // forward = +Z
                                    }
                                    else
                                    {
                                        z1a += fwdExtraA;  // forward = -Z
                                        z0a -= 200.0;
                                    }

                                    // Build cutter prism
                                    var cutter = MakeRectPrismCutter(local, x0a, y0a, z0a, x1a, y1a, z1a);

                                    if (addWeldRound && weldR > 0)
                                    {
                                        // Always place holes on the z1 face (the face touching the body),
                                        // regardless of reverseDirection. Reverse only changes how deep we extend z0a.
                                        double zFaceForHoles = z1a;

                                        // Sweep along local X across the full width
                                        double hX = Math.Max(0.0, x1a - x0a);

                                        // Two opposite corners at (y0a, z1a) and (y1a, z1a), sweep +X from x0a → x1a
                                        var cX1 = MakeAxisCylinder(local, 'X', x0a, y0a, zFaceForHoles, weldR, hX);
                                        var cX3 = MakeAxisCylinder(local, 'X', x0a, y1a, zFaceForHoles, weldR, hX);

                                        cutter.Unite(new[] { cX1, cX3 });
                                    }

                                    return cutter;
                                }
                                // ---------------------------------------------------------------------
                                // ANGLED CUT PATH (original logic)
                                // ---------------------------------------------------------------------
                                else
                                {
                                    // Direction of cut (forward): from target toward overlap, optionally reversed
                                    var dir = (centerOverlap - targetCenter).Direction;
                                    var dirOrig = D(dir);
                                    if (reverseDirection)
                                    {
                                        var v = dir.ToVector();
                                        dir = Direction.Create(-v.X, -v.Y, -v.Z);
                                    }

                                    // Split overlap by plane through center, normal 'dir'
                                    var splitPlane = Plane.Create(Frame.Create(centerOverlap, dir));
                                    var overlapForTarget = overlap.Copy();
                                    overlapForTarget.Split(splitPlane, null);
                                    var pieces = overlapForTarget.SeparatePieces();

                                    if (pieces.Count != 2)
                                    {
                                        Logger.Log($"[{opId}] Pair[{i}] {debugPrefix}: Unexpected piece count ({pieces.Count}). Aborting this half.");
                                        return null;
                                    }

                                    // Choose half closer to target along -dir
                                    Body chosen = null;
                                    double bestDot = double.NegativeInfinity;
                                    foreach (var piece in pieces)
                                    {
                                        var pc = CenterOf(piece.GetBoundingBox(Matrix.Identity, true));
                                        var v = pc - centerOverlap;
                                        double score = -Vector.Dot(v, dir.ToVector());
                                        if (score > bestDot) { bestDot = score; chosen = piece; }
                                    }
                                    if (chosen == null)
                                    {
                                        Logger.Log($"[{opId}] Pair[{i}] {debugPrefix}: No chosen piece. Aborting this half.");
                                        return null;
                                    }

                                    // Local frame: Z = cut direction, X/Y aligned with the target body's rotation
                                    Direction yHint;
                                    var gotY = TryPickIntersectingY(target, other, out yHint, $"[{opId}] Pair[{i}] {debugPrefix}");

                                    Frame local;
                                    if (gotY)
                                    {
                                        var z = VNorm(dir.ToVector());
                                        var y0 = VNorm(yHint.ToVector());

                                        // make y perpendicular to z (project off any small component)
                                        var yProj = y0 - Vector.Dot(y0, z) * z;
                                        var y = VNorm(yProj);
                                        if (Math.Sqrt(y.X * y.X + y.Y * y.Y + y.Z * y.Z) < 1e-9)
                                        {
                                            var owner = alignOwner;
                                            if (owner != null && TryGetOwnerFrame(owner, out var of2))
                                            {
                                                y0 = of2.DirY.ToVector();
                                                yProj = y0 - Vector.Dot(y0, z) * z;
                                                y = VNorm(yProj);
                                            }
                                            else
                                            {
                                                y = (Math.Abs(z.X) < 0.9) ? Vector.Create(1, 0, 0) : Vector.Create(0, 1, 0);
                                                y = VNorm(y - Vector.Dot(y, z) * z);
                                            }
                                        }

                                        var x = VNorm(VCross(y, z));
                                        y = VNorm(VCross(z, x));

                                        local = Frame.Create(
                                            centerOverlap,
                                            Direction.Create(x.X, x.Y, x.Z),
                                            Direction.Create(y.X, y.Y, y.Z)
                                        );
                                    }
                                    else
                                    {
                                        // fallback: owner-aligned XY with Z≈dir
                                        local = MakeAlignedFrame(
                                            centerOverlap,
                                            dir,
                                            alignOwner,
                                            $"[{opId}] Pair[{i}] {debugPrefix}"
                                        );
                                    }

                                    // Correct world->local mapping
                                    Matrix localToWorld, worldToLocal;
                                    try { localToWorld = Matrix.CreateMapping(local); worldToLocal = localToWorld.Inverse; }
                                    catch { localToWorld = Matrix.CreateMapping(local); worldToLocal = localToWorld.Inverse; }

                                    // Oriented bbox in local coords (do not mutate geometry)
                                    var bbLocal = chosen.GetBoundingBox(worldToLocal, true);
                                    var min = bbLocal.MinCorner; var max = bbLocal.MaxCorner;

                                    // Spans (informational for weld logic)
                                    double sx = max.X - min.X, sy = max.Y - min.Y, sz = max.Z - min.Z;

                                    // For welds only (tolerance is now independent)
                                    char perpAxis = (sy <= sx) ? 'Y' : 'X';
                                    char sideAxis = (perpAxis == 'X') ? 'Y' : 'X';

                                    // Start extents (no tol)
                                    double x0a = min.X, x1a = max.X;
                                    double y0a = min.Y * 1000.0, y1a = max.Y * 1000.0; // keep your current "huge Y" behavior
                                    double z0a = min.Z, z1a = max.Z;

                                    // TOLERANCE (updated): X±halfTol, Z+halfTol only if checkbox
                                    x0a -= halfTol * 2.0;
                                    x1a += halfTol * 2.0;

                                    double fwdExtraA = applyMiddleTolerance ? halfTol : 0.0;

                                    if (!reverseDirection)
                                    {
                                        z1a += fwdExtraA;  // forward = +Z
                                    }
                                    else
                                    {
                                        z1a += fwdExtraA;  // forward = -Z
                                        z0a -= 200.0;
                                    }

                                    // Build cutter prism
                                    var cutter = MakeRectPrismCutter(local, x0a, y0a, z0a, x1a, y1a, z1a);

                                    // Optional weld-round reliefs (unchanged)
                                    if (addWeldRound && weldR > 0)
                                    {
                                        // Sweep along Y from y0a -> y1a (full edge length)
                                        double hY = Math.Max(0.0, y1a - y0a);

                                        // Build at the far Z face (z1a), aligned with the Y-parallel edges at x=x0a and x=x1a
                                        double zFaceY = z1a;
                                        double yBase = y0a;   // base at min Y, extrusion goes +Y for hY

                                        // Left and right Y-edges on the z1 face
                                        var cYLeft = MakeAxisCylinder(local, 'Y', x0a, yBase, zFaceY, weldR, hY);
                                        var cYRight = MakeAxisCylinder(local, 'Y', x1a, yBase, zFaceY, weldR, hY);

                                        // Unite into cutter
                                        cutter.Unite(new[] { cYLeft, cYRight });
                                    }

                                    return cutter;
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.Log($"[{opId}] Pair[{i}] {debugPrefix}: BuildHalfLapCutter ERROR: {ex.Message}");
                                return null;
                            }
                        }

                        // Build BOTH cutters
                        var ownerA = TryFindOwner(origA);
                        var ownerB = TryFindOwner(origB);
                        Body cutterA;
                        Body cutterB;
                        if (perpendicularCut)
                        {
                            cutterA = BuildHalfLapCutter(origB, origA, ownerA, centerA, "A");
                            cutterB = BuildHalfLapCutter(origA, origB, ownerB, centerB, "B");
                        }
                        else
                        {
                            cutterA = BuildHalfLapCutter(origA, origB, ownerA, centerA, "A");
                            cutterB = BuildHalfLapCutter(origB, origA, ownerB, centerB, "B");
                        }

                        // -----------------------------
                        //   - cutterA is applied to body B
                        //   - cutterB is applied to body A
                        // -----------------------------
                        void ApplySubtraction(Body target, Body cutter, string applyTag)
                        {
                            if (target == null || cutter == null) return;

                            try
                            {
                                var cutterDb = DesignBody.Create(s_part, $"{applyTag}_Cutter", cutter.Copy());
                                cutterDb.SetColor(null, Color.Red);

                                var volBefore = target.Volume;
                                target.Subtract(new[] { cutterDb.Shape });
                                var volAfter = target.Volume;

                                try
                                {
                                    var owner = s_part.Bodies.FirstOrDefault(db => object.ReferenceEquals(db.Shape, target));
                                }
                                catch (Exception exOwn) { Logger.Log($"[{opId}] {applyTag}: Owner lookup ERROR: {exOwn.Message}"); }
                            }
                            catch (Exception exBool) { Logger.Log($"[{opId}] {applyTag}: Boolean subtract ERROR: {exBool.Message}"); }
                        }

                        ApplySubtraction(origB, cutterA, $"Pair[{i}] Apply cutterA->B");
                        ApplySubtraction(origA, cutterB, $"Pair[{i}] Apply cutterB->A");
                    }
                    catch (Exception exPair) { Logger.Log($"[{opId}] Pair[{i}] ERROR (outer): {exPair.Message}"); }
                }
            }
            catch (Exception ex) { Logger.Log($"[{opId}] RibCutOut.ProcessPairs FATAL: {ex.Message}"); }

            finally { s_part = null; }
        }
    }
}
