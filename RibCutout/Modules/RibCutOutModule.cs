/*
 RibCutOutModule contains the geometric core for creating half-lap rib cut-out joints.
 It builds local frames from body edges, constructs cutter prisms and optional weld-round cylinders,
 and subtracts them from body pairs to form the final joint geometry.
*/

using AESCConstruct2026.FrameGenerator.Utilities; // Logger
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using Point = SpaceClaim.Api.V242.Geometry.Point;

namespace AESCConstruct2026.RibCutout.Modules
{
    internal static class RibCutOutModule
    {
        // Parameters for half-lap joint
        public struct HalfLapParams
        {
            public double Width;
            public double Length;
            public double Thickness;
            public double Depth;
            public double Shoulder;
            public double Check;
        }

        static Part s_part;

        const double edgeOffset = 100.0;

        // Formats a double using invariant culture with up to 8 decimals.
        static string F(double d) => d.ToString("0.########", CultureInfo.InvariantCulture);

        // Returns the geometric center point of a box.
        static Point CenterOf(Box box)
        {
            var v = 0.5 * (box.MinCorner.Vector + box.MaxCorner.Vector);
            return Point.Create(v.X, v.Y, v.Z);
        }

        // Normalizes a vector; returns original if its magnitude is near zero.
        static Vector VNorm(Vector v)
        {
            var m = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            return (m > 1e-12) ? Vector.Create(v.X / m, v.Y / m, v.Z / m) : v;
        }

        // Computes the cross product of two vectors.
        static Vector VCross(Vector a, Vector b)
        {
            return Vector.Create(
                a.Y * b.Z - a.Z * b.Y,
                a.Z * b.X - a.X * b.Z,
                a.X * b.Y - a.Y * b.X
            );
        }

        // Draws debug axis lines (X, Y, Z) in the given frame inside the active part.
        static void AddAxes(Frame f, string name, double len = 0.01)
        {
            var o = f.Origin;
            AddPolyline(new[] { o, o + f.DirX.ToVector() * len }, $"{name}/X", Color.Red);
            AddPolyline(new[] { o, o + f.DirY.ToVector() * len }, $"{name}/Y", Color.LimeGreen);
            AddPolyline(new[] { o, o + f.DirZ.ToVector() * len }, $"{name}/Z", Color.Blue);
        }

        // Creates colored debug line segments in the active part between successive points.
        static void AddPolyline(IReadOnlyList<Point> pts, string name, Color color)
        {
            if (s_part == null || pts == null || pts.Count < 2)
            {
                return;
            }
            for (int i = 0; i < pts.Count - 1; i++)
            {
                var seg = CurveSegment.Create(pts[i], pts[i + 1]);
                var dc = DesignCurve.Create(s_part, seg);
                dc.Name = name;
                dc.SetColor(null, color);
            }
        }

        // Derives a local orthogonal frame for a body based on its edge directions and extents.
        static Frame CreateLocalFrameFromEdges(Body body)
        {
            // 1. Gather all edge directions in world space
            var edgeDirs = new List<Vector>();
            var edgesProp = body.GetType().GetProperty("Edges");
            var edgesObj = edgesProp?.GetValue(body, null) as System.Collections.IEnumerable;
            if (edgesObj == null)
                throw new InvalidOperationException("Body has no Edges property.");

            foreach (var edgeObj in edgesObj)
            {
                // Try to get start and end points of the edge
                var t = edgeObj.GetType();
                var sp = t.GetProperty("StartPoint") ?? t.GetProperty("PointStart");
                var ep = t.GetProperty("EndPoint") ?? t.GetProperty("PointEnd");
                if (sp != null && ep != null)
                {
                    var p0 = (Point)sp.GetValue(edgeObj, null);
                    var p1 = (Point)ep.GetValue(edgeObj, null);
                    var dir = p1 - p0;
                    if (dir.Magnitude > 1e-8)
                        edgeDirs.Add(VNorm(dir));
                }
            }

            if (edgeDirs.Count < 3)
                throw new InvalidOperationException("Not enough unique edges found.");

            // 2. Cluster edge directions (rectangular body: 3 unique directions)
            // Use a tolerance to group similar directions (opposite directions are considered the same)
            double tol = 1e-3;
            var uniqueDirs = new List<Vector>();
            foreach (var dir in edgeDirs)
            {
                bool found = false;
                foreach (var u in uniqueDirs)
                {
                    if (Math.Abs(Vector.Dot(dir, u)) > 1 - tol)
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                    uniqueDirs.Add(dir);
            }

            if (uniqueDirs.Count != 3)
                throw new InvalidOperationException($"Expected 3 unique edge directions, found {uniqueDirs.Count}.");

            // 3. For each direction, compute the total length of all edges in that direction
            var dirLengths = new double[3];
            for (int i = 0; i < 3; i++)
            {
                double sum = 0;
                foreach (var edgeObj in edgesObj)
                {
                    var t = edgeObj.GetType();
                    var sp = t.GetProperty("StartPoint") ?? t.GetProperty("PointStart");
                    var ep = t.GetProperty("EndPoint") ?? t.GetProperty("PointEnd");
                    if (sp != null && ep != null)
                    {
                        var p0 = (Point)sp.GetValue(edgeObj, null);
                        var p1 = (Point)ep.GetValue(edgeObj, null);
                        var dir = p1 - p0;
                        var norm = VNorm(dir);
                        if (Math.Abs(Vector.Dot(norm, uniqueDirs[i])) > 1 - tol)
                            sum += dir.Magnitude;
                    }
                }
                dirLengths[i] = sum;
            }

            // 4. Assign axes: X = longest, Z = shortest, Y = remaining (right-hand rule)
            int xIdx = Array.IndexOf(dirLengths, dirLengths.Max());
            int zIdx = Array.IndexOf(dirLengths, dirLengths.Min());
            int yIdx = 3 - xIdx - zIdx;

            Vector xDir = uniqueDirs[xIdx];
            Vector zDir = uniqueDirs[zIdx];
            Vector yDir = VNorm(VCross(zDir, xDir)); // right-hand rule

            // 5. Use the center of the bounding box as the origin
            var bb = body.GetBoundingBox(Matrix.Identity, true);
            var center = CenterOf(bb);

            return Frame.Create(center,
                Direction.Create(xDir.X, xDir.Y, xDir.Z),
                Direction.Create(yDir.X, yDir.Y, yDir.Z)
            );
        }

        // Builds a rectangular prism in a local frame, applying tolerance and optional middle tolerance.
        static Body MakeRectPrism(Frame local, double x0, double y0, double z0, double x1, double y1, double z1, double toleranceMM, bool middleTolerance, bool perpendicularCut)
        {
            // Apply tolerance in X and Z directions (widen symmetrically)
            double minX = perpendicularCut ? Math.Min(x0, x1) - toleranceMM / 1000.0 : Math.Min(x0, x1) - toleranceMM;
            double maxX = perpendicularCut ? Math.Max(x0, x1) + toleranceMM / 1000.0 : Math.Max(x0, x1) + toleranceMM;
            double minZ = Math.Min(z0, z1) - toleranceMM / 1000.0;
            double maxZ = Math.Max(z0, z1) + toleranceMM / 1000.0;
            double minY = middleTolerance ? Math.Min(y0, y1) - toleranceMM / 2000.0 : Math.Min(y0, y1);
            double maxY = middleTolerance ? Math.Max(y0, y1) + toleranceMM / 2000.0 : Math.Max(y0, y1);

            var p0 = ToWorld(local, minX, minY, minZ);
            var p1 = ToWorld(local, maxX, minY, minZ);
            var p2 = ToWorld(local, maxX, maxY, minZ);
            var p3 = ToWorld(local, minX, maxY, minZ);

            var baseOrigin = ToWorld(local, 0, 0, minZ);
            var baseFrame = Frame.Create(baseOrigin, local.DirX, local.DirY);
            var plane = Plane.Create(baseFrame);

            var segs = new[]
            {
                CurveSegment.Create(p0, p1),
                CurveSegment.Create(p1, p2),
                CurveSegment.Create(p2, p3),
                CurveSegment.Create(p3, p0),
            };

            var profile = new Profile(plane, segs);
            double height = Math.Max(0, maxZ - minZ);
            var body = Body.ExtrudeProfile(profile, height);
            return body;
        }

        // Transforms local frame coordinates (lx, ly, lz) into world coordinates.
        static Point ToWorld(Frame f, double lx, double ly, double lz)
        {
            var ox = f.Origin.X; var oy = f.Origin.Y; var oz = f.Origin.Z;
            var vx = f.DirX.ToVector(); var vy = f.DirY.ToVector(); var vz = f.DirZ.ToVector();
            var p = Point.Create(
                ox + lx * vx.X + ly * vy.X + lz * vz.X,
                oy + lx * vx.Y + ly * vy.Y + lz * vz.Y,
                oz + lx * vx.Z + ly * vy.Z + lz * vz.Z
            );
            return p;
        }

        // Creates weld-round cylinders at the split-side corners based on cutter frame and split line.
        // - 'local' is the frame used to create the cutter
        // - 'yMid' is the split line (near side) in that same local frame
        // - 'weldRoundRadius' is in millimeters (converted to meters here)
        static List<Body> MakeCornerCylindersAtSplitSide(Body cutter, Frame local, double yMid, double weldRoundRadius, bool perpendicularCut)
        {
            var result = new List<Body>();
            if (perpendicularCut)
            {
                if (cutter == null || weldRoundRadius <= 0) return result;

                var worldToLocal = Matrix.CreateMapping(local).Inverse;
                var bb = cutter.GetBoundingBox(worldToLocal, true);
                double minX = bb.MinCorner.X, maxX = bb.MaxCorner.X;
                double minY = bb.MinCorner.Y, maxY = bb.MaxCorner.Y;
                double minZ = bb.MinCorner.Z, maxZ = bb.MaxCorner.Z;

                // Pick the Y side closest to the split (yMid).
                double yNear = (Math.Abs(minY - yMid) <= Math.Abs(maxY - yMid)) ? minY : maxY;
                double height = Math.Max(0, maxZ - minZ);

                // Build two cylinders (axis = local.DirZ), one at each X corner on the near Y side.
                double r = weldRoundRadius / 1000.0; // mm -> m
                foreach (double x in new[] { minX, maxX })
                {
                    var baseCenterWorld = ToWorld(local, x, yNear, minZ);
                    var baseFrame = Frame.Create(baseCenterWorld, local.DirX, local.DirY);

                    // Circle on the base plane, then extrude along +Z by 'height'.
                    // (Circle + Profile + Extrude per SpaceClaim API pattern.) 
                    var circle = Circle.Create(baseFrame, r);
                    var plane = Plane.Create(baseFrame);
                    var profile = new Profile(plane, new[] { CurveSegment.Create(circle) });
                    var cyl = Body.ExtrudeProfile(profile, height);
                    result.Add(cyl);
                }
            }
            else
            {
                if (cutter == null || weldRoundRadius <= 0) return result;

                // Angled cuts live on the local XZ plane.
                var worldToLocal = Matrix.CreateMapping(local).Inverse;
                var bb = cutter.GetBoundingBox(worldToLocal, true);
                double minX = bb.MinCorner.X, maxX = bb.MaxCorner.X;
                double minY = bb.MinCorner.Y, maxY = bb.MaxCorner.Y;
                double minZ = bb.MinCorner.Z, maxZ = bb.MaxCorner.Z;

                // Near split side in local Y
                double yNear = (Math.Abs(minY - yMid) <= Math.Abs(maxY - yMid)) ? minY : maxY;

                // Span along counterpart's X and a small overshoot amount
                double dx = Math.Max(0, maxX - minX);
                if (dx <= 1e-9) return result;

                double r = weldRoundRadius / 1000.0;     // mm -> m
                double over = r;                         // shift axis a little outside
                double length = dx + 200.0;         // extrude past both ends

                foreach (double x in new[] { minX, maxX })
                {
                    bool isMinX = Math.Abs(x - minX) < 1e-12;

                    // Use diagonal near-Y corner per side to avoid overlapping the same spot
                    double zCorner = isMinX ? minZ : maxZ;

                    // Corner on the near split side
                    var cornerW = ToWorld(local, x, yNear, zCorner);

                    // Cylinder axis direction: +X for minX corner, -X for maxX corner
                    var axisDir = isMinX ? local.DirX.ToVector() : (-local.DirX.ToVector());

                    // Move the AXIS slightly OUTSIDE along its own direction, then extrude back through the body
                    var baseOrigin = cornerW - axisDir * 10.0;

                    // Build a right-handed circle plane whose normal is the axis (±X):
                    //  - For +X normal: ex = Y, ey = Z  (Y × Z = +X)
                    //  - For -X normal: ex = Z, ey = Y  (Z × Y = -X)
                    Direction ex, ey;
                    if (isMinX) { ex = local.DirY; ey = local.DirZ; }   // normal +X
                    else { ex = local.DirZ; ey = local.DirY; }   // normal -X

                    var baseFrame = Frame.Create(baseOrigin, ex, ey);

                    var circle = Circle.Create(baseFrame, r);
                    var plane = Plane.Create(baseFrame);
                    var profile = new Profile(plane, new[] { CurveSegment.Create(circle) });

                    // Extrude along the axis direction; because we shifted the start by 'over',
                    // using (dx + 2*over) covers both sides while keeping the axis outside.
                    var cyl = Body.ExtrudeProfile(profile, length);
                    result.Add(cyl);

                    // Debug: visualize placement and coverage
                    //if (s_part != null)
                    //{
                    //    var dbg = DesignBody.Create(s_part, $"DebugWeldCyl_Angled_{Guid.NewGuid()}", cyl.Copy());
                    //    dbg.SetColor(null, Color.FromArgb(128, 255, 165, 0));
                    //}
                }
            }

            return result;
        }

        // Main entry: creates half-lap joints for each body pair, including optional weld rounds and logging.
        public static void CreateHalfLapJoints(
            Document doc,
            List<(Body A, Body B)> pairs,
            HalfLapParams jointParams,
            bool perpendicularCut,
            bool reverseDirection,
            double toleranceMM,
            bool middleTolerance,
            bool addWeldRound,
            double weldRoundRadius
        )
        {
            s_part = doc?.MainPart;
            if (s_part == null)
            {
                return;
            }
            if (pairs == null || pairs.Count == 0)
            {
                return;
            }

            for (int i = 0; i < pairs.Count; i++)
            {
                var (bodyA, bodyB) = pairs[i];
                if (bodyA == null || bodyB == null)
                {
                    continue;
                }

                try
                {
                    var frameA = CreateLocalFrameFromEdges(bodyA);
                    var frameB = CreateLocalFrameFromEdges(bodyB);

                    //AddAxes(frameA, $"Pair[{i}]/A_LocalAxes");
                    //AddAxes(frameB, $"Pair[{i}]/B_LocalAxes");

                    var overlapA = bodyA.Copy();
                    overlapA.Intersect(new[] { bodyB.Copy() });

                    var overlapB = bodyB.Copy();
                    overlapB.Intersect(new[] { bodyA.Copy() });

                    if (overlapA.Volume < 1e-12 || overlapB.Volume < 1e-12)
                    {
                        continue; // No overlap, skip
                    }

                    Body cutterBodyA = null;
                    Body cutterBodyB = null;

                    // We'll capture the frame used to build each cutter and the yMid corresponding to that frame.
                    Frame cutterAFrame;
                    double cutterAYMid = 0.0;

                    Frame cutterBFrame;
                    double cutterBYMid = 0.0;
                    // --- Cutter creation for both A and B, works for both perpendicular and angled cuts ---
                    // Always lengthen the cutter away from the split (yMid) by at least edgeOffset (100.0) on local Y

                    // For A
                    //Body cutterBodyA = null;
                    {
                        var worldToLocalA = Matrix.CreateMapping(frameA).Inverse;
                        var bbLocalA = overlapA.GetBoundingBox(worldToLocalA, true);
                        double yMidA = 0.5 * (bbLocalA.MinCorner.Y + bbLocalA.MaxCorner.Y);

                        if (perpendicularCut)
                        {
                            // --- DO NOT CHANGE THIS BLOCK ---
                            double y0A, y1A;
                            if (reverseDirection)
                            {
                                y0A = bbLocalA.MinCorner.Y - edgeOffset;
                                y1A = yMidA;
                            }
                            else
                            {
                                y0A = yMidA;
                                y1A = bbLocalA.MaxCorner.Y + edgeOffset;
                            }

                            var halfA = MakeRectPrism(
                                frameA,
                                bbLocalA.MinCorner.X, y0A, bbLocalA.MinCorner.Z,
                                bbLocalA.MaxCorner.X, y1A, bbLocalA.MaxCorner.Z,
                                toleranceMM, middleTolerance, perpendicularCut
                            );
                            if (halfA != null && halfA.Volume > 1e-12)
                                cutterBodyA = halfA;

                            if (halfA != null && halfA.Volume > 1e-12)
                            {
                                cutterBodyA = halfA;
                                cutterAFrame = frameA;
                                cutterAYMid = yMidA;
                            }
                        }
                        else
                        {
                            // --- LENGTHEN THE CUTTER ON Y AXIS BY edgeOffset ---
                            var worldToLocalB = Matrix.CreateMapping(frameB).Inverse;
                            var bbLocalB = overlapA.GetBoundingBox(worldToLocalB, true);
                            double yMidB = 0.5 * (bbLocalB.MinCorner.Y + bbLocalB.MaxCorner.Y);

                            double y0B = reverseDirection ? bbLocalB.MinCorner.Y - edgeOffset : yMidB;
                            double y1B = reverseDirection ? yMidB : bbLocalB.MaxCorner.Y + edgeOffset;

                            var overlapA2 = overlapA.Copy();
                            var boxB = Box.Create(
                                Point.Create(bbLocalB.MinCorner.X, y0B, bbLocalB.MinCorner.Z),
                                Point.Create(bbLocalB.MaxCorner.X, y1B, bbLocalB.MaxCorner.Z)
                            );
                            var cutterB = MakeRectPrism(frameB, boxB.MinCorner.X, boxB.MinCorner.Y, boxB.MinCorner.Z, boxB.MaxCorner.X, boxB.MaxCorner.Y, boxB.MaxCorner.Z, toleranceMM, middleTolerance, perpendicularCut);
                            cutterBodyA = cutterB;
                            cutterAFrame = frameB;     // built in B's frame
                            cutterAYMid = yMidB;
                        }
                    }

                    // For B
                    //Body cutterBodyB = null;
                    {
                        var worldToLocalB = Matrix.CreateMapping(frameB).Inverse;
                        var bbLocalB = overlapB.GetBoundingBox(worldToLocalB, true);
                        double yMidB = 0.5 * (bbLocalB.MinCorner.Y + bbLocalB.MaxCorner.Y);

                        if (perpendicularCut)
                        {
                            // --- DO NOT CHANGE THIS BLOCK ---
                            double y0B, y1B;
                            if (reverseDirection)
                            {
                                y0B = yMidB;
                                y1B = bbLocalB.MaxCorner.Y + edgeOffset;
                            }
                            else
                            {
                                y0B = bbLocalB.MinCorner.Y - edgeOffset;
                                y1B = yMidB;
                            }

                            var halfB = MakeRectPrism(
                                frameB,
                                bbLocalB.MinCorner.X, y0B, bbLocalB.MinCorner.Z,
                                bbLocalB.MaxCorner.X, y1B, bbLocalB.MaxCorner.Z, toleranceMM, middleTolerance, perpendicularCut
                            );
                            if (halfB != null && halfB.Volume > 1e-12)
                                cutterBodyB = halfB;

                            if (halfB != null && halfB.Volume > 1e-12)
                            {
                                cutterBodyB = halfB;
                                cutterBFrame = frameB;
                                cutterBYMid = yMidB;
                            }
                        }
                        else
                        {
                            // --- LENGTHEN THE CUTTER ON Y AXIS BY edgeOffset ---
                            var worldToLocalA = Matrix.CreateMapping(frameA).Inverse;
                            var bbLocalA = overlapB.GetBoundingBox(worldToLocalA, true);
                            double yMidA = 0.5 * (bbLocalA.MinCorner.Y + bbLocalA.MaxCorner.Y);

                            double y0A = reverseDirection ? yMidA : bbLocalA.MinCorner.Y - edgeOffset;
                            double y1A = reverseDirection ? bbLocalA.MaxCorner.Y + edgeOffset : yMidA;

                            var overlapB2 = overlapB.Copy();
                            var boxA = Box.Create(
                                Point.Create(bbLocalA.MinCorner.X, y0A, bbLocalA.MinCorner.Z),
                                Point.Create(bbLocalA.MaxCorner.X, y1A, bbLocalA.MaxCorner.Z)
                            );
                            var cutterA = MakeRectPrism(frameA, boxA.MinCorner.X, boxA.MinCorner.Y, boxA.MinCorner.Z, boxA.MaxCorner.X, boxA.MaxCorner.Y, boxA.MaxCorner.Z, toleranceMM, middleTolerance, perpendicularCut);
                            cutterBodyB = cutterA;
                            cutterBFrame = frameA;     // built in A's frame
                            cutterBYMid = yMidA;
                        }
                    }

                    //Debug overlap body(optional, only one for clarity)
                    //var overlapDb = DesignBody.Create(s_part, $"Pair[{i}]_Overlap", overlapA.Copy());
                    //overlapDb.SetColor(null, Color.FromArgb(128, 0, 200, 255));
                    //Logger.Log($"CreateHalfLapJoints: Created debug overlap body for pair {i}");

                    //if (cutterBodyA != null)
                    //{
                    //    var debugCutterA = DesignBody.Create(s_part, $"Pair[{i}]_DebugCutterA", cutterBodyA.Copy());
                    //    debugCutterA.SetColor(null, Color.FromArgb(128, 255, 0, 0)); // Semi-transparent red
                    //    Logger.Log($"Debug: Created debug cutter body A for pair {i} (offset {edgeOffset})");
                    //}
                    //if (cutterBodyB != null)
                    //{
                    //    var debugCutterB = DesignBody.Create(s_part, $"Pair[{i}]_DebugCutterB", cutterBodyB.Copy());
                    //    debugCutterB.SetColor(null, Color.FromArgb(128, 0, 255, 0)); // Semi-transparent green
                    //    Logger.Log($"Debug: Created debug cutter body B for pair {i} (offset {edgeOffset})");
                    //}

                    // Subtract lower half from bodyA
                    if (cutterBodyA != null)
                    {
                        var tempCutterA = DesignBody.Create(s_part, $"Pair[{i}]_TempCutterA_{Guid.NewGuid()}", cutterBodyA.Copy());
                        try
                        {
                            bodyA.Subtract(new[] { tempCutterA.Shape });
                        }
                        catch (Exception ex)
                        {
                            Logger.Log($"CreateHalfLapJoints: Subtract exception (A): {ex} (tempCutterA.IsDeleted={tempCutterA.IsDeleted})");
                        }
                    }
                    else
                    {
                        Logger.Log($"CreateHalfLapJoints: Lower half prism for bodyA in pair {i} is null or has zero volume, skipping subtraction.");
                    }

                    // Subtract upper half from bodyB
                    if (cutterBodyB != null)
                    {
                        var tempCutterB = DesignBody.Create(s_part, $"Pair[{i}]_TempCutterB_{Guid.NewGuid()}", cutterBodyB.Copy());
                        try
                        {
                            bodyB.Subtract(new[] { tempCutterB.Shape });
                        }
                        catch (Exception ex)
                        {
                            Logger.Log($"CreateHalfLapJoints: Subtract exception (B): {ex} (tempCutterB.IsDeleted={tempCutterB.IsDeleted})");
                        }
                    }
                    else
                    {
                        Logger.Log($"CreateHalfLapJoints: Upper half prism for bodyB in pair {i} is null or has zero volume, skipping subtraction.");
                    }

                    if (addWeldRound && weldRoundRadius > 0)
                    {
                        // For A
                        if (cutterBodyA != null && cutterAFrame != null)
                        {
                            var cylsA = MakeCornerCylindersAtSplitSide(cutterBodyA, cutterAFrame, cutterAYMid, weldRoundRadius, perpendicularCut);
                            foreach (var cyl in cylsA)
                            {
                                var tmp = DesignBody.Create(s_part, $"Pair[{i}]_TempWeldCylA_{Guid.NewGuid()}", cyl);
                                try
                                {
                                    bodyA.Subtract(new[] { tmp.Shape });
                                }
                                catch (Exception ex)
                                {
                                    Logger.Log($"CreateHalfLapJoints: Weld round subtract exception (A): {ex}");
                                }
                            }
                        }

                        // For B
                        if (cutterBodyB != null && cutterBFrame != null)
                        {
                            var cylsB = MakeCornerCylindersAtSplitSide(cutterBodyB, cutterBFrame, cutterBYMid, weldRoundRadius, perpendicularCut);
                            foreach (var cyl in cylsB)
                            {
                                var tmp = DesignBody.Create(s_part, $"Pair[{i}]_TempWeldCylB_{Guid.NewGuid()}", cyl);
                                try
                                {
                                    bodyB.Subtract(new[] { tmp.Shape });
                                }
                                catch (Exception ex)
                                {
                                    Logger.Log($"CreateHalfLapJoints: Weld round subtract exception (B): {ex}");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"CreateHalfLapJoints: Exception in pair {i}: {ex}");
                    continue;
                }
            }

            s_part = null;
        }
    }
}
