using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;

namespace AESCConstruct25.RibCutout.Modules
{
    internal static class RibCutOutModule
    {
        public static void ProcessPairs(
    Document doc,
    List<(Body A, Body B)> pairs,
    bool perpendicularCut,
    double toleranceMM,
    bool applyMiddleTolerance
)
        {
            Logger.Log("RibCutOutModule: ProcessPairs() start.");
            var part = doc.MainPart;

            double tol = toleranceMM / 2000.0;

            Point GetCenter(Box box)
            {
                var vec = 0.5 * (box.MinCorner.Vector + box.MaxCorner.Vector);
                return Point.Create(vec.X, vec.Y, vec.Z);
            }

            foreach (var (origA, origB) in pairs)
            {
                try
                {
                    if (!perpendicularCut)
                    {
                        // Simple mutual subtraction
                        var toolB = DesignBody.Create(part, $"toolB_{Guid.NewGuid():N}", origB.Copy());
                        var toolA = DesignBody.Create(part, $"toolA_{Guid.NewGuid():N}", origA.Copy());

                        origA.Subtract(new[] { toolB.Shape });
                        origB.Subtract(new[] { toolA.Shape });

                        toolA.Delete();
                        toolB.Delete();

                        Logger.Log("Simple A-B and B-A subtraction completed.");
                        continue;
                    }

                    // Perpendicular logic below
                    var overlapBase = origA.Copy();
                    overlapBase.Intersect(new[] { origB.Copy() });

                    if (overlapBase.Volume <= 0)
                    {
                        Logger.Log("No overlap; skipping.");
                        continue;
                    }

                    var centerA = GetCenter(origA.GetBoundingBox(Matrix.Identity, true));
                    var centerB = GetCenter(origB.GetBoundingBox(Matrix.Identity, true));
                    var centerOverlap = GetCenter(overlapBase.GetBoundingBox(Matrix.Identity, true));

                    void SubtractHalfFrom(Body target, Point targetCenter, string debugPrefix)
                    {
                        var overlap = overlapBase.Copy();

                        var direction = (centerOverlap - targetCenter).Direction;
                        var plane = Plane.Create(Frame.Create(centerOverlap, direction));

                        overlap.Split(plane, null);
                        var pieces = overlap.SeparatePieces();

                        if (pieces.Count != 2)
                        {
                            Logger.Log("Split failed or incorrect number of pieces.");
                            return;
                        }

                        Body furthest = null;
                        double maxDist = double.MinValue;

                        foreach (var piece in pieces)
                        {
                            var mid = GetCenter(piece.GetBoundingBox(Matrix.Identity, true));
                            double dist = (mid - targetCenter).Magnitude;

                            if (dist > maxDist)
                            {
                                maxDist = dist;
                                furthest = piece;
                            }
                        }

                        if (furthest != null)
                        {
                            var bb = furthest.GetBoundingBox(Matrix.Identity, true);
                            var min = bb.MinCorner;
                            var max = bb.MaxCorner;

                            double inflateZ = applyMiddleTolerance ? tol : 0.0;

                            Point p0 = Point.Create(min.X - tol, min.Y - tol, min.Z - inflateZ);
                            Point p1 = Point.Create(max.X + tol, max.Y + tol, max.Z + inflateZ);

                            var basePts = new[]
                            {
                        Point.Create(p0.X, p0.Y, 0),
                        Point.Create(p1.X, p0.Y, 0),
                        Point.Create(p1.X, p1.Y, 0),
                        Point.Create(p0.X, p1.Y, 0)
                    };
                            var prof = new Profile(Plane.PlaneXY, new[]
                            {
                        CurveSegment.Create(basePts[0], basePts[1]),
                        CurveSegment.Create(basePts[1], basePts[2]),
                        CurveSegment.Create(basePts[2], basePts[3]),
                        CurveSegment.Create(basePts[3], basePts[0]),
                    });

                            double height = p1.Z - p0.Z;
                            var inflatedCutter = Body.ExtrudeProfile(prof, height);
                            inflatedCutter.Transform(Matrix.CreateTranslation(Vector.Create(0, 0, p0.Z)));

                            var debugCut = DesignBody.Create(part, $"{debugPrefix}_CutPreview", inflatedCutter.Copy());
                            debugCut.SetColor(null, System.Drawing.Color.Red);

                            var line = CurveSegment.Create(targetCenter, centerOverlap);
                            var dc = DesignCurve.Create(part, line);
                            dc.Name = $"{debugPrefix}_CutDirection";
                            dc.SetColor(null, System.Drawing.Color.Cyan);

                            target.Subtract(new[] { debugCut.Shape });
                            Logger.Log("  → Subtracted inflated cutter from target.");
                        }
                        else
                        {
                            Logger.Log("Could not determine furthest half.");
                        }
                    }

                    SubtractHalfFrom(origA, centerA, "DEBUG_A");
                    SubtractHalfFrom(origB, centerB, "DEBUG_B");
                }
                catch (Exception ex)
                {
                    Logger.Log($"RibCutOutModule: Exception: {ex}");
                }
            }

            Logger.Log("RibCutOutModule: ProcessPairs() end.");
        }
    }
}
