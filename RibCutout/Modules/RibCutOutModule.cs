using System;
using System.Collections.Generic;
using AESCConstruct25.FrameGenerator.Utilities;   // for Logger
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;               // Matrix, Vector, Point, Plane, Profile, CurveSegment, Frame
using SpaceClaim.Api.V242.Modeler;                // Body, DesignBody

namespace AESCConstruct25.RibCutout.Modules
{
    internal static class RibCutOutModule
    {
        public static void ProcessPairs(
            Document doc,
            List<(Body A, Body B)> pairs,
            bool perpendicularCut
        )
        {
            Logger.Log("RibCutOutModule: ProcessPairs() start.");
            var part = doc.MainPart;

            foreach (var (origA, origB) in pairs)
            {
                try
                {
                    if (!perpendicularCut)
                    {
                        // === simple subtraction ===
                        var rawB = origB.Copy();
                        var dbB = DesignBody.Create(part, $"toolB_{Guid.NewGuid():N}", rawB);
                        var rawA = origA.Copy();
                        var dbA = DesignBody.Create(part, $"toolA_{Guid.NewGuid():N}", rawA);

                        origA.Subtract(new[] { dbB.Shape });
                        origB.Subtract(new[] { dbA.Shape });

                        dbB.Delete();
                        dbA.Delete();
                    }
                    else
                    {
                        Logger.Log("Perpendicular mode: generating edge-anchored half-Y bboxes.");

                        // Compute overlap
                        var overlapTool = origB.Copy();
                        var overlap = origA.Copy();
                        overlap.Intersect(new[] { overlapTool });
                        overlapTool.Dispose();

                        if (overlap.Volume <= 0)
                        {
                            Logger.Log("No overlap; skipping perpendicular cut.");
                            continue;
                        }

                        // Compute bounds
                        var bb = overlap.GetBoundingBox(Matrix.Identity, true);
                        var min = bb.MinCorner;
                        var max = bb.MaxCorner;
                        double xMin = min.X, xMax = max.X;
                        double yMin = min.Y, yMax = max.Y;
                        double zMin = min.Z, zMax = max.Z;
                        double height = zMax - zMin;
                        double yMid = (yMax - yMin) / 2.0;
                        double centerZ = (zMin + zMax) / 2.0;
                        double centerX = (xMin + xMax) / 2.0;

                        // Decide which half for A and B (keep same logic as before)
                        bool useLowerHalf = true; // or derive from geometry if needed
                        double y0A = useLowerHalf ? yMin : yMin + yMid;
                        double y1A = useLowerHalf ? yMin + yMid : yMax;
                        double y0B = !useLowerHalf ? yMin : yMin + yMid;
                        double y1B = !useLowerHalf ? yMin + yMid : yMax;

                        // Always align cutters to global Y-axis
                        var globalFrame = Frame.Create(Point.Origin,            // origin
                                                       Direction.DirZ,       // Z-axis 
                                                       Direction.DirY);      // Y-axis

                        void BuildAndSubtract(double y0, double y1, Body target)
                        {
                            // 1) Create 2D profile in XY
                            var pts = new[]
                            {
                                Point.Create(xMin, y0, 0),
                                Point.Create(xMax, y0, 0),
                                Point.Create(xMax, y1, 0),
                                Point.Create(xMin, y1, 0)
                            };
                            var prof = new Profile(Plane.PlaneXY, new[]
                            {
                                CurveSegment.Create(pts[0], pts[1]),
                                CurveSegment.Create(pts[1], pts[2]),
                                CurveSegment.Create(pts[2], pts[3]),
                                CurveSegment.Create(pts[3], pts[0]),
                            });

                            // 2) Extrude up to match overlap height
                            var cutter = Body.ExtrudeProfile(prof, height);
                            cutter.Transform(Matrix.CreateTranslation(Vector.Create(0, 0, zMin)));

                            // 3) Rotate/translate cutter into the same global orientation
                            cutter.Transform(Matrix.CreateMapping(globalFrame));

                            // 4) Create a temporary DesignBody, subtract, and clean up
                            var db = DesignBody.Create(part, $"bbox_{Guid.NewGuid():N}", cutter);
                            target.Subtract(new[] { db.Shape });
                            db.Delete();

                            Logger.Log($"  Subtracted half-Y bbox from body.");
                        }

                        // Apply to both A and B
                        BuildAndSubtract(y0A, y1A, origA);
                        BuildAndSubtract(y0B, y1B, origB);
                    }
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
