using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using AESCConstruct25.FrameGenerator.Utilities;
using System.Linq;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;
using System.Globalization;
using System;
using System.Drawing;

namespace AESCConstruct25.FrameGenerator.Modules.Joints
{
    public sealed class MiterJoint : JointBase
    {
        public override string Name => "Miter";

        public override void Execute(
            Component componentA,
            Component componentB,
            double spacing,
            Body bodyA,
            Body bodyB
        )
        {
            WriteBlock.ExecuteTask("MiterJoint", () =>
            {
                //Logger.Log("MiterJoint.Execute() started");

                // 1) grab the two raw construction segments
                var rawA = componentA.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
                var rawB = componentB.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
                if (rawA == null || rawB == null)
                {
                    //Logger.Log("MiterJoint: ERROR – missing construction curve(s).");
                    return;
                }

                // 2) lift both into WORLD and get their directions
                Point wA0 = componentA.Placement * rawA.StartPoint;
                Point wA1 = componentA.Placement * rawA.EndPoint;
                Point wB0 = componentB.Placement * rawB.StartPoint;
                Point wB1 = componentB.Placement * rawB.EndPoint;

                Vector dirA = (wA1 - wA0).Direction.ToVector();
                Vector dirB = (wB1 - wB0).Direction.ToVector();

                // 3) world‐up = cross(dirA, dirB)
                Vector worldUp = Vector.Cross(dirA, dirB);
                if (worldUp.Magnitude < 1e-6)
                {
                    // fallback if nearly colinear
                    worldUp = Vector.Create(0, 1, 0);
                }
                else
                {
                    worldUp = worldUp.Direction.ToVector();
                }
                // hemisphere‐fix against global Y
                if (Vector.Dot(worldUp, Vector.Create(0, 1, 0)) < 0)
                    worldUp = -worldUp;

                // 4) offset edges & intersections (world‐space)
                double wA = JointCurveHelper.GetProfileWidth(componentA);
                double wB = JointCurveHelper.GetProfileWidth(componentB);
                double oA = JointCurveHelper.GetOffset(componentA, "offsetX");
                double oB = JointCurveHelper.GetOffset(componentB, "offsetX");

                var (innerA, outerA, _) = JointCurveHelper.GetOffsetEdges(
                    componentA, rawA, componentB, wA, oA
                );
                var (innerB, outerB, _) = JointCurveHelper.GetOffsetEdges(
                    componentB, rawB, componentA, wB, oB
                );

                Point? pIn = JointCurveHelper.IntersectLines(componentA, innerA, componentB, innerB);
                Point? pOut = JointCurveHelper.IntersectLines(componentA, outerA, componentB, outerB);
                if (pIn == null || pOut == null)
                {
                    //Logger.Log("MiterJoint: ERROR – cannot find intersections.");
                    return;
                }

                // 5) decide from/to for each
                const double tol = 1e-6;
                bool aEnd = ((componentA.Placement * rawA.EndPoint) - (componentB.Placement * rawB.StartPoint)).Magnitude < tol
                         || ((componentA.Placement * rawA.EndPoint) - (componentB.Placement * rawB.EndPoint)).Magnitude < tol;
                bool aStart = !aEnd;
                bool bEnd = ((componentB.Placement * rawB.EndPoint) - (componentA.Placement * rawA.StartPoint)).Magnitude < tol
                         || ((componentB.Placement * rawB.EndPoint) - (componentA.Placement * rawA.EndPoint)).Magnitude < tol;
                bool bStart = !bEnd;

                Point aFrom = aEnd ? pIn.Value : pOut.Value;
                Point aTo = aEnd ? pOut.Value : pIn.Value;
                Point bFrom = bEnd ? pOut.Value : pIn.Value;
                Point bTo = bEnd ? pIn.Value : pOut.Value;

                // 6) build & subtract on each
                SubtractLocalCutter(componentA, aFrom, aTo, worldUp, aStart, aEnd, spacing);
                SubtractLocalCutter(componentB, bFrom, bTo, worldUp, bStart, bEnd, spacing);

                //Logger.Log("MiterJoint: finished.");
            });
        }

        private void SubtractLocalCutter(
            Component comp,
            Point worldPIn,
            Point worldPOut,
            Vector worldUp,
            bool startConnected,
            bool endConnected,
            double spacing
        )
        {
            // map worldUp into component‐local
            var inv = comp.Placement.Inverse;
            Direction upDirLocal = (inv * worldUp.Direction);
            Vector upLocal = upDirLocal.ToVector();

            // map the two intersection points into local
            Point inLoc = inv * worldPIn;
            Point outLoc = inv * worldPOut;

            //Logger.Log($"Building cutter in LOCAL for '{comp.Name}'.");

            // build local cutter‐frame & loop
            var (planeLocal, loopLocal) = JointModule.BuildDebugCutterFrameAndLoop(
                inLoc, outLoc, upLocal, 500.0
            );
            if (planeLocal == null)
            {
                //Logger.Log($"  ERROR: BuildDebugCutterFrameAndLoop returned null for '{comp.Name}'.");
                return;
            }

            // pick forward/back in local
            var (fwd, back) = JointCurveHelper.PickDirection(
                planeLocal, comp, startConnected, endConnected,
                longLen: 200.0, shortLen: spacing / 2.0
            );

            // extrude
            var cutter = JointModule.CreateBidirectionalExtrudedBody(
                planeLocal, loopLocal, back, fwd
            );
            if (cutter == null)
            {
                //Logger.Log($"  ERROR: CreateBidirectionalExtrudedBody returned null for '{comp.Name}'.");
                return;
            }

            // subtract
            JointModule.SubtractCutter(comp, cutter);
            //Logger.Log($"  Subtraction complete for '{comp.Name}'.");
        }
    }

}
