using System;
using System.Globalization;
using System.Drawing;
using System.Linq;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using AESCConstruct25.FrameGenerator.Utilities;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;

namespace AESCConstruct25.FrameGenerator.Modules.Joints
{
    public sealed class StraightJoint : JointBase
    {
        public override string Name => "Straight";

        public override void Execute(
            Component componentA,
            Component componentB,
            double spacing,
            Body bodyA,
            Body bodyB
        )
        {
            WriteBlock.ExecuteTask("StraightJoint", () =>
            {
                //Logger.Log("StraightJoint.Execute() started");

                // 1) Fetch construction curves
                var rawA = componentA.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
                var rawB = componentB.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
                if (rawA == null || rawB == null)
                {
                    //Logger.Log("StraightJoint: ERROR – missing construction curve(s).");
                    return;
                }

                // 2) Lift into world & compute worldUp
                Point wA0 = componentA.Placement * rawA.StartPoint;
                Point wA1 = componentA.Placement * rawA.EndPoint;
                Point wB0 = componentB.Placement * rawB.StartPoint;
                Point wB1 = componentB.Placement * rawB.EndPoint;
                Vector dirA = (wA1 - wA0).Direction.ToVector();
                Vector dirB = (wB1 - wB0).Direction.ToVector();

                Vector worldUp = Vector.Cross(dirA, dirB);
                if (worldUp.Magnitude < 1e-6)
                    worldUp = Vector.Create(0, 1, 0);
                else
                    worldUp = worldUp.Direction.ToVector();
                if (Vector.Dot(worldUp, Vector.Create(0, 1, 0)) < 0)
                    worldUp = -worldUp;

                // 3) Compute offset‐edges + perp’s
                double wA = JointCurveHelper.GetProfileWidth(componentA);
                double wB = JointCurveHelper.GetProfileWidth(componentB);
                double oA = JointCurveHelper.GetOffset(componentA, "offsetX");
                double oB = JointCurveHelper.GetOffset(componentB, "offsetX");

                var (innerA, outerA, perpA) = JointCurveHelper.GetOffsetEdges(
                    componentA, rawA, componentB, wA, oA
                );
                var (innerB, outerB, perpB) = JointCurveHelper.GetOffsetEdges(
                    componentB, rawB, componentA, wB, oB
                );

                // 4) Intersection points:
                Point? pIn = JointCurveHelper.IntersectLines(componentA, innerA, componentB, innerB);
                Point? pOut = JointCurveHelper.IntersectLines(componentA, outerA, componentB, outerB);
                Point? pX = JointCurveHelper.IntersectLines(componentA, innerA, componentB, outerB);
                if (pIn == null || pOut == null || pX == null)
                {
                    //Logger.Log("StraightJoint: ERROR – cannot find intersections.");
                    return;
                }

                // 5) Determine connections
                const double tol = 1e-6;
                Point wAs = componentA.Placement * rawA.StartPoint;
                Point wAe = componentA.Placement * rawA.EndPoint;
                Point wBs = componentB.Placement * rawB.StartPoint;
                Point wBe = componentB.Placement * rawB.EndPoint;

                bool aStartConn = (wAs - wBs).Magnitude < tol || (wAs - wBe).Magnitude < tol;
                bool aEndConn = !aStartConn;
                bool bStartConn = (wBs - wAs).Magnitude < tol || (wBs - wAe).Magnitude < tol;
                bool bEndConn = !bStartConn;

                // 6) Cutter on A: base = pX, perp=perpA
                SubtractLocalCutter(
                    componentA,
                    pX.Value,
                    worldUp,
                    perpA,
                    aStartConn,
                    aEndConn,
                    longLen: 200.0,
                    shortLen: 0.0
                );

                // 7) Cutter on B: base = pIn, perp = worldDirA reprojected into B-local
                Vector worldDirA = dirA;
                Point originW = Point.Create(0, 0, 0);
                Point tipW = originW + worldDirA;
                var invB = componentB.Placement.Inverse;
                Vector perpB2 = (invB * tipW - invB * originW).Direction.ToVector();

                SubtractLocalCutter(
                    componentB,
                    pIn.Value,
                    worldUp,
                    perpB2,
                    bStartConn,
                    bEndConn,
                    longLen: 200.0,
                    shortLen: spacing
                );

                //Logger.Log("StraightJoint: finished.");
            });
        }

        private void SubtractLocalCutter(
            Component comp,
            Point worldPBase,
            Vector worldUp,
            Vector perpLocal,
            bool startConnected,
            bool endConnected,
            double longLen,
            double shortLen
        )
        {
            // map worldUp into local
            var inv = comp.Placement.Inverse;
            Direction upDirLocal = inv * worldUp.Direction;
            Vector upLocal = upDirLocal.ToVector();

            // map base point into local
            Point baseLoc = inv * worldPBase;

            // create a tiny second point along perpLocal
            const double eps = 0.01;
            Point offLoc = Point.Create(
                baseLoc.X + perpLocal.X * eps,
                baseLoc.Y + perpLocal.Y * eps,
                baseLoc.Z + perpLocal.Z * eps
            );

            //Logger.Log($"Building cutter in LOCAL for '{comp.Name}'.");

            // Build plane & loop
            var (planeLocal, loopLocal) = JointModule.BuildDebugCutterFrameAndLoop(
                baseLoc, offLoc, upLocal, 500.0
            );
            if (planeLocal == null)
            {
                //Logger.Log($"  ERROR: BuildDebugCutterFrameAndLoop returned null for '{comp.Name}'.");
                return;
            }

            // Pick local extrusion lengths
            var (fwd, back) = JointCurveHelper.PickDirection(
                planeLocal, comp, startConnected, endConnected,
                longLen: longLen, shortLen: shortLen
            );

            // Create and subtract cutter
            var cutterLocal = JointModule.CreateBidirectionalExtrudedBody(
                planeLocal, loopLocal, back, fwd
            );
            if (cutterLocal == null)
            {
                //Logger.Log($"  ERROR: CreateBidirectionalExtrudedBody returned null for '{comp.Name}'.");
                return;
            }

            JointModule.SubtractCutter(comp, cutterLocal);
            //Logger.Log($"  Subtraction complete for '{comp.Name}'.");
        }
    }
}
