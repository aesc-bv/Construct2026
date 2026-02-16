using AESCConstruct2026.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System.Linq;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;

namespace AESCConstruct2026.FrameGenerator.Modules.Joints
{
    public sealed class StraightJoint2 : JointBase
    {
        public override string Name => "Straight2";

        public override void Execute(
            Component componentA,
            Component componentB,
            double spacing,
            Body bodyA,
            Body bodyB
        )
        {
            WriteBlock.ExecuteTask("StraightJoint2", () =>
            {
                //Logger.Log("StraightJoint2.Execute() started");

                // 1) Fetch local construction curves
                var rawA = componentA.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
                var rawB = componentB.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
                if (rawA == null || rawB == null)
                {
                    //Logger.Log("StraightJoint2: ERROR – missing construction curve(s).");
                    return;
                }

                // 2) Lift into world & compute a stable world-up
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

                // 3) Compute offset‐edges + local perps
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
                Point? pY = JointCurveHelper.IntersectLines(componentA, innerA, componentB, outerB);
                Point? pZ = JointCurveHelper.IntersectLines(componentA, outerA, componentB, outerB);
                Point? pIn = JointCurveHelper.IntersectLines(componentA, innerA, componentB, innerB);
                if (pY == null || pZ == null || pIn == null)
                {
                    //Logger.Log("StraightJoint2: ERROR – cannot find required intersections.");
                    return;
                }

                // 5) Determine connection ends by shared local-points
                var jointPt = JointCurveHelper.FindSharedPoint(rawA, rawB);
                bool aStartConn = (rawA.StartPoint - jointPt).Magnitude < 1e-6;
                bool aEndConn = (rawA.EndPoint - jointPt).Magnitude < 1e-6;
                bool bStartConn = (rawB.StartPoint - jointPt).Magnitude < 1e-6;
                bool bEndConn = (rawB.EndPoint - jointPt).Magnitude < 1e-6;

                // 6) Cutter on A: base=pY, offset-pZ, perp=perpA
                SubtractLocalCutter(
                    componentA,
                    worldBase: pY.Value,
                    worldOffset: pZ.Value,
                    worldUp: worldUp,
                    perpLocal: perpA,
                    startConnected: aStartConn,
                    endConnected: aEndConn,
                    longLen: 200.0,
                    shortLen: 0.0
                );

                // 7) Cutter on B: base=pIn, offset = along A’s world-dir by ε
                Vector worldDirA = dirA;
                const double eps = 0.01;
                Point worldOffsetB = pIn.Value + worldDirA * eps;

                SubtractLocalCutter(
                    componentB,
                    worldBase: pIn.Value,
                    worldOffset: worldOffsetB,
                    worldUp: worldUp,
                    perpLocal: perpB,
                    startConnected: bStartConn,
                    endConnected: bEndConn,
                    longLen: 200.0,
                    shortLen: spacing
                );

                //Logger.Log("StraightJoint2: finished.");
            });
        }

        private void SubtractLocalCutter(
            Component comp,
            Point worldBase,
            Point worldOffset,
            Vector worldUp,
            Vector perpLocal,
            bool startConnected,
            bool endConnected,
            double longLen,
            double shortLen
        )
        {
            // 1) world→local reprojection
            var inv = comp.Placement.Inverse;
            Direction upDirLocal = inv * worldUp.Direction;
            Vector upLocal = upDirLocal.ToVector();
            Point baseLoc = inv * worldBase;
            Point offLoc = inv * worldOffset;

            //Logger.Log($"Building cutter in LOCAL for '{comp.Name}'.");

            // 2) build cutter‐plane + loop in local
            var (planeLocal, loopLocal) = JointModule.BuildDebugCutterFrameAndLoop(
                baseLoc, offLoc, upLocal, 500.0
            );
            if (planeLocal == null)
            {
                //Logger.Log($"  ERROR: BuildDebugCutterFrameAndLoop returned null for '{comp.Name}'.");
                return;
            }

            // 3) pick extrusion lengths in local
            var (fwd, back) = JointCurveHelper.PickDirection(
                planeLocal, comp, startConnected, endConnected,
                longLen: longLen, shortLen: shortLen
            );

            // 4) extrude & subtract
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
