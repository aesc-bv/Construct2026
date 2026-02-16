using AESCConstruct2026.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System.Linq;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;

namespace AESCConstruct2026.FrameGenerator.Modules.Joints
{
    public sealed class NoneJoint : JointBase
    {
        public override string Name => "None";

        public override void Execute(
            Component componentA,
            Component componentB,
            double spacing,
            Body bodyA,
            Body bodyB
        )
        {
            WriteBlock.ExecuteTask("NoneJoint", () =>
            {
                //Logger.Log("NoneJoint.Execute() started");

                // 1) Fetch the two raw construction segments…
                var rawA = componentA.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
                var rawB = componentB.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
                if (rawA == null || rawB == null)
                {
                    //Logger.Log("NoneJoint: ERROR – missing construction curve(s).");
                    return;
                }

                // 2) Lift into world and compute worldUp = cross(dirA, dirB)
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

                // Ensure consistent hemisphere
                if (Vector.Dot(worldUp, Vector.Create(0, 1, 0)) < 0)
                    worldUp = -worldUp;

                // 3) Compute offset-edges and their local perps
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

                // 4) Single intersection point
                Point? pIn = JointCurveHelper.IntersectLines(componentA, innerA, componentB, innerB);
                if (pIn == null)
                {
                    //Logger.Log("NoneJoint: ERROR – cannot find intersection.");
                    return;
                }

                // 5) Determine which end is connected (for pickDirection)
                const double tol = 1e-6;
                var wAstart = componentA.Placement * rawA.StartPoint;
                var wAend = componentA.Placement * rawA.EndPoint;
                var wBstart = componentB.Placement * rawB.StartPoint;
                var wBend = componentB.Placement * rawB.EndPoint;

                bool aStartConn = (wAstart - wBstart).Magnitude < tol || (wAstart - wBend).Magnitude < tol;
                bool aEndConn = !aStartConn;
                bool bStartConn = (wBstart - wAstart).Magnitude < tol || (wBstart - wAend).Magnitude < tol;
                bool bEndConn = !bStartConn;

                // 6) Compute a tiny world-space offset along each component’s perp
                const double eps = 0.01; // 1cm in world units
                var worldPerpA = (componentA.Placement * perpA.Direction).ToVector();
                var worldPerpB = (componentB.Placement * perpB.Direction).ToVector();

                Point worldAin = pIn.Value;
                Point worldAout = worldAin + worldPerpA * eps;
                Point worldBin = pIn.Value;
                Point worldBout = worldBin + worldPerpB * eps;

                // 7) Subtract cutter on each
                SubtractLocalCutter(componentA, worldAin, worldAout, worldUp, aStartConn, aEndConn, spacing);
                SubtractLocalCutter(componentB, worldBin, worldBout, worldUp, bStartConn, bEndConn, spacing);

                //Logger.Log("NoneJoint: finished.");
            });
        }

        /// <summary>
        /// Exactly the same SubtractLocalCutter logic as in MiterJoint.
        /// </summary>
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
            // 0) Map worldUp into local
            var inv = comp.Placement.Inverse;
            Direction upDirLocal = inv * worldUp.Direction;
            Vector upLocal = upDirLocal.ToVector();

            // 1) Map both world points into local
            Point inLoc = inv * worldPIn;
            Point outLoc = inv * worldPOut;

            //Logger.Log($"Building cutter in LOCAL for '{comp.Name}'.");

            // 2) Build local‐space cutter frame + loop
            var (planeLocal, loopLocal) = JointModule.BuildDebugCutterFrameAndLoop(
                inLoc, outLoc, upLocal, 500.0
            );
            if (planeLocal == null)
            {
                //Logger.Log($"  ERROR: BuildDebugCutterFrameAndLoop returned null for '{comp.Name}'.");
                return;
            }

            // 3) Pick extrusion distances (local)
            var (forwardDist, backwardDist) = JointCurveHelper.PickDirection(
                planeLocal,
                comp,
                startConnected,
                endConnected,
                longLen: 200.0,
                shortLen: spacing / 2.0
            );

            // 4) Create and subtract the bidirectional cutter
            var cutterLocal = JointModule.CreateBidirectionalExtrudedBody(
                planeLocal,
                loopLocal,
                backwardDist,
                forwardDist
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
