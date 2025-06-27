using SpaceClaim.Api.V251;
using SpaceClaim.Api.V251.Geometry;
using SpaceClaim.Api.V251.Modeler;
using AESCConstruct25.Utilities;
using AESCConstruct25.Commands;
using System.Linq;
using Point = SpaceClaim.Api.V251.Geometry.Point;
using Vector = SpaceClaim.Api.V251.Geometry.Vector;

namespace AESCConstruct25.Modules.Joints
{
    public sealed class TJoint : JointBase
    {
        public override string Name => "T";

        public override void Execute(
            Component componentA,
            Component componentB,
            double spacing,
            Body bodyA,
            Body bodyB
        )
        {
            WriteBlock.ExecuteTask("TJoint", () =>
            {
                const double tol = 1e-6;
                //Logger.Log($"TJoint.Execute: {componentA.Name} ⊥ {componentB.Name}");

                // 1) Fetch A & B construction curves
                var rawA = componentA.Template.Curves.OfType<DesignCurve>()
                              .FirstOrDefault()?.Shape as CurveSegment;
                var rawB = componentB.Template.Curves.OfType<DesignCurve>()
                              .FirstOrDefault()?.Shape as CurveSegment;
                if (rawA == null || rawB == null)
                {
                    //Logger.Log("TJoint: ERROR – missing construction curves.");
                    return;
                }

                // 2) World‐space endpoints
                Point wA0 = componentA.Placement * rawA.StartPoint;
                Point wA1 = componentA.Placement * rawA.EndPoint;
                Point wB0 = componentB.Placement * rawB.StartPoint;
                Point wB1 = componentB.Placement * rawB.EndPoint;

                // 3) Find the T‐point on B (or fallback to closest)
                bool startOn = ExecuteJointCommand.IsPointOnSegment(wB0, wA0, wA1);
                bool endOn = ExecuteJointCommand.IsPointOnSegment(wB1, wA0, wA1);
                Point pTworld;
                if (startOn)
                    pTworld = wB0;
                else if (endOn)
                    pTworld = wB1;
                else
                {
                    // fallback intersection
                    Point? pT = JointCurveHelper.IntersectTLines(
                        componentA, rawA,
                        componentB, rawB
                    );
                    if (pT == null)
                    {
                        //Logger.Log("TJoint: ERROR – fallback intersection failed.");
                        return;
                    }
                    pTworld = pT.Value;
                }

                // 4) Compute hemisphere‐stable worldUp = cross(A, B)
                Vector dA = (wA1 - wA0).Direction.ToVector();
                Vector dB = (wB1 - wB0).Direction.ToVector();
                Vector worldUp = Vector.Cross(dA, dB);
                if (worldUp.Magnitude < tol)
                    worldUp = Vector.Create(0, 1, 0);
                else
                    worldUp = worldUp.Direction.ToVector();
                if (Vector.Dot(worldUp, Vector.Create(0, 1, 0)) < 0)
                    worldUp = -worldUp;

                // 5) Compute world‐perp to A’s direction, flip if B is on the “wrong” side
                Vector WY = Vector.Create(0, 1, 0);
                Vector worldPerp = Vector.Cross(worldUp, dA).Magnitude > tol
                                   ? Vector.Cross(worldUp, dA).Direction.ToVector()
                                   : Vector.Cross(dA, WY).Direction.ToVector();
                Point wBmid = wB0 + (wB1 - wB0) * 0.5;
                if (Vector.Dot((wBmid - pTworld).Direction.ToVector(), worldPerp) < 0)
                {
                    worldPerp = -worldPerp;
                    //Logger.Log("  Flipped worldPerp because B lies on its negative side.");
                }

                // 6) Offset origin by half‐width + offsetX
                double halfOff = JointCurveHelper.GetProfileWidth(componentA) * 0.5
                                 + JointCurveHelper.GetOffset(componentA, "offsetX");
                Point cutterOrigin = pTworld + worldPerp * halfOff;

                // 7) Determine extrusion distances for T‐joint
                //    (we build a temporary plane here just to feed into the helper)
                var tempPlane = Plane.Create(Frame.Create(
                    cutterOrigin,
                    dA.Direction,
                    worldUp.Direction
                ));
                var (forwardDist, backwardDist) = JointCurveHelper.PickTJointDirectionWorld(
                    tempPlane,
                    componentB,
                    pTworld,
                    connectedAtStart: startOn,
                    connectedAtEnd: endOn,
                    attachLen: 200.0,
                    freeLen: spacing
                );

                // 8) Build & subtract the cutter in componentB’s local space
                SubtractLocalCutter(
                    componentB,
                    cutterOrigin,
                    cutterOrigin + dA.Direction.ToVector(),
                    worldUp,
                    startConnected: startOn,
                    endConnected: endOn,
                    forwardDist,
                    backwardDist
                );

                //Logger.Log("TJoint: finished.");
            });
        }

        /// <summary>
        /// Mirrors the same local cutter‐building & subtraction logic used by Miter/Straight/None
        /// </summary>
        private void SubtractLocalCutter(
            Component comp,
            Point worldPIn,
            Point worldPOut,
            Vector worldUp,
            bool startConnected,
            bool endConnected,
            double forwardDistance,
            double backwardDistance
        )
        {
            // 1) Map worldUp & points into local
            var inv = comp.Placement.Inverse;
            Direction upDirLocal = inv * worldUp.Direction;
            Vector upLocal = upDirLocal.ToVector();
            Point localPIn = inv * worldPIn;
            Point localPOut = inv * worldPOut;

            //Logger.Log($"Building cutter in LOCAL for '{comp.Name}'.");

            // 2) Build local cutter‐frame & square loop
            var (planeLocal, loopLocal) = JointModule.BuildDebugCutterFrameAndLoop(
                localPIn, localPOut, upLocal, 500.0
            );
            if (planeLocal == null)
            {
                //Logger.Log($"  ERROR: BuildDebugCutterFrameAndLoop returned null for '{comp.Name}'.");
                return;
            }

            // 3) Extrude bi-directionally using the T-joint distances
            var cutterLocal = JointModule.CreateBidirectionalExtrudedBody(
                planeLocal, loopLocal,
                forwardDistance, backwardDistance
            );
            if (cutterLocal == null)
            {
                //Logger.Log($"  ERROR: CreateBidirectionalExtrudedBody returned null for '{comp.Name}'.");
                return;
            }

            // 4) Subtract from the “ExtrudedProfile”
            JointModule.SubtractCutter(comp, cutterLocal);
            //Logger.Log($"  Subtraction complete for '{comp.Name}'.");
        }
    }
}
