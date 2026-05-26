using AESCConstruct2026.FrameGenerator.Commands;
using AESCConstruct2026.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System.Linq;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;

namespace AESCConstruct2026.FrameGenerator.Modules.Joints
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

                // 7) Determine extrusion distances for the T-joint.
                //    NOTE: PickTJointDirectionWorld decided direction against a
                //    `tempPlane` whose DirZ does NOT match the planeLocal that
                //    SubtractLocalCutter actually extrudes against (different frame
                //    construction) -> wrong side removed (same frame-mismatch class
                //    as the Miter bug). The deterministic far-end decision is now
                //    made INSIDE SubtractLocalCutter against the real planeLocal.
                SubtractLocalCutter(
                    componentB,
                    cutterOrigin,
                    cutterOrigin + dA.Direction.ToVector(),
                    worldUp,
                    startConnected: startOn,
                    endConnected: endOn,
                    attachLen: 200.0,
                    freeLen: spacing
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
            double attachLen,
            double freeLen
        )
        {
            // 1) Map worldUp & points into local
            var inv = comp.Placement.Inverse;
            Direction upDirLocal = inv * worldUp.Direction;
            Vector upLocal = upDirLocal.ToVector();
            Point localPIn = inv * worldPIn;
            Point localPOut = inv * worldPOut;

            // 2) Build local cutter‐frame & square loop
            var (planeLocal, loopLocal) = JointModule.BuildDebugCutterFrameAndLoop(
                localPIn, localPOut, upLocal, 500.0
            );
            if (planeLocal == null)
            {
                return;
            }

            // 3) Deterministic far-end rule against the REAL planeLocal (the plane
            //    we actually extrude against). For a T-joint the BRANCH member B
            //    must KEEP its free run and have the ATTACHED end (the part that
            //    would penetrate the through member) coped/removed. far end here =
            //    B's FREE (un-connected) end: Start if endConnected else End.
            //    comp.Template (Part) and planeLocal.Frame share the same local
            //    frame (planeLocal built from inv*world points), so dot is
            //    frame-consistent. Long slab (attachLen) goes on the side WITHOUT
            //    the free end (the attached/penetrating side); short (freeLen) on
            //    the keep/free side.
            var rawSeg = comp.Template.Curves
                             .OfType<DesignCurve>()
                             .FirstOrDefault()?.Shape as CurveSegment;

            double forwardDistance, backwardDistance;
            if (rawSeg == null)
            {
                // Robustness: never crash — fall back to prior helper behaviour.
                var tempPlane = planeLocal;
                var (fP, bP) = JointCurveHelper.PickTJointDirectionWorld(
                    tempPlane, comp, localPIn,
                    connectedAtStart: startConnected,
                    connectedAtEnd: endConnected,
                    attachLen: attachLen, freeLen: freeLen
                );
                forwardDistance = fP;
                backwardDistance = bP;
            }
            else
            {
                Point freeLocal = endConnected ? rawSeg.StartPoint : rawSeg.EndPoint;
                Vector planeZ = planeLocal.Frame.DirZ.ToVector();
                Point planeOrigin = planeLocal.Frame.Origin;
                Vector toFree = freeLocal - planeOrigin;
                double dot = Vector.Dot(toFree, planeZ);
                bool freeOnPlusZ = dot > 0;

                // CreateBidirectionalExtrudedBody: arg1=forwardDistance => +DirZ,
                // arg2=backwardDistance => -DirZ. Long slab on the side WITHOUT the
                // free end (attached/penetrating side); short on the free/keep side.
                forwardDistance = freeOnPlusZ ? freeLen : attachLen;
                backwardDistance = freeOnPlusZ ? attachLen : freeLen;
            }

            // 4) Extrude bi-directionally (fwd => +DirZ, back => -DirZ)
            var cutterLocal = JointModule.CreateBidirectionalExtrudedBody(
                planeLocal, loopLocal,
                forwardDistance, backwardDistance
            );
            if (cutterLocal == null)
            {
                //Logger.Log($"  ERROR: CreateBidirectionalExtrudedBody returned null for '{comp.Name}'.");
                return;
            }

            // 4) Subtract from the profile body (resolved via JointModule.FindProfileBody)
            JointModule.SubtractCutter(comp, cutterLocal);
            //Logger.Log($"  Subtraction complete for '{comp.Name}'.");
        }
    }
}
