using AESCConstruct2026.FrameGenerator.Utilities;
using AESCConstruct2026.Localization;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System.Linq;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;

namespace AESCConstruct2026.FrameGenerator.Modules.Joints
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
                // 1) grab the two raw construction segments
                var rawA = componentA.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
                var rawB = componentB.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
                if (rawA == null || rawB == null)
                {
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

                // PART 1 / LAYER 2 GUARD: a zero/degenerate profile width collapses the
                // miter cutter to a sliver and crashes the ACIS kernel. Reject early.
                if (wA <= 1e-6 || wB <= 1e-6)
                {
                    L.Status("Frame_Joint_Msg_DegenerateGeometry", StatusMessageType.Warning);
                    return;
                }

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
                    return;
                }

                // PART 1 / LAYER 2 GUARD: coincident inner/outer intersections mean a
                // zero-thickness cutter -> native ACIS AV. Use a real tolerance (1e-4 m),
                // NOT 1e-6, because a 1e-5..1e-4 sliver still crashes the kernel.
                double pInOutMag = (pIn.Value - pOut.Value).Magnitude;
                if (pInOutMag < 1e-4)
                {
                    L.Status("Frame_Joint_Msg_DegenerateGeometry", StatusMessageType.Warning);
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

            // build local cutter‐frame & loop
            var (planeLocal, loopLocal) = JointModule.BuildDebugCutterFrameAndLoop(
                inLoc, outLoc, upLocal, 500.0 // mm, cutter size
            );
            if (planeLocal == null)
            {
                return;
            }

            // Deterministic far-end rule (replaces the unreliable PickDirection
            // body-centroid heuristic + the confusing (back,fwd) arg swap):
            //
            // The miter cut must REMOVE the side of the cutting plane that does NOT
            // contain the member's FAR (un-connected) end, and KEEP the long run up
            // to the joint. Far end = Start if endConnected, else End. Both the
            // construction segment (comp.Template = the Part / local frame) and
            // planeLocal.Frame are in the SAME local frame (planeLocal was built
            // from inv*world points, inv = comp.Placement.Inverse), so the dot
            // product below is frame-consistent.
            double longLen = 200.0;            // mm-scale slab (kept large)
            double shortLen = spacing / 2.0;   // small keep-side allowance (gap)

            var rawSeg = comp.Template.Curves
                             .OfType<DesignCurve>()
                             .FirstOrDefault()?.Shape as CurveSegment;

            double fwdDist, backDist;
            if (rawSeg == null)
            {
                // Robustness: never crash — fall back to the previous behaviour.
                var (fwdP, backP) = JointCurveHelper.PickDirection(
                    planeLocal, comp, startConnected, endConnected,
                    longLen: longLen, shortLen: shortLen
                );
                fwdDist = backP;   // preserve the previous (back,fwd) call mapping
                backDist = fwdP;
            }
            else
            {
                Point farLocal = endConnected ? rawSeg.StartPoint : rawSeg.EndPoint;
                Vector planeZ = planeLocal.Frame.DirZ.ToVector();
                Point planeOrigin = planeLocal.Frame.Origin;
                Vector toFar = farLocal - planeOrigin;
                double dot = Vector.Dot(toFar, planeZ);
                bool farOnPlusZ = dot > 0;

                // Extrude the LONG cutter slab on the side WITHOUT the far end
                // (the waste side), the SHORT allowance on the keep side.
                // CreateBidirectionalExtrudedBody: arg1=forwardDistance => +DirZ,
                //                                  arg2=backwardDistance => -DirZ.
                fwdDist = farOnPlusZ ? shortLen : longLen;   // +DirZ extrusion
                backDist = farOnPlusZ ? longLen : shortLen;  // -DirZ extrusion
            }

            // extrude (fwdDist => +DirZ, backDist => -DirZ)
            var cutter = JointModule.CreateBidirectionalExtrudedBody(
                planeLocal, loopLocal, fwdDist, backDist
            );
            if (cutter == null)
            {
                return;
            }

            // subtract
            JointModule.SubtractCutter(comp, cutter);
        }
    }

}
