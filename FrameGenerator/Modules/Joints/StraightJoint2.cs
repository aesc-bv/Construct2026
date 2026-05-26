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

                // 5) Determine connection ends in WORLD space (per member).
                //    BUG (fixed): the previous code used FindSharedPoint(rawA,rawB)
                //    which compares each member's LOCAL construction points without
                //    lifting through Placement. Every Construct profile's local
                //    segment is (0,0,0)->(0,0,len), so rawA and rawB always
                //    "coincide" at local (0,0,0) -> it always returned
                //    aStartConn=bStartConn=True, aEndConn=bEndConn=False regardless
                //    of true connectivity. Both members then got the SAME keep
                //    side. Lift to world (as StraightJoint/MiterJoint do) so each
                //    member's joint end is its OWN true connected end.
                const double connTol = 1e-6;
                bool aStartConn = (wA0 - wB0).Magnitude < connTol || (wA0 - wB1).Magnitude < connTol;
                bool aEndConn = !aStartConn;
                bool bStartConn = (wB0 - wA0).Magnitude < connTol || (wB0 - wA1).Magnitude < connTol;
                bool bEndConn = !bStartConn;

                // ROLE MAPPING (proven from JointSelectionHelper.GetSelectedComponents
                // + GetConnectedPairs): panel Selection 1 = componentA = LONG part
                // (kept full, only trimmed to meet => shortLen 0); panel Selection 2
                // = componentB = SHORTENED part (=> shortLen spacing). The role
                // asymmetry is carried by these per-call shortLen values; the
                // keep/remove DIRECTION uses the sign-robust far-end rule.

                // 6) Cutter on A (Sel1, LONG): base=pY, offset-pZ, perp=perpA
                SubtractLocalCutter(
                    componentA,
                    worldBase: pY.Value,
                    worldOffset: pZ.Value,
                    worldUp: worldUp,
                    perpLocal: perpA,
                    startConnected: aStartConn,
                    endConnected: aEndConn,
                    longLen: 200.0,
                    shortLen: 0.0,
                    role: "Sel1-LONG"
                );

                // 7) Cutter on B (Sel2, SHORT): base=pIn, offset = along A’s world-dir by ε
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
                    shortLen: spacing,
                    role: "Sel2-SHORT"
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
            double shortLen,
            string role
        )
        {
            // 1) world→local reprojection
            var inv = comp.Placement.Inverse;
            Direction upDirLocal = inv * worldUp.Direction;
            Vector upLocal = upDirLocal.ToVector();
            Point baseLoc = inv * worldBase;
            Point offLoc = inv * worldOffset;

            // 2) build cutter‐plane + loop in local
            var (planeLocal, loopLocal) = JointModule.BuildDebugCutterFrameAndLoop(
                baseLoc, offLoc, upLocal, 500.0
            );
            if (planeLocal == null)
            {
                return;
            }

            // 3) Sign-/selection-order-robust keep/remove DIRECTION (the role
            //    asymmetry is already carried by longLen/shortLen from Execute:
            //    Sel1-LONG passes shortLen 0, Sel2-SHORT passes shortLen spacing).
            //    The cut plane for the SHORT member is built parallel to A's axis
            //    ("end cut parallel to selection 2") => DEGEN branch of
            //    PickEndCutDirection (member-axis vs plane-normal ~0) -> body-
            //    centroid fallback. far end = member's own un-connected end.
            double fwdDist, backDist;
            var rawSeg = comp.Template.Curves
                             .OfType<DesignCurve>()
                             .FirstOrDefault()?.Shape as CurveSegment;
            if (rawSeg == null)
            {
                // Robustness: never crash — fall back to prior behaviour exactly
                // (PickDirection + the previous (back,fwd) call mapping).
                var (fwdP, backP) = JointCurveHelper.PickDirection(
                    planeLocal, comp, startConnected, endConnected,
                    longLen: longLen, shortLen: shortLen
                );
                fwdDist = backP;
                backDist = fwdP;
            }
            else
            {
                var memberBody = JointModule.FindProfileBody(comp.Template)?.Shape;
                (fwdDist, backDist) = JointModule.PickEndCutDirection(
                    planeLocal, rawSeg, endConnected, memberBody,
                    longLen, shortLen, $"Straight2[{role}]:{comp?.Template?.Name}"
                );
            }

            // 4) extrude & subtract (fwdDist => +DirZ, backDist => -DirZ; no swap)
            var cutterLocal = JointModule.CreateBidirectionalExtrudedBody(
                planeLocal, loopLocal, fwdDist, backDist
            );
            if (cutterLocal == null)
            {
                return;
            }

            JointModule.SubtractCutter(comp, cutterLocal);
        }
    }
}
