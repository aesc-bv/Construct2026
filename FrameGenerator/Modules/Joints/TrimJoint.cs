using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System.Collections.Generic;
using System.Linq;

namespace AESCConstruct25.FrameGenerator.Modules.Joints
{
    public class TrimJoint : JointBase
    {
        public override string Name => "Trim";

        public override void Execute(
            Component componentA,
            Component componentB,
            double spacing,
            Body bodyA,
            Body bodyB
        )
        {
            WriteBlock.ExecuteTask("TrimJoint", () =>
            {
                //Logger.Log("TrimJoint.Execute() called");

                if (componentA == null)
                {
                    //Logger.Log("TrimJoint: ERROR - No componentA to trim.");
                    return;
                }

                var window = Window.ActiveWindow;
                var selFace = window.ActiveContext
                                   .Selection
                                   .OfType<IDesignFace>()
                                   .FirstOrDefault();
                if (selFace == null)
                {
                    //Logger.Log("TrimJoint: ERROR – no face selected.");
                    return;
                }

                // 2) Cast to ITrimmedSurface
                if (!(selFace.Shape is ITrimmedSurface trimmed))
                {
                    //Logger.Log("TrimJoint: ERROR – selected object is not a trimmed surface.");
                    return;
                }

                // 3) Get the untrimmed geometry and ensure it's planar
                var untrimmed = trimmed.Geometry;
                if (!(untrimmed is Plane plane))
                {
                    //Logger.Log("TrimJoint: ERROR – face is not planar.");
                    return;
                }

                // 4) Create a datum‐plane in the active part at that exact plane
                var part = window.Scene as Part;
                if (part == null)
                {
                    //Logger.Log("TrimJoint: ERROR – active scene is not a Part.");
                    return;
                }

                //var debugDatum = DatumPlane.Create(part, "DebugPlane", plane);
                //Logger.Log("TrimJoint: Debug datum‐plane created.");

                // ----- start of bidirectional‐cutter logic -----
                try
                {
                    // 1) Get the plane's origin & normal from its Frame
                    Point center = plane.Frame.Origin;     // Frame.Origin gives the plane's point in space :contentReference[oaicite:0]{index=0}
                    Vector normal = plane.Frame.DirZ.ToVector(); // Frame.DirZ is the plane's normal direction :contentReference[oaicite:1]{index=1}

                    if (normal.Magnitude < 1e-6)
                    {
                        //Logger.Log("TrimJoint: WARNING – invalid normal, aborting cutter.");
                        return;
                    }
                    normal = normal / normal.Magnitude;

                    // 2) Build a local X–Y frame on the plane
                    Vector x = Vector.Cross(Vector.Create(0, 1, 0), normal);
                    if (x.Magnitude < 1e-6)
                        x = Vector.Cross(Vector.Create(1, 0, 0), normal);
                    x = x / x.Magnitude;

                    Vector y = Vector.Cross(normal, x);
                    y = y / y.Magnitude;

                    // 3) Reconstruct a cutter‐plane and sketch a square loop
                    //Frame cutterFrame = Frame.Create(center, x.Direction, y.Direction);
                    //Plane cutterPlane = Plane.Create(cutterFrame);

                    double debugSize = 200.0; // 20 cm square
                    double half = debugSize * 0.5;

                    Point p1 = center + (-half) * x + (-half) * y;
                    Point p2 = center + (-half) * x + (half) * y;
                    Point p3 = center + (half) * x + (half) * y;
                    Point p4 = center + (half) * x + (-half) * y;

                    Frame cutterFrame = Frame.Create(center, x.Direction, y.Direction);
                    Plane cutterPlane = Plane.Create(cutterFrame);

                    // 2) Now map everything into componentA’s local space:
                    var inv = componentA.Placement.Inverse;

                    // 2a) Transform the plane’s origin + axes into local
                    Point originLocal = inv * center;

                    // transform the X‐axis direction into local, then to a Vector
                    Direction dirXLocal = inv * x.Direction;
                    Vector xLocal = dirXLocal.ToVector();
                    xLocal = xLocal / xLocal.Magnitude;

                    // same for Y‐axis
                    Direction dirYLocal = inv * y.Direction;
                    Vector yLocal = dirYLocal.ToVector();
                    yLocal = yLocal / yLocal.Magnitude;
                    Frame localFrame = Frame.Create(originLocal, xLocal.Direction, yLocal.Direction);
                    Plane localPlane = Plane.Create(localFrame);

                    // 2b) Transform your sketch loop points into local
                    Point lp1 = inv * p1;
                    Point lp2 = inv * p2;
                    Point lp3 = inv * p3;
                    Point lp4 = inv * p4;
                    var loopLocal = new List<ITrimmedCurve> {
                        CurveSegment.Create(lp1, lp2),
                        CurveSegment.Create(lp2, lp3),
                        CurveSegment.Create(lp3, lp4),
                        CurveSegment.Create(lp4, lp1)
                    };

                    // 3) Create the cutter in **local** coordinates
                    var cutterLocal = JointModule.CreateBidirectionalExtrudedBody(
                        localPlane,
                        loopLocal,
                        spacing,
                        200.0
                    );
                    if (cutterLocal == null)
                    {
                        //Logger.Log("TrimJoint: ERROR – Failed to create extruded cutter body.");
                        return;
                    }

                    // 4) Hand it straight into your SubtractCutter (which expects local bodies)
                    //Logger.Log("TrimJoint: Subtracting cutter from component...");
                    JointModule.SubtractCutter(componentA, cutterLocal);
                    //Logger.Log("TrimJoint: Cutter subtracted successfully.");
                }
                catch (System.Exception)
                {
                    //Logger.Log($"TrimJoint: ERROR during cutter creation: {ex.Message}");
                }

                //Logger.Log("TrimJoint: Execute complete.");
            });
        }
    }
}
