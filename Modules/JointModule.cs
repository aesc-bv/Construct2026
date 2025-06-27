using SpaceClaim.Api.V251;
using SpaceClaim.Api.V251.Geometry;
using SpaceClaim.Api.V251.Modeler;
using System.Collections.Generic;
using System.Linq;
using AESCConstruct25.Modules;
using Point = SpaceClaim.Api.V251.Geometry.Point;
using System;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace AESCConstruct25.Utilities
{
    public static class JointModule
    {
        // single paramater
        public static void ResetComponentGeometryOnly(List<Component> components)
        {
            // delegate into the full overload, letting it collect curves for us
            ResetComponentGeometryOnly(components, Vector.Create(0, 1, 0), null);
        }

        public static void ResetComponentGeometryOnly(
            List<Component> components,
            Vector forcedLocalUp,
            List<DesignCurve> allCurves = null
        )
        {
            WriteBlock.ExecuteTask("ResetComponentGeometryOnly", () =>
            {
                const double tol = 1e-6;

                // 1) Gather all construction curves if not provided
                if (allCurves == null)
                    allCurves = components
                        .SelectMany(c => c.Template?.Curves.OfType<DesignCurve>() ?? Enumerable.Empty<DesignCurve>())
                        .ToList();

                foreach (var component in components)
                {
                    var template = component.Template;
                    if (template == null)
                        continue;

                    // 2) Pull profile metadata
                    string profileType = null;
                    bool isHollow = false;
                    double offsetX = 0, offsetY = 0;
                    var profileData = new Dictionary<string, string>();
                    foreach (var prop in template.CustomProperties)
                    {
                        switch (prop.Key)
                        {
                            case "Type":
                                profileType = prop.Value.Value?.ToString();
                                break;
                            case "Hollow":
                                isHollow = prop.Value.Value?.ToString() == "true";
                                break;
                            case "offsetX":
                                offsetX = (double)prop.Value.Value;
                                break;
                            case "offsetY":
                                offsetY = (double)prop.Value.Value;
                                break;
                            default:
                                if (prop.Key.StartsWith("Construct_"))
                                    profileData[prop.Key.Substring("Construct_".Length)]
                                        = prop.Value.Value?.ToString();
                                break;
                        }
                    }
                    if (string.IsNullOrEmpty(profileType))
                        continue;

                    // 3) Grab stored construction curve (no extension!)
                    var storedDc = template.Curves.OfType<DesignCurve>().FirstOrDefault();
                    if (storedDc?.Shape is not CurveSegment segOrig)
                        continue;

                    // 4) Apply any X/Y offsets baked into the curve
                    Point origStart = Point.Create(
                        segOrig.StartPoint.X + offsetX,
                        segOrig.StartPoint.Y + offsetY,
                        segOrig.StartPoint.Z
                    );
                    Point origEnd = Point.Create(
                        segOrig.EndPoint.X + offsetX,
                        segOrig.EndPoint.Y + offsetY,
                        segOrig.EndPoint.Z
                    );

                    // 5) Clone back to origin so extrude runs from (0,0,0)
                    var delta = origStart;
                    var originSeg = CurveSegment.Create(
                        Point.Create(0, 0, 0),
                        Point.Create(
                            origEnd.X - delta.X,
                            origEnd.Y - delta.Y,
                            origEnd.Z - delta.Z
                        )
                    );

                    // 6) Stabilize localUp (we don’t recompute from neighbour here)
                    Vector localUp = forcedLocalUp.Magnitude > tol
                                   ? forcedLocalUp
                                   : Vector.Create(0, 1, 0);
                    var WY = Vector.Create(0, 1, 0);
                    var WX = Vector.Create(1, 0, 0);
                    var WZ = Vector.Create(0, 0, 1);
                    if (Math.Abs(Vector.Dot(localUp, WY)) > tol && Vector.Dot(localUp, WY) < 0) localUp = -localUp;
                    else if (Math.Abs(Vector.Dot(localUp, WX)) > tol && Vector.Dot(localUp, WX) < 0) localUp = -localUp;
                    else if (Math.Abs(Vector.Dot(localUp, WZ)) > tol && Vector.Dot(localUp, WZ) < 0) localUp = -localUp;

                    // 7) Re-extrude along the original curve (no extension)
                    ProfileModule.ExtrudeProfile(
                        Window.ActiveWindow,
                        profileType,
                        originSeg,
                        isHollow,
                        profileData,
                        offsetX,
                        offsetY,
                        localUp: localUp,
                        dxfFilePath: null,
                        dxfContours: null,
                        reuseComponent: component
                    );
                }
            });
        }



        public static void ResetComponentGeometryAndExtend(
            List<Component> components,
            Vector forcedLocalUp,
            List<DesignCurve> allCurves = null,
            string connectionSideOld = ""
        )
        {
            WriteBlock.ExecuteTask("ResetComponentGeometryAndExtend", () =>
            {
                //Logger.Log("ResetComponentGeometryAndExtend: Resetting and extending multiple components");
                const double tol = 1e-6;
                const double extendAmount = 200.0;

                // 1) Gather all construction curves if not provided
                if (allCurves == null)
                    allCurves = components
                        .SelectMany(c => c.Template?.Curves.OfType<DesignCurve>() ?? Enumerable.Empty<DesignCurve>())
                        .ToList();

                foreach (var component in components)
                {
                    //Logger.Log($"  Processing '{component.Name}'");
                    var template = component.Template;
                    if (template == null)
                    {
                        //Logger.Log("    ERROR – No template found. Skipping.");
                        continue;
                    }

                    // --- pull profile metadata ---
                    string profileType = null;
                    bool isHollow = false;
                    double offsetX = 0, offsetY = 0;
                    var profileData = new Dictionary<string, string>();
                    foreach (var prop in template.CustomProperties)
                    {
                        switch (prop.Key)
                        {
                            case "Type":
                                profileType = prop.Value.Value?.ToString();
                                break;
                            case "Hollow":
                                isHollow = prop.Value.Value?.ToString() == "true";
                                break;
                            case "offsetX":
                                offsetX = (double)prop.Value.Value;
                                break;
                            case "offsetY":
                                offsetY = (double)prop.Value.Value;
                                break;
                            default:
                                if (prop.Key.StartsWith("Construct_"))
                                    profileData[prop.Key.Substring("Construct_".Length)] = prop.Value.Value?.ToString();
                                break;
                        }
                    }
                    if (string.IsNullOrEmpty(profileType))
                    {
                        //Logger.Log("    ERROR – No profileType found. Skipping.");
                        continue;
                    }

                    // 2) Grab the stored (local) construction curve
                    var storedDc = template.Curves.OfType<DesignCurve>().FirstOrDefault();
                    if (storedDc?.Shape is not CurveSegment segOrig)
                    {
                        //Logger.Log("    ERROR – No valid stored construction CurveSegment. Skipping.");
                        continue;
                    }
                    Point origStart = Point.Create(
                        segOrig.StartPoint.X + offsetX,
                        segOrig.StartPoint.Y + offsetY,
                        segOrig.StartPoint.Z
                    );
                    Point origEnd = Point.Create(
                        segOrig.EndPoint.X + offsetX,
                        segOrig.EndPoint.Y + offsetY,
                        segOrig.EndPoint.Z
                    );
                    //Logger.Log($"    [ORIGINAL] start={PointToString(origStart)}, end={PointToString(origEnd)}");

                    // 3) **EXTEND BOTH DIRECTIONS** by 200 in local space
                    Vector segDir = (origEnd - origStart).Direction.ToVector();
                    Point newStart = origStart - segDir * extendAmount;
                    Point newEnd = origEnd + segDir * extendAmount;
                    //Logger.Log($"    [EXTENDED] start={PointToString(newStart)}, end={PointToString(newEnd)}");
                    var segThis = CurveSegment.Create(newStart, newEnd);

                    // 4) Compute a stable localUp
                    Component neighbourForUp = components.FirstOrDefault(c =>
                    {
                        if (c == component) return false;
                        if (c.Template?.Curves.OfType<DesignCurve>().FirstOrDefault()?.Shape is not CurveSegment otherSeg)
                            return false;
                        return (segThis.StartPoint - otherSeg.StartPoint).Magnitude < tol
                            || (segThis.StartPoint - otherSeg.EndPoint).Magnitude < tol
                            || (segThis.EndPoint - otherSeg.StartPoint).Magnitude < tol
                            || (segThis.EndPoint - otherSeg.EndPoint).Magnitude < tol;
                    });

                    Vector localUp;
                    if (neighbourForUp != null)
                    {
                        double w = JointCurveHelper.GetProfileWidth(component);
                        var offs = JointCurveHelper.GetOffsetEdges(component, segThis, neighbourForUp, w, 0.0);
                        Vector perp = offs.perp;
                        Vector rawCross = Vector.Cross((segThis.EndPoint - segThis.StartPoint).Direction.ToVector(), perp);
                        localUp = rawCross.Magnitude > tol
                                  ? rawCross.Direction.ToVector()
                                  : Vector.Create(0, 1, 0);
                        //Logger.Log($"    Computed localUp from neighbour '{neighbourForUp.Name}' = {PointToString(Point.Create(localUp.X, localUp.Y, localUp.Z))}");
                    }
                    else
                    {
                        localUp = forcedLocalUp;
                        //Logger.Log($"    No neighbour for Up; using forcedLocalUp = {PointToString(Point.Create(localUp.X, localUp.Y, localUp.Z))}");
                    }

                    // 5) Hemisphere stabilization
                    var WY = Vector.Create(0, 1, 0);
                    var WX = Vector.Create(1, 0, 0);
                    var WZ = Vector.Create(0, 0, 1);
                    if (Math.Abs(Vector.Dot(localUp, WY)) > tol && Vector.Dot(localUp, WY) < 0)
                    {
                        localUp = -localUp;
                        //Logger.Log("    Flipped localUp to align with WY.");
                    }
                    else if (Math.Abs(Vector.Dot(localUp, WX)) > tol && Vector.Dot(localUp, WX) < 0)
                    {
                        localUp = -localUp;
                        //Logger.Log("    Flipped localUp to align with WX.");
                    }
                    else if (Math.Abs(Vector.Dot(localUp, WZ)) > tol && Vector.Dot(localUp, WZ) < 0)
                    {
                        localUp = -localUp;
                        //Logger.Log("    Flipped localUp to align with WZ.");
                    }
                    //Logger.Log($"    localUp (final) = {PointToString(Point.Create(localUp.X, localUp.Y, localUp.Z))}");

                    // 6) Extrude the profile along the fully extended segThis
                    //Logger.Log($"    Calling ProfileModule.ExtrudeProfile for '{component.Name}'...");
                    try
                    {
                        ProfileModule.ExtrudeProfile(
                            Window.ActiveWindow,
                            profileType,
                            segThis,
                            isHollow,
                            profileData,
                            offsetX,
                            offsetY,
                            localUp: localUp,
                            dxfFilePath: null,
                            dxfContours: null,
                            reuseComponent: component
                        );
                        //Logger.Log($"    ExtrudeProfile succeeded for '{component.Name}'.");
                    }
                    catch (Exception ex)
                    {
                        //Logger.Log($"    ERROR in ExtrudeProfile for '{component.Name}': {ex.Message}");
                    }
                }

                //Logger.Log("ResetComponentGeometryAndExtend: All components processed");
            });
        }

        public static (Plane, IList<ITrimmedCurve>) BuildDebugCutterFrameAndLoop(
            Point start,     // WORLD‐space intersection of inner lines
            Point end,       // WORLD‐space intersection of outer lines
            Vector worldUp,  // WORLD‐space “up” vector
            double size = 0.05
        )
        {
            // 1) “between” vector in world coordinates
            Vector between = end - start;
            if (between.Magnitude < 1e-6)
            {
                //Logger.Log("BuildDebugCutterFrameAndLoop (world): Skipped due to zero‐length vector.");
                return (null, null);
            }

            // 2) axisX = normalized (end – start) in world
            Vector axisX = between / between.Magnitude;
            //Logger.Log($"BuildDebugCutterFrameAndLoop (world): axisX = {axisX}");

            // 3) Gram–Schmidt: take worldUp, remove any component along axisX, then normalize → axisY
            Vector rawUp = worldUp;
            Vector projOnX = Vector.Dot(rawUp, axisX) * axisX;
            Vector candidateY = rawUp - projOnX;

            if (candidateY.Magnitude < 1e-6)
            {
                // If worldUp is nearly parallel to axisX, pick any vector perpendicular to axisX
                Vector test = Vector.Create(1, 0, 0);
                candidateY = Vector.Cross(axisX, test).Magnitude > 1e-6
                             ? Vector.Cross(axisX, test)
                             : Vector.Cross(axisX, Vector.Create(0, 1, 0));
            }

            Vector axisY = candidateY / candidateY.Magnitude;
            //Logger.Log($"BuildDebugCutterFrameAndLoop (world): axisY = {axisY}");

            // 4) axisZ = axisX × axisY (world normal)
            Vector axisZ = Vector.Cross(axisX, axisY);
            axisZ = axisZ / axisZ.Magnitude;
            //Logger.Log($"BuildDebugCutterFrameAndLoop (world): axisZ = {axisZ}");

            // 5) Build a world‐space Frame so DirX = axisX, DirY = axisY (DirZ computed automatically)
            Point center = start + between * 0.5;
            Frame frame = Frame.Create(center, axisX.Direction, axisY.Direction);
            //Logger.Log(
            //  $"BuildDebugCutterFrameAndLoop (world): Frame origin={center}, " +
            //  $"dirX={frame.DirX}, dirY={frame.DirY}, dirZ={frame.DirZ}"
            //);

            Plane plane = Plane.Create(frame);

            // 6) Build a small square loop of side “size” in that plane’s local X–Y (all in world)
            double half = size / 2;
            Point p1 = center + (-half) * axisX + (-half) * axisY;
            Point p2 = center + (-half) * axisX + (+half) * axisY;
            Point p3 = center + (+half) * axisX + (+half) * axisY;
            Point p4 = center + (+half) * axisX + (-half) * axisY;

            IList<ITrimmedCurve> loop = new List<ITrimmedCurve>
            {
                CurveSegment.Create(p1, p2),
                CurveSegment.Create(p2, p3),
                CurveSegment.Create(p3, p4),
                CurveSegment.Create(p4, p1)
            };

            return (plane, loop);
        }

        public static Body CreateBidirectionalExtrudedBody(
            Plane plane,
            IList<ITrimmedCurve> loop,
            double forwardDistance,
            double backwardDistance
        )
        {
            //Logger.Log($"CreateBidirectionalExtrudedBody: forwardDistance={forwardDistance}, backwardDistance={backwardDistance}");

            if (forwardDistance < 1e-6 && backwardDistance < 1e-6)
            {
                //Logger.Log("  Both distances are zero — skipping extrusion.");
                return null;
            }

            var profile = new Profile(plane, loop);
            //Logger.Log("  Created Profile object.");

            Body bodyForward = null;
            if (forwardDistance > 1e-6)
            {
                //Logger.Log($"  Extruding forward by {forwardDistance}.");
                bodyForward = Body.ExtrudeProfile(profile, forwardDistance);
                //Logger.Log($"  Finished forward extrusion; result != null? {bodyForward != null}");
            }

            Body bodyBackward = null;
            if (backwardDistance > 1e-6)
            {
                //Logger.Log($"  Extruding backward by {backwardDistance}.");
                Direction xDir = plane.Frame.DirX;
                Direction yDir = plane.Frame.DirY;
                Point origin = plane.Frame.Origin;
                Frame flippedFrame = Frame.Create(origin, xDir, -yDir);
                Plane flippedPlane = Plane.Create(flippedFrame);
                //Logger.Log($"  Created flippedPlane at origin={origin}, dirX={xDir}, dirY={-yDir}.");

                var flippedProfile = new Profile(flippedPlane, loop);
                bodyBackward = Body.ExtrudeProfile(flippedProfile, backwardDistance);
                //Logger.Log($"  Finished backward extrusion; result != null? {bodyBackward != null}");
            }

            if (bodyForward != null && bodyBackward != null)
            {
                //Logger.Log("  Uniting forward and backward bodies.");
                bodyForward.Unite(new[] { bodyBackward });
                return bodyForward;
            }
            else if (bodyForward != null)
            {
                //Logger.Log("  Returning only forward body.");
                return bodyForward;
            }
            else
            {
                //Logger.Log("  Returning only backward body.");
                return bodyBackward;
            }
        }

        public static void SubtractCutter(
            Component component,
            Body cutter
        )
        {
            if (component == null || cutter == null) return;

            Part part = component.Template;
            DesignBody target = part.Bodies.FirstOrDefault(b => b.Name == "ExtrudedProfile");
            if (target == null)
            {
                //Logger.Log($"JointModule: No 'ExtrudedProfile' body found in {component.Name}");
                return;
            }

            //Logger.Log($"  Target BB local center before subtract = {PointToString(target.Shape.GetBoundingBox(Matrix.Identity, true).Center)}");

            // Insert cutter into component’s Part (local)
            DesignBody temp = DesignBody.Create(part, "TempCutter", cutter.Copy());
            //DesignBody temp2 = DesignBody.Create(part, "TempCutter2", cutter.Copy());

            temp.Layer = part.Document.GetLayer("Frames");
            temp.SetVisibility(null, true);

            try
            {
                //Logger.Log($"  Subtracting cutter from {component.Name} …");
                target.Shape.Subtract(new[] { temp.Shape });
                //Logger.Log($"  Subtraction succeeded. New target BB local center = {PointToString(target.Shape.GetBoundingBox(Matrix.Identity, true).Center)}");
            }
            catch (System.Exception ex)
            {
                //Logger.Log($"JointModule: ERROR during subtraction: {ex.Message}");
            }
        }
        public static (double forward, double backward) DetermineExtrusionDirection(
            Point curveJointPoint, Point curveFarPoint,
            Plane plane,
            double longLength,
            double shortLength)
        {
            // Step 1: Direction from far end toward the joint
            Vector curveDirection = (curveJointPoint - curveFarPoint);

            // Step 2: Compare both Z+ and Z- of the extrusion plane
            Vector z = plane.Frame.DirZ.ToVector();
            Vector minusZ = -z;

            double dotPlus = Vector.Dot(curveDirection, z);
            double dotMinus = Vector.Dot(curveDirection, minusZ);

            // Step 3: Choose the extrusion direction (which is more aligned with curve)
            bool usePositiveZ = Math.Abs(dotPlus) >= Math.Abs(dotMinus);

            return usePositiveZ
                ? (longLength, shortLength)
                : (shortLength, longLength);
        }

        //public static (double forward, double backward) DetermineTExtrusionDirection(
        //    Point curveJointPoint, Point curveFarPoint,
        //    Plane plane,
        //    double longLength,
        //    double shortLength)
        //{
        //    // Now measure from the joint TOWARD the free end
        //    Vector curveDirection = (curveFarPoint - curveJointPoint);

        //    Vector z = plane.Frame.DirZ.ToVector();
        //    Vector minusZ = -z;

        //    double dotPlus = Vector.Dot(curveDirection, z);
        //    double dotMinus = Vector.Dot(curveDirection, minusZ);

        //    bool usePositiveZ = Math.Abs(dotPlus) >= Math.Abs(dotMinus);

        //    // If the free‐end direction aligns better with +Z, +Z gets longLength (spacing),
        //    // and –Z gets shortLength (200 through the other part), else swap.
        //    return usePositiveZ
        //        ? (longLength, shortLength)
        //        : (shortLength, longLength);
        //}


        /// <summary>
        /// Splits the ExtrudedProfile (or merged halves) at its midpoint,
        /// using `localUp` to define the cutter plane orientation.
        /// </summary>
        public static (Body bodyStart, Body bodyEnd) SplitBodyAtMidpoint(
             Component component,
             Vector localUp
         )
        {
            try
            {
                var dc = component.Template
                                  .Curves
                                  .OfType<DesignCurve>()
                                  .FirstOrDefault();
                var rawSeg = dc?.Shape as CurveSegment;
                if (rawSeg == null)
                {
                    //Logger.Log("SplitBodyAtMidpoint: ERROR – no construction curve.");
                    return (null, null);
                }

                // --- NEW: half-width shift along local X
                double halfW = JointCurveHelper.GetProfileWidth(component) * 0.5;
                Vector shiftX = Vector.Create(1, 0, 0) * halfW;

                // apply that shift in local
                Point localStart = rawSeg.StartPoint + shiftX;
                Point localEnd = rawSeg.EndPoint + shiftX;

                // midpoint
                Point mid = Point.Create(
                    0.5 * (localStart.X + localEnd.X),
                    0.5 * (localStart.Y + localEnd.Y),
                    0.5 * (localStart.Z + localEnd.Z)
                );

                // validate localUp
                Vector zDir = (localEnd - localStart).Direction.ToVector();
                if (localUp.Magnitude < 1e-6 ||
                    Math.Abs(Vector.Dot(localUp, zDir)) > 0.99)
                    localUp = Vector.Create(0, 1, 0);

                // build local plane
                Direction dZ = (localEnd - localStart).Direction;
                Direction dX = Vector.Cross(localUp, zDir).Direction;
                Direction dY = Direction.Cross(dZ, dX);
                Frame frame = Frame.Create(mid, dX, dY);
                Plane plane = Plane.Create(frame);

                // big square loop
                double half = 200.0;
                Vector vx = plane.Frame.DirX.ToVector() * half;
                Vector vy = plane.Frame.DirY.ToVector() * half;
                var loop = new List<ITrimmedCurve> {
                    CurveSegment.Create(mid - vx - vy, mid - vx + vy),
                    CurveSegment.Create(mid - vx + vy, mid + vx + vy),
                    CurveSegment.Create(mid + vx + vy, mid + vx - vy),
                    CurveSegment.Create(mid + vx - vy, mid - vx - vy),
                };

                // fetch or merge existing ExtrudedProfile
                Body original = component.Template
                                    .Bodies
                                    .FirstOrDefault(b => b.Name == "ExtrudedProfile")
                                    ?.Shape.Copy();
                if (original == null)
                {
                    var hs = component.Template.Bodies
                                .FirstOrDefault(b => b.Name == "HalfStart")?.Shape.Copy();
                    var he = component.Template.Bodies
                                .FirstOrDefault(b => b.Name == "HalfEnd")?.Shape.Copy();
                    if (hs != null && he != null)
                    {
                        hs.Unite(new[] { he.Copy() });
                        original = hs;
                    }
                }
                if (original == null)
                {
                    //Logger.Log("SplitBodyAtMidpoint: ERROR – no profile to split.");
                    return (null, null);
                }

                // create two cutters
                var (fwdE, backE) = DetermineExtrusionDirection(mid, localEnd, plane, 200, 0);
                var cutterEnd = CreateBidirectionalExtrudedBody(plane, loop, fwdE, backE);
                var (fwdS, backS) = DetermineExtrusionDirection(mid, localStart, plane, 0, 200);
                var cutterStart = CreateBidirectionalExtrudedBody(plane, loop, fwdS, backS);
                if (cutterEnd == null || cutterStart == null)
                {
                    //Logger.Log("SplitBodyAtMidpoint: ERROR – cutter creation failed.");
                    return (null, null);
                }

                // boolean‐intersect to get halves
                Body halfEnd = original.Copy();
                Body halfStart = original.Copy();
                halfEnd.Intersect(new[] { cutterEnd.Copy() });
                halfStart.Intersect(new[] { cutterStart.Copy() });

                // delete old halves
                foreach (var name in new[] { "HalfStart", "HalfEnd" })
                    component.Template.Bodies
                             .FirstOrDefault(b => b.Name == name)
                             ?.Delete();

                // create new DesignBodies
                var framesLayer = component.Template.Document.GetLayer("Frames");
                DesignBody.Create(component.Template, "HalfStart", halfStart)
                         .Layer = framesLayer;
                DesignBody.Create(component.Template, "HalfEnd", halfEnd)
                         .Layer = framesLayer;

                return (halfStart, halfEnd);
            }
            catch (Exception ex)
            {
                //Logger.Log($"SplitBodyAtMidpoint: EXCEPTION – {ex.Message}");
                return (null, null);
            }
        }


        public static void ResetHalfForJoint(
            Component component,
            string connectionSide,
            bool extendProfile,
            Vector _ignoredLocalUp,
            List<DesignCurve> allCurves,
            List<Component> selectedComponents
        )
        {
            Logger.Log($"ResetHalfForJoint: Starting for '{component?.Name}', side={connectionSide}");
            if (component?.Template == null)
            {
                Logger.Log("ResetHalfForJoint: ERROR – invalid component or template; aborting.");
                return;
            }
            var template = component.Template;

            // 1) Grab & shift the construction segment
            var dc = template.Curves.OfType<DesignCurve>().FirstOrDefault();
            if (dc?.Shape is not CurveSegment seg)
            {
                Logger.Log("ResetHalfForJoint: ERROR – no valid construction curve.");
                return;
            }
            double halfW = JointCurveHelper.GetProfileWidth(component) * 0.5;
            Vector shiftX = Vector.Create(1, 0, 0) * halfW;
            Point localStart = seg.StartPoint - shiftX;
            Point localEnd = seg.EndPoint - shiftX;

            // 2) compute localUp, xDir, yDir
            Vector sweepDir = (localEnd - localStart).Direction.ToVector();
            if (sweepDir.Magnitude < 1e-6)
            {
                Logger.Log("ResetHalfForJoint: ERROR – zero-length segment.");
                return;
            }
            Vector localUp = Vector.Create(0, 1, 0);
            Vector xDir = Vector.Cross(localUp, sweepDir).Direction.ToVector();
            if (xDir.Magnitude < 1e-6)
            {
                localUp = Vector.Create(1, 0, 0);
                xDir = Vector.Cross(localUp, sweepDir).Direction.ToVector();
            }
            Vector yDir = Vector.Cross(sweepDir, xDir).Direction.ToVector();

            // 3) First split
            var (halfStart, halfEnd) = SplitBodyAtMidpoint(component, localUp);
            if (halfStart == null || halfEnd == null)
            {
                Logger.Log("ResetHalfForJoint: ERROR – initial split failed.");
                return;
            }

            // 4) preserve non-corner half
            bool preserveEnd = connectionSide == "HalfStart";
            Body preserved = preserveEnd ? halfEnd.Copy() : halfStart.Copy();

            // 5) delete old & stash preserved
            foreach (var name in new[] { "HalfStart", "HalfEnd", "ExtrudedProfile" })
                template.Bodies.FirstOrDefault(b => b.Name == name)?.Delete();
            DesignBody.Create(template, "preservedHalf", preserved)
                     .Layer = template.Document.GetLayer("Frames");

            // 6) regenerate full profile
            if (extendProfile)
            {
                Logger.Log("  Regenerating extended profile...");
                ResetComponentGeometryAndExtend(
                    new List<Component> { component },
                    localUp,
                    allCurves,
                    connectionSide
                );
            }
            else
            {
                Logger.Log("  Regenerating original (unextended) profile...");
                ResetComponentGeometryOnly(
                    new List<Component> { component },
                    localUp,
                    allCurves
                );
            }
            if (template.Bodies.All(b => b.Name != "ExtrudedProfile"))
            {
                Logger.Log("ResetHalfForJoint: ERROR – no 'ExtrudedProfile' after extend.");
                return;
            }

            // 7) second split
            var (halfStart2, halfEnd2) = SplitBodyAtMidpoint(component, localUp);
            if (halfStart2 == null || halfEnd2 == null)
            {
                Logger.Log("ResetHalfForJoint: ERROR – second split failed.");
                return;
            }

            // 8) isolate corner half
            Body corner = preserveEnd ? halfStart2 : halfEnd2;

            // 9) copy & delete temps
            Body presCopy = preserved.Copy();
            Body corCopy = corner.Copy();
            foreach (var name in new[] { "HalfStart", "HalfEnd", "ExtrudedProfile", "preservedHalf" })
                template.Bodies.FirstOrDefault(b => b.Name == name)?.Delete();

            // 10) unite into final
            presCopy.Unite(new[] { corCopy });
            DesignBody finalDb = DesignBody.Create(template, "ExtrudedProfile", presCopy);
            finalDb.Layer = template.Document.GetLayer("Frames");

            Logger.Log("ResetHalfForJoint: Completed successfully; 'ExtrudedProfile' is ready.");
        }



        private static string PointToString(Point p)
        {
            return $"({p.X:F4}, {p.Y:F4}, {p.Z:F4})";
        }
    }
}
