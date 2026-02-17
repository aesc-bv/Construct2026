/*
 JointModule encapsulates joint-related operations for frame components.

 It is responsible for:
 - Rebuilding profile geometry from stored metadata (DXF/CSV/built-in) using ProfileModule.
 - Optionally extending profiles along their path for joint construction and splitting them into halves.
 - Building cutter frames and bidirectional cutter bodies, and applying boolean subtraction.
 - Providing joint-specific reset logic (per half/side) used by miter and straight joints.
*/

using AESCConstruct2026.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Application = SpaceClaim.Api.V242.Application;
using Component = SpaceClaim.Api.V242.Component;
using Point = SpaceClaim.Api.V242.Geometry.Point;

namespace AESCConstruct2026.FrameGenerator.Modules
{
    public static class JointModule
    {
        // Holds the metadata parsed from a component's custom properties.
        private struct ComponentMetadata
        {
            public string ProfileType;
            public bool IsHollow;
            public double OffsetX;
            public double OffsetY;
            public string DxfPath;
            public string RawCsv;
            public Dictionary<string, string> ProfileData;
        }

        // Reads profile metadata (type, hollow, offsets, DXF/CSV paths, Construct_ properties) from a component's template.
        private static bool TryReadComponentMetadata(Component component, out ComponentMetadata meta)
        {
            meta = default;
            var template = component.Template;
            if (template == null) return false;

            string profileType = null;
            bool isHollow = false;
            double offsetX = 0, offsetY = 0;
            string dxfPath = null;
            string rawCsv = null;
            var profileData = new Dictionary<string, string>();

            foreach (var prop in template.CustomProperties)
            {
                var key = prop.Key;
                var val = prop.Value.Value?.ToString();
                switch (key)
                {
                    case "Type":
                        profileType = val;
                        break;
                    case "Hollow":
                        isHollow = string.Equals(val, "true", StringComparison.OrdinalIgnoreCase);
                        break;
                    case "offsetX":
                        double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out offsetX);
                        break;
                    case "offsetY":
                        double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out offsetY);
                        break;
                    case "DXFPath":
                        dxfPath = val;
                        break;
                    case "RawCSV":
                        rawCsv = val;
                        break;
                    default:
                        if (key.StartsWith("Construct_"))
                            profileData[key.Substring("Construct_".Length)] = val;
                        break;
                }
            }

            if (string.IsNullOrEmpty(profileType)) return false;

            meta = new ComponentMetadata
            {
                ProfileType = profileType,
                IsHollow = isHollow,
                OffsetX = offsetX,
                OffsetY = offsetY,
                DxfPath = dxfPath,
                RawCsv = rawCsv,
                ProfileData = profileData
            };
            return true;
        }

        // Resolves DXF/CSV contours from metadata, returning null for built-in profile types.
        private static List<ITrimmedCurve> ResolveContours(ComponentMetadata meta, out bool failed)
        {
            failed = false;
            List<ITrimmedCurve> contours = null;

            if (string.Equals(meta.ProfileType, "DXF", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(meta.DxfPath))
            {
                if (!DXFImportHelper.ImportDXFContours(meta.DxfPath, out contours))
                {
                    failed = true;
                    return null;
                }
            }
            else if (string.Equals(meta.ProfileType, "CSV", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(meta.RawCsv))
            {
                contours = new List<ITrimmedCurve>();
                foreach (var loop in meta.RawCsv.Split('&'))
                    foreach (var chunk in loop.Trim().Split(' '))
                        if (!string.IsNullOrEmpty(chunk))
                            contours.Add(DXFImportHelper.CurveFromString(chunk));
            }

            return contours;
        }

        // Convenience overload that resets component geometry using a default up-vector and inferred construction curves.
        public static void ResetComponentGeometryOnly(List<Component> components)
        {
            ResetComponentGeometryOnly(components, Vector.Create(0, 1, 0), null);
        }

        // Rebuilds component profile geometry (ExtrudedProfile) from stored metadata, using a forced local up-vector and optional curve cache.
        public static void ResetComponentGeometryOnly(
            List<Component> components,
            Vector forcedLocalUp,
            List<DesignCurve> allCurves = null
        )
        {
            try
            {
                WriteBlock.ExecuteTask("ResetComponentGeometryOnly", () =>
                {
                    const double tol = 1e-6;

                    if (allCurves == null)
                        allCurves = components
                            .SelectMany(c => c.Template?.Curves.OfType<DesignCurve>() ?? Enumerable.Empty<DesignCurve>())
                            .ToList();

                    foreach (var component in components)
                    {
                        try
                        {
                            if (!TryReadComponentMetadata(component, out var meta)) continue;

                            var storedDc = component.Template.Curves.OfType<DesignCurve>().FirstOrDefault();
                            if (storedDc?.Shape is not CurveSegment segOrig) continue;

                            var origStart = Point.Create(segOrig.StartPoint.X + meta.OffsetX, segOrig.StartPoint.Y + meta.OffsetY, segOrig.StartPoint.Z);
                            var origEnd = Point.Create(segOrig.EndPoint.X + meta.OffsetX, segOrig.EndPoint.Y + meta.OffsetY, segOrig.EndPoint.Z);
                            var delta = origStart;
                            var originSeg = CurveSegment.Create(
                                Point.Create(0, 0, 0),
                                Point.Create(origEnd.X - delta.X, origEnd.Y - delta.Y, origEnd.Z - delta.Z)
                            );

                            Vector localUp = forcedLocalUp.Magnitude > tol
                                           ? forcedLocalUp
                                           : Vector.Create(0, 1, 0);
                            Vector WY = Vector.Create(0, 1, 0), WX = Vector.Create(1, 0, 0), WZ = Vector.Create(0, 0, 1);
                            if (Math.Abs(Vector.Dot(localUp, WY)) > tol && Vector.Dot(localUp, WY) < 0) localUp = -localUp;
                            else if (Math.Abs(Vector.Dot(localUp, WX)) > tol && Vector.Dot(localUp, WX) < 0) localUp = -localUp;
                            else if (Math.Abs(Vector.Dot(localUp, WZ)) > tol && Vector.Dot(localUp, WZ) < 0) localUp = -localUp;

                            var contours = ResolveContours(meta, out bool contourFailed);
                            if (contourFailed) continue;

                            ProfileModule.ExtrudeProfile(
                                Window.ActiveWindow,
                                meta.ProfileType,
                                originSeg,
                                meta.IsHollow,
                                meta.ProfileData,
                                meta.OffsetX,
                                meta.OffsetY,
                                localUp,
                                dxfFilePath: meta.ProfileType == "DXF" ? meta.DxfPath : null,
                                dxfContours: contours,
                                reuseComponent: component
                            );
                        }
                        catch (Exception ex)
                        {
                            Application.ReportStatus($"Error resetting geometry for component '{component?.Name}':\n{ex.Message}", StatusMessageType.Error, null);
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Unexpected error during geometry reset:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Rebuilds and extends component profiles along their path to give extra length for joint construction.
        public static void ResetComponentGeometryAndExtend(
            List<Component> components,
            Vector forcedLocalUp,
            List<DesignCurve> allCurves = null,
            string connectionSideOld = ""
        )
        {
            try
            {
                const double tol = 1e-6;
                const double extendAmount = 200.0; // mm

                if (allCurves == null)
                    allCurves = components
                        .SelectMany(c => c.Template?.Curves.OfType<DesignCurve>() ?? Enumerable.Empty<DesignCurve>())
                        .ToList();

                foreach (var component in components)
                {
                    try
                    {
                        if (!TryReadComponentMetadata(component, out var meta)) continue;

                        var storedDc = component.Template.Curves.OfType<DesignCurve>()
                                          .FirstOrDefault(dc => dc.Shape is CurveSegment);
                        if (storedDc?.Shape is not CurveSegment segOrig)
                            continue;

                        var origStart = Point.Create(segOrig.StartPoint.X + meta.OffsetX, segOrig.StartPoint.Y + meta.OffsetY, segOrig.StartPoint.Z);
                        var origEnd = Point.Create(segOrig.EndPoint.X + meta.OffsetX, segOrig.EndPoint.Y + meta.OffsetY, segOrig.EndPoint.Z);

                        Vector segDir = (origEnd - origStart).Direction.ToVector();
                        var newStart = origStart - segDir * extendAmount;
                        var newEnd = origEnd + segDir * extendAmount;
                        var segThis = CurveSegment.Create(newStart, newEnd);

                        Vector localUp = forcedLocalUp.Magnitude > tol ? forcedLocalUp : Vector.Create(0, 1, 0);
                        if (Math.Abs(Vector.Dot(localUp, Vector.Create(0, 1, 0))) > tol && Vector.Dot(localUp, Vector.Create(0, 1, 0)) < 0) localUp = -localUp;

                        var contours = ResolveContours(meta, out bool contourFailed);
                        if (contourFailed) continue;

                        ProfileModule.ExtrudeProfile(
                            Window.ActiveWindow,
                            meta.ProfileType,
                            segThis,
                            meta.IsHollow,
                            meta.ProfileData,
                            meta.OffsetX,
                            meta.OffsetY,
                            localUp,
                            dxfFilePath: meta.ProfileType == "DXF" ? meta.DxfPath : null,
                            dxfContours: contours,
                            reuseComponent: component
                        );
                    }
                    catch (Exception ex)
                    {
                        Application.ReportStatus($"Error extending geometry for component '{component?.Name}':\n{ex.Message}", StatusMessageType.Error, null);
                    }
                }
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Unexpected error during geometry reset:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }



        // Builds a planar frame and small rectangular loop between two points, used to position debug or cutter geometry.
        public static (Plane, IList<ITrimmedCurve>) BuildDebugCutterFrameAndLoop(
            Point start,
            Point end,
            Vector worldUp,
            double size = 0.05
        )
        {
            // 1) “between” vector in world coordinates
            Vector between = end - start;
            if (between.Magnitude < 1e-6)
            {
                return (null, null);
            }

            // 2) axisX = normalized (end – start) in world
            Vector axisX = between / between.Magnitude;

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

            // 4) axisZ = axisX × axisY (world normal)
            Vector axisZ = Vector.Cross(axisX, axisY);
            axisZ = axisZ / axisZ.Magnitude;

            // 5) Build a world‐space Frame so DirX = axisX, DirY = axisY (DirZ computed automatically)
            Point center = start + between * 0.5;
            Frame frame = Frame.Create(center, axisX.Direction, axisY.Direction);

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

        // Creates a body by extruding a loop forward and backward from a plane, merging the two resulting volumes if needed.
        public static Body CreateBidirectionalExtrudedBody(
            Plane plane,
            IList<ITrimmedCurve> loop,
            double forwardDistance,
            double backwardDistance
        )
        {
            if (forwardDistance < 1e-6 && backwardDistance < 1e-6)
            {
                return null;
            }

            var profile = new Profile(plane, loop);

            Body bodyForward = null;
            if (forwardDistance > 1e-6)
            {
                bodyForward = Body.ExtrudeProfile(profile, forwardDistance);
            }

            Body bodyBackward = null;
            if (backwardDistance > 1e-6)
            {
                Direction xDir = plane.Frame.DirX;
                Direction yDir = plane.Frame.DirY;
                Point origin = plane.Frame.Origin;
                Frame flippedFrame = Frame.Create(origin, xDir, -yDir);
                Plane flippedPlane = Plane.Create(flippedFrame);

                var flippedProfile = new Profile(flippedPlane, loop);
                bodyBackward = Body.ExtrudeProfile(flippedProfile, backwardDistance);
            }

            if (bodyForward != null && bodyBackward != null)
            {
                bodyForward.Unite(new[] { bodyBackward });
                return bodyForward;
            }
            else if (bodyForward != null)
            {
                return bodyForward;
            }
            else
            {
                return bodyBackward;
            }
        }

        // Subtracts the cutter body from the component’s ExtrudedProfile body inside the component’s local Part.
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
                return;
            }

            // Insert cutter into component’s Part (local)
            DesignBody temp = DesignBody.Create(part, "TempCutter", cutter.Copy());

            temp.Layer = part.Document.GetLayer("Frames");
            temp.SetVisibility(null, true);

            try
            {
                target.Shape.Subtract(new[] { temp.Shape });
            }
            catch (System.Exception ex)
            {
                Application.ReportStatus($"Joint error: {ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Determines which side of the plane should receive the long/short extrusion based on the curve direction toward the joint.
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

        /// <summary>
        /// Splits the ExtrudedProfile (or merged halves) at its midpoint,
        /// using `localUp` to define the cutter plane orientation.
        /// </summary>
        // Splits the component’s profile body into start/end halves around the path midpoint using a local up-vector.
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
                double half = 200.0; // mm, half-size of splitting square
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
                    return (null, null);
                }

                // create two cutters
                var (fwdE, backE) = DetermineExtrusionDirection(mid, localEnd, plane, 200, 0);
                var cutterEnd = CreateBidirectionalExtrudedBody(plane, loop, fwdE, backE);
                var (fwdS, backS) = DetermineExtrusionDirection(mid, localStart, plane, 0, 200);
                var cutterStart = CreateBidirectionalExtrudedBody(plane, loop, fwdS, backS);
                if (cutterEnd == null || cutterStart == null)
                {
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
            catch (Exception)
            {
                return (null, null);
            }
        }

        // Rebuilds the component body so only one half is modified for the joint, preserving the far half and recombining the result.
        public static void ResetHalfForJoint(
            Component component,
            string connectionSide,
            bool extendProfile,
            Vector _ignoredLocalUp,
            List<DesignCurve> allCurves,
            List<Component> selectedComponents
        )
        {
            if (component?.Template == null)
            {
                return;
            }
            var template = component.Template;

            // 1) Grab & shift the construction segment
            var dc = template.Curves.OfType<DesignCurve>().FirstOrDefault();
            if (dc?.Shape is not CurveSegment seg)
            {
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
                return;
            }

            // 4) preserve non-corner half
            bool preserveEnd = connectionSide == "HalfStart";
            Body preserved = preserveEnd ? halfEnd.Copy() : halfStart.Copy();

            // 5) delete old halves & existing ExtrudedProfile
            foreach (var name in new[] { "HalfStart", "HalfEnd", "ExtrudedProfile" })
            {
                var old = template.Bodies.FirstOrDefault(b => b.Name == name);
                if (old != null)
                {
                    old.Delete();
                }
            }

            DesignBody.Create(template, "preservedHalf", preserved)
                     .Layer = template.Document.GetLayer("Frames");

            // 6) regenerate full profile
            if (extendProfile)
            {
                ResetComponentGeometryAndExtend(
                    new List<Component> { component },
                    localUp,
                    allCurves,
                    connectionSide
                );
            }
            else
            {
                ResetComponentGeometryOnly(
                    new List<Component> { component },
                    localUp,
                    allCurves
                );
            }

            if (template.Bodies.All(b => b.Name != "ExtrudedProfile"))
            {
                return;
            }

            // 7) second split
            var (halfStart2, halfEnd2) = SplitBodyAtMidpoint(component, localUp);
            if (halfStart2 == null || halfEnd2 == null)
            {
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
            var finalDb = DesignBody.Create(template, "ExtrudedProfile", presCopy);
            finalDb.Layer = template.Document.GetLayer("Frames");
        }
    }
}
