/*
 JointModule encapsulates joint-related operations for frame components.

 It is responsible for:
 - Rebuilding profile geometry from stored metadata (DXF/CSV/built-in) using ProfileModule.
 - Optionally extending profiles along their path for joint construction and splitting them into halves.
 - Building cutter frames and bidirectional cutter bodies, and applying boolean subtraction.
 - Providing joint-specific reset logic (per half/side) used by miter and straight joints.
*/

using AESCConstruct2026.FrameGenerator.Utilities;
using AESCConstruct2026.Localization;
using AESCConstruct2026.Properties;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Drawing;
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
                        NumberParsing.TryParseUserInput(val, out offsetX);
                        break;
                    case "offsetY":
                        NumberParsing.TryParseUserInput(val, out offsetY);
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
                // SpaceClaim API units are metres; origStart/origEnd below are metric.
                // 0.2 m = 200 mm of extra length for joint construction. (Was 200.0,
                // which extended each end by 200 m and exploded the body ~1000x.)
                const double extendAmount = 0.2; // metres (= 200 mm)

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
                // PART 1 GUARD: never feed a degenerate body to the ACIS Unite.
                if (!IsUsableBody(bodyForward, "cbeb.fwd") || !IsUsableBody(bodyBackward, "cbeb.back"))
                {
                    return IsUsableBody(bodyForward, "cbeb.fwd.fallback") ? bodyForward
                         : IsUsableBody(bodyBackward, "cbeb.back.fallback") ? bodyBackward
                         : null;
                }
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
            DesignBody target = FindProfileBody(part);
            if (target == null)
            {
                return;
            }

            // PART 1 GUARD: never feed a degenerate body to the ACIS Subtract.
            // A degenerate / zero-volume body causes a native AV in SpaACIS.dll
            // that no managed try/catch can recover from.
            if (!IsUsableBody(target.Shape, "subtract.target") || !IsUsableBody(cutter, "subtract.cutter"))
            {
                L.Status("Frame_Joint_Msg_InvalidBody", StatusMessageType.Warning);
                return;
            }

            // Insert cutter into component’s Part (local)
            DesignBody temp = DesignBody.Create(part, "TempCutter", cutter.Copy());

            temp.Layer = GetOrCreateFramesLayer(part.Document);
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

                // fetch or merge existing profile body (Part-name rule; see FindProfileBody)
                Body original = FindProfileBody(component.Template)?.Shape.Copy();
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

                // PART 1 GUARD: never feed a degenerate body to the ACIS Intersect.
                if (!IsUsableBody(halfEnd, "split.halfEnd.src") || !IsUsableBody(cutterEnd, "split.cutterEnd"))
                {
                    L.Status("Frame_Joint_Msg_InvalidBody", StatusMessageType.Warning);
                    return (null, null);
                }
                halfEnd.Intersect(new[] { cutterEnd.Copy() });

                if (!IsUsableBody(halfStart, "split.halfStart.src") || !IsUsableBody(cutterStart, "split.cutterStart"))
                {
                    L.Status("Frame_Joint_Msg_InvalidBody", StatusMessageType.Warning);
                    return (null, null);
                }
                halfStart.Intersect(new[] { cutterStart.Copy() });

                // delete old halves
                foreach (var name in new[] { "HalfStart", "HalfEnd" })
                    component.Template.Bodies
                             .FirstOrDefault(b => b.Name == name)
                             ?.Delete();

                // create new DesignBodies
                var framesLayer = GetOrCreateFramesLayer(component.Template.Document);
                DesignBody.Create(component.Template, "HalfStart", halfStart)
                         .Layer = framesLayer;
                DesignBody.Create(component.Template, "HalfEnd", halfEnd)
                         .Layer = framesLayer;

                return (halfStart, halfEnd);
            }
            catch (Exception ex)
            {
                Logger.Log("SplitBodyAtMidpoint failed: " + ex.ToString());
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

            // Reference scale = true profile length from the construction curve.
            // A correctly-built body's bbox diagonal is ~this; a unit-bug-exploded
            // body is ~1000x this. Used for scale-aware IsUsableBody checks.
            double refProfileDiag = (seg.EndPoint - seg.StartPoint).Magnitude;

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

            // 5) delete old halves & legacy-named profile body. The literal
            //    "ExtrudedProfile" is intentionally kept to purge legacy-document
            //    bodies; the current Part-named profile body (and everything else)
            //    is wiped by the regen in step 6 (ProfileModule.ExtrudeProfile).
            foreach (var name in new[] { "HalfStart", "HalfEnd", "ExtrudedProfile" })
            {
                var old = template.Bodies.FirstOrDefault(b => b.Name == name);
                if (old != null)
                {
                    old.Delete();
                }
            }

            DesignBody.Create(template, "preservedHalf", preserved)
                     .Layer = GetOrCreateFramesLayer(template.Document);

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

            // Post-regen the profile body is named after the Part (see FindProfileBody /
            // ProfileModule.cs:169-172), not the legacy literal "ExtrudedProfile".
            if (FindProfileBody(template) == null)
            {
                return;
            }

            // FIX 4b: the original `preserved` Body was a transient Copy() invalidated by
            // the regen body-wipe (preservedVol came back as -1 at R5a). Re-acquire it
            // from the "preservedHalf" DesignBody that FIX 4a kept alive across the wipe,
            // so `preserved` is a valid, document-anchored body for the rest of the method.
            var preservedDb = template.Bodies.FirstOrDefault(b => b.Name == "preservedHalf");
            if (preservedDb?.Shape != null)
            {
                preserved = preservedDb.Shape.Copy();
            }

            // 7) second split
            var (halfStart2, halfEnd2) = SplitBodyAtMidpoint(component, localUp);
            if (halfStart2 == null || halfEnd2 == null)
            {
                return;
            }

            // 8) isolate corner half
            Body corner = preserveEnd ? halfStart2 : halfEnd2;

            // RELOCATED PART 1 GUARD: the original R6 guard ran AFTER the Copy()/Delete
            // region; the native AV fires there. Validate the source bodies BEFORE the
            // first ACIS call (.Copy()) in this region, scale-aware against the true
            // profile length.
            if (!IsUsableBody(preserved, "reset.preserved.src", refProfileDiag) ||
                !IsUsableBody(corner, "reset.corner.src", refProfileDiag))
            {
                L.Status("Frame_Joint_Msg_InvalidBody", StatusMessageType.Warning);
                return;
            }

            // 9) copy & delete temps
            Body presCopy = preserved.Copy();
            Body corCopy = corner.Copy();
            // NOTE: the legacy literal "ExtrudedProfile" is intentionally kept in this
            // delete loop so old-named bodies from legacy documents are also purged.
            // The current-named profile body (Part name) is removed separately below.
            foreach (var name in new[] { "HalfStart", "HalfEnd", "ExtrudedProfile", "preservedHalf" })
                template.Bodies.FirstOrDefault(b => b.Name == name)?.Delete();
            FindProfileBody(template)?.Delete();

            // 10) unite into final
            // PART 1 GUARD: validate operands before the ACIS Unite.
            if (!IsUsableBody(presCopy, "reset.preserved") || !IsUsableBody(corCopy, "reset.corner"))
            {
                L.Status("Frame_Joint_Msg_InvalidBody", StatusMessageType.Warning);
                Body keep = IsUsableBody(presCopy, "reset.preserved.fallback") ? presCopy
                          : IsUsableBody(corCopy, "reset.corner.fallback") ? corCopy
                          : null;
                if (keep == null)
                {
                    return;
                }
                var nameKeep = !string.IsNullOrWhiteSpace(template.Name) ? template.Name : "ExtrudedProfile";
                DesignBody.Create(template, nameKeep, keep).Layer = GetOrCreateFramesLayer(template.Document);
                return;
            }
            presCopy.Unite(new[] { corCopy });
            // Re-create the profile body under the Part-name rule (see FindProfileBody /
            // ProfileModule.cs:169-172) so subsequent joints and BOM/STEP export find it.
            var finalName = !string.IsNullOrWhiteSpace(template.Name) ? template.Name : "ExtrudedProfile";
            var finalDb = DesignBody.Create(template, finalName, presCopy);
            finalDb.Layer = GetOrCreateFramesLayer(template.Document);
        }

        /// <summary>
        /// Resolves the profile solid body of a Construct part.
        ///
        /// Naming contract: <see cref="ProfileModule"/> (ProfileModule.cs:169-172) names
        /// the extruded profile body after the owning Part (e.g. "Rect_20x20_426"), and
        /// only falls back to the literal "ExtrudedProfile" when the Part has no name.
        /// All joint/export code must therefore resolve the body through THIS method
        /// rather than hard-coding the literal "ExtrudedProfile", otherwise a normally
        /// created profile body is never found (the lookup misses, the body appears
        /// "missing", and the caller crashes on the empty result).
        ///
        /// Resolution order:
        ///   1. Part-named, NOT a joint scratch/cutter body
        ///      (HalfStart/HalfEnd/preservedHalf/TempCutter), Volume &gt; 1e-9;
        ///      when several match, the SMALLEST sane volume wins (the true profile
        ///      is far smaller than a cutter/merged/exploded body that may transiently
        ///      share the Part name);
        ///   2. body literally named "ExtrudedProfile" (legacy / unnamed-part documents);
        ///   3. the single real solid body, excluding joint scratch bodies and
        ///      zero-volume bodies.
        /// Returns null if no suitable body exists.
        ///
        /// This still resolves the normal Part-named profile body (no Bug-A /
        /// KeyNotFound regression), and the rule-2 legacy fallback still finds
        /// old documents whose body is literally named "ExtrudedProfile".
        /// </summary>
        private static bool IsScratchBodyName(string n)
            => n == "HalfStart" || n == "HalfEnd" || n == "preservedHalf" || n == "TempCutter";

        public static DesignBody FindProfileBody(Part part)
        {
            if (part == null) return null;

            // Rule 1: Part-named, non-scratch, sane volume; smallest volume wins.
            var byPartName = part.Bodies
                .Where(b => b.Name == part.Name
                         && !IsScratchBodyName(b.Name)
                         && b.Shape != null && b.Shape.Volume > 1e-9)
                .OrderBy(b => b.Shape.Volume)
                .FirstOrDefault();
            if (byPartName != null) return byPartName;

            // Rule 2: legacy literal "ExtrudedProfile" (unnamed-part or older documents).
            var byLegacy = part.Bodies.FirstOrDefault(b => b.Name == "ExtrudedProfile");
            if (byLegacy != null) return byLegacy;

            // Rule 3: any real solid body, excluding joint scratch/cutter bodies.
            return part.Bodies
                .Where(b => !IsScratchBodyName(b.Name)
                         && b.Shape != null && b.Shape.Volume > 1e-9)
                .OrderBy(b => b.Shape.Volume)
                .FirstOrDefault();
        }

        /// <summary>
        /// Returns a finite bounding-box diagonal length (m) for a body, or -1 if it
        /// cannot be measured. Used by IsUsableBody for degenerate-geometry validation.
        /// </summary>
        private static double BboxDiagonal(Body b)
        {
            try
            {
                if (b == null) return -1.0;
                var bb = b.GetBoundingBox(Matrix.Identity, tight: true);
                var d = (bb.MaxCorner - bb.MinCorner).Magnitude;
                return (double.IsNaN(d) || double.IsInfinity(d)) ? -1.0 : d;
            }
            catch
            {
                return -1.0;
            }
        }

        /// <summary>Volume of a body, or -1 if it cannot be read (null / disposed).</summary>
        private static double SafeVol(Body b)
        {
            try { return b == null ? -1.0 : b.Volume; }
            catch { return -1.0; }
        }

        /// <summary>
        /// Sign- and selection-order-robust cutter direction for end-cut joints
        /// (Straight / None / any "keep the long run, remove the overlap past the
        /// joint" case).
        ///
        /// Returns (forwardDistance, backwardDistance) for
        /// CreateBidirectionalExtrudedBody (arg1 => +planeLocal.DirZ,
        /// arg2 => -planeLocal.DirZ): the LONG slab on the waste side, the SHORT
        /// allowance on the keep side.
        ///
        /// Why this is robust where (farLocal-origin).DirZ was not: that earlier
        /// test reads a value that is NEAR ZERO and sign-unstable when the cut
        /// plane is built nearly parallel to the member axis ("end cut parallel
        /// to selection 2"), because BuildDebugCutterFrameAndLoop's DirZ sign
        /// flips with perp/up-vector/selection order. Here we instead align the
        /// (well-defined, ±) plane normal with the member's OWN construction axis
        /// (joint -> far, never degenerate for a real member). If that alignment
        /// is itself degenerate (cut ~parallel to the member), we fall back to a
        /// body-centroid half-space test that does not depend on the member's
        /// run direction. Caller passes the rebuilt profile body for the fallback.
        /// </summary>
        public static (double forward, double backward) PickEndCutDirection(
            Plane planeLocal,
            CurveSegment rawSegLocal,
            bool endConnected,
            Body memberBodyLocal,
            double longLen,
            double shortLen,
            string diagTag)
        {
            const double degenDot = 1e-3; // |normal . memberAxis| below this = parallel

            Vector nZ = planeLocal.Frame.DirZ.ToVector();
            Point origin = planeLocal.Frame.Origin;

            Point connectedLocal = endConnected ? rawSegLocal.EndPoint : rawSegLocal.StartPoint;
            Point farLocal = endConnected ? rawSegLocal.StartPoint : rawSegLocal.EndPoint;

            Vector keepVec = farLocal - connectedLocal;     // joint -> far (keeper side)
            double keepMag = keepVec.Magnitude;
            if (keepMag < 1e-9)
            {
                return (longLen, shortLen);
            }
            Vector keepDir = keepVec / keepMag;

            double alongKeep = Vector.Dot(nZ, keepDir);

            if (Math.Abs(alongKeep) >= degenDot)
            {
                // Primary (covers perpendicular butt and all well-conditioned cuts).
                // +DirZ toward keeper => push long slab on -DirZ (backward) and vice versa.
                bool plusZIsKeeper = alongKeep > 0;
                double fwd = plusZIsKeeper ? shortLen : longLen;
                double back = plusZIsKeeper ? longLen : shortLen;
                return (fwd, back);
            }

            // Degenerate: cut plane ~parallel to member axis (the "parallel to
            // selection 2" failing case). The plane normal cannot separate
            // keep/waste along the member. Use the member body's centroid: the
            // long slab must go to the half-space NOT containing the body mass.
            if (memberBodyLocal != null && SafeVol(memberBodyLocal) > 1e-9)
            {
                var bb = memberBodyLocal.GetBoundingBox(Matrix.Identity, tight: true);
                Vector toBody = bb.Center - origin;
                bool bodyOnPlusZ = Vector.Dot(toBody, nZ) > 0;
                // body on +DirZ => keep +DirZ, push long slab on -DirZ (backward).
                return bodyOnPlusZ ? (shortLen, longLen) : (longLen, shortLen);
            }

            return (longLen, shortLen);
        }

        /// <summary>
        /// Validates that a body is safe to hand to an ACIS boolean
        /// (Subtract/Intersect/Unite). A degenerate / zero-volume / disposed body
        /// fed to the ACIS kernel causes a native Access Violation (0xc0000005 in
        /// SpaACIS.dll) that no managed try/catch can recover from — so every
        /// boolean MUST validate its operands through this method first.
        ///
        /// Criteria: non-null, Shape readable, Volume &gt; 1e-9 m^3, and a finite
        /// bounding-box diagonal within a sane range (1e-5 m .. 1e4 m).
        ///
        /// Scale-awareness: when <paramref name="referenceDiag"/> &gt; 0 is supplied
        /// (the source profile's bbox diagonal), a body whose diagonal exceeds
        /// MaxScaleFactor x referenceDiag is rejected as "wrong-scale garbage"
        /// (e.g. a 400 m body for a 0.4 m profile = a unit-bug explosion).
        /// </summary>
        private const double MaxScaleFactor = 100.0;   // body may be at most 100x the source profile

        public static bool IsUsableBody(Body b, string label = null, double referenceDiag = -1.0)
        {
            if (b == null) return false;

            double vol = SafeVol(b);
            if (vol <= 1e-9) return false;

            double diag = BboxDiagonal(b);
            if (diag < 1e-5 || diag > 1e4) return false;

            // Scale-aware rejection (only when a reference profile size is known):
            // reject a body whose bbox-diag exceeds MaxScaleFactor x the source
            // profile size (e.g. a 400 m body for a 0.4 m profile = wrong-scale
            // explosion). Tunable via MaxScaleFactor.
            if (referenceDiag > 0 && diag > MaxScaleFactor * referenceDiag) return false;

            return true;
        }

        private static Layer GetOrCreateFramesLayer(Document doc)
        {
            // Setting key: Settings.Default.FrameColor (Properties/Settings.settings,
            // System.String, default "#006d8b"). Body color is layer-driven via the
            // "Frames" layer. The new-profile path (ProfileModule.CreateComponent)
            // applies FrameColor to this layer; the joint/regen path never did, so
            // jointed bodies kept whatever (stale/teal) color the layer already had.
            // Fix: ALWAYS re-sync the existing/new Frames layer to the configured
            // FrameColor so jointed/regenerated bodies match a freshly-created
            // profile's color.
            string hex = Settings.Default.FrameColor ?? "";
            bool blank = string.IsNullOrWhiteSpace(hex);
            string effHex = blank ? "#006d8b" : hex;

            var layer = doc.GetLayer("Frames");
            if (layer != null)
            {
                // Re-apply the configured color when explicitly set (do NOT force
                // the teal fallback onto a layer the user/doc may have customised
                // when FrameColor is blank).
                if (!blank)
                {
                    try { layer.SetColor(null, ColorTranslator.FromHtml(effHex)); }
                    catch { /* invalid hex – leave existing layer color unchanged */ }
                }
                return layer;
            }

            try
            {
                layer = Layer.Create(doc, "Frames", ColorTranslator.FromHtml(effHex));
            }
            catch (Exception ex)
            {
                Logger.Log("GetOrCreateFramesLayer color parse failed: " + ex.ToString());
                layer = Layer.Create(doc, "Frames", ColorTranslator.FromHtml("#006d8b"));
            }
            return layer;
        }
    }
}
