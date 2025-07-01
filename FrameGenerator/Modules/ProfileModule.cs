using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Drawing;
using AESCConstruct25.FrameGenerator.Modules.Profiles;
using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System.Windows.Media;
using Color = System.Drawing.Color;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Matrix = SpaceClaim.Api.V242.Geometry.Matrix;

namespace AESCConstruct25.FrameGenerator.Modules
{
    public static class ProfileModule
    {
        public static void ExtrudeProfile(
            Window window,
            string profileType,
            ITrimmedCurve selectedCurve,
            bool isHollow,
            Dictionary<string, string> profileData,
            double offsetX,
            double offsetY,
            Vector localUp,
            string dxfFilePath = "",
            List<ITrimmedCurve> dxfContours = null,
            Component reuseComponent = null
        )
        {
            //Logger.Log($"ExtrudeProfile started for {profileType}");

            if (reuseComponent != null)
            {
                //Logger.Log($" → reuseComponent provided: '{reuseComponent.Name}' with existing Placement:\n    {MatrixToString(reuseComponent.Placement)}");
            }

            var doc = window?.Document;
            if (doc?.MainPart == null)
            {
                //Logger.Log("ERROR - No valid document");
                return;
            }

            // 1) Compute world-space start/end and sweep direction
            Point worldStart = selectedCurve.EndPoint;
            Point worldEnd = selectedCurve.StartPoint;
            double length = (worldEnd - worldStart).Magnitude;
            Direction sweepZ = (worldEnd - worldStart).Direction;

            // 2) Build world-space frame axes (X,Y,Z)
            Vector zv = sweepZ.ToVector();
            Direction xDir = Math.Abs(Vector.Dot(zv, localUp)) > 0.99
                           ? Vector.Create(1, 0, 0).Direction
                           : Vector.Cross(localUp, zv).Direction;
            Direction yDir = Direction.Cross(sweepZ, xDir);

            Vector shift = xDir.ToVector() * offsetX
             + yDir.ToVector() * offsetY;
            Point shiftedStart = worldStart + shift;

            // now build the world‐frame on which the component will sit
            Frame worldFrame = Frame.Create(shiftedStart, xDir, yDir);
            Matrix compPlacement = Matrix.CreateMapping(worldFrame);
            //Logger.Log($"Computed new compPlacement (world-frame) for extrusion:\n    {MatrixToString(compPlacement)}");

            // 4) In-part (local) always sketch on pure XY plane
            Frame localFrame;
            if (reuseComponent != null)
            {
                localFrame = Frame.Create(
                    selectedCurve.StartPoint,
                    Direction.Create(1, 0, 0),
                    Direction.Create(0, 1, 0)
                );
            }
            else
            {
                localFrame = Frame.Create(
                    Point.Create(0, 0, 0),
                    Direction.Create(1, 0, 0),
                    Direction.Create(0, 1, 0)
                );
            }
            Plane localPlane = Plane.Create(localFrame);

            // 5) Build profile loops in that localPlane
            string[] args = GetArgs(profileType, profileData);
            var profile = ProfileBase.CreateProfile(
                profileType, args, isHollow,
                offsetX, offsetY,
                dxfFilePath, dxfContours
            );
            if (profile == null)
            {
                //Logger.Log($"ERROR – CreateProfile returned null for {profileType}");
                return;
            }
            var outerLoop = profile.GetProfileCurves(localPlane).ToList();
            var innerLoop = isHollow
                          ? profile.GetInnerProfile(localPlane).ToList()
                          : null;

            // 6) Extrude along local +Z by 'length'
            var outerBody = Body.ExtrudeProfile(new Profile(localPlane, outerLoop), length);
            if (outerBody == null)
            {
                //Logger.Log($"ERROR – failed to extrude outer for {profileType}");
                return;
            }
            if (innerLoop?.Count > 0)
            {
                var innerBody = Body.ExtrudeProfile(new Profile(localPlane, innerLoop), length);
                if (innerBody != null)
                    outerBody.Subtract(innerBody);
            }

            // 7) Insert (or reuse) the component, then (optionally) place it
            var comp = reuseComponent
                     ?? CreateComponent(
                            doc,
                            profileType,
                            profileData,
                            selectedCurve,
                            length,
                            isHollow,
                            offsetX,
                            offsetY,
                            dxfContours,
                            dxfFilePath
                        );

            // 8) Wipe out any old "ExtrudedProfile" bodies
            foreach (var old in comp.Template.Bodies
                                             .Where(b => b.Name == "ExtrudedProfile")
                                             .ToList())
                old.Delete();

            // 9) If we're creating a brand-new component, set its Placement now;
            //    but if we're re-using an existing component, do NOT touch its Placement.
            if (reuseComponent == null)
            {
                //Logger.Log($"Before setting Placement, new component '{comp.Name}' had default:\n    {MatrixToString(comp.Placement)}");
                comp.Placement = compPlacement;
                //Logger.Log($"After setting Placement, component '{comp.Name}' now has:\n    {MatrixToString(comp.Placement)}");
            }
            else
            {
                //Logger.Log($"ReuseComponent = '{comp.Name}', preserving placement:\n    {MatrixToString(comp.Placement)}");
            }

            // 10) Finally add the fresh body at local (0,0,0)→+Z
            var framesLayer = doc.GetLayer("Frames");
            var db = DesignBody.Create(comp.Template, "ExtrudedProfile", outerBody);
            db.Layer = framesLayer;

            // 11) Log bounding-box center in local and world to verify
            Point localCenter = BodyCenterLocal(outerBody);
            //Logger.Log($"Created ExtrudedProfile body: center local = {PointToString(localCenter)}, world = {PointToString(comp.Placement * localCenter)}");

            //Logger.Log($"Successfully extruded {profileType} at origin/+Z; " +
                       //(reuseComponent == null
                       //    ? "placed in world via Component.Placement."
                       //    : "body remains in component’s existing local frame."));
        }

        private static string[] GetArgs(string profileType, Dictionary<string, string> pd)
        {
            string g(string k) => pd != null && pd.TryGetValue(k, out var v) ? v : "0";
            switch (profileType)
            {
                case "Circular": return new[] { g("D"), g("t") };
                case "L": return new[] { g("a"), g("b"), g("t"), g("r2"), g("r1") };
                case "Rectangular": return new[] { g("h"), g("w"), g("t"), g("r1"), g("r2") };
                case "H": return new[] { g("h"), g("w"), g("s"), g("t"), g("r1"), g("r2") };
                case "T": return new[] { g("h"), g("w"), g("s"), g("t"), g("r1"), g("r2"), g("r3") };
                case "U": return new[] { g("h"), g("w"), g("s"), g("t"), g("r1"), g("r2") };
                case "DXF": return new[] { g("w") };
                default: return pd?.Values.ToArray() ?? Array.Empty<string>();
            }
        }


        // --- your existing CreateComponent (unchanged) ---
        private static Component CreateComponent(
            Document doc,
            string profileType,
            Dictionary<string, string> profileData,
            ITrimmedCurve pathCurve,
            double len,
            bool isHollow,
            double offsetX,
            double offsetY,
            List<ITrimmedCurve> dxfContourVal,
            string dxfFilePath)
        {
            // clone so we don’t mutate caller’s dictionary
            profileData = profileData != null
                ? new Dictionary<string, string>(profileData)
                : new Dictionary<string, string>();

            //// build unique part/component name
            //string baseName = profileType;
            //if (profileData.Count > 0)
            //    baseName += "_" + string.Join("_", profileData.Values);
            //baseName += "_" + ((int)(len * 1000)).ToString();

            //string partName = baseName;
            //int n = 1;
            //while (doc.Parts.Any(p => p.Name == partName))
            //    partName = baseName + "-" + (n++);

            Part part = Part.Create(doc, "Temp");
            Component comp = Component.Create(doc.MainPart, part);
            CompNameHelper.SetNameAndLength(
                comp,
                profileType,
                profileData,
                len
            );

            Layer framesLayer = doc.GetLayer("Frames");

            // …if it wasn’t there, create it
            if (framesLayer == null)
            {
                Color myCustomColor = ColorTranslator.FromHtml("#007AFF");
                framesLayer = Layer.Create(doc, "Frames", myCustomColor);
            }

            // store metadata
            CustomPartProperty.Create(part, "Type", profileType);
            CustomPartProperty.Create(part, "Hollow", isHollow.ToString().ToLower());
            CustomPartProperty.Create(part, "offsetX", offsetX);
            CustomPartProperty.Create(part, "offsetY", offsetY);

            // helper to parse a numeric parameter
            double GetNum(string key)
            {
                if (!profileData.TryGetValue(key, out var raw))
                    return 0.0;

                // 1) replace any comma decimal-separator with a dot
                var normalized = raw.Replace(',', '.');

                // 2) parse with InvariantCulture
                if (double.TryParse(
                        normalized,
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out var v
                    ))
                    return v;

                return 0.0;
            }

            // ensure we always have a "w" value (in mm) for joint logic
            if (!profileData.ContainsKey("w"))
            {
                double w = 0.0;
                switch (profileType)
                {
                    case "Circular":
                        w = GetNum("D");
                        break;
                    case "L":
                        w = GetNum("b");
                        break;
                    case "Rectangular":
                    case "H":
                    case "U":
                    case "T":
                        w = GetNum("w");
                        break;
                    case "DXF":
                        var (widthMm, heightMm) = DXFImportHelper.GetDXFSize(dxfContourVal);
                        w = widthMm * 1000.0;
                        CustomPartProperty.Create(part, "DXFPath", dxfFilePath);
                        break;
                    default:
                        // fallback: try any of the common keys
                        w = GetNum("w");
                        if (w == 0.0) w = GetNum("b");
                        if (w == 0.0) w = GetNum("D");
                        break;
                }
                //Logger.Log($"Width from profiledata: {w}.");
                if (w > 0.0)
                {
                    // store in profileData in millimetres
                    profileData["w"] = w.ToString(CultureInfo.InvariantCulture);
                }
            }

            CustomPartProperty.Create(part, "AESC_Construct", true);

            // write out all Construct_ properties
            foreach (var kv in profileData)
                CustomPartProperty.Create(part, "Construct_" + kv.Key, kv.Value);

            // record the construction line for joints
            if (pathCurve != null)
            {
                // draw from (0,0,0) to (len,0,0)
                var flatSeg = CurveSegment.Create(
                    Point.Create(-offsetX, -offsetY, 0),
                    Point.Create(-offsetX, -offsetY, len)
                );
                var dc = DesignCurve.Create(part, flatSeg);
                dc.Name = "ConstructCurve";
                dc.Layer = framesLayer;
                dc.SetVisibility(null, true);
            }

            foreach (var dc in comp.Content.Curves.OfType<DesignCurve>())
            {
                //Logger.Log($"Adding layer to dc");
                dc.Layer = framesLayer;
            }

            return comp;
        }

        private static Point BodyCenterLocal(Body body)
        {
            // Returns the local‐space center of the body’s bounding box.
            var bb = body.GetBoundingBox(Matrix.Identity, tight: true);
            return bb.Center;
        }
    }
}
