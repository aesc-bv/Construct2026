/*
 ProfileModule generates and manages frame/profile components along a selected path in SpaceClaim.

 It is responsible for:
 - Building a stable 3D frame from a path curve, offset and local up vector.
 - Creating the correct cross-section geometry (built-in or DXF/CSV based) and extruding it along the path.
 - Creating or reusing SpaceClaim components, assigning metadata + layers and constructing hidden helper curves.
 - Providing utilities for profile argument extraction, axis selection, naming and custom property management.
*/

using AESCConstruct2026.FrameGenerator.Modules.Profiles;
using AESCConstruct2026.FrameGenerator.Utilities;
using AESCConstruct2026.Properties;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Body = SpaceClaim.Api.V242.Modeler.Body;
using Color = System.Drawing.Color;
using Document = SpaceClaim.Api.V242.Document;
using Frame = SpaceClaim.Api.V242.Geometry.Frame;
using Matrix = SpaceClaim.Api.V242.Geometry.Matrix;
using Point = SpaceClaim.Api.V242.Geometry.Point;

namespace AESCConstruct2026.FrameGenerator.Modules
{
    public static class ProfileModule
    {
        // Builds the profile cross-section, extrudes it along the selected curve and inserts or reuses a component in the document.
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
            Component reuseComponent = null,
            string csvProfileString = null,
            List<DesignCurve> createdCurves = null
        )
        {
            var doc = window?.Document;
            if (doc?.MainPart == null)
            {
                return;
            }

            // 1) Compute world-space start/end and sweep direction
            Point worldStart = selectedCurve.StartPoint;
            Point worldEnd = selectedCurve.EndPoint;

            double length = (worldEnd - worldStart).Magnitude;
            Direction sweepZ = (worldEnd - worldStart).Direction;

            // 2) Build world-space frame axes (X, Y, Z) – deterministic, plane-agnostic
            Vector zv = sweepZ.ToVector();

            // If localUp ≈ z, choose the world axis most orthogonal to z for x
            Direction xDir;
            if (Math.Abs(Vector.Dot(zv, localUp)) > 0.99)
            {
                xDir = MostOrthogonalWorldAxis(zv).Direction;
            }
            else
            {
                xDir = Vector.Cross(localUp, zv).Direction;
            }
            Direction yDir = Direction.Cross(sweepZ, xDir);

            // Apply offsets in that XY
            Vector shift = xDir.ToVector() * offsetX + yDir.ToVector() * offsetY;
            Point shiftedStart = worldStart + shift;

            // Build the world frame and placement
            Frame worldFrame = Frame.Create(shiftedStart, xDir, yDir);
            Matrix compPlacement = Matrix.CreateMapping(worldFrame);

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
                            dxfFilePath,
                            csvProfileString,
                            createdCurves
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
                comp.Placement = compPlacement;
            }

            // 10) Finally add the fresh body at local (0,0,0)→+Z
            var db = DesignBody.Create(comp.Template, "ExtrudedProfile", outerBody);
            string frameColor = Settings.Default.FrameColor ?? "";
            if (!string.IsNullOrWhiteSpace(frameColor))
            {
                var framesLayer = doc.GetLayer("Frames");
                if (framesLayer != null)
                    db.Layer = framesLayer;
            }
        }

        // Returns the world axis vector that is most orthogonal to the given vector v.
        private static Vector MostOrthogonalWorldAxis(Vector v)
        {
            var axes = new[] { Vector.Create(0, 0, 1), Vector.Create(0, 1, 0), Vector.Create(1, 0, 0) };
            Vector best = axes[0];
            double bestScore = 1.0 - Math.Abs(Vector.Dot(v, best));
            for (int i = 1; i < axes.Length; i++)
            {
                double score = 1.0 - Math.Abs(Vector.Dot(v, axes[i]));
                if (score > bestScore) { bestScore = score; best = axes[i]; }
            }
            return best;
        }

        // Maps the profile type + profileData dictionary into the ordered argument list expected by the profile classes.
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
                case "U": return new[] { g("h"), g("w"), g("s"), g("t"), g("r1"), g("r2"), g("Name") };
                case "DXF": return new[] { g("w") };
                default: return pd?.Values.ToArray() ?? Array.Empty<string>();
            }
        }

        // Returns the logical profile name from profileData["Name"], or an empty string if not present.
        private static string GetProfileName(Dictionary<string, string> profileData)
        {
            if (profileData != null
             && profileData.TryGetValue("Name", out var nm)
             && !string.IsNullOrWhiteSpace(nm))
            {
                return nm;
            }
            else
            {
                return string.Empty;
            }
        }


        // Creates a new Part/Component for the profile, assigns metadata, layers and helper curves, and returns the component.
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
            string dxfFilePath,
            string csvProfileString,
            List<DesignCurve> createdCurves = null
        )
        {
            // clone so we don’t mutate caller’s dictionary
            profileData = profileData != null
                ? new Dictionary<string, string>(profileData)
                : new Dictionary<string, string>();

            string dataString = string.Join("_",
            profileData
              .OrderBy(kv => kv.Key)
              .Select(kv => $"{kv.Key}{kv.Value}"));

            // 2) Express length in millimetres (integer)
            string lengthMm = ((int)(len * 1000)).ToString();

            // 3) Base partName
            string baseName = $"{profileType}_{dataString}_{lengthMm}";

            // 4) Make sure it’s unique in this document
            string partName = baseName;
            int suffix = 1;
            while (doc.MainPart
                    .GetChildren<Part>()
                    .Any(p => p.Name.Equals(partName, StringComparison.OrdinalIgnoreCase)))
            {
                partName = $"{baseName}_{suffix++}";
            }

            // 5) Finally create it
            Part part = Part.Create(doc, partName);
            Component comp = Component.Create(doc.MainPart, part);
            CompNameHelper.SetNameAndLength(
                comp,
                profileType,
                profileData,
                len
            );

            string frameColorHex = Settings.Default.FrameColor ?? "";
            if (!string.IsNullOrWhiteSpace(frameColorHex))
            {
                try
                {
                    Color parsedColor = ColorTranslator.FromHtml(frameColorHex);
                    Layer framesLayer = doc.GetLayer("Frames");
                    if (framesLayer == null)
                    {
                        framesLayer = Layer.Create(doc, "Frames", parsedColor);
                    }
                    else
                    {
                        framesLayer.SetColor(null, parsedColor);
                    }
                }
                catch { /* invalid hex – skip layer creation */ }
            }

            Layer hiddenFramesLayer = doc.GetLayer("Construct (hidden)");
            if (hiddenFramesLayer == null)
            {
                var gray = ColorTranslator.FromHtml("#9ea0a1");
                hiddenFramesLayer = Layer.Create(doc, "Construct (hidden)", gray);
                hiddenFramesLayer.SetVisible(null, false);
            }
            string rawName = GetProfileName(profileData) ?? "";
            string profileName = Regex.Replace(rawName.Trim(), @"\s+", "_");
            // store metadata
            CreateCustomProperty(part, "Type", profileType);
            CreateCustomProperty(part, "Hollow", isHollow.ToString().ToLower());
            CreateCustomProperty(part, "offsetX", offsetX);
            CreateCustomProperty(part, "offsetY", offsetY);
            CreateCustomProperty(part, "Name", profileName);

            // helper to parse a numeric parameter
            // Parses a numeric value from profileData[key] using invariant culture, returning 0.0 on failure.
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
                        var (dxfW, dxfH) = DXFImportHelper.GetDXFSize(dxfContourVal);
                        w = dxfW * 1000.0;
                        if (dxfH > 0)
                            profileData["h"] = (dxfH * 1000.0).ToString(CultureInfo.InvariantCulture);
                        CreateCustomProperty(part, "DXFPath", dxfFilePath);
                        break;
                    case "CSV":
                        var (csvWidth, csvHeight) = DXFImportHelper.GetDXFSize(dxfContourVal);
                        w = csvWidth * 1000.0;
                        if (csvHeight > 0)
                            profileData["h"] = (csvHeight * 1000.0).ToString(CultureInfo.InvariantCulture);
                        CreateCustomProperty(part, "RawCSV", csvProfileString);
                        break;
                    default:
                        // fallback: try any of the common keys
                        w = GetNum("w");
                        if (w == 0.0) w = GetNum("b");
                        if (w == 0.0) w = GetNum("D");
                        break;
                }
                if (w > 0.0)
                {
                    // store in profileData in millimetres
                    profileData["w"] = w.ToString(CultureInfo.InvariantCulture);
                }
            }

            CustomPartProperty.Create(part, "AESC_Construct", true);

            // write out all Construct_ properties
            foreach (var kv in profileData)
            {
                CreateCustomProperty(part, "Construct_" + kv.Key, kv.Value);
            }

            // record the construction line for joints
            if (pathCurve != null)
            {
                //var window = Window.ActiveWindow;
                //var ctx = window?.ActiveContext as IAppearanceContext;
                // draw from (0,0,0) to (len,0,0)
                var flatSeg = CurveSegment.Create(
                    Point.Create(-offsetX, -offsetY, 0),
                    Point.Create(-offsetX, -offsetY, len)
                );
                var dc = DesignCurve.Create(part, flatSeg);
                dc.Name = "ConstructCurve";
                dc.Layer = hiddenFramesLayer;
                dc.SetVisibility(null, false);

                if (createdCurves != null)
                    createdCurves.Add(dc);
            }

            return comp;
        }

        // Creates or updates a custom property on the given Part, storing newValue as a string.
        private static void CreateCustomProperty(Part part, string key, object newValue)
        {
            // 1) Convert newValue to a string
            string valueString;
            switch (newValue)
            {
                case bool b:
                    valueString = b.ToString().ToLowerInvariant();
                    break;
                case IFormattable formattable:
                    valueString = formattable.ToString(null, CultureInfo.InvariantCulture);
                    break;
                default:
                    valueString = newValue?.ToString() ?? string.Empty;
                    break;
            }

            // 2) Look up the existing KV pair by its key
            var existingKV = part.CustomProperties
                .FirstOrDefault(kv => kv.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

            // 3) If found (i.e. Key != null), update; otherwise create
            if (existingKV.Key != null)
            {
                existingKV.Value.Value = valueString;
            }
            else
            {
                CustomPartProperty.Create(part, key, valueString);
            }
        }
    }
}
