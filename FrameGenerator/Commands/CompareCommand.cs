using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Application = SpaceClaim.Api.V242.Application;

namespace AESCConstruct25.FrameGenerator.Commands
{
    class CompareCommand
    {
        public static void CompareSimple()
        {
            try
            {
                var window = Window.ActiveWindow;
                if (window == null)
                {
                    Application.ReportStatus("No active window found.", StatusMessageType.Warning, null);
                    return;
                }

                Part mainPart = window.Document.MainPart;

                // 1) Collect unique master bodies from the model
                var bodies = mainPart
                    .GetDescendants<IDesignBody>()
                    .Select(idb => (IDesignBody)idb.Master)
                    .Distinct()
                    .ToList();

                if (bodies.Count == 0)
                {
                    Application.ReportStatus("No bodies found in the main part.", StatusMessageType.Warning, null);
                    return;
                }

                // 2) Refresh names from template + set Construct_Length (based on measured dims)
                foreach (var idb in bodies)
                {
                    try
                    {
                        IPart comp = idb.Parent;
                        Part part = comp.Master;

                        // Find the Component instance bound to this template Part
                        var comp1 = mainPart
                            .GetDescendants<IComponent>()
                            .OfType<Component>()
                            .FirstOrDefault(c => c.Template == part);

                        if (comp1 == null)
                            continue;

                        var bb = idb.Shape.GetBoundingBox(Matrix.Identity, tight: true);
                        var size = bb.MaxCorner - bb.MinCorner;

                        // Convention here: length is the largest dimension (meters),
                        // width/height are the other two (reported in mm for naming)
                        double lengthM = Math.Max(size.X, Math.Max(size.Y, size.Z));
                        double[] dims = new[] { size.X, size.Y, size.Z }.OrderByDescending(d => d).ToArray();
                        double widthMm = dims[1] * 1000.0;
                        double heightMm = dims[2] * 1000.0;

                        if (!part.CustomProperties.ContainsKey("Type"))
                            continue;

                        string profileType = part.CustomProperties["Type"].Value?.ToString() ?? "";

                        var profileData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var kv in part.CustomProperties.Where(kv => kv.Key.StartsWith("Construct_") && kv.Key != "Construct_Length"))
                        {
                            var v = kv.Value?.Value?.ToString();
                            if (!string.IsNullOrEmpty(v))
                                profileData[kv.Key.Substring("Construct_".Length)] = v;
                        }

                        // Overwrite width/height with measured values (rounded inputs)
                        profileData["w"] = widthMm.ToString(CultureInfo.InvariantCulture);
                        profileData["h"] = heightMm.ToString(CultureInfo.InvariantCulture);

                        CompNameHelper.SetNameAndLength(comp1, profileType, profileData, lengthM);
                    }
                    catch (Exception ex)
                    {
                        Application.ReportStatus($"Error processing body for name/length:\n{ex.Message}", StatusMessageType.Error, null);
                    }
                }

                // 3) Stamp Construct_Tubelength (max dimension) on each template Part
                foreach (var masterBody in bodies)
                {
                    try
                    {
                        var part = ((DesignBody)masterBody.Master).Parent;
                        var bb = masterBody.Shape.GetBoundingBox(Matrix.Identity, tight: true);
                        var size = bb.MaxCorner - bb.MinCorner;
                        double length = Math.Max(size.X, Math.Max(size.Y, size.Z));
                        string lengthString = length.ToString("F4", CultureInfo.InvariantCulture);

                        if (part.CustomProperties.ContainsKey("Construct_Tubelength"))
                            part.CustomProperties["Construct_Tubelength"].Value = lengthString;
                        else
                            CustomPartProperty.Create(part, "Construct_Tubelength", lengthString);
                    }
                    catch (Exception ex)
                    {
                        Application.ReportStatus($"Error updating Construct_Tubelength:\n{ex.Message}", StatusMessageType.Error, null);
                    }
                }

                // 4) Renaming phase with normalized rules
                double tol = 1e-12;

                // Consider only Construct parts (those that carry the AESC_Construct flag)
                var constructBodies = bodies
                    .Where(b => b?.Parent?.Master?.CustomProperties != null &&
                                b.Parent.Master.CustomProperties.TryGetValue("AESC_Construct", out _))
                    .ToList();

                // Build (Body, Part) items for renaming
                var items = constructBodies
                    .Select(b => new
                    {
                        Body = b,
                        Part = ((DesignBody)b.Master).Parent
                    })
                    .Where(x => x.Part != null && x.Body?.Shape != null)
                    .ToList();

                // Group by base name (strip trailing -n and (k))
                var groups = items
                    .GroupBy(x => BaseFromName(x.Part.Name))
                    .ToList();

                WriteBlock.ExecuteTask("CompareConstruct_Rename", () =>
                {
                    foreach (var g in groups)
                    {
                        var members = g.ToList();
                        if (members.Count == 0) continue;

                        // Partition this base-name group into equivalence classes by geometry
                        var classes = new List<List<(IDesignBody body, Part part)>>();
                        var used = new HashSet<IDesignBody>();

                        for (int i = 0; i < members.Count; i++)
                        {
                            var a = members[i];
                            if (!used.Add(a.Body))
                                continue;

                            var cls = new List<(IDesignBody, Part)>();
                            cls.Add((a.Body, a.Part));

                            for (int j = i + 1; j < members.Count; j++)
                            {
                                var b = members[j];
                                if (used.Contains(b.Body)) continue;

                                if (AreCongruentShapes(a.Body, b.Body, tol))
                                {
                                    cls.Add((b.Body, b.Part));
                                    used.Add(b.Body);
                                }
                            }

                            classes.Add(cls);
                        }

                        // Assign names inside this base group per the rules:
                        //   - If only one item in group → clean base name (strip any old suffix).
                        //   - Distinct-geometry classes beyond the first get "-n" (n = 1,2,...).
                        //   - Within a class, duplicates after the first get "(k)" (k = 1,2,...).
                        string baseName = g.Key;

                        for (int ci = 0; ci < classes.Count; ci++)
                        {
                            int hyphenIdx = (ci == 0) ? 0 : ci; // class 1 → -1, class 2 → -2, ...
                            string hyphenBase = ApplyHyphen(baseName, hyphenIdx);

                            var cls = classes[ci];

                            for (int k = 0; k < cls.Count; k++)
                            {
                                var (body, part) = cls[k];

                                string finalName = (k == 0)
                                    ? hyphenBase
                                    : ApplyParen(hyphenBase, k);

                                part.Name = finalName;
                            }
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Compare failed:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        // ───────────────────────────────────────────────────────────────────────
        // Helpers for name normalization
        // ───────────────────────────────────────────────────────────────────────

        private static string StripTrailing(string name, string pattern)
        {
            if (string.IsNullOrEmpty(name)) return string.Empty;
            return Regex.Replace(name, pattern, "");
        }

        private static string BaseFromName(string name)
        {
            // Remove "(number)" then "-number" at the very end
            var noParen = StripTrailing(name, @"\(\d+\)$");
            return StripTrailing(noParen, @"-\d+$");
        }

        private static string ApplyHyphen(string baseName, int idx)
        {
            return (idx <= 0) ? baseName : $"{baseName}-{idx}";
        }

        private static string ApplyParen(string name, int k)
        {
            return (k <= 0) ? name : $"{name}({k})";
        }

        // ───────────────────────────────────────────────────────────────────────
        // Geometry equivalence (rotation-robust)
        // 1) Quick filter by volume/surface area
        // 2) Canonicalize both: align their LONGEST axis to +Z and translate min-corner to origin
        // 3) For s2 try optional 180° Y-flip and 0/90/180/270 Z-rotations; after each, re-translate to origin
        // 4) Unite and compare volume to original
        // ───────────────────────────────────────────────────────────────────────

        private static bool AreCongruentShapes(IDesignBody idb1, IDesignBody idb2, double tol)
        {
            if (idb1 == null || idb2 == null || idb1.Shape == null || idb2.Shape == null)
                return false;

            // Quick numeric filters
            if (Math.Abs(idb1.Shape.Volume - idb2.Shape.Volume) > tol ||
                Math.Abs(idb1.Shape.SurfaceArea - idb2.Shape.SurfaceArea) > tol)
                return false;

            // Canonicalize s1 once
            using (var s1Orig = idb1.Master.Shape.Copy())
            using (var s1 = CanonicalizeToZAndOrigin(s1Orig))
            {
                var volS1 = s1.Volume;

                // For s2, try each axis as "length" candidate (in case of ties / different local frames)
                for (int axisCandidate = 0; axisCandidate < 3; axisCandidate++)
                {
                    Body s2AxisAligned = null;
                    try
                    {
                        using (var s2Orig = idb2.Master.Shape.Copy())
                        {
                            s2AxisAligned = CanonicalizeToZAndOrigin(s2Orig, axisCandidate);
                        }

                        // Skip if dimensions do not roughly match after canonicalization
                        var bb1 = s1.GetBoundingBox(Matrix.Identity, true);
                        var bb2 = s2AxisAligned.GetBoundingBox(Matrix.Identity, true);
                        var d1 = bb1.MaxCorner - bb1.MinCorner;
                        var d2 = bb2.MaxCorner - bb2.MinCorner;

                        if (!DimsClose(d1, d2, tol))
                        {
                            s2AxisAligned.Dispose();
                            continue;
                        }

                        // Try mirror (Y 180) and 0/90/180/270 around Z
                        for (int phase = 0; phase < 2; phase++)
                        {
                            Body s2Phase = null;
                            try
                            {
                                s2Phase = s2AxisAligned.Copy();

                                if (phase == 1)
                                {
                                    var center = s2Phase.GetBoundingBox(Matrix.Identity, true).Center;
                                    var rotY180 = Matrix.CreateRotation(Line.Create(center, Direction.DirY), Math.PI);
                                    s2Phase.Transform(rotY180);
                                    // Re-anchor to origin
                                    var bbp = s2Phase.GetBoundingBox(Matrix.Identity, true);
                                    s2Phase.Transform(Matrix.CreateTranslation(Vector.Create(-bbp.MinCorner.X, -bbp.MinCorner.Y, -bbp.MinCorner.Z)));
                                }

                                for (int rotZ = 0; rotZ < 4; rotZ++)
                                {
                                    Body s1Copy = null, s2Copy = null;
                                    try
                                    {
                                        s1Copy = s1.Copy();
                                        s2Copy = s2Phase.Copy();

                                        if (rotZ > 0)
                                        {
                                            var axis = Line.Create(Point.Origin, Direction.DirZ);
                                            var matRotZ = Matrix.CreateRotation(axis, (90 * rotZ) * Math.PI / 180.0);
                                            s2Copy.Transform(matRotZ);

                                            // Re-anchor s2Copy to origin after rotation
                                            var bb2r = s2Copy.GetBoundingBox(Matrix.Identity, true);
                                            s2Copy.Transform(Matrix.CreateTranslation(Vector.Create(-bb2r.MinCorner.X, -bb2r.MinCorner.Y, -bb2r.MinCorner.Z)));
                                        }

                                        // Unite and compare
                                        s1Copy.Unite(s2Copy);
                                        if (Math.Abs(s1Copy.Volume - volS1) <= tol)
                                            return true;
                                    }
                                    finally
                                    {
                                        s1Copy?.Dispose();
                                        s2Copy?.Dispose();
                                    }
                                }
                            }
                            finally
                            {
                                s2Phase?.Dispose();
                            }
                        }
                    }
                    finally
                    {
                        s2AxisAligned?.Dispose();
                    }
                }
            }

            return false;
        }

        private static bool DimsClose(Vector d1, Vector d2, double tol)
        {
            // Allow tiny tolerance and ordering difference in X/Y (since we rotate in Z anyway)
            bool closeZ = Math.Abs(d1.Z - d2.Z) <= tol;
            bool xyOrder1 = Math.Abs(d1.X - d2.X) <= tol && Math.Abs(d1.Y - d2.Y) <= tol;
            bool xyOrder2 = Math.Abs(d1.X - d2.Y) <= tol && Math.Abs(d1.Y - d2.X) <= tol;
            return closeZ && (xyOrder1 || xyOrder2);
        }

        private static Body CanonicalizeToZAndOrigin(Body src, int axisPref = -1)
        {
            var s = src.Copy();

            // Choose which world axis to map to +Z (longest extent or axisPref)
            var bb = s.GetBoundingBox(Matrix.Identity, true);
            var size = bb.MaxCorner - bb.MinCorner;

            int majorAxis = axisPref;
            if (majorAxis < 0 || majorAxis > 2)
            {
                // 0:X, 1:Y, 2:Z
                if (size.X >= size.Y && size.X >= size.Z) majorAxis = 0;
                else if (size.Y >= size.X && size.Y >= size.Z) majorAxis = 1;
                else majorAxis = 2;
            }

            // Rotate so that majorAxis -> Z
            var rot = AxisToZ(majorAxis);
            if (!rot.IsIdentity) s.Transform(rot);

            // Translate min corner to origin, so s1 and s2 are co-anchored
            var bba = s.GetBoundingBox(Matrix.Identity, true);
            var t = Matrix.CreateTranslation(Vector.Create(-bba.MinCorner.X, -bba.MinCorner.Y, -bba.MinCorner.Z));
            s.Transform(t);

            return s;
        }

        private static Matrix AxisToZ(int axisIndex)
        {
            // 0:X->Z (rotate +90° about Y)
            // 1:Y->Z (rotate -90° about X)
            // 2:Z->Z (identity)
            switch (axisIndex)
            {
                case 0:
                    return Matrix.CreateRotation(Line.Create(Point.Origin, Direction.DirY), Math.PI / 2.0);
                case 1:
                    return Matrix.CreateRotation(Line.Create(Point.Origin, Direction.DirX), -Math.PI / 2.0);
                default:
                    return Matrix.Identity;
            }
        }
    }
}
