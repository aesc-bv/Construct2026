using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace AESCConstruct25.FrameGenerator.Commands
{
    class CompareCommand
    {
        public static void CompareSimple()
        {
            var window = Window.ActiveWindow;
            if (window == null)
                return;

            Part mainPart = window.Document.MainPart;

            // 1) Collect unique master bodies from the model
            var bodies = mainPart
              .GetDescendants<IDesignBody>()
              .Select(idb => (IDesignBody)idb.Master)
              .Distinct()
              .ToList();

            // Build a lookup: each body -> its original component name
            var originalNames = bodies.ToDictionary(
              idb => idb,
              idb => ((DesignBody)idb.Master).Parent.Name
            );

            // ─── Strip any existing "(n)" suffix before applying new numbering ───
            //foreach (var idb in originalNames.Keys.ToList())
            //{
            //    originalNames[idb] = Regex.Replace(
            //        originalNames[idb],
            //        @"\(\d+\)$",
            //        ""
            //    );
            //}
            foreach (var idb in bodies)
            {
                // grab the Component
                //var comp = ((DesignBody)idb.Master).Parent;               // this is a Component
                //var part = comp.Template;                                 // the underlying Part

                IPart comp = idb.Parent;
                Part part = comp.Master;
                var comp1 = mainPart
                  .GetDescendants<IComponent>()
                  .OfType<Component>()
                  .FirstOrDefault(c => c.Template == part);

                if (comp1 == null)
                    continue;

                // compute bbox dims
                var bb = idb.Shape.GetBoundingBox(Matrix.Identity, tight: true);
                var size = bb.MaxCorner - bb.MinCorner;
                double lengthM = size.Z;       // meters
                double widthMm = size.X * 1000.0;    // mm
                double heightMm = size.Y * 1000.0;    // mm

                // pull Type + existing Construct_… props
                string profileType = part.CustomProperties["Type"].Value.ToString();
                var profileData = part.CustomProperties
                    .Where(kv => kv.Key.StartsWith("Construct_") && kv.Key != "Construct_Length")
                    .ToDictionary(
                        kv => kv.Key.Substring("Construct_".Length),
                        kv => kv.Value.ToString()
                    );

                // override w/h so helper sees the real numbers
                profileData["w"] = widthMm.ToString(CultureInfo.InvariantCulture);
                profileData["h"] = heightMm.ToString(CultureInfo.InvariantCulture);

                // **this** wipes any old “(n)” off the name and writes the new length
                CompNameHelper.SetNameAndLength(comp1, profileType, profileData, lengthM);
            }
            //var existingNames = new HashSet<string>(
            //    mainPart.GetDescendants<IComponent>().Select(c => c.Name)
            //);
            foreach (var masterBody in bodies)
            {
                // the Part that owns this master
                var part = ((DesignBody)masterBody).Parent;
                // use the IDesignBody.GetBoundingBox overload with tight = true
                var bb = masterBody.Shape.GetBoundingBox(Matrix.Identity, tight: true);
                var size = bb.MaxCorner - bb.MinCorner;
                double length = Math.Max(size.X, Math.Max(size.Y, size.Z));
                string lengthString = length.ToString("F4", CultureInfo.InvariantCulture);

                //Logger.Log($"bbox length = {length}");

                if (part.CustomProperties.ContainsKey("Construct_Tubelength"))
                    part.CustomProperties["Construct_Tubelength"].Value = lengthString;
                else
                    CustomPartProperty.Create(part, "Construct_Tubelength", lengthString);
            }
            // ─────────────────────────────────────────────────────────────────────

            double tol = 1e-12;
            var existingNames = new HashSet<string>(
              mainPart.GetDescendants<IComponent>()
                      .Select(c => c.Name)
            );
            var renameCounts = new Dictionary<IDesignBody, int>();
            var matchedBodies = new HashSet<IDesignBody>();
            var refinedPairs = new List<Tuple<IDesignBody, IDesignBody>>();
            int nrMatches = 0;

            WriteBlock.ExecuteTask("CompareConstruct", () =>
            {
                for (int i = 0; i < bodies.Count; i++)
                {
                    var idb1 = bodies[i];
                    IPart ipart = idb1.Parent;
                    Part partMaster = ipart.Master;

                    // Get length data and rename part
                    Box bbox = idb1.Master.Shape.GetBoundingBox(Matrix.Identity, true);
                    double length = bbox.Size.Z * 1000;

                    // To do: Save length as custom property and Rename part according to name setting.

                    if (!idb1.Parent.Master.CustomProperties.TryGetValue("AESC_Construct", out _))
                        continue;
                    if (matchedBodies.Contains(idb1))
                        continue;

                    if (!renameCounts.ContainsKey(idb1))
                        renameCounts[idb1] = 1;

                    for (int j = i + 1; j < bodies.Count; j++)
                    {
                        var idb2 = bodies[j];
                        if (matchedBodies.Contains(idb2))
                            continue;
                        if (!idb2.Parent.Master.CustomProperties.TryGetValue("AESC_Construct", out _))
                            continue;

                        // a) Skip if both design bodies reference the same Shape instance
                        if (ReferenceEquals(idb1.Shape, idb2.Shape))
                            continue;

                        // b) Quick check: volume and surface area must be within tolerance
                        if (Math.Abs(idb1.Shape.Volume - idb2.Shape.Volume) > tol ||
                            Math.Abs(idb1.Shape.SurfaceArea - idb2.Shape.SurfaceArea) > tol)
                            continue;

                        // c) Try up to 8 union tests: 4 around Z, then 180° about Y + 4 around Z
                        bool foundMatch = false;

                        // Loop two phases: phase 0 = no Y-rotation, phase 1 = apply 180° Y-rotation
                        for (int phase = 0; phase < 2 && !foundMatch; phase++)
                        {
                            for (int rotZ = 0; rotZ < 4 && !foundMatch; rotZ++)
                            {
                                // Copy original shapes so we don't modify live geometry
                                Body s1 = idb1.Master.Shape.Copy();
                                Body s2 = idb2.Master.Shape.Copy();

                                // Translate both shapes so their bounding boxes sit on the XY plane
                                Box bb1 = s1.GetBoundingBox(Matrix.Identity, true);
                                Box bb2 = s2.GetBoundingBox(Matrix.Identity, true);
                                if (Math.Abs(bb1.MinCorner.Z) > tol)
                                    s1.Transform(Matrix.CreateTranslation(Vector.Create(0, 0, -bb1.MinCorner.Z)));
                                if (Math.Abs(bb2.MinCorner.Z) > tol)
                                    s2.Transform(Matrix.CreateTranslation(Vector.Create(0, 0, -bb2.MinCorner.Z)));

                                // If phase 1, apply a 180° rotation around the Y-axis at the shape center
                                if (phase == 1)
                                {
                                    // center of s1's bounding box
                                    var center = s1.GetBoundingBox(Matrix.Identity, true).Center;
                                    var matRotY = Matrix.CreateRotation(
                                        Line.Create(center, Direction.DirY),
                                        Math.PI // 180° in radians
                                    );
                                    s2.Transform(matRotY);
                                }

                                // Apply Z-axis rotation by rotZ * 90° around the origin
                                if (rotZ > 0)
                                {
                                    var axis = Line.Create(Point.Origin, Direction.DirZ);
                                    var angleDeg = 90 * rotZ;
                                    var matRotZ = Matrix.CreateRotation(
                                        axis,
                                        angleDeg * Math.PI / 180.0
                                    );
                                    s2.Transform(matRotZ);
                                }


                                //DesignBody.Create(mainPart, $"s2_{phase}_{rotZ}", s2.Copy());

                                // Unite and test volume equality
                                s1.Unite(s2);
                                //DesignBody.Create(mainPart, $"s1_Unite_{phase}_{rotZ}", s1.Copy());

                                if (Math.Abs(s1.Volume - idb1.Shape.Volume) <= tol)
                                {
                                    // Match found: rename the component holding idb2
                                    var comp2 = ((DesignBody)idb2.Master).Parent;

                                    // Use the original (pre-rename) name for suffixing
                                    string baseName = originalNames[idb1];

                                    // Determine the next available suffix index
                                    int k = renameCounts[idb1]++;
                                    string newName;
                                    do
                                    {
                                        newName = string.Format("{0}({1})", baseName, k);
                                        k++;
                                    }
                                    while (existingNames.Contains(newName));

                                    // Reserve and apply the new unique name
                                    existingNames.Add(newName);
                                    comp2.Name = newName;

                                    // Mark idb2 as consumed so it won't be reused
                                    matchedBodies.Add(idb2);
                                    foundMatch = true;
                                }

                                // Clean up temporary bodies
                                s1.Dispose();
                                s2.Dispose();
                            }
                        }

                        // d) If no union/rotation succeeded, queue this pair for reporting
                        if (!foundMatch)
                            refinedPairs.Add(Tuple.Create(idb1, idb2));
                        else
                            nrMatches++;
                    }
                }
            });

            //Application.ReportStatus($"Found {nrMatches} duplicates", StatusMessageType.Information, null);

            // 4) Display any remaining distinct pairs via message boxes
            //foreach (var pair in refinedPairs)
            //{
            //    MessageBox.Show(
            //      $"{pair.Item1.Master.Parent.Name} ≈ {pair.Item2.Master.Parent.Name}  " +
            //      $"(V:{pair.Item1.Shape.Volume:F6}/{pair.Item2.Shape.Volume:F6}, " +
            //      $"A:{pair.Item1.Shape.SurfaceArea:F6}/{pair.Item2.Shape.SurfaceArea:F6})"
            //    );
            //}
        }
    }
}
