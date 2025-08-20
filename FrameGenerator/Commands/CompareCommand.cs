using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

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
                    MessageBox.Show("No active window found.", "Compare Bodies", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show("No bodies found in the main part.", "Compare Bodies", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Build a lookup: each body -> its original component name
                var originalNames = bodies.ToDictionary(
                  idb => idb,
                  idb => ((DesignBody)idb.Master).Parent.Name
                );

                foreach (var idb in bodies)
                {
                    try
                    {
                        IPart comp = idb.Parent;
                        Part part = comp.Master;
                        var comp1 = mainPart
                          .GetDescendants<IComponent>()
                          .OfType<Component>()
                          .FirstOrDefault(c => c.Template == part);

                        if (comp1 == null)
                            continue;

                        var bb = idb.Shape.GetBoundingBox(Matrix.Identity, tight: true);
                        var size = bb.MaxCorner - bb.MinCorner;
                        double lengthM = size.Z;
                        double widthMm = size.X * 1000.0;
                        double heightMm = size.Y * 1000.0;

                        string profileType = null;
                        if (!part.CustomProperties.ContainsKey("Type"))
                        {
                            continue;
                        }
                        profileType = part.CustomProperties["Type"].Value?.ToString() ?? "";

                        var profileData = new Dictionary<string, string>();
                        foreach (var kv in part.CustomProperties.Where(kv => kv.Key.StartsWith("Construct_") && kv.Key != "Construct_Length"))
                        {
                            var v = kv.Value?.Value?.ToString();
                            if (!string.IsNullOrEmpty(v))
                                profileData[kv.Key.Substring("Construct_".Length)] = v;
                        }

                        profileData["w"] = widthMm.ToString(CultureInfo.InvariantCulture);
                        profileData["h"] = heightMm.ToString(CultureInfo.InvariantCulture);

                        CompNameHelper.SetNameAndLength(comp1, profileType, profileData, lengthM);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error processing body for name/length:\n{ex.Message}", "Compare Bodies", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                foreach (var masterBody in bodies)
                {
                    try
                    {
                        var part = ((DesignBody)masterBody).Parent;
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
                        MessageBox.Show($"Error setting tube length:\n{ex.Message}", "Compare Bodies", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

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

                        Box bbox = idb1.Master.Shape.GetBoundingBox(Matrix.Identity, true);
                        double length = bbox.Size.Z * 1000;

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

                            if (ReferenceEquals(idb1.Shape, idb2.Shape))
                                continue;

                            if (Math.Abs(idb1.Shape.Volume - idb2.Shape.Volume) > tol ||
                                Math.Abs(idb1.Shape.SurfaceArea - idb2.Shape.SurfaceArea) > tol)
                                continue;

                            bool foundMatch = false;

                            for (int phase = 0; phase < 2 && !foundMatch; phase++)
                            {
                                for (int rotZ = 0; rotZ < 4 && !foundMatch; rotZ++)
                                {
                                    Body s1 = null, s2 = null;
                                    try
                                    {
                                        s1 = idb1.Master.Shape.Copy();
                                        s2 = idb2.Master.Shape.Copy();

                                        Box bb1 = s1.GetBoundingBox(Matrix.Identity, true);
                                        Box bb2 = s2.GetBoundingBox(Matrix.Identity, true);
                                        if (Math.Abs(bb1.MinCorner.Z) > tol)
                                            s1.Transform(Matrix.CreateTranslation(Vector.Create(0, 0, -bb1.MinCorner.Z)));
                                        if (Math.Abs(bb2.MinCorner.Z) > tol)
                                            s2.Transform(Matrix.CreateTranslation(Vector.Create(0, 0, -bb2.MinCorner.Z)));

                                        if (phase == 1)
                                        {
                                            var center = s1.GetBoundingBox(Matrix.Identity, true).Center;
                                            var matRotY = Matrix.CreateRotation(
                                                Line.Create(center, Direction.DirY),
                                                Math.PI
                                            );
                                            s2.Transform(matRotY);
                                        }

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

                                        s1.Unite(s2);

                                        if (Math.Abs(s1.Volume - idb1.Shape.Volume) <= tol)
                                        {
                                            var comp2 = ((DesignBody)idb2.Master).Parent;
                                            string baseName = originalNames[idb1];
                                            int k = renameCounts[idb1]++;
                                            string newName;
                                            do
                                            {
                                                newName = string.Format("{0}({1})", baseName, k);
                                                k++;
                                            }
                                            while (existingNames.Contains(newName));

                                            existingNames.Add(newName);
                                            comp2.Name = newName;
                                            matchedBodies.Add(idb2);
                                            foundMatch = true;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($"Error during shape comparison:\n{ex.Message}", "Compare Bodies", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    finally
                                    {
                                        s1?.Dispose();
                                        s2?.Dispose();
                                    }
                                }
                            }

                            if (!foundMatch)
                                refinedPairs.Add(Tuple.Create(idb1, idb2));
                            else
                                nrMatches++;
                        }
                    }
                });

                // Optionally, show a summary message
                //MessageBox.Show($"Found {nrMatches} duplicate bodies.", "Compare Bodies", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Optionally, display any remaining distinct pairs
                //foreach (var pair in refinedPairs)
                //{
                //    MessageBox.Show(
                //      $"{pair.Item1.Master.Parent.Name} ≈ {pair.Item2.Master.Parent.Name}  " +
                //      $"(V:{pair.Item1.Shape.Volume:F6}/{pair.Item2.Shape.Volume:F6}, " +
                //      $"A:{pair.Item1.Shape.SurfaceArea:F6}/{pair.Item2.Shape.SurfaceArea:F6})",
                //      "Distinct Pair",
                //      MessageBoxButtons.OK,
                //      MessageBoxIcon.Information
                //    );
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error during comparison:\n{ex.Message}", "Compare Bodies", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
