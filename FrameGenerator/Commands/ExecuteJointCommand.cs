/*
 ExecuteJointCommand orchestrates creation and management of frame joints between selected components.

 It:
 - Validates selection for trim and joint operations.
 - Builds connected component pairs and applies different JointBase implementations (miter, straight, T, trim).
 - Handles geometry splitting / resetting for joints and executes profile cut-outs.
 - Provides helper predicates to detect connectivity and point-on-segment relationships in world space.
*/

using AESCConstruct2026.FrameGenerator.Modules;
using AESCConstruct2026.FrameGenerator.Modules.Joints;
using AESCConstruct2026.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Linq;
using Application = SpaceClaim.Api.V242.Application;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;

namespace AESCConstruct2026.FrameGenerator.Commands
{
    class ExecuteJointCommand
    {
        public const string CommandName = "AESCConstruct2026.ExecuteJoint";

        // Validates that the current selection contains at least one target profile and one cutter face, otherwise reports a warning.
        private static bool ValidateTrimSelectionOrAlert(Window window)
        {
            var ctx = window?.ActiveContext;
            var sel = ctx?.Selection;

            if (sel == null || sel.Count == 0)
            {
                Application.ReportStatus("Trim requires: select at least one profile (body/component/edge) and one face for the cutter.", StatusMessageType.Warning, null);
                return false;
            }

            // Count faces and detect at least one valid target
            int faceCount = 0;
            bool hasTarget = false;

            foreach (var o in sel)
            {
                if (o is IDesignFace) { faceCount++; continue; }
                if (o is IDesignBody || o is IComponent || o is IDesignEdge || o is IDesignCurve)
                    hasTarget = true;
            }

            if (!hasTarget || faceCount < 1)
            {
                Application.ReportStatus("Trim requires: at least one profile (body/component/edge of target body) and one face for the cutter.\nPlease adjust your selection and try again.", StatusMessageType.Warning, null);
                return false;
            }

            return true;
        }

        // Entry point for applying a joint of the given type to selected components, handling trimming, pairing, and joint execution.
        public static void ExecuteJoint(Window window, double spacing, string jointType, bool updateBOM)
        {
            // 1) Gather selected components
            List<Component> selectedComponents = JointSelectionHelper.GetSelectedComponents(window);
            if (selectedComponents.Count < 1)
            {
                Application.ReportStatus("Select at least two components to apply a joint.", StatusMessageType.Warning, null);
                return;
            }

            string msgTemplate = $"Applying {jointType} joint ";

            // 2) Pre‐cache each component’s single construction CurveSegment
            var curveMap = new Dictionary<Component, CurveSegment>();
            foreach (var comp in selectedComponents)
            {
                CurveSegment seg = null;
                foreach (var dc in comp.Template.Curves)              // no LINQ
                    if (dc is DesignCurve design && design.Shape is CurveSegment cs)
                    {
                        seg = cs;
                        break;
                    }
                curveMap[comp] = seg;
            }

            // 3) Collect all curves (used by ResetHalfForJoint)
            List<DesignCurve> allCurves = new List<DesignCurve>();
            foreach (var c in selectedComponents)
                foreach (var dc in c.Template.Curves)
                    if (dc is DesignCurve design)
                        allCurves.Add(design);

            var halves = new Dictionary<Component, (Body Start, Body End)>();
            var alreadyProcessed = new HashSet<(Component, Component)>();
            const double tol = 1e-6;

            // 4) Trim joint is special
            if (jointType == "Trim")
            {
                if (!ValidateTrimSelectionOrAlert(window))
                    return;

                Component compA = selectedComponents.FirstOrDefault();
                if (compA == null) return;

                JointBase trimJoint = CreateJoint(jointType);
                trimJoint.Execute(compA, null, spacing, null, null);
            }
            else if (jointType == "CutOut")
            {
                ExecuteCutOut(window);
                return;
            }
            else
            {
                // 5) Build all valid A–B pairs
                var workItems = GetConnectedPairs(selectedComponents, jointType);
                int total = workItems.Count();
                if (total > 0)
                {
                    using (var PT = ProgressTracker.Create(workItems.Count()))
                    {
                        int i = 0;
                        foreach (var (compA, compB) in workItems)
                        {
                            i++;
                            PT.Progress = i;
                            PT.Message = msgTemplate + i + "/" + total;

                            if (alreadyProcessed.Contains((compA, compB)) ||
                                alreadyProcessed.Contains((compB, compA)))
                                continue;

                            //6) Connection tests
                            bool aStartConnected = ArePointsConnected(compA, compB, "start");
                            bool aEndConnected = ArePointsConnected(compA, compB, "end");
                            bool bStartConnected = ArePointsConnected(compB, compA, "start");
                            bool bEndConnected = ArePointsConnected(compB, compA, "end");

                            // 7) Lookup cached CurveSegments
                            var locSegA = curveMap[compA];
                            var locSegB = curveMap[compB];
                            if (locSegA == null || locSegB == null)
                                continue;

                            // 8) Local directions
                            var dA = (locSegA.EndPoint - locSegA.StartPoint).Direction.ToVector();
                            var dB = (locSegB.EndPoint - locSegB.StartPoint).Direction.ToVector();

                            // 9) Stable localUp (hemisphere flip)
                            var rawUp = Vector.Cross(dA, dB);
                            Vector localUp = rawUp.Magnitude > tol
                                           ? rawUp.Direction.ToVector()
                                           : Vector.Create(0, 1, 0);
                            Vector WY = Vector.Create(0, 1, 0),
                                   WX = Vector.Create(1, 0, 0),
                                   WZ = Vector.Create(0, 0, 1);
                            if (Math.Abs(Vector.Dot(localUp, WY)) > tol && Vector.Dot(localUp, WY) < 0) localUp = -localUp;
                            else if (Math.Abs(Vector.Dot(localUp, WX)) > tol && Vector.Dot(localUp, WX) < 0) localUp = -localUp;
                            else if (Math.Abs(Vector.Dot(localUp, WZ)) > tol && Vector.Dot(localUp, WZ) < 0) localUp = -localUp;

                            // 10) Instantiate joint
                            JointBase joint = CreateJoint(jointType);

                            if (jointType == "T")
                            {
                                // ensure compB split once
                                if (!halves.ContainsKey(compB))
                                {
                                    var (sB, eB) = JointModule.SplitBodyAtMidpoint(compB, localUp);
                                    if (sB != null && eB != null)
                                        halves[compB] = (sB, eB);
                                }

                                // lift A/B into world
                                Point wA0 = compA.Placement * locSegA.StartPoint;
                                Point wA1 = compA.Placement * locSegA.EndPoint;
                                Point wB0 = compB.Placement * locSegB.StartPoint;
                                Point wB1 = compB.Placement * locSegB.EndPoint;

                                bool cutStart = IsPointOnSegment(wB0, wA0, wA1);
                                bool cutEnd = IsPointOnSegment(wB1, wA0, wA1);

                                string jointHalfName = cutStart ? "HalfStart"
                                                       : cutEnd ? "HalfEnd"
                                                       : null;
                                if (jointHalfName == null) continue;

                                JointModule.ResetHalfForJoint(
                                    compB,
                                    jointHalfName,
                                    extendProfile: true,
                                    localUp,
                                    allCurves,
                                    new List<Component> { compA, compB }
                                );

                                Body resetHalf = null;
                                foreach (var b in compB.Template.Bodies)
                                    if (b.Name == "ExtrudedProfile")
                                    {
                                        resetHalf = b.Shape;
                                        break;
                                    }
                                if (resetHalf == null) continue;

                                joint.Execute(compA, compB, spacing, null, resetHalf);
                            }
                            else
                            {
                                if (aStartConnected || aEndConnected)
                                    // Non-T joints: ensure both halves exist
                                    if (!halves.ContainsKey(compA))
                                    {
                                        var (sA, eA) = JointModule.SplitBodyAtMidpoint(compA, localUp);
                                        if (sA != null && eA != null)
                                            halves[compA] = (sA, eA);
                                    }
                                if (!halves.ContainsKey(compB))
                                {
                                    var (sB, eB) = JointModule.SplitBodyAtMidpoint(compB, localUp);
                                    if (sB != null && eB != null)
                                        halves[compB] = (sB, eB);
                                }

                                var (aStart, aEnd) = halves[compA];
                                var (bStart, bEnd) = halves[compB];

                                if (aEndConnected && bStartConnected)
                                {
                                    JointModule.ResetHalfForJoint(compA, "HalfEnd", true, localUp, allCurves, new List<Component> { compA, compB });
                                    JointModule.ResetHalfForJoint(compB, "HalfStart", true, localUp, allCurves, new List<Component> { compA, compB });
                                    joint.Execute(compA, compB, spacing, aEnd, bStart);
                                }
                                else if (aStartConnected && bEndConnected)
                                {
                                    JointModule.ResetHalfForJoint(compA, "HalfStart", true, localUp, allCurves, new List<Component> { compA, compB });
                                    JointModule.ResetHalfForJoint(compB, "HalfEnd", true, localUp, allCurves, new List<Component> { compA, compB });
                                    joint.Execute(compA, compB, spacing, aStart, bEnd);
                                }
                                else if (aStartConnected && bStartConnected)
                                {
                                    JointModule.ResetHalfForJoint(compA, "HalfStart", true, localUp, allCurves, new List<Component> { compA, compB });
                                    JointModule.ResetHalfForJoint(compB, "HalfStart", true, localUp, allCurves, new List<Component> { compA, compB });
                                    joint.Execute(compA, compB, spacing, aStart, bStart);
                                }
                                else if (aEndConnected && bEndConnected)
                                {
                                    JointModule.ResetHalfForJoint(compA, "HalfEnd", true, localUp, allCurves, new List<Component> { compA, compB });
                                    JointModule.ResetHalfForJoint(compB, "HalfEnd", true, localUp, allCurves, new List<Component> { compA, compB });
                                    joint.Execute(compA, compB, spacing, aEnd, bEnd);
                                }
                            }

                            alreadyProcessed.Add((compA, compB));
                        }
                    }
                }
                else
                {
                    Application.ReportStatus("No joint pair found, check line connections.", StatusMessageType.Warning, null);
                }
            }
            if (updateBOM == true)
                ExportCommands.ExportBOM(Window.ActiveWindow, update: true);
            else
                CompareCommand.CompareSimple();
        }

        // Restores the original (un-jointed) geometry for the selected components by re-extruding their profiles.
        public static void RestoreGeometry(Window window)
        {
            List<Component> selectedComponents = JointSelectionHelper.GetSelectedComponents(window);
            if (selectedComponents.Count == 0)
            {
                Application.ReportStatus("Select at least one component to restore geometry.", StatusMessageType.Warning, null);
                return;
            }

            JointModule.ResetComponentGeometryOnly(selectedComponents);
        }

        // Attempts to restore previously created joints between selected components by resetting the appropriate halves.
        public static void RestoreJoint(Window window)
        {
            WriteBlock.ExecuteTask("RestoreJoint", () =>
            {
                // 1) Gather selection
                var selected = JointSelectionHelper.GetSelectedComponents(window);
                if (selected.Count < 2)
                {
                    Application.ReportStatus("Select at least two components to restore joints.", StatusMessageType.Warning, null);
                    return;
                }

                const double tol = 1e-6;

                // 2) Cache each component’s CurveSegment
                var curveMap = new Dictionary<Component, CurveSegment>();
                foreach (var comp in selected)
                {
                    CurveSegment seg = null;
                    foreach (var dc in comp.Template.Curves)
                        if (dc is DesignCurve design && design.Shape is CurveSegment cs)
                        {
                            seg = cs;
                            break;
                        }
                    curveMap[comp] = seg;
                }

                // 3) Collect all DesignCurves once
                var allCurves = selected
                    .SelectMany(c => c.Template.Curves.OfType<DesignCurve>())
                    .ToList();

                // 4) Build “reset” actions
                var resetActions = new List<(Component comp, string half, Vector localUp, List<Component> partners)>();
                var seenPairs = new HashSet<(Component, Component)>();

                foreach (var (a, b) in GetConnectedPairs(selected, "RestoreJoint"))
                {
                    if (seenPairs.Contains((b, a))) continue;
                    seenPairs.Add((a, b));

                    bool aStart = ArePointsConnected(a, b, "start");
                    bool bStart = ArePointsConnected(b, a, "start");
                    bool aEnd = ArePointsConnected(a, b, "end");
                    bool bEnd = ArePointsConnected(b, a, "end");
                    if (!((aStart || aEnd) && (bStart || bEnd))) continue;

                    string halfA = aStart ? "HalfStart" : "HalfEnd";
                    string halfB = bStart ? "HalfStart" : "HalfEnd";

                    // compute stable localUp
                    var segA = curveMap[a];
                    var segB = curveMap[b];
                    if (segA == null || segB == null) continue;
                    var dA = (segA.EndPoint - segA.StartPoint).Direction.ToVector();
                    var dB = (segB.EndPoint - segB.StartPoint).Direction.ToVector();
                    Vector rawUp = Vector.Cross(dA, dB);
                    Vector localUp = rawUp.Magnitude > tol
                                   ? rawUp.Direction.ToVector()
                                   : Vector.Create(0, 1, 0);
                    if (Vector.Dot(localUp, Vector.Create(0, 1, 0)) < 0
                     || Vector.Dot(localUp, Vector.Create(1, 0, 0)) < 0
                     || Vector.Dot(localUp, Vector.Create(0, 0, 1)) < 0)
                        localUp = -localUp;

                    resetActions.Add((a, halfA, localUp, new List<Component> { a, b }));
                    resetActions.Add((b, halfB, localUp, new List<Component> { a, b }));
                }

                // 5) Execute resets (always extended profile)
                foreach (var (comp, half, up, partners) in resetActions)
                {
                    try
                    {
                        JointModule.ResetHalfForJoint(
                            comp,
                            half,
                            extendProfile: false,   // <— always regenerate via the “extended” path
                            _ignoredLocalUp: up,   // passed but ignored when extendProfile=true
                            allCurves,
                            partners
                        );
                    }
                    catch (Exception ex)
                    {
                        Application.ReportStatus($"Joint error: {ex.Message}", StatusMessageType.Error, null);
                    }
                }
            });
        }

        // Cuts the last selected profile component by all previously selected profiles and cleans up the resulting bodies.
        public static void ExecuteCutOut(Window window)
        {
            var comps = JointSelectionHelper.GetSelectedComponents(window);
            if (comps == null || comps.Count < 2)
            {
                Application.ReportStatus("Select at least two profiles. The last selected will be cut by the others.", StatusMessageType.Warning, null);
                return;
            }

            var target = comps.Last();
            var cutters = comps.Take(comps.Count - 1).Where(c => c != target).ToList();
            if (cutters.Count == 0) return;

            WriteBlock.ExecuteTask("CutOut", () =>
            {
                var part = target.Template;
                if (part == null) return;

                // Snapshot original target bodies and main-body metrics (fallback: largest volume)
                var preBodies = part.Bodies.Where(b => b.Shape != null && b.Shape.Volume > 0).ToList();
                var preMain = preBodies.OrderByDescending(b => b.Shape.Volume).FirstOrDefault();

                Point preCenter = Point.Origin;
                double preVol = 0.0;
                double lenScale = 1.0;

                if (preMain != null)
                {
                    var bb = preMain.Shape.GetBoundingBox(Matrix.Identity, true);
                    preCenter = bb.Center;
                    preVol = preMain.Shape.Volume;
                    var size = bb.MaxCorner - bb.MinCorner;
                    lenScale = Math.Max(1e-6, Math.Abs(size.X) + Math.Abs(size.Y) + Math.Abs(size.Z));
                }

                // Subtract all cutters (using copies transformed into target-local)
                foreach (var cutterComp in cutters)
                {
                    try
                    {
                        if (cutterComp?.Template == null) continue;

                        var bodies = cutterComp.Template.Bodies
                            .Where(b => b.Shape != null && b.Shape.Volume > 0)
                            .ToList();
                        if (bodies.Count == 0) continue;

                        Matrix toTargetLocal = target.Placement.Inverse * cutterComp.Placement;

                        foreach (var db in bodies)
                        {
                            Body copy = db.Shape.Copy();          // keep original cutter intact
                            copy.Transform(toTargetLocal);         // place into target's local frame
                            JointModule.SubtractCutter(target, copy);
                            copy.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log("ExecuteCutOut: cutter subtraction failed: " + ex.ToString());
                    }
                }

                // Cleanup: retain the post-cut body that is closest to the original main body.
                var postBodies = part.Bodies.Where(b => b.Shape != null && b.Shape.Volume > 0).ToList();
                if (postBodies.Count <= 1) return;

                DesignBody keeper = null;

                if (preMain != null)
                {
                    // Score by normalized center distance + volume difference ratio
                    double bestScore = double.MaxValue;

                    foreach (var b in postBodies)
                    {
                        var bb = b.Shape.GetBoundingBox(Matrix.Identity, true);
                        var center = bb.Center;
                        double dist = (center - preCenter).Magnitude / lenScale;

                        double v = b.Shape.Volume;
                        double volDiff = (preVol > 0.0 || v > 0.0)
                            ? Math.Abs(v - preVol) / Math.Max(preVol, v)
                            : 0.0;

                        double score = dist + volDiff; // simple composite score
                        if (score < bestScore)
                        {
                            bestScore = score;
                            keeper = b;
                        }
                    }
                }
                else
                {
                    // Fallback: keep largest volume if we had no pre-main
                    keeper = postBodies.OrderByDescending(b => b.Shape.Volume).FirstOrDefault();
                }

                if (keeper != null)
                {
                    foreach (var extra in postBodies.Where(b => !object.ReferenceEquals(b, keeper)).ToList())
                    {
                        try { extra.Delete(); } catch (Exception ex) { Logger.Log("ExecuteCutOut: cleanup delete failed: " + ex.ToString()); }
                    }
                }
            });
        }

        // Factory that maps jointType strings to concrete JointBase implementations (with Miter as default).
        private static JointBase CreateJoint(string jointType)
        {
            switch (jointType)
            {
                case "Miter": return new MiterJoint();
                case "Straight": return new StraightJoint();
                case "Straight2": return new StraightJoint2();
                case "T": return new TJoint();
                case "None": return new NoneJoint();
                case "Trim": return new TrimJoint();
                default: return new MiterJoint();
            }
        }

        // Returns true when component B’s start or end world point lies on component A’s swept world segment.
        private static bool IsTConnected(Component a, Component b)
        {
            var segA = a.Template.Curves.OfType<DesignCurve>().FirstOrDefault()?.Shape as CurveSegment;
            var segB = b.Template.Curves.OfType<DesignCurve>().FirstOrDefault()?.Shape as CurveSegment;
            if (segA == null || segB == null)
                return false;

            Point wA0 = a.Placement * segA.StartPoint;
            Point wA1 = a.Placement * segA.EndPoint;
            Point wB0 = b.Placement * segB.StartPoint;
            Point wB1 = b.Placement * segB.EndPoint;

            return IsPointOnSegment(wB0, wA0, wA1) ||
                   IsPointOnSegment(wB1, wA0, wA1);
        }

        // Enumerates all component pairs that are either physically connected or valid T-connections (depending on jointType).
        private static IEnumerable<(Component, Component)> GetConnectedPairs(
            List<Component> comps,
            string jointType
        )
        {
            for (int i = 0; i < comps.Count; i++)
            {
                for (int j = i + 1; j < comps.Count; j++)
                {
                    var a = comps[i];
                    var b = comps[j];

                    bool phys = ArePhysicallyConnected(a, b);
                    bool tconn = jointType == "T" && IsTConnected(a, b);

                    if (phys || tconn)
                        yield return (a, b);
                }
            }
        }

        // Checks if a given start or end point of component A coincides (in world space) with either endpoint of component B.
        private static bool ArePointsConnected(Component a, Component b, string side)
        {
            var segA = a.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
            var segB = b.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
            if (segA == null || segB == null) return false;
            Point pA = (side == "start" ? segA.StartPoint : segA.EndPoint);
            // lift into world
            var wA = a.Placement * pA;
            var wB0 = b.Placement * segB.StartPoint;
            var wB1 = b.Placement * segB.EndPoint;
            const double tol = 1e-6;
            return (wA - wB0).Magnitude < tol || (wA - wB1).Magnitude < tol;
        }

        // Returns true when any endpoint of A’s segment coincides with any endpoint of B’s segment in world space.
        private static bool ArePhysicallyConnected(Component a, Component b)
        {
            var segA = a.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
            var segB = b.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
            if (segA == null || segB == null) return false;
            var a0 = a.Placement * segA.StartPoint;
            var a1 = a.Placement * segA.EndPoint;
            var b0 = b.Placement * segB.StartPoint;
            var b1 = b.Placement * segB.EndPoint;
            const double tol = 1e-6;
            return (a0 - b0).Magnitude < tol ||
                   (a0 - b1).Magnitude < tol ||
                   (a1 - b0).Magnitude < tol ||
                   (a1 - b1).Magnitude < tol;
        }

        // Tests whether a point lies on a finite segment (within tolerance) using colinearity and projection checks.
        public static bool IsPointOnSegment(Point p, Point segStart, Point segEnd)
        {
            Vector ab = segEnd - segStart;
            Vector ap = p - segStart;

            double cross = Vector.Cross(ab, ap).Magnitude;
            if (cross > 1e-6)
                return false; // Not colinear

            double dot = Vector.Dot(ap, ab);
            if (dot < -1e-6)
                return false; // Before segment start

            if (dot > Vector.Dot(ab, ab) + 1e-6)
                return false; // Beyond segment end

            return true;
        }
    }
}
