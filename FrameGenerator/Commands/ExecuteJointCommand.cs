using SpaceClaim.Api.V251;
using SpaceClaim.Api.V251.Geometry;
using SpaceClaim.Api.V251.Modeler;
using AESCConstruct25.FrameGenerator.Modules.Joints;
using AESCConstruct25.FrameGenerator.Utilities;
using AESCConstruct25.FrameGenerator.Modules;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System;
using SpaceClaim.Geometry;
using Vector = SpaceClaim.Api.V251.Geometry.Vector;

namespace AESCConstruct25.FrameGenerator.Commands
{
    class ExecuteJointCommand
    {
        public const string CommandName = "AESCConstruct25.ExecuteJoint";

        //public static void ExecuteJoint(Window window, double spacing, string jointType)
        //{
        //    //Logger.Log("AESCConstruct25: ExecuteJoint started");

        //    List<Component> selectedComponents = JointSelectionHelper.GetSelectedComponents(window);
        //    if (selectedComponents.Count < 1)
        //    {
        //        MessageBox.Show("Select at least one component to apply a joint.", "Selection Error");
        //        return;
        //    }


        //    //Logger.Log($"Number of selected components: {selectedComponents.Count}");

        //    //RotateComponentCommand.ApplyStoredRotation(window, selectedComponents);

        //    List<DesignCurve> allCurves = selectedComponents
        //   .SelectMany(c => c.Template.Curves.OfType<DesignCurve>())
        //   .ToList();

        //    var halves = new Dictionary<Component, (Body Start, Body End)>();

        //    var alreadyProcessed = new HashSet<(Component, Component)>();

        //    // Apply joint logic for each valid connection
        //    if (jointType == "Trim")
        //    {
        //        //Logger.Log("Trim: Start.");
        //        // Trim only needs one component
        //        Component compA = selectedComponents.FirstOrDefault();
        //        if (compA == null)
        //        {
        //            //Logger.Log("Trim: No component selected.");
        //            return;
        //        }

        //        JointBase joint = CreateJoint(jointType);
        //        joint.Execute(compA, null, spacing, null, null);

        //        //Logger.Log("Trim: Execute complete.");
        //    }
        //    else
        //    {
        //        IEnumerable<(Component A, Component B)> workItems = GetConnectedPairs(selectedComponents, jointType);
        //        using (var PT = ProgressTracker.Create(workItems.Count()))
        //        {
        //            int i = 0;
        //            foreach (var (compA, compB) in workItems)
        //            {
        //                i++;
        //                //Logger.Log($"i = {i}");
        //                PT.Progress = i;
        //                PT.Message = $"Applying {jointType} joint {i}/{workItems.Count()}";

        //                if (alreadyProcessed.Contains((compA, compB)) || alreadyProcessed.Contains((compB, compA)))
        //                    continue;

        //                bool aStartConnected = ArePointsConnected(compA, compB, "start");
        //                bool aEndConnected = ArePointsConnected(compA, compB, "end");
        //                bool bStartConnected = ArePointsConnected(compB, compA, "start");
        //                bool bEndConnected = ArePointsConnected(compB, compA, "end");

        //                //Logger.Log($"connections: {aStartConnected},: {aEndConnected},: {bStartConnected},: {bEndConnected}.");

        //                var locSegA = compA.Template.Curves.First().Shape as CurveSegment;
        //                var locSegB = compB.Template.Curves.First().Shape as CurveSegment;
        //                if (locSegA == null || locSegB == null)
        //                {
        //                    //Logger.Log("Missing construction curves.");
        //                    continue;
        //                }

        //                // compute purely‐local directions:
        //                var dA = (locSegA.EndPoint - locSegA.StartPoint).Direction.ToVector();
        //                var dB = (locSegB.EndPoint - locSegB.StartPoint).Direction.ToVector();

        //                // --- stable localUp computation (local) ---
        //                var rawUp = Vector.Cross(dA, dB);
        //                Vector localUp = rawUp.Magnitude > 1e-6
        //                              ? rawUp.Direction.ToVector()
        //                              : Vector.Create(0, 1, 0);

        //                // now flip to keep a consistent hemisphere (still local)
        //                Vector WY = Vector.Create(0, 1, 0),
        //                       WX = Vector.Create(1, 0, 0),
        //                       WZ = Vector.Create(0, 0, 1);
        //                if (Math.Abs(Vector.Dot(localUp, WY)) > 1e-6 && Vector.Dot(localUp, WY) < 0) localUp = -localUp;
        //                else if (Math.Abs(Vector.Dot(localUp, WX)) > 1e-6 && Vector.Dot(localUp, WX) < 0) localUp = -localUp;
        //                else if (Math.Abs(Vector.Dot(localUp, WZ)) > 1e-6 && Vector.Dot(localUp, WZ) < 0) localUp = -localUp;


        //                // --- stable localUp computation ---
        //                //var rawUp = Vector.Cross(dA, dB);
        //                //Vector localUp = rawUp.Magnitude > 1e-6
        //                //              ? rawUp.Direction.ToVector()
        //                //              : Vector.Create(0, 1, 0);

        //                // now flip to keep a consistent hemisphere
        //                //Vector WY = Vector.Create(0, 1, 0), WX = Vector.Create(1, 0, 0), WZ = Vector.Create(0, 0, 1);
        //                //if (Math.Abs(Vector.Dot(localUp, WY)) > 1e-6 && Vector.Dot(localUp, WY) < 0) localUp = -localUp;
        //                //else if (Math.Abs(Vector.Dot(localUp, WX)) > 1e-6 && Vector.Dot(localUp, WX) < 0) localUp = -localUp;
        //                //else if (Math.Abs(Vector.Dot(localUp, WZ)) > 1e-6 && Vector.Dot(localUp, WZ) < 0) localUp = -localUp;

        //                JointBase joint = CreateJoint(jointType);
        //                bool shouldExtend = true;// (jointType == "Miter" || jointType == "Straight" || jointType == "Straight2");

        //                if (jointType == "T")
        //                {
        //                    //Logger.Log($"{compA.Name} ⊥ {compB.Name}");

        //                    // make sure compB is split once
        //                    if (!halves.ContainsKey(compB))
        //                    {
        //                        var (sB, eB) = JointModule.SplitBodyAtMidpoint(compB, localUp);
        //                        if (sB != null && eB != null)
        //                            halves[compB] = (sB, eB);
        //                    }

        //                    // fetch local construction segments
        //                    var segA = compA.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
        //                    var segB = compB.Template.Curves.FirstOrDefault()?.Shape as CurveSegment;
        //                    if (segA == null || segB == null)
        //                    {
        //                        //Logger.Log("TJoint: Missing construction curves.");
        //                        continue;
        //                    }

        //                    // lift into world
        //                    Point wA0 = compA.Placement * segA.StartPoint;
        //                    Point wA1 = compA.Placement * segA.EndPoint;
        //                    Point wB0 = compB.Placement * segB.StartPoint;
        //                    Point wB1 = compB.Placement * segB.EndPoint;

        //                    // determine which B-end lies on A’s world line
        //                    bool cutStart = IsPointOnSegment(wB0, wA0, wA1);
        //                    bool cutEnd = IsPointOnSegment(wB1, wA0, wA1);

        //                    string jointHalfName;
        //                    if (cutStart)
        //                    {
        //                        // B’s start sits on A → we need to reset B’s opposite half (HalfEnd)
        //                        jointHalfName = "HalfStart";
        //                        //Logger.Log($"{compB.Name}: T-joint — segB.Start lies on A; resetting HalfEnd.");
        //                    }
        //                    else if (cutEnd)
        //                    {
        //                        // B’s end sits on A → reset B’s HalfStart
        //                        jointHalfName = "HalfEnd";
        //                        //Logger.Log($"{compB.Name}: T-joint — segB.End lies on A; resetting HalfStart.");
        //                    }
        //                    else
        //                    {
        //                        //Logger.Log($"{compB.Name}: WARNING — no valid T connection detected; skipping.");
        //                        continue;
        //                    }

        //                    // reset just that half of B
        //                    //Logger.Log($"Start ResetHalfForJoint(…, \"{jointHalfName}\", …) on {compB.Name}");
        //                    JointModule.ResetHalfForJoint(
        //                        compB,
        //                        jointHalfName,
        //                        extendProfile: true,
        //                        localUp,
        //                        allCurves,
        //                        new List<Component> { compA, compB }
        //                    );
        //                    //Logger.Log($"Finished ResetHalfForJoint on {compB.Name}");

        //                    // grab the freshly-reset half as ExtrudedProfile
        //                    Body resetHalf = compB
        //                      .Template
        //                      .Bodies
        //                      .FirstOrDefault(b => b.Name == "ExtrudedProfile")
        //                      ?.Shape;
        //                    if (resetHalf == null)
        //                    {
        //                        //Logger.Log($"{compB.Name}: ERROR — reset half not found as ExtrudedProfile.");
        //                        continue;
        //                    }

        //                    // now execute the cutter on *only* that half
        //                    joint.Execute(compA, compB, spacing, null, resetHalf);
        //                }
        //                else
        //                {
        //                    // Non-T joints: split A and B if not split yet
        //                    if (!halves.ContainsKey(compA))
        //                    {
        //                        var (startA, endA) = JointModule.SplitBodyAtMidpoint(compA, localUp);
        //                        if (startA != null && endA != null)
        //                            halves[compA] = (startA, endA);
        //                    }
        //                    if (!halves.ContainsKey(compB))
        //                    {
        //                        var (startB, endB) = JointModule.SplitBodyAtMidpoint(compB, localUp);
        //                        if (startB != null && endB != null)
        //                            halves[compB] = (startB, endB);
        //                    }

        //                    var (aStart, aEnd) = halves[compA];
        //                    var (bStart, bEnd) = halves[compB];

        //                    if (aEndConnected && bStartConnected)
        //                    {
        //                        //Logger.Log($"{compA.Name}.end ↔ {compB.Name}.start");

        //                        JointModule.ResetHalfForJoint(compA, "HalfEnd", shouldExtend, localUp, allCurves, new List<Component> { compA, compB });
        //                        JointModule.ResetHalfForJoint(compB, "HalfStart", shouldExtend, localUp, allCurves, new List<Component> { compA, compB });

        //                        aEnd = compA.Template.Bodies.FirstOrDefault(b => b.Name == "HalfEnd")?.Shape;
        //                        bStart = compB.Template.Bodies.FirstOrDefault(b => b.Name == "HalfStart")?.Shape;

        //                        joint.Execute(compA, compB, spacing, aEnd, bStart);
        //                    }
        //                    else if (aStartConnected && bEndConnected)
        //                    {
        //                        //Logger.Log($"{compA.Name}.start ↔ {compB.Name}.end");

        //                        JointModule.ResetHalfForJoint(compA, "HalfStart", shouldExtend, localUp, allCurves, new List<Component> { compA, compB });
        //                        JointModule.ResetHalfForJoint(compB, "HalfEnd", shouldExtend, localUp, allCurves, new List<Component> { compA, compB });

        //                        aStart = compA.Template.Bodies.FirstOrDefault(b => b.Name == "HalfStart")?.Shape;
        //                        bEnd = compB.Template.Bodies.FirstOrDefault(b => b.Name == "HalfEnd")?.Shape;

        //                        joint.Execute(compA, compB, spacing, aStart, bEnd);
        //                    }
        //                    else if (aStartConnected && bStartConnected)
        //                    {
        //                        //Logger.Log($"{compA.Name}.start ↔ {compB.Name}.start");

        //                        JointModule.ResetHalfForJoint(compA, "HalfStart", shouldExtend, localUp, allCurves, new List<Component> { compA, compB });
        //                        JointModule.ResetHalfForJoint(compB, "HalfStart", shouldExtend, localUp, allCurves, new List<Component> { compA, compB });

        //                        aStart = compA.Template.Bodies.FirstOrDefault(b => b.Name == "HalfStart")?.Shape;
        //                        bStart = compB.Template.Bodies.FirstOrDefault(b => b.Name == "HalfStart")?.Shape;

        //                        joint.Execute(compA, compB, spacing, aStart, bStart);
        //                    }
        //                    else if (aEndConnected && bEndConnected)
        //                    {
        //                        //Logger.Log($"{compA.Name}.end ↔ {compB.Name}.end");

        //                        JointModule.ResetHalfForJoint(compA, "HalfEnd", shouldExtend, localUp, allCurves, new List<Component> { compA, compB });
        //                        JointModule.ResetHalfForJoint(compB, "HalfEnd", shouldExtend, localUp, allCurves, new List<Component> { compA, compB });

        //                        aEnd = compA.Template.Bodies.FirstOrDefault(b => b.Name == "HalfEnd")?.Shape;
        //                        bEnd = compB.Template.Bodies.FirstOrDefault(b => b.Name == "HalfEnd")?.Shape;

        //                        joint.Execute(compA, compB, spacing, aEnd, bEnd);
        //                    }
        //                    else
        //                    {
        //                        //Logger.Log($"{compA.Name} and {compB.Name}: WARNING - No matching connection found for joint.");
        //                    }
        //                }
        //                alreadyProcessed.Add((compA, compB));
        //            }
        //        }
        //    }
        //    //Logger.Log("ExecuteJoint: Complete.");
        //}

        public static void ExecuteJoint(Window window, double spacing, string jointType, bool updateBOM)
        {
            // 1) Gather selected components
            List<Component> selectedComponents = JointSelectionHelper.GetSelectedComponents(window);
            if (selectedComponents.Count < 1)
            {
                MessageBox.Show("Select at least one component to apply a joint.", "Selection Error");
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
                Component compA = selectedComponents.FirstOrDefault();
                if (compA == null) return;

                JointBase trimJoint = CreateJoint(jointType);
                trimJoint.Execute(compA, null, spacing, null, null);
            }
            else
            {
                // 5) Build all valid A–B pairs
                var workItems = GetConnectedPairs(selectedComponents, jointType);
                int total = workItems.Count();
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
                            if(aStartConnected || aEndConnected)
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
            if (updateBOM == true)
                //CompareCommand.CompareSimple();
                ExportCommands.ExportBOM(Window.ActiveWindow, update: true);
            else
                CompareCommand.CompareSimple();
        }


        public static void RestoreGeometry(Window window)
        {
            //Logger.Log("RestoreGeometry: started");

            List<Component> selectedComponents = JointSelectionHelper.GetSelectedComponents(window);
            if (selectedComponents.Count == 0)
            {
                MessageBox.Show("Select at least one component to restore geometry.", "Selection Error");
                return;
            }

            JointModule.ResetComponentGeometryOnly(selectedComponents);
            //Logger.Log("RestoreGeometry: complete");
        }

        public static void RestoreJoint(Window window)
        {
            Logger.Log("==== RestoreJoint START ====");
            WriteBlock.ExecuteTask("RestoreJoint", () =>
            {
                // 1) Gather selection
                var selected = JointSelectionHelper.GetSelectedComponents(window);
                if (selected.Count < 2)
                {
                    MessageBox.Show(
                        "Select at least two components to restore joints.",
                        "Restore Joint",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
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
                    Logger.Log($"RestoreJoint: ResetHalfForJoint({comp.Name}, {half}, extendProfile: true)");
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
                        Logger.Log($"RestoreJoint: ERROR on {comp.Name}.{half} → {ex.Message}");
                    }
                }

                MessageBox.Show(
                    "Joints restored: only the corner halves have been rebuilt (extended path).",
                    "Restore Joint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                Logger.Log("RestoreJoint: complete");
            });
            Logger.Log("==== RestoreJoint END ====");
        }



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

        /// <summary>
        /// True if B’s start or end world‐point lies on A’s world‐segment.
        /// </summary>
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

        /// <summary>
        /// Yields only physically‐connected or true T‐connected pairs.
        /// </summary>
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

                    //Logger.Log($"Pair [{a.Name},{b.Name}]: phys={phys}, Tconn={tconn}");
                    if (phys || tconn)
                        yield return (a, b);
                }
            }
        }

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
