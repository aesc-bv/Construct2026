/*
 ExtrudeProfileCommand wires the SpaceClaim "Extrude Profile" command into the UI.

 It gathers the selected straight curves, resolves whether the profile comes from
 CSV/built-in data or DXF contours, computes a deterministic local up vector per
 path segment, and delegates the actual solid creation to ProfileModule.ExtrudeProfile.
 Optionally it also updates the BOM after all extrusions are created.
*/

using AESCConstruct2026.FrameGenerator.Modules;
using AESCConstruct2026.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Extensibility;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using Application = SpaceClaim.Api.V242.Application; // MessageBox

namespace AESCConstruct2026.FrameGenerator.Commands
{
    class ExtrudeProfileCommand : CommandCapsule
    {
        // Registers the "Extrude Profile" command with SpaceClaim and sets its label/description.
        public const string CommandName = "AESCConstruct2026.ExtrudeProfile";

        public ExtrudeProfileCommand()
            : base(CommandName, "Extrude Profile", null, "Extrudes a selected profile along selected lines or edges") { }

        // Drives the full extrusion flow: reads selection, resolves profile source (CSV/DXF), and extrudes along each selected segment.
        public static void ExecuteExtrusion(
            string profileType,
            object profileDataOrContours,
            bool isHollow,
            double offsetX,
            double offsetY,
            string dxfFilePath = "",
            bool updateBOM = false,
            string selectedProfileString = ""
        )
        {
            try
            {
                var win = Window.ActiveWindow;
                if (win == null)
                {
                    Application.ReportStatus("No active window found.", StatusMessageType.Error, null);
                    return;
                }

                var replace = ProfileReplaceHelper.PromptAndReplaceIfAny(win);
                if (!replace.proceed)
                    return;

                // 1) Get the selected CurveSegment(s) in the active SpaceClaim view:
                var trimmed = ProfileSelectionHelper.GetSelectedCurves(win);
                var rawCurves = trimmed
                    .Select(tc => tc as CurveSegment
                                ?? CurveSegment.Create(tc.StartPoint, tc.EndPoint))
                    .ToList();
                var createdCurves = new List<DesignCurve>();

                if (rawCurves.Count == 0)
                {
                    Application.ReportStatus("Select at least one straight line or edge.", StatusMessageType.Warning, null);
                    return;
                }

                // 2) Decide which “mode” we’re in:
                List<ITrimmedCurve> dxfContours = null;
                Dictionary<string, string> profileData = null;

                if (profileType == "DXF" && profileDataOrContours is List<ITrimmedCurve> dc)
                {
                    dxfContours = dc;
                }
                else if (profileType == "CSV" && profileDataOrContours is List<ITrimmedCurve> contoursFromCsv)
                {
                    dxfContours = contoursFromCsv;
                }
                else if (profileDataOrContours is Dictionary<string, string> fields)
                {
                    profileData = fields.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                }
                else
                {
                    Application.ReportStatus("Error: invalid profile type or data. Cannot create extrusion.", StatusMessageType.Error, null);
                    return;
                }

                const double tol = 1e-6;

                using (var PT = ProgressTracker.Create(rawCurves.Count))
                {
                    int i = 0;
                    foreach (var seg in rawCurves)
                    {
                        i++;
                        PT.Progress = i;
                        PT.Message = $"Creating profile {i}/{rawCurves.Count}";

                        try
                        {
                            var dataForThisCurve = profileData != null
                                                 ? new Dictionary<string, string>(profileData)
                                                 : null;

                            var neighbour = FindConnectedSegment(seg, rawCurves);

                            //var dA = (seg.EndPoint - seg.StartPoint).Direction.ToVector();
                            //Vector localUp;
                            //if (neighbour != null)
                            //{
                            //    var dB = (neighbour.EndPoint - neighbour.StartPoint).Direction.ToVector();
                            //    localUp = Vector.Cross(dA, dB).Magnitude > tol
                            //            ? Vector.Cross(dA, dB).Direction.ToVector()
                            //            : Vector.Create(0, 1, 0);
                            //}
                            //else
                            //{
                            //    localUp = Vector.Create(0, 1, 0);
                            //}

                            //Vector WY = Vector.Create(0, 1, 0),
                            //       WX = Vector.Create(1, 0, 0),
                            //       WZ = Vector.Create(0, 0, 1);

                            //if (Math.Abs(Vector.Dot(localUp, WY)) > tol && Vector.Dot(localUp, WY) < 0)
                            //    localUp = -localUp;
                            //else if (Math.Abs(Vector.Dot(localUp, WX)) > tol && Vector.Dot(localUp, WX) < 0)
                            //    localUp = -localUp;
                            //else if (Math.Abs(Vector.Dot(localUp, WZ)) > tol && Vector.Dot(localUp, WZ) < 0)
                            //    localUp = -localUp;

                            var dA = (seg.EndPoint - seg.StartPoint).Direction.ToVector();

                            // Always compute a canonical, deterministic up
                            var localUp = CanonicalUpForLine(dA);

                            // Proceed as before
                            ProfileModule.ExtrudeProfile(
                                win,
                                profileType,
                                seg,
                                isHollow,
                                dataForThisCurve,
                                offsetX,
                                offsetY,
                                localUp,
                                dxfFilePath,
                                dxfContours,
                                reuseComponent: null,
                                csvProfileString: selectedProfileString,
                                createdCurves: createdCurves
                            );

                            var selectedCurves = win.ActiveContext.Selection
                                .OfType<DesignCurve>()
                                .ToHashSet();

                            var p1 = seg.StartPoint;
                            var p2 = seg.EndPoint;

                            var selectedDesignCurves = win.ActiveContext.Selection
                                .OfType<DesignCurve>()
                                .ToList();

                            foreach (var dcurve in selectedDesignCurves)
                            {
                                dcurve.SetVisibility(null, false);
                            }

                            // Local helper: compares two points with a distance tolerance in model units.
                            bool PointsEqual(Point a, Point b, double tolerance)
                            {
                                return (a - b).Magnitude < tolerance;
                            }

                            var match = selectedCurves.FirstOrDefault(dc =>
                            {
                                var line = dc.Shape.Geometry as Line;
                                if (line == null) return false;

                                var lsp = dc.Shape.StartPoint;
                                var lep = dc.Shape.EndPoint;

                                return (PointsEqual(p1, lsp, tol) && PointsEqual(p2, lep, tol)) ||
                                       (PointsEqual(p1, lep, tol) && PointsEqual(p2, lsp, tol));
                            });
                        }
                        catch (Exception ex)
                        {
                            Application.ReportStatus($"Error extruding profile segment {i}:\n{ex.Message}", StatusMessageType.Error, null);
                        }
                    }
                }

                try
                {
                    if (updateBOM == true)
                        ExportCommands.ExportBOM(Window.ActiveWindow, update: true);
                    //else
                    //CompareCommand.CompareSimple();
                }
                catch (Exception ex)
                {
                    Application.ReportStatus($"Error updating BOM or comparing bodies:\n{ex.Message}", StatusMessageType.Error, null);
                }
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Unexpected error during extrusion:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Computes a stable "up" vector for a given line direction, preferring +Z and falling back to +Y/+X if needed.
        static Vector CanonicalUpForLine(Vector dirVec)
        {
            const double tol = 1e-9;

            // Canonical world axes
            Vector UZ = Vector.Create(0, 0, 1);
            Vector UY = Vector.Create(0, 1, 0);
            Vector UX = Vector.Create(1, 0, 0);

            // Ensure we work with a unit direction
            var d = dirVec.Direction.ToVector();

            // Projects an up-vector onto the plane perpendicular to the given axis (removes axis component).
            Vector ProjectOntoPlane(Vector up, Vector axisUnit)
            {
                // Reject the component along the axis; result is perpendicular to the line
                return up - Vector.Dot(up, axisUnit) * axisUnit;
            }

            // 1) Prefer world-Z (so anything in the XY plane will keep +Z as up)
            var up = ProjectOntoPlane(UZ, d);
            if (up.Magnitude <= tol)
            {
                // 2) If the line is parallel to Z, use world-Y
                up = ProjectOntoPlane(UY, d);
                if (up.Magnitude <= tol)
                {
                    // 3) Degenerate edge case: fall back to world-X
                    up = ProjectOntoPlane(UX, d);
                }
            }

            // Force a consistent orientation: always toward +Z when possible
            if (Vector.Dot(up, UZ) < 0) up = -up;

            return up.Direction.ToVector();
        }

        /// <summary>
        /// Returns another segment in 'all' that shares an endpoint with 'seg', or null.
        /// </summary>
        // Finds the first other CurveSegment in the collection that touches seg within a small tolerance.
        private static CurveSegment FindConnectedSegment(CurveSegment seg, IEnumerable<CurveSegment> all)
        {
            const double tol = 1e-6;
            foreach (var o in all)
            {
                if (o == seg) continue;
                if ((seg.StartPoint - o.StartPoint).Magnitude < tol ||
                    (seg.StartPoint - o.EndPoint).Magnitude < tol ||
                    (seg.EndPoint - o.StartPoint).Magnitude < tol ||
                    (seg.EndPoint - o.EndPoint).Magnitude < tol)
                    return o;
            }
            return null;
        }
    }
}
