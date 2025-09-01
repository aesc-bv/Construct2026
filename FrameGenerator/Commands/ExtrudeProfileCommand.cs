using AESCConstruct25.FrameGenerator.Modules;
using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Extensibility;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using Application = SpaceClaim.Api.V242.Application; // MessageBox

namespace AESCConstruct25.FrameGenerator.Commands
{
    class ExtrudeProfileCommand : CommandCapsule
    {
        public const string CommandName = "AESCConstruct25.ExtrudeProfile";

        public ExtrudeProfileCommand()
            : base(CommandName, "Extrude Profile", null, "Extrudes a selected profile along selected lines or edges") { }

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

                // 1) Get the selected CurveSegment(s) in the active SpaceClaim view:
                var trimmed = ProfileSelectionHelper.GetSelectedCurves(win);
                var rawCurves = trimmed
                    .Select(tc => tc as CurveSegment
                                ?? CurveSegment.Create(tc.StartPoint, tc.EndPoint))
                    .ToList();
                var createdCurves = new List<DesignCurve>();

                if (rawCurves.Count == 0)
                {
                    //MessageBox.Show(
                    //    "Select at least one straight line or edge.",
                    //    "Selection Required",
                    //    MessageBoxButtons.OK,
                    //    MessageBoxIcon.Warning
                    //);
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
                    //MessageBox.Show(
                    //    "Error: invalid profile type or data. Cannot create extrusion.",
                    //    "Error",
                    //    MessageBoxButtons.OK,
                    //    MessageBoxIcon.Error
                    //);
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

                            var dA = (seg.EndPoint - seg.StartPoint).Direction.ToVector();
                            Vector localUp;
                            if (neighbour != null)
                            {
                                var dB = (neighbour.EndPoint - neighbour.StartPoint).Direction.ToVector();
                                localUp = Vector.Cross(dA, dB).Magnitude > tol
                                        ? Vector.Cross(dA, dB).Direction.ToVector()
                                        : Vector.Create(0, 1, 0);
                            }
                            else
                            {
                                localUp = Vector.Create(0, 1, 0);
                            }

                            Vector WY = Vector.Create(0, 1, 0),
                                   WX = Vector.Create(1, 0, 0),
                                   WZ = Vector.Create(0, 0, 1);

                            if (Math.Abs(Vector.Dot(localUp, WY)) > tol && Vector.Dot(localUp, WY) < 0)
                                localUp = -localUp;
                            else if (Math.Abs(Vector.Dot(localUp, WX)) > tol && Vector.Dot(localUp, WX) < 0)
                                localUp = -localUp;
                            else if (Math.Abs(Vector.Dot(localUp, WZ)) > tol && Vector.Dot(localUp, WZ) < 0)
                                localUp = -localUp;

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
                    else
                        CompareCommand.CompareSimple();
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

        /// <summary>
        /// Returns another segment in 'all' that shares an endpoint with 'seg', or null.
        /// </summary>
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
