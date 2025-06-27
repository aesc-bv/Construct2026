using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;            // MessageBox
using SpaceClaim.Api.V251;
using SpaceClaim.Api.V251.Geometry;
using AESCConstruct25.Modules;
using AESCConstruct25.Utilities;
using SpaceClaim.Api.V251.Extensibility;

namespace AESCConstruct25.Commands
{
    class ExtrudeProfileCommand : CommandCapsule
    {
        public const string CommandName = "AESCConstruct25.ExtrudeProfile";
        private static readonly string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AESCConstruct25_Log.txt");

        public ExtrudeProfileCommand()
            : base(CommandName, "Extrude Profile", null, "Extrudes a selected profile along selected lines or edges") { }

        public static void ExecuteExtrusion(
            string profileType,
            object profileDataOrContours,
            bool isHollow,
            double offsetX,
            double offsetY,
            string dxfFilePath = ""
        )
        {
            //Logger.Log($"ExtrudeProfileCommand: {profileType}, hollow = {isHollow}");

            var win = Window.ActiveWindow;
            if (win == null)
                return;

            // 1) Get the selected CurveSegment(s) in the active SpaceClaim view:
            var trimmed = ProfileSelectionHelper.GetSelectedCurves(win);
            var rawCurves = trimmed
                .Select(tc => tc as CurveSegment
                            ?? CurveSegment.Create(tc.StartPoint, tc.EndPoint))
                .ToList();
            if (rawCurves.Count == 0)
            {
                MessageBox.Show(
                    "Select at least one straight line or edge.",
                    "Selection Required",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            // 2) Decide which “mode” we’re in:
            List<ITrimmedCurve> dxfContours = null;
            Dictionary<string, string> profileData = null;

            if (profileType == "DXF" && profileDataOrContours is List<ITrimmedCurve> dc)
            {
                // True DXF mode: we expect a real DXF file on disk
                //if (string.IsNullOrEmpty(dxfFilePath))
                //{
                //    //Logger.Log("ERROR – DXF profileType but missing file path");
                //    return;
                //}
                dxfContours = dc;
            }
            else if (profileType == "CSV" && profileDataOrContours is List<ITrimmedCurve> contoursFromCsv)
            {
                // “CSV” mode: user‐saved profile, already reconstructed as a list of ITrimmedCurve
                dxfContours = contoursFromCsv;

            }
            else if (profileDataOrContours is Dictionary<string, string> fields)
            {
                // Built-in shape mode (Rectangular, Circular, etc.)
                profileData = fields.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            }
            else
            {
                //Logger.Log($"ERROR – unrecognized profileType or data: “{profileType}”");
                MessageBox.Show(
                    "Error: invalid profile type or data. Cannot create extrusion.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            const double tol = 1e-6;

            //IEnumerable<(Component A, Component B)> workItems = GetConnectedPairs(selectedComponents, "RestoreJoint");
            using (var PT = ProgressTracker.Create(rawCurves.Count()))
            {
                int i = 0;
                foreach (var seg in rawCurves)
                {
                    i++;
                    PT.Progress = i;
                    PT.Message = $"Creating profile {i}/{rawCurves.Count()}";

                    // 3) For each selected segment, compute localUp and call ProfileModule.ExtrudeProfile:
                    //foreach (var seg in rawCurves)
                    //{
                    // If we’re in built-in mode, profileData is non-null; otherwise, profileData stays null.
                    var dataForThisCurve = profileData != null
                                         ? new Dictionary<string, string>(profileData)
                                         : null;

                    // Find a “neighbor” so we can compute a stable localUp direction:
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

                    // “Stabilize” the sign of localUp so it’s roughly pointing up in world space:
                    Vector WY = Vector.Create(0, 1, 0),
                           WX = Vector.Create(1, 0, 0),
                           WZ = Vector.Create(0, 0, 1);

                    if (Math.Abs(Vector.Dot(localUp, WY)) > tol && Vector.Dot(localUp, WY) < 0)
                        localUp = -localUp;
                    else if (Math.Abs(Vector.Dot(localUp, WX)) > tol && Vector.Dot(localUp, WX) < 0)
                        localUp = -localUp;
                    else if (Math.Abs(Vector.Dot(localUp, WZ)) > tol && Vector.Dot(localUp, WZ) < 0)
                        localUp = -localUp;

                    // 4) Now hand off to ProfileModule.ExtrudeProfile:
                    ProfileModule.ExtrudeProfile(
                        win,
                        profileType,
                        seg,
                        isHollow,
                        dataForThisCurve,
                        offsetX,
                        offsetY,
                        localUp,
                        dxfFilePath,   // only non-empty if profileType == "DXF"
                        dxfContours,   // non-null if DXF or CSV
                        reuseComponent: null
                    );
                }
            }
            CompareCommand.CompareSimple();
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
