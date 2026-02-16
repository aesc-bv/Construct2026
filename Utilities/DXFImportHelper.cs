using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using Application = SpaceClaim.Api.V242.Application;
using Point = SpaceClaim.Api.V242.Geometry.Point;

namespace AESCConstruct2026.FrameGenerator.Utilities
{
    /// <summary>
    /// Simple container for a DXF profile (name, profile‐string, and preview image).
    /// </summary>
    public class DXFProfile
    {
        public string Name { get; set; }
        public string ProfileString { get; set; }
        public string ImgString { get; set; }
    }

    public static class DXFImportHelper
    {
        //
        // ─── SESSION STORAGE FOR ALL DXFProfile INSTANCES ──────────────────────────────
        //
        /// <summary>
        /// Holds all DXFProfile objects created during this session.
        /// Used by “Save CSV” / “Load CSV” commands.
        /// </summary>
        public static List<DXFProfile> SessionProfiles { get; } = new List<DXFProfile>();

        //
        // ─── 1) IMPORT DXF CONTOURS ──────────────────────────────────────────────────────
        //
        /// <summary>
        /// Opens the specified .dxf file, extracts all trimmed curves from its first DatumPlane,
        /// and returns them as a List<ITrimmedCurve>. Returns true if at least 3 curves were found.
        /// </summary>
        public static bool ImportDXFContours(string filePath, out List<ITrimmedCurve> contours)
        {
            contours = new List<ITrimmedCurve>();
            var localContours = new List<ITrimmedCurve>();
            var originalWindow = Window.ActiveWindow;
            try
            {
                WriteBlock.ExecuteTask("Import DXF", () =>
                {
                    Document.Open(filePath, null);
                    var dxfWindow = Window.ActiveWindow;
                    var mainPartDXF = dxfWindow.Document.MainPart;

                    // Extract curves from first DatumPlane
                    DatumPlane dp = mainPartDXF.DatumPlanes.First();
                    foreach (DesignCurve dc in dp.Curves)
                    {
                        localContours.Add(dc.Shape);
                    }
                });
                contours.AddRange(localContours);
                return contours.Count > 2; // Return true if ≥ 3 curves found
            }
            catch
            {
                return false;
            }
        }

        //
        // ─── 2) GET DXF BOUNDING‐BOX SIZE ─────────────────────────────────────────────────
        //
        /// <summary>
        /// Given a list of ITrimmedCurve (assumed to lie in XY), returns (width, height).
        /// </summary>
        public static (double width, double height) GetDXFSize(List<ITrimmedCurve> contours)
        {
            if (contours == null || contours.Count == 0)
                return (0, 0);

            Point min = Point.Origin;
            Point max = Point.Origin;
            bool initialized = false;

            foreach (var curve in contours)
            {
                if (curve is CurveSegment segment)
                {
                    Point[] pts = { segment.StartPoint, segment.EndPoint };
                    foreach (var p in pts)
                    {
                        if (!initialized)
                        {
                            min = max = p;
                            initialized = true;
                        }
                        else
                        {
                            min = Point.Create(
                                Math.Min(min.X, p.X),
                                Math.Min(min.Y, p.Y),
                                0);
                            max = Point.Create(
                                Math.Max(max.X, p.X),
                                Math.Max(max.Y, p.Y),
                                0);
                        }
                    }
                }
            }

            double width = max.X - min.X;
            double height = max.Y - min.Y;
            return (width, height);
        }

        //
        // ─── 3) “DXF → PROFILE” (GENERATE PROFILE STRING + PREVIEW IMAGE) ─────────────────
        //
        /// <summary>
        /// Assumes the active window is a DXF. Builds a planar Body from all IDesignCurve contours,
        /// converts that Body into a single “profile string,” creates a small PNG‐preview (Base64),
        /// and returns a new DXFProfile (Name,ProfileString,ImgString). Also appends to SessionProfiles.
        /// </summary>
        public static DXFProfile DXFtoProfile()
        {
            DXFProfile dxfProfile = null;

            // Ensure SpaceClaim API is initialized
            if (Api.Session == null)
                initializeSC();

            var window = Window.ActiveWindow;
            if (window == null)
            {
                Application.ReportStatus("No active window.", StatusMessageType.Warning, null);
                return null;
            }

            var doc = window.Document;
            var mainPart = doc.MainPart;

            // Must have exactly one datum plane
            if (mainPart.DatumPlanes.Count() != 1)
            {
                Application.ReportStatus("DXF should have a single datum plane", StatusMessageType.Warning, null);
                return null;
            }

            var datumPlane = mainPart.DatumPlanes.First();
            // Ensure that plane is aligned with XY
            if (!((Plane)datumPlane.Shape.Geometry).Frame.DirZ.IsParallel(Direction.DirZ))
            {
                Application.ReportStatus("Plane is not aligned to XY frame", StatusMessageType.Warning, null);
                return null;
            }

            // Gather all contours (design curves) from that plane
            var itcList = new List<ITrimmedCurve>();
            foreach (IDesignCurve dc in mainPart.GetDescendants<IDesignCurve>())
                itcList.Add(dc.Shape);

            Body body = null;
            WriteBlock.ExecuteTask("Create DXF profile string and image", () =>
            {
                try
                {
                    body = Body.CreatePlanarBody(Plane.PlaneXY, itcList);
                }
                catch
                {
                }

                if (body != null)
                {
                    // 1) Build the profile‐string
                    string bodyString = StringFromBody(body);

                    // 2) Generate a small PNG preview (Base64) by creating a temporary document
                    var docImg = Document.Create();
                    DesignBody db = DesignBody.Create(docImg.MainPart, "ProfileBody", body);
                    Color myCustomColor = ColorTranslator.FromHtml("#006d8b");
                    db.SetColor(null, myCustomColor);
                    string imgString = getImgBase64(docImg.MainPart, 200, 150, Frame.Create(Point.Origin, Direction.DirZ));

                    // 4) Build the DXFProfile object
                    dxfProfile = new DXFProfile
                    {
                        Name = datumPlane.Name,
                        ProfileString = bodyString,
                        ImgString = imgString
                    };

                    // 5) (Optionally) rebuild it in the mainPart for verification
                    {
                        var bodyFromString = BodyFromString(bodyString);
                        DesignBody.Create(mainPart, "bodyFromString", bodyFromString);
                    }

                    // 6) Store in session for later “Save CSV”
                    SessionProfiles.Add(dxfProfile);
                }
            });

            if (body == null)
            {
                Application.ReportStatus("Invalid profile—check the DXF contours.", StatusMessageType.Error, null);
            }
            return dxfProfile;
        }

        //
        // ─── 4) CONVERT BODY → PROFILE STRING ────────────────────────────────────────────
        //
        /// <summary>
        /// Converts the very first face of <paramref name="body"/> (assumed planar) into a single
        /// string. Each loop's edges become substrings of the form:
        ///   For a line:    Sx_yEy_z
        ///   For an arc:     Sx_yEy_zMmx_my
        /// Loops are joined by “&”, outer‐loops first.
        /// </summary>

        public static string StringFromBody(Body body)
        {
            string profileString = "";
            var face = body.Faces.First();

            foreach (Loop loop in face.Loops)
            {
                string loopString = "";

                foreach (Edge edge in loop.Edges)
                {
                    var geomLine = edge.Geometry as Line;
                    var geomArc = edge.Geometry as Circle;
                    bool isLine = (geomLine != null);
                    bool isCircle = (geomArc != null);

                    // Format endpoints ×1000 as "Sx_yEy_z" using invariant‐culture:
                    string sx = (1000 * edge.StartPoint.X)
                                    .ToString("0.#####", CultureInfo.InvariantCulture);
                    string sy = (1000 * edge.StartPoint.Y)
                                    .ToString("0.#####", CultureInfo.InvariantCulture);
                    string ex = (1000 * edge.EndPoint.X)
                                    .ToString("0.#####", CultureInfo.InvariantCulture);
                    string ey = (1000 * edge.EndPoint.Y)
                                    .ToString("0.#####", CultureInfo.InvariantCulture);

                    string startEnd = $"S{sx}_{sy}E{ex}_{ey}";
                    string subString = "";

                    if (isLine)
                    {
                        subString = startEnd;
                    }
                    else if (isCircle)
                    {
                        // Get circle midpoint, also scaled ×1000 using invariant‐culture
                        string mx = (1000 * geomArc.Axis.Origin.X)
                                        .ToString("0.#####", CultureInfo.InvariantCulture);
                        string my = (1000 * geomArc.Axis.Origin.Y)
                                        .ToString("0.#####", CultureInfo.InvariantCulture);

                        if ((edge.StartPoint - edge.EndPoint).Magnitude > 1e-5)
                        {
                            var testArc = CurveSegment.CreateArc(
                                geomArc.Axis.Origin,
                                edge.StartPoint,
                                edge.EndPoint,
                                -Direction.DirZ);
                            bool coincident = edge.IsCoincident(testArc);
                            if (!coincident)
                                startEnd = $"S{ex}_{ey}E{sx}_{sy}";
                        }

                        subString = $"{startEnd}M{mx}_{my}";
                    }
                    else
                    {
                        continue;
                    }

                    if (loopString == "")
                        loopString = subString;
                    else
                        loopString += " " + subString;
                }

                if (loop.IsOuter)
                {
                    profileString = (profileString == "")
                                  ? loopString
                                  : loopString + "&" + profileString;
                }
                else
                {
                    profileString = (profileString == "")
                                  ? loopString
                                  : profileString + "&" + loopString;
                }
            }

            return profileString;
        }


        //
        // ─── 5) PARSING POINT & CURVE STRINGS ────────────────────────────────────────────
        //
        /// <summary>
        /// Parse a single substring (either "S…E…" or "S…E…M…") into an ITrimmedCurve (line, arc, or full circle).
        /// </summary>
        public static ITrimmedCurve CurveFromString(string curveString)
        {

            const double tol = 1e-6;

            // 1) Check for midpoint ('M') tag
            int mIndex = curveString.IndexOf('M');
            bool hasMidpoint = (mIndex >= 0);

            // 2) Extract "S…E…" portion (everything before 'M' if present)
            string sePart = hasMidpoint
                ? curveString.Substring(0, mIndex)
                : curveString;

            // 3) Parse start/end points out of sePart: it always starts with 'S', then "x_y", then 'E', then "x_y".
            //    e.g. "Ssx_syExx_yy"
            if (!sePart.StartsWith("S") || !sePart.Contains("E"))
            {
                // Malformed → return a degenerate zero‐length line at origin
                return CurveSegment.Create(Point.Create(0, 0, 0), Point.Create(0, 0, 0));
            }

            string afterS = sePart.Substring(1);        // drop leading 'S'
            string[] seTokens = afterS.Split('E');       // ["sx_sy", "ex_ey"]
            if (seTokens.Length != 2)
            {
                // Malformed → degenerate line
                return CurveSegment.Create(Point.Create(0, 0, 0), Point.Create(0, 0, 0));
            }

            // 4) Convert "sx_sy" → ps, "ex_ey" → pe
            Point ps = PointFromString(seTokens[0]);
            Point pe = PointFromString(seTokens[1]);

            // 5) If there is no 'M', this is just a straight line
            if (!hasMidpoint)
            {
                return CurveSegment.Create(ps, pe);
            }

            // 6) Otherwise, parse midpoint "Mmx_my" (everything after 'M')
            string midPart = curveString.Substring(mIndex + 1);
            Point pm = PointFromString(midPart);

            // 7) If start and end coincide (within tolerance), build a full circle
            if ((ps - pe).Magnitude < tol)
            {
                double radius = (ps - pm).Magnitude;
                var circleGeo = Circle.Create(Frame.Create(pm, Direction.DirZ), radius);
                return CurveSegment.Create(circleGeo);
            }

            // 8) Otherwise, build a circular arc from ps → pe around center pm.
            //    Try CCW (DirZ) first. If endpoints don't match, flip normal to -DirZ.
            try
            {
                var testArc = CurveSegment.CreateArc(pm, ps, pe, Direction.DirZ);
                // Compare endpoints by distance instead of .IsAlmostEqual
                if (((testArc.StartPoint - ps).Magnitude < tol) &&
                    ((testArc.EndPoint - pe).Magnitude < tol))
                {
                    return testArc;
                }
                else
                {
                    // Reverse winding
                    var reversedArc = CurveSegment.CreateArc(pm, ps, pe, -Direction.DirZ);
                    return reversedArc;
                }
            }
            catch
            {
                // If arc creation fails, fall back to straight line
                return CurveSegment.Create(ps, pe);
            }
        }

        /// <summary>
        /// Parse "x_y" (scaled×1000 and using dot‐decimal) back into a SpaceClaim Point (Z=0).
        /// </summary>
        /// 
        public static Point PointFromString(string pointString)
        {
            // pointString is like “13.66025_-6.83013” (dot decimal)
            string[] parts = pointString.Split('_');
            double x = double.Parse(parts[0], CultureInfo.InvariantCulture) * 0.001;
            double y = double.Parse(parts[1], CultureInfo.InvariantCulture) * 0.001;
            return Point.Create(x, y, 0);
        }

        //
        // ─── 6) BUILD A PLANAR BODY FROM A PROFILE STRING ───────────────────────────────
        //
        /// <summary>
        /// Given a single “profileString” (loops separated by '&', edges by spaces), rebuild a planar Body.
        /// </summary>
        public static Body BodyFromString(string profileString)
        {
            var allCurves = new List<ITrimmedCurve>();
            foreach (string loopStr in profileString.Split('&'))
            {
                foreach (string curveStr in loopStr.Split(' '))
                {
                    allCurves.Add(CurveFromString(curveStr));
                }
            }
            return Body.CreatePlanarBody(Plane.PlaneXY, allCurves);
        }

        //
        // ─── 7) RENDER A PART TO A SMALL PNG (BASE64) ────────────────────────────────────
        //
        /// <summary>
        /// Create a hidden SpaceClaim window, render <paramref name="part"/> at size (sizeX×sizeY),
        /// using the given viewFrame, and return a Base64-encoded PNG string.
        /// </summary>
        public static string getImgBase64(Part part, int sizeX, int sizeY, Frame viewFrame)
        {
            try
            {
                var imageSize = new Size(sizeX, sizeY);
                var _window = Window.Create(part, false);
                _window.SetProjection(Matrix.CreateMapping(viewFrame), true, false);
                _window.InteractionMode = InteractionMode.Solid;

                Bitmap bmp = _window.CreateBitmap(imageSize);
                bmp.MakeTransparent();

                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    byte[] bytes = ms.ToArray();
                    return Convert.ToBase64String(bytes);
                }
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"getImgBase64 error:\n{ex}", StatusMessageType.Error, null);
                return "";
            }
        }

        //
        // ─── 8) CSV-SAVE / CSV-LOAD FOR DXFProfile ───────────────────────────────────────
        //
        public static class DXFProfileCsvHandler
        {
            /// <summary>
            /// Save the given list of DXFProfile objects to a semicolon-delimited CSV.
            /// Columns: Name;ProfileString;ImgString
            /// </summary>
            public static void SaveDXFProfiles(string filePath, List<DXFProfile> profiles)
            {
                using (var sw = new StreamWriter(filePath))
                {
                    sw.WriteLine("Name;ProfileString;ImgString");

                    foreach (var profile in profiles)
                    {
                        sw.WriteLine($"{profile.Name};{profile.ProfileString};{profile.ImgString}");
                    }
                }
            }

            /// <summary>
            /// Load all DXFProfile rows from a semicolon-delimited CSV and return them.
            /// Expects either a header row “Name;ProfileString;ImgString” or directly data rows.
            /// </summary>
            public static List<DXFProfile> LoadDXFProfiles(string filePath)
            {
                var profiles = new List<DXFProfile>();
                if (!File.Exists(filePath))
                    throw new FileNotFoundException("The specified file does not exist.", filePath);

                using (var sr = new StreamReader(filePath))
                {
                    string firstLine = sr.ReadLine();
                    bool hasHeader = firstLine != null && firstLine.StartsWith("Name;");
                    if (!hasHeader && firstLine != null)
                    {
                        ProcessCsvLine(firstLine, profiles);
                    }

                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        ProcessCsvLine(line, profiles);
                    }
                }

                return profiles;
            }

            private static void ProcessCsvLine(string line, List<DXFProfile> profiles)
            {
                var parts = line.Split(';');
                if (parts.Length < 3)
                    return;
                profiles.Add(new DXFProfile
                {
                    Name = parts[0],
                    ProfileString = parts[1],
                    ImgString = parts[2]
                });
            }
        }

        //
        // ─── 9) ENSURE SPACECLAIM API IS INITIALIZED ─────────────────────────────────────
        //
        private static void initializeSC()
        {
            Api.Initialize();
        }
    }
}
