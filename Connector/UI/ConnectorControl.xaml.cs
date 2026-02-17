/* 
 * ConnectorControl hosts the UI and preview logic for configuring and creating connector geometry
 * in SpaceClaim: it reads presets from CSV, drives a WinForms drawing surface, builds ConnectorModel
 * instances from the WPF UI, and applies 3D boolean edits to owner bodies and their neighbours.
 */

// SpaceClaim APIs
using AESCConstruct2026.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Application = SpaceClaim.Api.V242.Application;
using Body = SpaceClaim.Api.V242.Modeler.Body;
using CheckBox = System.Windows.Controls.CheckBox;
using Color = System.Drawing.Color;
using Component = SpaceClaim.Api.V242.Component;
using ConnectorModel = AESCConstruct2026.Connector.Connector;
using Frame = SpaceClaim.Api.V242.Geometry.Frame;
using KeyEventHandler = System.Windows.Input.KeyEventHandler;
using Line = SpaceClaim.Api.V242.Geometry.Line;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using RadioButton = System.Windows.Controls.RadioButton;
using Settings = AESCConstruct2026.Properties.Settings;
using TextBox = System.Windows.Controls.TextBox;
using UserControl = System.Windows.Controls.UserControl;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;
using WF = System.Windows.Forms;
using Window = SpaceClaim.Api.V242.Window;


namespace AESCConstruct2026.UI
{
    public partial class ConnectorControl : UserControl
    {
        private WF.PictureBox picDrawing;
        private readonly TimeSpan UiDebounceInterval = TimeSpan.FromMilliseconds(150);
        private DispatcherTimer _uiDebounce;

        private static readonly Dictionary<DesignBody, NeighbourIndepChoice> s_neighbourDecisionCache
            = new Dictionary<DesignBody, NeighbourIndepChoice>();

        private class ConnectorPresetRecord
        {
            public string Name { get; set; } = "";
            public double? Height { get; set; }
            public double? Width1 { get; set; }
            public double? Width2 { get; set; }
            public double? Tolerance { get; set; }
            public double EndRelief { get; set; }
            public double? RadiusChamfer { get; set; }
            public string Location { get; set; } = "";

            public bool? UseCustom { get; set; }
            public bool? DynamicHeight { get; set; }
            public bool? CornerCutout { get; set; }
            public double? CornerCutoutValue { get; set; }
            public bool? CornerCutoutRadiusEnabled { get; set; }
            public double? CornerCutoutRadius { get; set; }
            public bool? ClickLocation { get; set; }
            public bool? ShowTolerance { get; set; }
            public bool? connectorStraight { get; set; }

            // "Radius" or "Chamfer"
            public string CornerStyle { get; set; } = "";
        }

        public static int nameIndex = 1;

        private enum NeighbourIndepChoice { MakeIndependent, EditShared, Skip }

        private readonly List<ConnectorPresetRecord> _presets = new List<ConnectorPresetRecord>();

        // Constructor wires up the WPF control, WinForms drawing host and localization, then initializes the connector UI.
        public ConnectorControl()
        {
            try
            {
                InitializeComponent();
                InitializeDrawingHost();

                // ensure we actually run once
                this.Loaded += ConnectorControl_Loaded;

                DataContext = this;
                Localization.Language.LocalizeFrameworkElement(this);
                LocalizeUI();
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to initialize ProfileSelectionControl:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        // InitializeDrawingHost creates and embeds the WinForms PictureBox used to preview the connector profile.
        private void InitializeDrawingHost()
        {
            try
            {
                picDrawing = new WF.PictureBox
                {
                    Dock = WF.DockStyle.Fill,
                    BackColor = System.Drawing.Color.White
                };
                picDrawing.Paint += PicDrawing_Paint;
                picDrawing.Resize += (_, __) =>
                {
                    try
                    {
                        InvalidateDrawing();
                    }
                    catch (Exception)
                    {
                    }
                };

                // host from XAML
                picDrawingHost.Child = picDrawing;

                // keep drawing responsive to WPF size changes too
                this.SizeChanged += (_, __) =>
                {
                    try { InvalidateDrawing(); }
                    catch (Exception)
                    {
                    }
                };
            }
            catch (Exception)
            {
            }
        }

        // PicDrawing_Paint renders the current ConnectorModel instance onto the PictureBox graphics surface.
        private void PicDrawing_Paint(object sender, WF.PaintEventArgs e)
        {
            try
            {
                if (connector != null)
                    DrawConnector(e.Graphics, connector);
            }
            catch (Exception ex)
            {
                ////Logger.Log($"PicDrawing_Paint error: {ex}");
            }
        }

        // LocalizeUI re-applies localization to the ConnectorControl UI elements.
        private void LocalizeUI()
        {
            Localization.Language.LocalizeFrameworkElement(this);
        }

        // === Legacy field (now strongly-typed via alias) ===
        private ConnectorModel connector;

        // drawConnector rebuilds the ConnectorModel from the current UI values and refreshes the drawing.
        private void drawConnector()
        {
            try
            {
                connector = ConnectorModel.CreateConnector(this);
                InvalidateDrawing();
            }
            catch (Exception ex)
            {
            }
        }

        // btnCreate_connector is the click handler that starts 3D connector creation from the current UI.
        private void btnCreate_connector(object sender, RoutedEventArgs e) => createConnector();

        // ConnectorControl_Loaded runs once when the control is loaded to wire UI events, load presets, and draw the initial connector.
        private void ConnectorControl_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                WireUiChangeHandlers();
                LoadConnectorPresets();
                drawConnector();
                UpdateGenerateEnabled();
            }
            catch (Exception ex) { }
        }

        // LoadConnectorPresets reads connector presets from a CSV file and binds them into the preset combobox.
        private void LoadConnectorPresets()
        {
            try
            {
                string csvPath = Settings.Default.ConnectorProperties;
                if (string.IsNullOrWhiteSpace(csvPath) || !File.Exists(csvPath))
                {
                    // Fallback: CommonApplicationData\AESCConstruct\Connector\ConnectorProperties.csv
                    var fallback = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                        "AESCConstruct", "Connector", "ConnectorProperties.csv");
                    csvPath = File.Exists(csvPath) ? csvPath : fallback;
                }

                if (!File.Exists(csvPath))
                    throw new FileNotFoundException("ConnectorProperties.csv not found", csvPath);

                _presets.Clear();

                using var sr = new StreamReader(csvPath, DetectEncoding(csvPath));
                if (sr.EndOfStream)
                    throw new InvalidDataException("ConnectorProperties.csv is empty.");

                var headerLine = sr.ReadLine();
                if (headerLine == null)
                    throw new InvalidDataException("ConnectorProperties.csv header row missing.");

                char delimiter = headerLine.Contains(';') ? ';' : ',';
                var headers = headerLine.Split(delimiter)
                                        .Select((h, i) => new { h = h.Trim(), i })
                                        .ToDictionary(x => x.h, x => x.i, StringComparer.OrdinalIgnoreCase);

                string line;
                int lineNum = 1;
                while ((line = sr.ReadLine()) != null)
                {
                    lineNum++;
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var cols = SplitCsv(line, delimiter);

                    // Only read these columns: Height, Width1, Width2, RadiusButton, Radius, Tolerance (+ optional Name for display)
                    var p = new ConnectorPresetRecord
                    {
                        Name = Get<string>(headers, cols, "Name", ""),
                        Height = GetDouble(headers, cols, "Height"),
                        Width1 = GetDouble(headers, cols, "Width1"),
                        Width2 = GetDouble(headers, cols, "Width2"),
                        // CSV uses "Radius" (not "RadiusChamfer"); map to the existing Radius/Chamfer textbox
                        RadiusChamfer = GetDouble(headers, cols, "Radius"),
                        Tolerance = GetDouble(headers, cols, "Tolerance"),
                    };

                    // Interpret "RadiusButton": Y/1/true => Radius radio, N/0/false => Chamfer radio
                    var radiusBtn = GetBool(headers, cols, "RadiusButton");
                    if (radiusBtn.HasValue)
                        p.CornerStyle = radiusBtn.Value ? "radius" : "chamfer";
                    else
                        p.CornerStyle = ""; // leave as-is if not provided

                    // Keep rows that at least have one meaningful value
                    if (p.Height.HasValue || p.Width1.HasValue || p.Width2.HasValue || p.RadiusChamfer.HasValue || p.Tolerance.HasValue || !string.IsNullOrWhiteSpace(p.Name))
                        _presets.Add(p);
                }

                // Bind to existing combobox. Use Name when present; otherwise show a compact synthesized label.
                if (ConnectorShapeCombobox != null)
                {
                    string Label(ConnectorPresetRecord r)
                    {
                        if (!string.IsNullOrWhiteSpace(r.Name)) return r.Name;
                        string f(double? v) => v.HasValue ? v.Value.ToString("0.###", CultureInfo.InvariantCulture) : "";
                        return $"H{f(r.Height)} W1{f(r.Width1)} W2{f(r.Width2)} R{f(r.RadiusChamfer)} T{f(r.Tolerance)}".Trim();
                    }

                    var items = _presets.Select(Label).ToList();
                    ConnectorShapeCombobox.ItemsSource = items;

                    // Rewire handler
                    ConnectorShapeCombobox.SelectionChanged -= ConnectorPresetCombo_SelectionChanged;
                    ConnectorShapeCombobox.SelectionChanged += ConnectorPresetCombo_SelectionChanged;

                    if (ConnectorShapeCombobox.Items.Count > 0 && ConnectorShapeCombobox.SelectedIndex < 0)
                        ConnectorShapeCombobox.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to load connector presets:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        // ConnectorPresetCombo_SelectionChanged applies the selected preset values into the connector UI and redraws.
        private void ConnectorPresetCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var name = ConnectorShapeCombobox?.SelectedItem as string;
                if (string.IsNullOrWhiteSpace(name)) return;

                var preset = _presets.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
                if (preset == null) return;

                // Temporarily stop debounce while we overwrite many fields
                _uiDebounce?.Stop();

                ApplyPresetToUi(preset);

                // Force immediate redraw from final state
                drawConnector();
            }
            catch (Exception ex)
            {
            }
        }

        // ApplyPresetToUi writes preset values into the bound TextBoxes/radios and keeps corner state consistent.
        private void ApplyPresetToUi(ConnectorPresetRecord p)
        {
            string F(double? v) => v.HasValue ? v.Value.ToString("0.###", CultureInfo.InvariantCulture) : "";

            // Only overwrite the requested fields from CSV
            if (p.Height.HasValue) connectorHeight.Text = F(p.Height);
            if (p.Width1.HasValue) connectorWidth1.Text = F(p.Width1);
            if (p.Width2.HasValue) connectorWidth2.Text = F(p.Width2);
            if (p.Tolerance.HasValue) connectorTolerance.Text = F(p.Tolerance);

            // Radius/Chamfer: set the shared value textbox and the radio choice
            if (p.RadiusChamfer.HasValue) connectorRadiusChamfer.Text = F(p.RadiusChamfer);

            var style = (p.CornerStyle ?? "").Trim().ToLowerInvariant();
            if (style == "radius")
            {
                connectorRadius.IsChecked = true;
                connectorChamfer.IsChecked = false;
            }
            else if (style == "chamfer")
            {
                connectorRadius.IsChecked = false;
                connectorChamfer.IsChecked = true;
            }

            EnforceCornerCoupling();
            UpdateGenerateEnabled();
        }

        // DetectEncoding inspects the file BOM and returns the appropriate text encoding, defaulting to UTF8 without BOM.
        private static Encoding DetectEncoding(string path)
        {
            using var fs = File.OpenRead(path);
            using var br = new BinaryReader(fs, Encoding.UTF8, true);
            var bom = br.ReadBytes(4);
            if (bom.Length >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) return Encoding.UTF8;
            if (bom.Length >= 2 && bom[0] == 0xFF && bom[1] == 0xFE) return Encoding.Unicode;
            if (bom.Length >= 2 && bom[0] == 0xFE && bom[1] == 0xFF) return Encoding.BigEndianUnicode;
            if (bom.Length >= 4 && bom[0] == 0xFF && bom[1] == 0xFE && bom[2] == 0x00 && bom[3] == 0x00) return Encoding.UTF32;
            return new UTF8Encoding(false);
        }

        // SplitCsv splits a simple delimiter-separated line into trimmed field strings.
        private static string[] SplitCsv(string line, char delimiter)
        {
            // simple split (fields expected to be plain values)
            return line.Split(delimiter).Select(s => s.Trim()).ToArray();
        }

        // Get reads a typed value from the CSV header/column map or returns a fallback if not present or invalid.
        private static T Get<T>(Dictionary<string, int> headers, string[] cols, string key, T fallback)
        {
            if (!headers.TryGetValue(key, out var i) || i < 0 || i >= cols.Length) return fallback;
            var s = cols[i];
            if (typeof(T) == typeof(string)) return (T)(object)s;
            try
            {
                return (T)Convert.ChangeType(s, typeof(T), CultureInfo.InvariantCulture);
            }
            catch { return fallback; }
        }

        // GetDouble parses a nullable double from a CSV column using invariant culture semantics.
        private static double? GetDouble(Dictionary<string, int> headers, string[] cols, string key)
        {
            if (!headers.TryGetValue(key, out var i) || i < 0 || i >= cols.Length) return null;
            if (double.TryParse(cols[i], NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) return v;
            return null;
        }

        // GetBool parses a tolerant boolean flag from a CSV column (supports 1/0, y/n, yes/no, true/false).
        private static bool? GetBool(Dictionary<string, int> headers, string[] cols, string key)
        {
            if (!headers.TryGetValue(key, out var i) || i < 0 || i >= cols.Length) return null;
            var s = (cols[i] ?? "").Trim().ToLowerInvariant();
            if (bool.TryParse(s, out var b)) return b;
            if (s == "1" || s == "y" || s == "yes") return true;
            if (s == "0" || s == "n" || s == "no") return false;
            return null;
        }

        // getFacesFromSelection selects the larger and smaller adjacent faces for a given edge based on area.
        private static void getFacesFromSelection(IDesignEdge selectedEdge, out DesignFace bigFace, out DesignFace smallFace)
        {
            bigFace = null;
            smallFace = null;
            DesignEdge desEdge = selectedEdge.Master;

            DesignFace face0 = desEdge.Faces.ElementAt(0);
            DesignFace face1 = desEdge.Faces.ElementAt(1);

            if (desEdge.Faces.Count != 2)
                return;

            if (face0.Area > face1.Area)
            {
                bigFace = face0;
                smallFace = face1;
            }
            else if (face0.Area < face1.Area)
            {
                bigFace = face1;
                smallFace = face0;
            }
            else
                return;
        }

        // FindExitDistance_NoCopy walks a ray through a body in its master space to find the exit distance from a starting point.
        private static double FindExitDistance_NoCopy(
            Body masterBody,             // body in its own MASTER space
            Point pStartMaster,          // start in same MASTER space
            Direction dirMaster,         // ray direction in MASTER space
            double maxT,
            double tol = 5e-6,
            int maxIters = 40)
        {
            Vector dV = dirMaster.ToVector();
            bool inside0 = false;
            try { inside0 = masterBody.ContainsPoint(pStartMaster); } catch { return double.PositiveInfinity; }

            double t = 0.0;
            double step = Math.Max(0.0005, maxT / 50.0);
            const double GROW = 1.75;

            bool lastInside = inside0;
            while (t <= maxT)
            {
                t += step;
                Point p = pStartMaster + t * dV;
                bool nowInside = lastInside;
                try { nowInside = masterBody.ContainsPoint(p); } catch { step *= GROW; continue; }

                if (nowInside != lastInside)
                {
                    double a = Math.Max(0.0, t - step);
                    double b = t;
                    for (int i = 0; i < maxIters && (b - a) > tol; i++)
                    {
                        double m = 0.5 * (a + b);
                        Point pm = pStartMaster + m * dV;
                        bool inM = nowInside;
                        try { inM = masterBody.ContainsPoint(pm); } catch { }
                        if (inM == nowInside) b = m; else a = m;
                    }
                    return 0.5 * (a + b);
                }
                lastInside = nowInside;
                step *= GROW;
            }
            return double.PositiveInfinity;
        }

        // FindExitDistance_MasterSpace transforms a world-space ray into another body's master space and delegates to FindExitDistance_NoCopy.
        private static double FindExitDistance_MasterSpace(
            Body otherMaster,
            Matrix worldToOtherMaster,
            Point pStartWorld,
            Direction dirWorld,
            double maxT,
            double tol = 5e-6,
            int maxIters = 40)
        {
            var pStartM = worldToOtherMaster * pStartWorld;
            var dirM = (worldToOtherMaster * dirWorld.ToVector()).Direction;
            return FindExitDistance_NoCopy(otherMaster, pStartM, dirM, maxT, tol, maxIters);
        }

        // ComputeAvailableHeightAlongRay determines how much connector height fits along a ray before colliding with surrounding bodies.
        internal static double ComputeAvailableHeightAlongRay(
            Part searchRoot,
            DesignBody ownerBody,
            IDesignBody ownerOcc,
            Point p1,
            Direction dir,
            double userHeight,
            double thickness,
            Direction shiftThickness,
            double projectTol = 1e-5)
        {
            double req = Math.Max(0.0, userHeight);
            if (req <= 0) return 0.0;

            double CheckOneDirection(Direction testDir)
            {
                const double nudge = 2e-6;
                double probeLen = Math.Max(1.0, req * 2);
                double probeR = 0.0005;

                Point pStartWorld = p1 + nudge * testDir;

                // Build the tiny probe ONCE in WORLD, with its extrusion axis aligned to testDir
                Direction x = Direction.Cross(testDir, Vector.Create(1, 0, 0).Direction);
                if (x.IsZero) x = Direction.Cross(testDir, Vector.Create(0, 1, 0).Direction);
                if (x.IsZero) x = Direction.Cross(testDir, Vector.Create(0, 0, 1).Direction);

                // Choose y so that Cross(x, y) equals testDir (plane normal drives the extrude direction)
                Direction y = Direction.Cross(testDir, x);

                var probePlaneWorld = Plane.Create(Frame.Create(pStartWorld, x, y));
                using var probeWorld = Body.ExtrudeProfile(new CircleProfile(probePlaneWorld, probeR), probeLen);


                double best = double.PositiveInfinity;

                // Iterate occurrences WITHOUT copying their masters
                foreach (IDesignBody idb in searchRoot.GetDescendants<IDesignBody>())
                {
                    var master = idb.Master?.Shape;
                    if (master == null) continue;

                    // Map the light probe into the other body's MASTER space
                    using var probeInOther = probeWorld.Copy();
                    probeInOther.Transform(idb.TransformToMaster);

                    Collision cstat;
                    try { cstat = master.GetCollision(probeInOther); }
                    catch { continue; }

                    if (cstat == Collision.Intersect)
                    {
                        double tExit = FindExitDistance_MasterSpace(
                            master, idb.TransformToMaster, pStartWorld, testDir, Math.Max(req, 1.0));
                        if (tExit > 0 && !double.IsInfinity(tExit) && !double.IsNaN(tExit))
                            best = Math.Min(best, tExit);
                    }
                }

                return best;
            }

            double bestFwd = CheckOneDirection(dir);
            double bestRev = CheckOneDirection(-dir);

            double best = double.PositiveInfinity;
            if (bestFwd > 0 && !double.IsInfinity(bestFwd)) best = bestFwd;
            if (bestRev > 0 && !double.IsInfinity(bestRev)) best = Math.Min(best, bestRev);

            if (double.IsInfinity(best)) return req;

            double available = Math.Max(0.0, best);
            return Math.Max(0.0, Math.Min(req, available));
        }

        // createConnector is the main entry that builds connector geometries on selected edges and applies all booleans and neighbour cuts.
        private void createConnector()
        {
            static bool IsAttachedLocal(DesignBody ownerMaster, Body tool)
            {
                if (ownerMaster?.Shape == null || tool == null) return false;
                try
                {
                    var c = ownerMaster.Shape.GetCollision(tool);
                    return c == Collision.Intersect;// || c == Collision.Touch;
                }
                catch { return false; }
            }
            static bool IsAttachedLocalCyl(DesignBody ownerMaster, Body tool)
            {
                if (ownerMaster?.Shape == null || tool == null) return false;
                try
                {
                    var c = ownerMaster.Shape.GetCollision(tool);
                    return c == Collision.Intersect || c == Collision.Touch;
                }
                catch { return false; }
            }

            static bool ShouldExtendPlanarLocal(
                Part mainPart,
                IDesignBody ownerOcc,
                DesignBody ownerMaster,
                Body connectorBody,
                Body cutBody)
            {
                try
                { 
                    using var ownerCopy = ownerMaster.Shape.Copy();
                    int ownerfaces = ownerCopy.Faces.Count;
                    using var connCopy = connectorBody.Copy();
                    ownerCopy.Unite(new[] { connCopy });
                    Logger.Log("ShouldExtendPlanarLocal: No extension needed");
                    // If this try doesn't revert to catch, unite() was successful. No extension needed.
                    return false;
                }
                catch
                {
                    Logger.Log("ShouldExtendPlanarLocal: Unite failed, extend");
                    return true;
                }
            }

            s_neighbourDecisionCache.Clear();

            string rid = Guid.NewGuid().ToString("N").Substring(0, 8);
            var swAll = System.Diagnostics.Stopwatch.StartNew();
            bool rectangularCut = connectorRectangularCut?.IsChecked == true;

            const bool LOG_CYL = true;

            string P(Point p) => $"({p.X:0.######},{p.Y:0.######},{p.Z:0.######})";
            string D(Direction d)
            {
                var v = d.ToVector();
                return $"({v.X:0.###},{v.Y:0.###},{v.Z:0.###})";
            }

            void LogRun(string msg)
            {
                if (!LOG_CYL) return;
                try { Logger.Log($"[Connector][{rid}] {msg}"); } catch { }
            }

            void LogCyl(int edgeNo, int placeNo, string msg)
            {
                if (!LOG_CYL) return;
                try { Logger.Log($"[Connector][{rid}][E{edgeNo} P{placeNo}][CYL] {msg}"); } catch { }
            }

            try
            {
                if (!IsReliefAllowed(out var why))
                {
                    Application.ReportStatus(why, StatusMessageType.Warning, null);
                    return;
                }

                var win = Window.ActiveWindow;
                if (win == null) return;

                var ctx = win.ActiveContext;
                var doc = win.Document;
                var mainPart = doc.MainPart;

                if (!CheckSelectedEdgeWidth())
                {
                    Application.ReportStatus("Bottom width can't be larger than edge width.", StatusMessageType.Error, null);
                    return;
                }

                var c = ConnectorModel.CreateConnector(this);
                if (c == null)
                {
                    Application.ReportStatus("Please fill valid numeric values first.", StatusMessageType.Warning, null);
                    return;
                }

                LogRun(
                    $"UI W1={c.Width1}mm W2={c.Width2}mm H={c.Height}mm Tol={c.Tolerance}mm EndRelief={c.EndRelief}mm " +
                    $"DynamicHeight={c.DynamicHeight} RectCut={rectangularCut} Straight={c.ConnectorStraight} " +
                    $"CornerCutout={c.HasCornerCutout} CornerR={c.CornerCutoutRadius} " +
                    $"TopPair={c.RadiusInCutOut} TopPairR={c.RadiusInCutOut_Radius}"
                );

                bool patternEnabled = (connectorPattern?.IsChecked == true);

                int patternQty = 0;
                if (patternEnabled)
                    int.TryParse(connectorPatternValue?.Text ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out patternQty);
                patternEnabled = patternEnabled && patternQty > 0;
                if (patternEnabled && patternQty <= 0)
                {
                    Application.ReportStatus("Pattern quantity must be at least 1.", StatusMessageType.Error, null);
                    return;
                }

                // --- collect edges (same logic, but only one scene scan when needed) ---
                var iEdges = ctx.Selection.OfType<IDesignEdge>().ToList();
                if (iEdges.Count == 0)
                {
                    // gather any DesignEdge masters
                    var masters = ctx.Selection.OfType<DesignEdge>().ToList();
                    if (masters.Count > 0)
                    {
                        // single pass over all IDesignEdge occurrences
                        var occEdges = mainPart.GetDescendants<IDesignEdge>().ToList();
                        foreach (var de in masters)
                        {
                            var occ = occEdges.FirstOrDefault(e => e.Master == de);
                            if (occ != null) iEdges.Add(occ);
                        }
                    }
                    else
                    {
                        foreach (var sel in ctx.Selection)
                            if (sel is IDesignEdge ide) iEdges.Add(ide);
                    }
                }
                if (iEdges.Count == 0)
                {
                    Application.ReportStatus("Select one or more edges.", StatusMessageType.Information, null);
                    return;
                }

                // --- precompute placements (unchanged logic) ---
                var allPlacements = new List<(IDesignEdge edge, IDesignBody iDesignBody, DesignBody designBody, List<(Point pOcc, Direction tanOcc)> centers)>();
                foreach (var iEdge in iEdges)
                {
                    try
                    {
                        if (iEdge == null || iEdge.Parent == null) continue;
                        var iDesignBody = iEdge.Parent as IDesignBody;
                        var designBody = iDesignBody?.Master;
                        var partMaster = designBody?.Parent;
                        if (iDesignBody == null || designBody == null || partMaster == null) continue;

                        var centers = new List<(Point pOcc, Direction tanOcc)>();
                        if (!patternEnabled)
                        {
                            if (!checkDesignEdge(iEdge, c.ClickPosition, out var pOcc))
                                continue;

                            var tan = iEdge.Shape.ProjectPoint(pOcc).Derivative.Direction;
                            pOcc = pOcc + (0.001 * c.Location) * tan;
                            centers.Add((pOcc, tan));
                        }
                        else
                        {
                            centers.AddRange(ComputePatternCentersOcc(iEdge, c, patternQty));

                        }

                        if (centers.Count > 0)
                            allPlacements.Add((iEdge, iDesignBody, designBody, centers));

                    }
                    catch {}
                }

                if (allPlacements.Count == 0)
                {
                    Application.ReportStatus("No valid edges found to place connectors.", StatusMessageType.Information, null);
                    return;
                }

                // --- Single undo step (unchanged) ---
                WriteBlock.ExecuteTask("Connectors (per-edge commit)", () =>
                {
                    int edgeCounter = 0;
                    foreach (var entry in allPlacements)
                    {
                        edgeCounter++;
                        var iEdge = entry.edge;
                        var iDesignBody = entry.iDesignBody;
                        var ownerMaster = iDesignBody.Master;
                        var parentPart = ownerMaster.Parent;
                        var winMainPart = Window.ActiveWindow.Document.MainPart;

                        var occToMaster = iDesignBody.TransformToMaster;
                        var masterToOcc = occToMaster.Inverse;

                        // per-edge accumulators with capacity (reduce reallocations)
                        var uniteList = new List<Body>(entry.centers.Count);
                        var cutPropList = new List<Body>(entry.centers.Count);
                        var collList = new List<Body>(entry.centers.Count);
                        var ownerCuts = new List<Body>(entry.centers.Count);

                        // --- compute faces once per edge (no functional change) ---
                        getFacesFromSelection(iEdge, out DesignFace bigFaceEdge, out DesignFace smallFaceEdge);
                        if (bigFaceEdge == null || smallFaceEdge == null) continue;

                        bool isPlaneEdge = bigFaceEdge.Shape.Geometry is Plane;

                        LogRun($"Edge #{edgeCounter} placements={entry.centers.Count} isPlaneEdge={isPlaneEdge} owner={ownerMaster?.Name}");

                        // Cylindrical data per edge (used only in cylindrical branch)
                        DesignFace innerFaceEdge = null, outerFaceEdge = null;
                        double thicknessEdge = 0, outerRadiusEdge = 0;
                        Line axisEdge = null;
                        if (!isPlaneEdge)
                        {
                            var cyl = CylInfo.GetCoaxialCylPair(iDesignBody, bigFaceEdge);
                            innerFaceEdge = cyl.InnerFace;
                            outerFaceEdge = cyl.OuterFace;
                            thicknessEdge = cyl.Thickness;
                            outerRadiusEdge = cyl.OuterRadius;
                            axisEdge = cyl.Axis;

                            var bigCyl = bigFaceEdge.Shape.Geometry as Cylinder;
                            var innerCyl = innerFaceEdge?.Shape?.Geometry as Cylinder;
                            var outerCyl = outerFaceEdge?.Shape?.Geometry as Cylinder;

                            LogRun(
                                $"CylPair axisO={P(axisEdge?.Origin ?? Point.Origin)} axisD={D(axisEdge?.Direction ?? Direction.DirZ)} " +
                                $"outerR={outerRadiusEdge:0.######}m thick={thicknessEdge:0.######}m " +
                                $"bigR={(bigCyl != null ? bigCyl.Radius.ToString("0.######") : "na")} " +
                                $"innerR={(innerCyl != null ? innerCyl.Radius.ToString("0.######") : "na")} " +
                                $"outerFaceR={(outerCyl != null ? outerCyl.Radius.ToString("0.######") : "na")}"
                            );

                            if (innerFaceEdge == null || outerFaceEdge == null || axisEdge == null)
                            {
                                LogRun("CylPair invalid: innerFace or outerFace or axis is null");
                                Application.ReportStatus($"No inner/outer Face found", StatusMessageType.Information, null);
                                continue;
                            }
                        }

                        // Cylindrical shell prototypes (build once per edge; copy per placement)
                        Body protoCylInner = null, protoCylOuter = null, protoCylAnnulus = null;
                        if (!isPlaneEdge)
                        {
                            Plane shellPlane = Plane.Create(Frame.Create(axisEdge.Origin - 10 * axisEdge.Direction, axisEdge.Direction));
                            protoCylInner = Body.ExtrudeProfile(new CircleProfile(shellPlane, outerRadiusEdge - thicknessEdge), 20);
                            protoCylOuter = Body.ExtrudeProfile(new CircleProfile(shellPlane, outerRadiusEdge), 20);
                            protoCylAnnulus = Body.ExtrudeProfile(new CircleProfile(shellPlane, outerRadiusEdge + 1), 20);

                            try { protoCylAnnulus.Subtract(protoCylOuter.Copy()); }
                            catch (Exception ex) { LogRun($"Proto annulus subtract failed: {ex.Message}"); }
                        }

                        int placementIdx = 0;

                        foreach (var (pCenterOcc, tanOcc) in entry.centers)
                        {
                            placementIdx++;
                            //Logger.Log("placementId " + placementIdx.ToString());

                            Body connectorBody = null, cutBody = null, collisionBody = null;
                            var cutBodiesSource = new List<Body>();

                            if (isPlaneEdge)
                            {
                                //Logger.Log("planar");
                                // === PLANAR ===
                                var pCenterM = occToMaster * pCenterOcc;

                                Direction nBigM = ((Plane)bigFaceEdge.Shape.Geometry).Frame.DirZ;
                                if (bigFaceEdge.Shape.IsReversed) nBigM = -nBigM;

                                Direction tanM_raw = (occToMaster * tanOcc.ToVector()).Direction;

                                Vector tanM_vec = tanM_raw.ToVector();
                                Vector nM_vec = nBigM.ToVector();
                                double tanDotN = Vector.Dot(tanM_vec, nM_vec);
                                Vector tanProj = (tanM_vec - tanDotN * nM_vec);

                                Direction dirX_m;
                                if (tanProj.Magnitude < 1e-12)
                                {
                                    var tryX = Vector.Create(1, 0, 0);
                                    if (Math.Abs(Vector.Dot(tryX, nM_vec)) > 0.95) tryX = Vector.Create(0, 1, 0);
                                    var xProj = (tryX - Vector.Dot(tryX, nM_vec) * nM_vec).Direction;
                                    dirX_m = xProj;
                                }
                                else
                                {
                                    dirX_m = tanProj.Direction;
                                }

                                Direction dirY_m = nBigM;                  // sheet/extrude normal
                                Direction dirZ_m = Direction.Cross(dirX_m, dirY_m);
                                dirX_m = Direction.Cross(dirY_m, dirZ_m);

                                Direction dirZ_w = (masterToOcc * dirZ_m.ToVector()).Direction;
                                try
                                {
                                    var bodyM = ownerMaster?.Shape;
                                    if (bodyM != null && bodyM.ContainsPoint(pCenterM + 1e-5 * dirZ_m))
                                    {
                                        dirY_m = Direction.Cross(dirZ_m, dirX_m);
                                        dirZ_m = -dirZ_m;
                                        dirX_m = Direction.Cross(dirY_m, dirZ_m);
                                    }
                                }
                                catch { }

                                // compute thickness once per edge via opposite face (small<->big pairing unchanged)
                                var opp = getOppositeFace(smallFaceEdge, bigFaceEdge, isPlaneEdge, out double thickness, out Direction _);
                                if (opp == null) continue;

                                double usedHeight = 0.001 * c.Height; // UI height (m)
                                if (c.DynamicHeight)
                                {
                                    usedHeight = ComputeAvailableHeightAlongRay(
                                        winMainPart, ownerMaster, iDesignBody, pCenterOcc,
                                        -dirZ_w, usedHeight, thickness, dirY_m
                                    );
                                }

                                double tolM = Math.Max(0.0, c.Tolerance) * 0.001;
                                double gapM = Math.Max(0.0, c.EndRelief) * 0.001;

                                double connectorHeightM = Math.Max(0.0, usedHeight - gapM);

                                double baseRectHeightM = rectangularCut
                                    ? (c.DynamicHeight ? (connectorHeightM + gapM) : (0.001 * c.Height))
                                    : usedHeight;

                                double cutterHeightM = Math.Max(0.0, baseRectHeightM);

                                // reuse existing instance 'c' (no new Connector allocations)
                                c.CreateGeometry(
                                    parentPart, dirX_m, dirY_m, dirZ_m, pCenterM,
                                    cutterHeightM, connectorHeightM, thickness,
                                    out connectorBody, out cutBodiesSource, out cutBody, out collisionBody, false,
                                    null, rectangularCut
                                );

                                bool doExtend = ShouldExtendPlanarLocal(
                                    winMainPart,
                                    iDesignBody,
                                    ownerMaster,
                                    connectorBody,
                                    cutBody);

                                // Always build both base and extended geometries, then merge results (unchanged policy)
                                var extCenter = pCenterM;                 // no downward shift
                                var extConnHeight = connectorHeightM;     // same height

                                Body extConnector = null, extCut = null, extCollision = null;
                                var extOwnerCuts = new List<Body>();

                                if (doExtend)
                                {
                                    c.CreateGeometry(
                                        parentPart,
                                        dirX_m,
                                        dirY_m,
                                        -dirZ_m,            // ← flipped direction only
                                        extCenter,
                                        cutterHeightM,
                                        extConnHeight,
                                        thickness,
                                        out extConnector,
                                        out extOwnerCuts,
                                        out extCut,
                                        out extCollision,
                                        false,
                                        null,
                                        rectangularCut,
                                        false // disable corner features for extended pass
                                    );
                                }

                                try
                                {
                                    if (extConnector != null && connectorBody != null)
                                    {
                                        connectorBody.Unite(extConnector.Copy());
                                    }
                                    else if (connectorBody == null && extConnector != null)
                                    {
                                        connectorBody = extConnector.Copy();
                                    }

                                    if (extCut != null && cutBody != null)
                                    {
                                        cutBody.Unite(extCut.Copy());
                                    }
                                    else if (cutBody == null && extCut != null)
                                    {
                                        cutBody = extCut.Copy();
                                    }

                                    if (extCollision != null && collisionBody != null)
                                    {
                                        collisionBody.Unite(extCollision.Copy());
                                    }
                                    else if (collisionBody == null && extCollision != null)
                                    {
                                        collisionBody = extCollision.Copy();
                                    }

                                    foreach (var b in extOwnerCuts)
                                    {
                                        cutBodiesSource.Add(b.Copy());
                                    }
                                }
                                catch (Exception exUnite)
                                {
                                }
                                finally
                                {
                                    try { extConnector?.Dispose(); } catch { }
                                    try { extCut?.Dispose(); } catch { }
                                    try { extCollision?.Dispose(); } catch { }
                                    foreach (var b in extOwnerCuts) { try { b?.Dispose(); } catch { } }
                                }

                                if (!rectangularCut && connectorBody != null && !c.RadiusInCutOut)
                                {
                                    var inflated = connectorBody.Copy();
                                    try
                                    {
                                        double grow = Math.Max(1e-6, tolM);
                                        inflated.OffsetFaces(inflated.Faces, grow);
                                        cutBody?.Dispose();
                                        cutBody = inflated;
                                        collisionBody?.Dispose();
                                        collisionBody = inflated.Copy();
                                    }
                                    catch { inflated?.Dispose(); }
                                }
                            }
                            // AFTER (full cylindrical branch, no omissions)
                            // Fix: force c.Height to connectorHeightM only while building the connector loft,
                            // then force c.Height to cutterHeightM while building rectLoft and rectCutLoft,
                            // so the cutter becomes (height + tol) and is not influenced by end relief.

                            else
                            {
                                // === CYLINDRICAL === (face pair computed once per edge)
                                try
                                {
                                    double w1m = c.Width1 / 1000.0;
                                    double w2m = c.Width2 / 1000.0;

                                    LogCyl(edgeCounter, placementIdx, $"Widths w1m={w1m:0.######} w2m={w2m:0.######} outerDiamMm={outerRadiusEdge * 2000.0:0.###}");

                                    if (outerRadiusEdge * 2 < w1m)
                                    {
                                        LogCyl(edgeCounter, placementIdx, "Abort: width1 exceeds diameter check");
                                        Application.ReportStatus($"Too wide, width should be less than: {outerRadiusEdge * 2000} mm", StatusMessageType.Information, null);
                                        continue;
                                    }

                                    var pCenter = occToMaster * pCenterOcc;
                                    LogCyl(edgeCounter, placementIdx, $"pCenterM={P(pCenter)} pCenterOcc={P(pCenterOcc)}");

                                    double distSelectionPoint2Axis = (axisEdge.ProjectPoint(pCenter).Point - pCenter).Magnitude;
                                    bool selectedInnerFace = outerRadiusEdge - distSelectionPoint2Axis > 1e-5;
                                    DesignFace selectedFace = selectedInnerFace ? innerFaceEdge : outerFaceEdge;

                                    LogCyl(edgeCounter, placementIdx, $"distToAxis={distSelectionPoint2Axis:0.######} selectedFace={(selectedInnerFace ? "INNER" : "OUTER")}");

                                    Point pAxis = axisEdge.ProjectPoint(pCenter).Point;
                                    Direction dirPoint2Axis = (pAxis - pCenter).Direction;
                                    Direction dirZ = axisEdge.Direction;

                                    bool insideTest = false;
                                    try
                                    {
                                        insideTest = ownerMaster.Shape.ContainsPoint(pCenter + 1e-5 * dirZ);
                                    }
                                    catch (Exception ex)
                                    {
                                        LogCyl(edgeCounter, placementIdx, $"Owner ContainsPoint test failed: {ex.Message}");
                                    }

                                    LogCyl(edgeCounter, placementIdx, $"Axis pAxis={P(pAxis)} dirP2A={D(dirPoint2Axis)} dirZ_pre={D(dirZ)} containsTest={insideTest}");

                                    if (insideTest) dirZ = -dirZ;

                                    LogCyl(edgeCounter, placementIdx, $"dirZ_post={D(dirZ)}");

                                    Direction dirX = Direction.Cross(dirPoint2Axis, dirZ);

                                    double dXZ = Math.Abs(Vector.Dot(dirX.ToVector(), dirZ.ToVector()));
                                    double dXP = Math.Abs(Vector.Dot(dirX.ToVector(), dirPoint2Axis.ToVector()));
                                    double dZP = Math.Abs(Vector.Dot(dirZ.ToVector(), dirPoint2Axis.ToVector()));
                                    LogCyl(edgeCounter, placementIdx, $"Frame dot(X,Z)={dXZ:0.######} dot(X,P2A)={dXP:0.######} dot(Z,P2A)={dZP:0.######}");

                                    double maxWidth = Math.Max(w1m, w2m);
                                    double under = outerRadiusEdge * outerRadiusEdge - (0.5 * maxWidth) * (0.5 * maxWidth);
                                    LogCyl(edgeCounter, placementIdx, $"maxWidth={maxWidth:0.######} underSqrt={under:0.######}");

                                    double distWidth = Math.Sqrt(under);
                                    Point pWidth = pAxis - distWidth * dirPoint2Axis;
                                    Point pWidth_A = pAxis - distWidth * dirPoint2Axis + 0.5 * maxWidth * dirX;
                                    Point pWidth_B = pAxis - distWidth * dirPoint2Axis - 0.5 * maxWidth * dirX;

                                    LogCyl(edgeCounter, placementIdx, $"pWidth={P(pWidth)} pWidthA={P(pWidth_A)} pWidthB={P(pWidth_B)}");

                                    Direction dir_A = c.ConnectorStraight ? dirPoint2Axis : (pAxis - pWidth_A).Direction;
                                    Direction dir_B = c.ConnectorStraight ? dirPoint2Axis : (pAxis - pWidth_B).Direction;
                                    Point pInner_A = pWidth_A + thicknessEdge * dir_A;
                                    Point pInner_B = pWidth_B + thicknessEdge * dir_B;

                                    Point pInnerMid = pInner_A - 0.5 * (pInner_A - pInner_B).Magnitude * dirX;
                                    Point p_InnerFace = pAxis - (outerRadiusEdge - thicknessEdge) * dirPoint2Axis;
                                    Point p_OuterFace = pAxis - (outerRadiusEdge) * dirPoint2Axis;

                                    LogCyl(edgeCounter, placementIdx, $"dirA={D(dir_A)} dirB={D(dir_B)} pInnerA={P(pInner_A)} pInnerB={P(pInner_B)} pInnerMid={P(pInnerMid)}");
                                    LogCyl(edgeCounter, placementIdx, $"pInnerFaceRef={P(p_InnerFace)} pOuterFaceRef={P(p_OuterFace)}");

                                    double alpha = Math.Acos(distWidth / outerRadiusEdge);
                                    double distOuter = outerRadiusEdge / Math.Cos(alpha);

                                    LogCyl(edgeCounter, placementIdx, $"alphaDeg={(alpha * 180.0 / Math.PI):0.###} distOuter={distOuter:0.######}");

                                    Point pOuter_A = pAxis - distOuter * dir_A;
                                    Point pOuter_B = pAxis - distOuter * dir_B;

                                    var (success, innerEdge, outerEdge) = CylInfo.GetEdges(innerFaceEdge, p_InnerFace, outerFaceEdge, p_OuterFace);
                                    LogCyl(edgeCounter, placementIdx, $"GetEdges success={success} innerEdgeNull={(innerEdge == null)} outerEdgeNull={(outerEdge == null)}");

                                    var (successInnerA, p_InnerEdge_A) = CylInfo.GetClosestPoint(innerEdge, pInner_A, dirZ);
                                    var (successInnerB, p_InnerEdge_B) = CylInfo.GetClosestPoint(innerEdge, pInner_B, dirZ);
                                    var (successOuterA, p_OuterEdge_A) = CylInfo.GetClosestPoint(outerEdge, pWidth_A, dirZ);
                                    var (successOuterB, p_OuterEdge_B) = CylInfo.GetClosestPoint(outerEdge, pWidth_B, dirZ);

                                    LogCyl(edgeCounter, placementIdx, $"ClosestPoints innerA={successInnerA} innerB={successInnerB} outerA={successOuterA} outerB={successOuterB}");

                                    double maxDistanceBottom = 0;

                                    if (successInnerA)
                                    {
                                        var dv = pInner_A - p_InnerEdge_A;
                                        bool ok = dv.Direction == dirZ;
                                        LogCyl(edgeCounter, placementIdx, $"InnerA snap dist={dv.Magnitude:0.######} dirOk={ok}");
                                        if (ok) maxDistanceBottom = Math.Max(maxDistanceBottom, dv.Magnitude);
                                    }
                                    if (successInnerB)
                                    {
                                        var dv = pInner_B - p_InnerEdge_B;
                                        bool ok = dv.Direction == dirZ;
                                        LogCyl(edgeCounter, placementIdx, $"InnerB snap dist={dv.Magnitude:0.######} dirOk={ok}");
                                        if (ok) maxDistanceBottom = Math.Max(maxDistanceBottom, dv.Magnitude);
                                    }
                                    if (successOuterA)
                                    {
                                        var dv = pWidth_A - p_OuterEdge_A;
                                        bool ok = dv.Direction == dirZ;
                                        LogCyl(edgeCounter, placementIdx, $"OuterA snap dist={dv.Magnitude:0.######} dirOk={ok}");
                                        if (ok) maxDistanceBottom = Math.Max(maxDistanceBottom, dv.Magnitude);
                                    }
                                    if (successOuterB)
                                    {
                                        var dv = pWidth_B - p_OuterEdge_B;
                                        bool ok = dv.Direction == dirZ;
                                        LogCyl(edgeCounter, placementIdx, $"OuterB snap dist={dv.Magnitude:0.######} dirOk={ok}");
                                        if (ok) maxDistanceBottom = Math.Max(maxDistanceBottom, dv.Magnitude);
                                    }

                                    LogCyl(edgeCounter, placementIdx, $"maxDistanceBottom={maxDistanceBottom:0.######}");

                                    Plane planeInner = Plane.Create(Frame.Create(pInnerMid, dirX, dirZ));
                                    Plane planeOuter = Plane.Create(Frame.Create(p_OuterFace, dirX, dirZ));

                                    double widthDiffInner = c.ConnectorStraight ? 0 : (pWidth_A - pWidth_B).Magnitude - (pInner_A - pInner_B).Magnitude;
                                    double widthDiffOuter = c.ConnectorStraight ? 0 : (pWidth_A - pWidth_B).Magnitude - (pOuter_A - pOuter_B).Magnitude;

                                    Direction dirZ_world = (masterToOcc * dirZ.ToVector()).Direction;

                                    // Thickness shift direction must go into the wall material
                                    // Outer face selected: into wall is toward axis
                                    // Inner face selected: into wall is away from axis
                                    Direction thickIntoWall_master = selectedInnerFace ? -dirPoint2Axis : dirPoint2Axis;
                                    Direction shiftThicknessWorld = (masterToOcc * thickIntoWall_master.ToVector()).Direction;

                                    double userHeightM = 0.001 * c.Height;
                                    double usedHeightM = c.DynamicHeight
                                        ? ComputeAvailableHeightAlongRay(
                                            Window.ActiveWindow.Document.MainPart,
                                            ownerMaster, iDesignBody,
                                            pCenterOcc,
                                            -dirZ_world,
                                            userHeightM, thicknessEdge, shiftThicknessWorld)
                                        : userHeightM;

                                    double tolM = Math.Max(0.0, c.Tolerance) * 0.001;
                                    double gapM = Math.Max(0.0, c.EndRelief) * 0.001;
                                    double connectorHeightM = Math.Max(0.0, usedHeightM - gapM);
                                    double cutterHeightM = usedHeightM;

                                    LogCyl(edgeCounter, placementIdx,
                                        $"Heights user={userHeightM} used={usedHeightM} gap={gapM} tol={tolM} " +
                                        $"connectorH={connectorHeightM} cutterH={cutterHeightM} dyn={c.DynamicHeight}"
                                    );

                                    double savedHeightMm = c.Height;

                                    try
                                    {
                                        LogCyl(edgeCounter, placementIdx, $"widthDiffInner={widthDiffInner:0.######} widthDiffOuter={widthDiffOuter:0.######} Straight={c.ConnectorStraight}");

                                        // Connector uses (usedHeight - gap)
                                        c.Height = connectorHeightM * 1000.0;

                                        var boundary_Outer = c.CreateBoundary(
                                            dirX, dirPoint2Axis, dirZ, p_OuterFace,
                                            widthDiffOuter, maxDistanceBottom, connectorHeightM);

                                        var boundary_Inner = c.CreateBoundary(
                                            dirX, dirPoint2Axis, dirZ, pInnerMid,
                                            widthDiffInner, maxDistanceBottom, connectorHeightM);

                                        LogCyl(edgeCounter, placementIdx,
                                            $"BoundaryHeight connectorHeightM={connectorHeightM:0.######} c.HeightMmNow={c.Height:0.###}");

                                        LogCyl(edgeCounter, placementIdx, $"Boundary counts outer={boundary_Outer?.Count ?? 0} inner={boundary_Inner?.Count ?? 0}");

                                        if (true && c.HasCornerCutout && c.CornerCutoutRadius > 0)
                                        {
                                            double rCC = 0.001 * c.CornerCutoutRadius;

                                            double dbgLen = thicknessEdge * 2.0;

                                            Direction toolAxisA = (pInner_A - pWidth_A).Direction;
                                            Direction toolAxisB = (pInner_B - pWidth_B).Direction;

                                            double shiftIn = -0.5 * thicknessEdge;

                                            Point pToolA = pWidth_A + shiftIn * toolAxisA;
                                            Point pToolB = pWidth_B + shiftIn * toolAxisB;

                                            Plane ccPlaneA = Plane.Create(Frame.Create(pToolA, toolAxisA));
                                            Plane ccPlaneB = Plane.Create(Frame.Create(pToolB, toolAxisB));

                                            // Build the actual subtraction tools (these are the ones we will subtract)
                                            Body cornerToolA = Body.ExtrudeProfile(new CircleProfile(ccPlaneA, rCC), dbgLen);
                                            Body cornerToolB = Body.ExtrudeProfile(new CircleProfile(ccPlaneB, rCC), dbgLen);

                                            // This is what makes them subtract from the owner AFTER unite
                                            // Only add these two bodies, and nothing else.
                                            cutBodiesSource.Add(cornerToolA);
                                            cutBodiesSource.Add(cornerToolB);
                                        }

                                        var connectorBodyLocal = c.CreateLoft(boundary_Outer, planeOuter, boundary_Inner, planeInner);

                                        // Shell copies used for subtraction
                                        var cylInnerCopy = protoCylInner?.Copy();
                                        var cylAnnulCopy = protoCylAnnulus?.Copy();

                                        try { if (cylInnerCopy != null) connectorBodyLocal.Subtract(cylInnerCopy); } catch { }
                                        try { if (cylAnnulCopy != null) connectorBodyLocal.Subtract(cylAnnulCopy); } catch { }

                                        connectorBody = connectorBodyLocal.Copy();

                                        // IMPORTANT FIX: switch c.Height back to full cutter height before building cutter/collision
                                        c.Height = cutterHeightM * 1000.0;
                                        LogCyl(edgeCounter, placementIdx,
                                            $"BoundaryHeight cutterHeightM={cutterHeightM:0.######} c.HeightMmNow={c.Height:0.###}");

                                        var savedRadius = c.Radius;
                                        var savedHasRounding = c.HasRounding;
                                        Body rectLoft = null;

                                        try
                                        {
                                            c.Radius = 0.0;
                                            c.HasRounding = false;

                                            var rectOuter = c.CreateBoundary(
                                                dirX, dirPoint2Axis, dirZ, p_OuterFace,
                                                widthDiffOuter, maxDistanceBottom, cutterHeightM);

                                            var rectInner = c.CreateBoundary(
                                                dirX, dirPoint2Axis, dirZ, pInnerMid,
                                                widthDiffInner, maxDistanceBottom, cutterHeightM);

                                            rectLoft = c.CreateLoft(rectOuter, planeOuter, rectInner, planeInner);
                                        }
                                        finally
                                        {
                                            c.Radius = savedRadius;
                                            c.HasRounding = savedHasRounding;
                                        }

                                        // shell subtraction for cutter/collision using prototypes
                                        var cylInnerCopy2 = protoCylInner?.Copy();
                                        var cylAnnulCopy2 = protoCylAnnulus?.Copy();
                                        if (cylInnerCopy2 != null) rectLoft.Subtract(cylInnerCopy2);
                                        if (cylAnnulCopy2 != null) rectLoft.Subtract(cylAnnulCopy2);

                                        // Collision stays the true rectangular loft without tolerance
                                        collisionBody = rectLoft;

                                        // Cut body depends on rectangularCut, matching planar behavior
                                        if (rectangularCut)
                                        {
                                            double halfW = 0.5 * maxWidth;
                                            double halfW_g = halfW + tolM;

                                            // Planar logic equivalent:
                                            // dynamic: top at used height
                                            // not dynamic: top at user height + tol
                                            double topH_g = cutterHeightM;

                                            // Keep c.Height full for the cutter tool
                                            c.Height = cutterHeightM * 1000.0;

                                            c.Radius = 0.0;
                                            c.HasRounding = false;

                                            var rectOuter = c.CreateBoundary(
                                                dirX, dirPoint2Axis, dirZ, p_OuterFace,
                                                widthDiffOuter, maxDistanceBottom, cutterHeightM);

                                            var rectInner = c.CreateBoundary(
                                                dirX, dirPoint2Axis, dirZ, pInnerMid,
                                                widthDiffInner, maxDistanceBottom, cutterHeightM);

                                            var rectCutLoft = c.CreateLoft(rectOuter, planeOuter, rectInner, planeInner);
                                            rectCutLoft.OffsetFaces(rectCutLoft.Faces, tolM);

                                            // Keep prototype subtractions consistent with your existing cutter behavior
                                            var cylInnerCopyCut = protoCylInner?.Copy();
                                            var cylAnnulCopyCut = protoCylAnnulus?.Copy();
                                            if (cylInnerCopyCut != null) rectCutLoft.Subtract(cylInnerCopyCut);
                                            if (cylAnnulCopyCut != null) rectCutLoft.Subtract(cylAnnulCopyCut);

                                            cutBody = rectCutLoft;

                                            LogCyl(edgeCounter, placementIdx,
                                                $"RectCut: halfW={halfW:0.######} halfW_g={halfW_g:0.######} topH_g={topH_g:0.######} maxDistanceBottom={maxDistanceBottom:0.######}");
                                        }
                                        else
                                        {
                                            // Non rectangular cut: inflate by face offset, same as before
                                            cutBody = rectLoft.Copy();
                                            try { cutBody.OffsetFaces(cutBody.Faces, tolM); } catch { }
                                        }

                                    }
                                    finally
                                    {
                                        c.Height = savedHeightMm;
                                    }

                                    // ---- Ensure attachment; else extend once (cylindrical) ----
                                    if (connectorBody != null)
                                    {
                                        try
                                        {
                                            var col = ownerMaster.Shape.GetCollision(connectorBody);
                                            LogCyl(edgeCounter, placementIdx, $"AttachCheck collision={col}");
                                        }
                                        catch (Exception ex)
                                        {
                                            LogCyl(edgeCounter, placementIdx, $"AttachCheck collision query failed: {ex.Message}");
                                        }
                                    }
                                    Logger.Log($"isnull{connectorBody == null}");
                                    Logger.Log($"isattached{IsAttachedLocalCyl(ownerMaster, connectorBody)}");
                                    if (connectorBody != null && !IsAttachedLocalCyl(ownerMaster, connectorBody))
                                    {
                                        LogCyl(edgeCounter, placementIdx, "Not attached, attempting extension");

                                        var extConnHeight = 2.0 * connectorHeightM;

                                        var p_OuterFace_ext = p_OuterFace - connectorHeightM * dirZ;
                                        var pInnerMid_ext = pInnerMid - connectorHeightM * dirZ;

                                        LogCyl(edgeCounter, placementIdx,
                                            $"Extension extConnHeight={extConnHeight:0.######} pOuterExt={P(p_OuterFace_ext)} pInnerMidExt={P(pInnerMid_ext)} dirZ={D(dirZ)}"
                                        );

                                        Plane planeInner_ext = Plane.Create(Frame.Create(pInnerMid_ext, dirX, dirZ));
                                        Plane planeOuter_ext = Plane.Create(Frame.Create(p_OuterFace_ext, dirX, dirZ));

                                        double savedHeightMm_Ext = c.Height;
                                        try
                                        {
                                            c.Height = extConnHeight * 1000.0;

                                            var boundary_Outer_ext = c.CreateBoundary(
                                                dirX, dirPoint2Axis, dirZ, p_OuterFace_ext,
                                                widthDiffOuter, maxDistanceBottom, extConnHeight);

                                            var boundary_Inner_ext = c.CreateBoundary(
                                                dirX, dirPoint2Axis, dirZ, pInnerMid_ext,
                                                widthDiffInner, maxDistanceBottom, extConnHeight);

                                            LogCyl(edgeCounter, placementIdx, $"Ext boundary counts outer={boundary_Outer_ext?.Count ?? 0} inner={boundary_Inner_ext?.Count ?? 0}");

                                            var connectorBodyLocal_ext = c.CreateLoft(boundary_Outer_ext, planeOuter_ext, boundary_Inner_ext, planeInner_ext);

                                            var cylInnerCopy3 = protoCylInner?.Copy();
                                            var cylAnnulCopy3 = protoCylAnnulus?.Copy();

                                            try { if (cylInnerCopy3 != null) connectorBodyLocal_ext.Subtract(cylInnerCopy3); }
                                            catch (Exception ex) { LogCyl(edgeCounter, placementIdx, $"Ext subtract protoInner failed: {ex.Message}"); }

                                            try { if (cylAnnulCopy3 != null) connectorBodyLocal_ext.Subtract(cylAnnulCopy3); }
                                            catch (Exception ex) { LogCyl(edgeCounter, placementIdx, $"Ext subtract protoAnnulus failed: {ex.Message}"); }

                                            var extConnector = connectorBodyLocal_ext.Copy();

                                            if (extConnector != null && IsAttachedLocalCyl(ownerMaster, extConnector))
                                            {
                                                try
                                                {
                                                    var col2 = ownerMaster.Shape.GetCollision(extConnector);
                                                    LogCyl(edgeCounter, placementIdx, $"Extension attached collision={col2}");
                                                }
                                                catch { }

                                                connectorBody?.Dispose();
                                                connectorBody = extConnector;
                                            }
                                            else
                                            {
                                                try
                                                {
                                                    var col2 = ownerMaster.Shape.GetCollision(extConnector);
                                                    LogCyl(edgeCounter, placementIdx, $"Extension still not attached collision={col2}");
                                                }
                                                catch { }

                                                try { extConnector?.Dispose(); } catch { }
                                                throw new InvalidOperationException("Connector could not attach to the owner body (after one extension).");
                                            }
                                        }
                                        finally
                                        {
                                            c.Height = savedHeightMm_Ext;
                                        }
                                    }

                                    if (!rectangularCut && connectorBody != null)
                                    {
                                        var inflated = connectorBody.Copy();
                                        try
                                        {
                                            double grow = Math.Max(1e-6, tolM);
                                            inflated.OffsetFaces(inflated.Faces, grow);
                                            cutBody?.Dispose(); cutBody = inflated;
                                            collisionBody?.Dispose(); collisionBody = inflated.Copy();
                                        }
                                        catch { inflated?.Dispose(); }
                                    }
                                }
                                catch (Exception exCyl)
                                {
                                    Application.ReportStatus($"Connector cylindrical geometry failed: {exCyl}", StatusMessageType.Error, null);
                                    continue;
                                }
                            }


                            // --- Accumulate this placement (kept identical) ---
                            if (connectorBody != null) uniteList.Add(connectorBody.Copy());
                            if (cutBody != null) cutPropList.Add(cutBody.Copy());
                            if (collisionBody != null) collList.Add(collisionBody.Copy());
                            foreach (var cb in cutBodiesSource) ownerCuts.Add(cb.Copy());
                        } // placements

                        // cleanup per-edge prototypes
                        try { protoCylInner?.Dispose(); } catch { }
                        try { protoCylOuter?.Dispose(); } catch { }
                        try { protoCylAnnulus?.Dispose(); } catch { }

                        try
                        {
                            var ownerOccSourceForTools = iDesignBody;

                            if (!ApplyOwnerEditsWithChoice(
                                mainPart,
                                iDesignBody,
                                ownerMaster,
                                parentPart,
                                uniteList,
                                ownerCuts,
                                out var choiceForNeighbors,
                                out var ownerOccForMapping))
                            {
                                continue;
                            }
                            if (collList.Count > 0 && cutPropList.Count > 0)
                            {
                                // neighbours are found around ownerOccForMapping,
                                // tool mapping uses the ORIGINAL owner occurrence
                                PropagateCutsForChoice(
                                    mainPart,
                                    ownerOccForMapping,
                                    ownerOccSourceForTools,
                                    parentPart,
                                    collList,
                                    cutPropList,
                                    choiceForNeighbors);
                            }
                        }
                        catch
                        {
                            // keep processing other edges
                        }
                    } // foreach edge
                }); // WriteBlock
            }
            catch (Exception ex)
            {
                Application.ReportStatus(ex.ToString(), StatusMessageType.Error, null);
            }
            finally
            {
                swAll.Stop();
            }
        }

        // MakeIndependentOcc executes MakeIndependent and finds the new occurrence corresponding to the original body.
        private static IDesignBody MakeIndependentOcc(IDesignBody occ)
        {
            if (occ == null) return null;

            var win = Window.ActiveWindow;
            var ctx = win?.ActiveContext;
            var mainPart = win?.Document?.MainPart;
            if (ctx == null || mainPart == null)
                return occ;

            var originalMaster = occ.Master;
            var originalPart = originalMaster?.Parent as Part;
            var originalComponent = originalPart?.Parent as Component;

            // Prefer the component name as stem source, fall back to body name
            var stemSourceName = originalComponent?.Name ?? originalMaster?.Name ?? string.Empty;

            // snapshot
            var originalHashCodes = new HashSet<int>();
            foreach (var idb in mainPart.GetDescendants<IDesignBody>())
                originalHashCodes.Add(idb.GetHashCode());

            var occInCtx = ctx.MapToContext<IDesignBody>(occ) ?? occ;
            ctx.SingleSelection = occInCtx;

            Command.Execute("MakeIndependent");

            IDesignBody newOcc = null;
            foreach (var idb in mainPart.GetDescendants<IDesignBody>())
            {
                if (!originalHashCodes.Contains(idb.GetHashCode()))
                {
                    newOcc = idb;
                    break;
                }
            }

            if (newOcc != null)
            {
                // Normalise all components that share the same stem
                NormalizeStemSuffixes(newOcc);
                return newOcc;
            }

            return occ;
        }

        // NormalizeStemSuffixes renames component masters with a numeric suffix to avoid duplicate names.
        private static void NormalizeStemSuffixes(IDesignBody iBod)
        {
            if (iBod == null) return;
            var win = Window.ActiveWindow;
            var mainPart = win?.Document?.MainPart;
            string currName = iBod.Parent.Master.Name;
            var allComps = mainPart.GetDescendants<IDesignBody>().ToList();
            foreach (var comp in allComps)
            {
                Logger.Log(comp.Parent.Master.Name);
            }

            string newName = $"{currName.Remove(currName.Length - 1)}_{nameIndex}";
            Logger.Log($"Newname = {newName}");
            bool exists = allComps.Any(c => c.Parent.Master.Name.Equals(newName, StringComparison.OrdinalIgnoreCase));

            if (exists)
            {
                nameIndex++;
                NormalizeStemSuffixes(iBod);
            }
            else
            {
                iBod.Parent.Master.Name = newName;
            }
        }

        // ApplyOwnerEditsWithChoice applies connector unite/subtract operations to the owner master, respecting linked-body choices.
        private static bool ApplyOwnerEditsWithChoice(
            Part mainPart,
            IDesignBody iDesignBody,
            DesignBody ownerMaster,
            Part parentPart,
            IEnumerable<Body> uniteList,
            IEnumerable<Body> ownerCuts,
            out LinkedChoice choice,
            out IDesignBody ownerOccForMapping)
        {
            choice = LinkedChoice.ThisOnly;
            ownerOccForMapping = iDesignBody;

            try
            {
                // Log bounding boxes of each unite body
                int uidx = 0;

                bool isLinked = IsLinked(iDesignBody, mainPart);

                if (!isLinked)
                {
                    UniteEachIntoOwner(ownerMaster, parentPart, uniteList);
                    SubtractEachFromOwner(ownerMaster, parentPart, ownerCuts);
                    return true;
                }

                choice = AskLinkedChoice(iDesignBody);

                if (choice == LinkedChoice.Cancel)
                {
                    return false;
                }

                if (choice == LinkedChoice.AllLinked)
                {
                    UniteEachIntoOwner(ownerMaster, parentPart, uniteList);
                    SubtractEachFromOwner(ownerMaster, parentPart, ownerCuts);
                    ownerOccForMapping = iDesignBody;
                    return true;
                }

                nameIndex = 1;
                var indepOcc = MakeIndependentOcc(iDesignBody);

                //NormalizeStemSuffixes(iDesignBody);
                if (indepOcc == null)
                {
                    Application.ReportStatus("Could not make this body independent.", StatusMessageType.Warning, null);
                    ownerOccForMapping = iDesignBody;
                    return false;
                }

                var targetMaster = indepOcc.Master;
                var targetPart = targetMaster.Parent;

                UniteEachIntoOwner(targetMaster, targetPart, uniteList);
                SubtractEachFromOwner(targetMaster, targetPart, ownerCuts);

                ownerOccForMapping = indepOcc;
                return true;
            }
            catch (Exception ex)
            {
                choice = LinkedChoice.Cancel;
                ownerOccForMapping = iDesignBody;
                return false;
            }
        }

        private enum LinkedChoice { ThisOnly, AllLinked, Cancel }

        // IsLinked checks whether a given occurrence shares its master with at least one other occurrence.
        private static bool IsLinked(IDesignBody occ, Part mainPart)
        {
            if (occ?.Master == null || mainPart == null) return false;
            return mainPart.GetDescendants<IDesignBody>().Count(x => x.Master == occ.Master) > 1;
        }

        // AskLinkedChoice prompts the user on how to treat linked owner bodies for connector edits.
        private static LinkedChoice AskLinkedChoice(IDesignBody occ)
        {
            const string caption = "Linked bodies detected";
            const string text =
                "This body is linked to others. Do you want to adjust all linked bodies?\n\n" +
                "  • Yes  — Adjust ALL linked bodies\n" +
                "  • No   — Adjust ONLY this body\n" +
                "  • Cancel — Abort";

            var buttons = System.Windows.Forms.MessageBoxButtons.YesNoCancel;
            var icon = System.Windows.Forms.MessageBoxIcon.Question;

            // Default to "No" (Single body) to be conservative.
            var res = System.Windows.Forms.MessageBox.Show(
                text, caption, buttons, icon, System.Windows.Forms.MessageBoxDefaultButton.Button2);

            if (res == System.Windows.Forms.DialogResult.Yes) return LinkedChoice.AllLinked; // Yes = All linked
            if (res == System.Windows.Forms.DialogResult.No) return LinkedChoice.ThisOnly;  // No = This only
            return LinkedChoice.Cancel;                                                     // Cancel or closed
        }

        // GetAllLinkedOccurrences returns all occurrences in mainPart that reference the same master as ownerOcc.
        private static List<IDesignBody> GetAllLinkedOccurrences(Part mainPart, IDesignBody ownerOcc)
        {
            if (mainPart == null || ownerOcc?.Master == null) return new List<IDesignBody>();
            return mainPart.GetDescendants<IDesignBody>().Where(x => x.Master == ownerOcc.Master).ToList();
        }

        private enum NeighbourBatchChoice { MakeIndependentAll, EditSharedAll, SkipAll }

        // BoxesIntersect performs an AABB intersection test between two SpaceClaim boxes.
        private static bool BoxesIntersect(Box a, Box b)
        {
            var A0 = a.MinCorner; var A1 = a.MaxCorner;
            var B0 = b.MinCorner; var B1 = b.MaxCorner;
            return !(A1.X < B0.X || A0.X > B1.X
                  || A1.Y < B0.Y || A0.Y > B1.Y
                  || A1.Z < B0.Z || A0.Z > B1.Z);
        }

        // IntersectsNeighbour checks whether a mapped tool body actually intersects a neighbour occurrence.
        private static bool IntersectsNeighbour(IDesignBody ownerOcc, IDesignBody neighbourOcc, Body toolInOwnerMaster)
        {
            if (neighbourOcc?.Master?.Shape == null || toolInOwnerMaster == null) return false;
            try
            {
                using var mapped = MapTool(toolInOwnerMaster, ownerOcc, neighbourOcc);
                var c = neighbourOcc.Master.Shape.GetCollision(mapped);
                return c == Collision.Intersect; // strict: no Touch
            }
            catch { return false; }
        }

        // PropagateCutsForChoice applies connector cuts to neighbour bodies around the owner.
        // Behaviour:
        // - No prompts for neighbours.
        // - If the owner was edited "ThisOnly" (made independent), then any linked neighbour hit by
        //   a cut is also made independent before subtracting, so cuts stay local.
        // - If the owner was edited "AllLinked", neighbours are cut on their shared master.
        private static void PropagateCutsForChoice(
            Part mainPart,
            IDesignBody ownerOccForNeighbours,   // enumerate neighbours / pick ring
            IDesignBody ownerOccSourceForTools,  // ORIGINAL owner; tool geometry lives here (bodies are in its master space)
            Part parentPart,
            IEnumerable<Body> collBodies,
            IEnumerable<Body> cutBodies,
            LinkedChoice choice)
        {
            if (mainPart == null || ownerOccForNeighbours == null || ownerOccSourceForTools == null) return;

            var collList = (collBodies ?? Enumerable.Empty<Body>()).Where(b => b != null).ToList();
            var cutList = (cutBodies ?? Enumerable.Empty<Body>()).Where(b => b != null).ToList();
            if (collList.Count == 0 && cutList.Count == 0) return;

            // Prefer real cutters
            var ownersToProcess = (choice == LinkedChoice.AllLinked)
                ? GetAllLinkedOccurrences(mainPart, ownerOccForNeighbours)
                : new List<IDesignBody> { ownerOccForNeighbours };

            foreach (var ownerOcc in ownersToProcess)
            {
                if (ownerOcc?.Master?.Shape == null) continue;

                var allTools = cutList.Count > 0 ? cutList : collList;
                if (!allTools.Any()) continue;

                // WORLD AABBs of tools placed at THIS owner occurrence
                var toolWorldBoxes = new List<Box>();
                foreach (var t in allTools)
                {
                    try
                    {
                        toolWorldBoxes.Add(t.GetBoundingBox(ownerOcc.TransformToMaster.Inverse));
                    }
                    catch
                    {
                        // ignore this tool for box prefiltering
                    }
                }
                if (toolWorldBoxes.Count == 0) continue;

                var neighbours = mainPart.GetDescendants<IDesignBody>()
                                         .Where(idb => idb != null && idb.Master != ownerOcc.Master)
                                         .ToList();

                // Collect colliding neighbours first
                var colliding = new List<IDesignBody>();
                foreach (var nOcc in neighbours)
                {
                    var nMaster = nOcc.Master;
                    var nShape = nMaster?.Shape;
                    if (nShape == null) continue;

                    // coarse AABB
                    Box nWorldBox;
                    try
                    {
                        nWorldBox = nShape.GetBoundingBox(nOcc.TransformToMaster.Inverse);
                    }
                    catch
                    {
                        continue;
                    }

                    if (!toolWorldBoxes.Any(tb => BoxesIntersect(nWorldBox, tb)))
                        continue;

                    // strict intersection with at least one tool
                    if (!allTools.Any(t => IntersectsNeighbour(ownerOcc, nOcc, t)))
                        continue;

                    colliding.Add(nOcc);
                }

                if (colliding.Count == 0) continue;

                // Apply cuts to all colliding neighbours.
                // If the owner was made independent (ThisOnly), then also make any linked neighbours
                // independent before cutting them, so we do not modify their shared master.
                foreach (var nOcc0 in colliding)
                {
                    IDesignBody targetOcc = nOcc0;

                    // Owner choice == ThisOnly means we unlinked the owner;
                    // mirror that behaviour onto any linked neighbours we are about to cut.
                    if (choice == LinkedChoice.ThisOnly && IsLinked(nOcc0, mainPart))
                    {
                        try
                        {
                            nameIndex = 1; // reset suffix counter for a clean series of names
                            var newOcc = MakeIndependentOcc(nOcc0);
                            if (newOcc == null)
                            {
                                // Could not make this neighbour independent; skip it rather than
                                // unexpectedly modifying the shared master.
                                continue;
                            }
                            targetOcc = newOcc;
                        }
                        catch
                        {
                            // If independence fails for any reason, skip this neighbour to avoid
                            // unintended edits on a shared master.
                            continue;
                        }
                    }

                    var targetMaster = targetOcc.Master;
                    var targetPart = targetMaster.Parent;

                    int toolIdx = 0;
                    foreach (var toolSrc in allTools)
                    {
                        if (toolSrc == null) continue;

                        try
                        {
                            if (!IntersectsNeighbour(ownerOcc, targetOcc, toolSrc))
                                continue;

                            using var mapped = MapTool(toolSrc, ownerOcc, targetOcc);
                            var tmp = DesignBody.Create(targetPart, $"_cut_tmp_{toolIdx++}", mapped.Copy());
                            try
                            {
                                targetMaster.Shape.Subtract(new[] { tmp.Shape });
                            }
                            finally
                            {
                                try { if (!tmp.IsDeleted) tmp.Delete(); } catch { }
                            }
                        }
                        catch
                        {
                            // swallow and continue with the next tool / neighbour
                        }
                    }
                }
            }
        }


        // UniteIntoOwner wraps a tool in a temporary DesignBody and unites it into the owner DesignBody.
        private static void UniteIntoOwner(DesignBody owner, Part part, Body tool)
        {
            if (owner == null || part == null || tool == null) return;

            DesignBody tmp = null;
            try
            {
                tmp = DesignBody.Create(part, "_tmp_unite", tool.Copy());
                owner.Shape.Unite(new[] { tmp.Shape });
            }
            catch (Exception ex)
            {
            }
        }

        // SubtractFromOwner wraps a tool in a temporary DesignBody and subtracts it from the owner DesignBody.
        private static void SubtractFromOwner(DesignBody owner, Part part, Body tool)
        {
            if (owner == null || part == null || tool == null) return;

            DesignBody tmp = null;
            try
            {
                tmp = DesignBody.Create(part, "_tmp_unite", tool.Copy());
                owner.Shape.Subtract(new[] { tmp.Shape });
            }
            catch (Exception ex)
            {
            }
        }

        // UniteEachIntoOwner iterates over all connector bodies and unites each into the owner master.
        private static void UniteEachIntoOwner(DesignBody owner, Part part, IEnumerable<Body> connectors)
        {
            if (owner == null || part == null || connectors == null) return;

            int idx = 0;
            foreach (var b in connectors)
            {
                if (b == null) continue;
                try
                {
                    var cp = b.Copy();
                    UniteIntoOwner(owner, part, cp);
                }
                catch (Exception ex)
                {
                }
            }
        }

        // SubtractEachFromOwner iterates over all cutter bodies and subtracts each from the owner master.
        private static void SubtractEachFromOwner(DesignBody owner, Part part, IEnumerable<Body> cutters)
        {
            if (owner == null || part == null || cutters == null) return;

            int idx = 0;
            foreach (var b in cutters)
            {
                if (b == null) continue;
                try
                {
                    var cp = b.Copy();
                    SubtractFromOwner(owner, part, cp);
                }
                catch (Exception ex)
                {
                }
            }
        }

        // ComputePatternCentersOcc computes evenly spaced connector centers and tangents along an edge in occurrence space.
        private static List<(Point pOcc, Direction tanOcc)> ComputePatternCentersOcc(
            IDesignEdge iEdge,
            ConnectorModel c,
            int n)
        {
            //Logger.Log($"ComputePatternCentersOcc called. Looking for {n} connectors");
            if (iEdge == null || iEdge.Shape == null) return new List<(Point, Direction)>();
            if (n <= 0) throw new ArgumentException("Pattern amount must be greater than 0.");

            //Logger.Log($"ComputePatternCentersOcc cont");
            // Occupied width along the edge in metres
            double occWidthM = (Math.Max(c.Width1, c.Width2) + 2.0 * c.Tolerance) / 1000.0;
            double dLoc = 0.001 * c.Location; // user location shift (m), applied along local tangent

            var shape = iEdge.Shape;
            var geom = shape.Geometry;
            var p0 = shape.StartPoint;
            var p1 = shape.EndPoint;

            var centers = new List<(Point, Direction)>();

            // --- Straight segment -----------------------------------------------------
            if (geom is Line)
            {
                //Logger.Log($"geom is Line");
                double L = (p1 - p0).Magnitude; // metres

                // capacity check
                if (L < n * occWidthM - 1e-12)
                    throw new InvalidOperationException(
                        $"Requested pattern ({n}) does not fit on this edge. Length={L * 1000:0.###} mm, needed={(n * occWidthM) * 1000:0.###} mm.");

                // even spacing with “gaps” at ends
                double gap = (L - n * occWidthM) / (n + 1);
                var dir = (p1 - p0).Direction;

                for (int k = 0; k < n; k++)
                {
                    //Logger.Log($"for loop 'n'");
                    double s = gap + 0.5 * occWidthM + k * (occWidthM + gap);
                    // clamp softly inside the edge
                    s = Math.Max(1e-9, Math.Min(L - 1e-9, s));
                    var p = p0 + s * dir.ToVector();

                    // apply Location along tangent
                    var tan = dir;
                    p = p + dLoc * tan;

                    centers.Add((p, tan));

                    //Logger.Log($"cemters added {centers.Count}");
                }
                return centers;
            }

            // --- Circular arc / full circle (equal-angle / equal-arc spacing) ------------
            if (geom is Circle circle)
            {
                var sShape = iEdge.Shape;

                // Edge parameter range (in radians for circles)
                double ps = sShape.ProjectPoint(sShape.StartPoint).Param;
                double pe = sShape.ProjectPoint(sShape.EndPoint).Param;

                // Signed span; detect full circle if start≈end
                const double EPS = 1e-9;
                double dpar = pe - ps;
                bool isFullCircle = Math.Abs(dpar) < EPS;    // closed edge → use full 2π
                double sign = (dpar >= 0) ? 1.0 : -1.0;      // preserve traversal direction

                double R = circle.Radius; // metres

                // Helper to evaluate at parameter t and add Location shift
                (Point p, Direction tan) EvalAt(double t)
                {
                    var ev = circle.Evaluate(t);
                    var p = ev.Point;
                    var tan = ev.Derivative.Direction;
                    // Respect traversal sign so pattern follows edge orientation
                    if (sign < 0) tan = -tan;
                    // Apply user Location (mm) along tangent
                    p = p + dLoc * tan;
                    return (p, tan);
                }

                if (isFullCircle)
                {
                    // Equal-angle spacing around the loop: t_k = ps + sign * k * (2π/n)
                    double step = (2.0 * Math.PI) / n;
                    for (int k = 0; k < n; k++)
                    {
                        double t = ps + sign * k * step;
                        var (p, tan) = EvalAt(t);
                        centers.Add((p, tan));
                    }
                }
                else
                {
                    // Open arc: even distribution across the actual span.
                    // Include endpoints for n>1; center for n==1.
                    if (n == 1)
                    {
                        double tMid = ps + 0.5 * (pe - ps);
                        var (p, tan) = EvalAt(tMid);
                        centers.Add((p, tan));
                    }
                    else
                    {
                        for (int k = 0; k < n; k++)
                        {
                            double frac = (double)k / (n - 1);     // 0 → 1 inclusive
                            double t = ps + frac * (pe - ps);      // follows edge direction
                            var (p, tan) = EvalAt(t);
                            centers.Add((p, tan));
                        }
                    }
                }

                return centers;
            }
            return centers;
        }

        // getOppositeFace finds the opposite face to bigFace and returns thickness and direction for planar or cylindrical sheets.
        private DesignFace getOppositeFace(DesignFace smallFace, DesignFace bigFace, bool isPlanar, out double thickness, out Direction direction)
        {
            DesignFace returnFace = null;
            thickness = 999;
            direction = Direction.Zero;

            foreach (var face in smallFace.AdjacentFaces)
            {
                if (face == bigFace)
                    continue;


                var facePlane = face.Shape.Geometry as Plane;
                var faceCylinder = face.Shape.Geometry as Cylinder;

                bool isPlane = facePlane != null;
                bool isCylinder = faceCylinder != null;

                if (!isPlanar && isCylinder)
                {
                    var cyl = (Cylinder)bigFace.Shape.Geometry;
                    double dist = (cyl.Axis.ProjectPoint(faceCylinder.Axis.Origin).Point - faceCylinder.Axis.Origin).Magnitude;
                    if (dist > 1e-6)
                        continue;

                    double _thickness = cyl.Radius - faceCylinder.Radius;

                    if (_thickness < thickness)
                    {
                        thickness = _thickness;
                        returnFace = face;
                    }

                    //return face;
                }
                else if (isPlanar && isPlane)
                {
                    var pl = (Plane)bigFace.Shape.Geometry;

                    Direction dir1 = pl.Frame.DirZ;
                    Direction dir2 = facePlane.Frame.DirZ;
                    bool test = pl.Frame.DirZ.IsParallel(facePlane.Frame.DirZ);
                    if (!pl.Frame.DirZ.IsParallel(facePlane.Frame.DirZ))
                        continue;
                    Point p = pl.Evaluate(new PointUV(0, 0)).Point;
                    Point p1 = facePlane.ProjectPoint(p).Point;
                    double _thickness = (p - p1).Magnitude;

                    if (_thickness < thickness)
                    {
                        thickness = _thickness;
                        direction = (p1 - p).Direction;
                        returnFace = face;
                    }
                }
            }
            return returnFace;

        }

        // checkDesignEdge determines a stable placement point on the edge (either click location or curve midpoint).
        private bool checkDesignEdge(IDesignEdge iDesEdge, bool clickPosition, out Point midPointOcc)
        {
            midPointOcc = Point.Origin;
            if (iDesEdge == null || iDesEdge.Shape == null) return false;

            // We will work in occurrence space via iDesEdge.Shape
            var geom = iDesEdge.Shape.Geometry;

            if (clickPosition)
            {
                // Selection point is world/master; project into OCCurrence space
                var selWorld = (Point)Window.ActiveWindow.ActiveContext.GetSelectionPoint(iDesEdge);
                midPointOcc = iDesEdge.Shape.ProjectPoint(selWorld).Point;
                return true;
            }

            // Compute midpoint parameter robustly for common curve types
            var line = geom as Line;
            var circ = geom as Circle;
            var nurb = geom as NurbsCurve;
            var ell = geom as Ellipse;
            var proc = geom as ProceduralCurve;

            if (line != null)
            {
                // Use the shape's own parameter range
                var ps = iDesEdge.Shape.ProjectPoint(iDesEdge.Shape.StartPoint).Param;
                var pe = iDesEdge.Shape.ProjectPoint(iDesEdge.Shape.EndPoint).Param;
                midPointOcc = line.Evaluate(ps + 0.5 * (pe - ps)).Point;
                return true;
            }
            if (circ != null)
            {
                var ps = iDesEdge.Shape.ProjectPoint(iDesEdge.Shape.StartPoint).Param;
                var pe = iDesEdge.Shape.ProjectPoint(iDesEdge.Shape.EndPoint).Param;
                var pm = ps + 0.5 * (pe - ps);
                midPointOcc = circ.Evaluate(pm).Point;
                return true;
            }
            if (nurb != null)
            {
                var ps = iDesEdge.Shape.ProjectPoint(iDesEdge.Shape.StartPoint).Param;
                var pe = iDesEdge.Shape.ProjectPoint(iDesEdge.Shape.EndPoint).Param;
                var pm = ps + 0.5 * (pe - ps);
                midPointOcc = nurb.Evaluate(pm).Point;
                return true;
            }
            if (ell != null)
            {
                var ps = iDesEdge.Shape.ProjectPoint(iDesEdge.Shape.StartPoint).Param;
                var pe = iDesEdge.Shape.ProjectPoint(iDesEdge.Shape.EndPoint).Param;
                var pm = (Math.Abs(pe) < 1e-12) ? 0.5 * pe : ps + 0.5 * (pe - ps);
                midPointOcc = ell.Evaluate(pm).Point;
                return true;
            }
            if (proc != null)
            {
                midPointOcc = proc.Evaluate(0.5).Point;
                return true;
            }

            return false;
        }

        // WireUiChangeHandlers attaches debounced change handlers to all relevant connector UI controls.
        private void WireUiChangeHandlers()
        {
            TextChangedEventHandler onText = OnTextChanged;
            RoutedEventHandler onRouted = OnRoutedChanged;
            KeyEventHandler onKey = OnTextBoxKeyDown;
            SelectionChangedEventHandler onSel = OnSelectionChanged;

            var textBoxes = new TextBox[]
            {
                connectorHeight,
                connectorWidth1,
                connectorWidth2,
                connectorTolerance,
                connectorSpacing,
                connectorRadiusChamfer,
                connectorLocation,
                connectorCornerCutoutValue,
                connectorCornerCutoutRadiusValue
            };
            foreach (var tb in textBoxes)
            {
                if (tb == null) continue;
                tb.TextChanged += onText;   // while typing
                tb.LostFocus += onRouted;   // focus-out commit
                tb.KeyDown += onKey;        // Enter
            }

            var checkBoxes = new CheckBox[]
            {
                //connectorUseCustom,
                connectorDynamicHeight,
                connectorCornerCutout,
                connectorCornerCutoutRadius,
                connectorClickLocation,
                connectorShowTolerance,
                connectorRectangularCut
                //connectorStraight
            };
            foreach (var cb in checkBoxes)
            {
                if (cb == null) continue;
                cb.Checked += onRouted;
                cb.Unchecked += onRouted;
                cb.LostFocus += onRouted;
                cb.KeyDown += onKey;
            }

            var radios = new RadioButton[]
            {
                connectorRadius,
                connectorChamfer
            };
            foreach (var rb in radios)
            {
                if (rb == null) continue;
                rb.Checked += onRouted;
                rb.Unchecked += onRouted;
                rb.LostFocus += onRouted;
                rb.KeyDown += onKey;
            }

            if (ConnectorShapeCombobox != null)
            {
                ConnectorShapeCombobox.SelectionChanged += onSel;
                ConnectorShapeCombobox.LostFocus += onRouted;
                ConnectorShapeCombobox.KeyDown += onKey;
            }
        }

        // OnTextChanged reacts to live text edits by scheduling a redraw and enforcing coupled UI rules.
        private void OnTextChanged(object sender, TextChangedEventArgs e)
        {
            DebouncedRedraw();
            EnforceCornerCoupling();
            EnforceRectCutExclusivity();
            UpdateGenerateEnabled();
        }

        // OnRoutedChanged reacts to routed events (focus changes, etc.) by redrawing and validating layout options.
        private void OnRoutedChanged(object sender, RoutedEventArgs e)
        {
            DebouncedRedraw();
            EnforceCornerCoupling();
            EnforceRectCutExclusivity();
            UpdateGenerateEnabled();
        }

        // OnSelectionChanged responds to preset combobox changes by redrawing and toggling the generate button state.
        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DebouncedRedraw();
            UpdateGenerateEnabled();
        }

        private const double EPS = 1e-6;

        // TryReadDouble parses a double from a TextBox using current or invariant cultures, with last-chance normalization.
        private bool TryReadDouble(TextBox tb, out double v)
        {
            v = 0;
            if (tb == null) return false;
            var s = (tb.Text ?? "").Trim();
            // current culture OR invariant
            if (double.TryParse(s, NumberStyles.Float, CultureInfo.CurrentCulture, out v)) return true;
            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return true;
            // last-chance normalize
            s = s.Replace(',', '.');
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v);
        }

        // IsReliefAllowed enforces End Relief constraints against current width settings and rectangular cut options.
        private bool IsReliefAllowed(out string reason)
        {
            reason = null;

            if (!TryReadDouble(connectorWidth1, out var w1)) { reason = "Enter a valid Width1."; return false; }
            if (!TryReadDouble(connectorWidth2, out var w2)) { reason = "Enter a valid Width2."; return false; }
            if (!TryReadDouble(connectorSpacing, out var relief)) { reason = "Enter a valid End Relief."; return false; }

            if ((connectorRectangularCut?.IsChecked == false) && relief > 0)
            {
                reason = "Rectangular cut can’t be used while End Relief > 0.";
                return false;
            }

            // Only block when relief > 0 AND widths differ
            if (relief > 0 && Math.Abs(w1 - w2) > EPS)
            {
                reason = "End Relief can only be used if top and bottom width are equal.";
                return false;
            }

            return true;
        }

        // EnforceRectCutExclusivity keeps rectangular-cut and End Relief UI states mutually consistent.
        private void EnforceRectCutExclusivity()
        {
            // If Rectangular cut is ON, force End Relief to 0 and disable it.
            if (connectorRectangularCut?.IsChecked == false)
            {
                if (connectorSpacing != null)
                {
                    connectorSpacing.Text = "0";
                    connectorSpacing.IsEnabled = false;
                    connectorSpacing.ToolTip = "End Relief is disabled while Rectangular cut is on.";
                }
            }
            else
            {
                if (connectorSpacing != null)
                {
                    connectorSpacing.IsEnabled = true;
                    connectorSpacing.ToolTip = null;
                }
            }

            // If user types End Relief > 0, auto-turn Off Rectangular cut (vice-versa)
            if (connectorSpacing != null && double.TryParse(connectorSpacing.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var rel) && rel > 0)
            {
                if (connectorRectangularCut != null)
                    connectorRectangularCut.IsChecked = true;
            }
        }

        // UpdateGenerateEnabled toggles the create button based on whether current settings are valid (including relief rules).
        private void UpdateGenerateEnabled()
        {
            var btn = this.FindName("btnCreateConnector") as System.Windows.Controls.Button
                      ?? this.FindName("btnCreate") as System.Windows.Controls.Button
                      ?? this.FindName("generateButton") as System.Windows.Controls.Button;

            if (btn == null) return; // if your XAML uses a different name, adjust the FindName above

            if (IsReliefAllowed(out var reason))
            {
                btn.IsEnabled = true;
                btn.ToolTip = null;
                // optional: clear visual warning
                if (connectorSpacing != null) connectorSpacing.Background = System.Windows.Media.Brushes.White;
            }
            else
            {
                btn.IsEnabled = false;
                btn.ToolTip = reason;
                // optional: subtle visual cue on the End Relief box
                if (connectorSpacing != null) connectorSpacing.Background = System.Windows.Media.Brushes.MistyRose;
            }
        }

        // OnTextBoxKeyDown intercepts Enter presses to trigger a debounced redraw instead of default behavior.
        private void OnTextBoxKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                DebouncedRedraw();
            }
        }

        // DebouncedRedraw coalesces UI changes and triggers a delayed connector rebuild and UI validation.
        private void DebouncedRedraw()
        {
            if (_uiDebounce == null)
            {
                _uiDebounce = new DispatcherTimer { Interval = UiDebounceInterval };
                _uiDebounce.Tick += (_, __) =>
                {
                    _uiDebounce.Stop();
                    // drawConnector() rebuilds the ConnectorModel from UI and invalidates the picture
                    drawConnector();
                    EnforceRectCutExclusivity();
                    UpdateGenerateEnabled();
                };
            }
            _uiDebounce.Stop();
            _uiDebounce.Start();
        }

        // ComputeBoundsMm calculates a tight bounding box in mm around the connector with optional tolerance and top radius features.
        private static (double minX, double maxX, double minY, double maxY) ComputeBoundsMm(ConnectorModel c, bool showTolerance)
        {
            // Inputs in mm
            double w1 = Math.Max(0, c.Width1);
            double w2 = Math.Max(0, c.Width2);
            double h = Math.Max(0, c.Height);

            double rCorner = (c.HasCornerCutout ? Math.Max(0, c.CornerCutoutRadius) : 0.0); // mm
            double tol = (showTolerance ? Math.Max(0, c.Tolerance) : 0.0);              // mm

            // Extra top pair (connectorCornerCutoutRadius)
            bool hasTopPair = c.RadiusInCutOut && c.RadiusInCutOut_Radius > 0;
            double rTop = hasTopPair ? Math.Max(0, c.RadiusInCutOut_Radius) : 0.0;    // mm

            // --- Horizontal extents ---
            // Bottom: bottom width + optional corner cutout that adds outward at the corners
            double halfBottomX = w1 * 0.5 + rCorner + tol;

            // Top: top width + tolerance; top pair adds outward by its radius
            double halfTopX = w2 * 0.5 + tol + rTop;

            double halfX = Math.Max(halfBottomX, halfTopX);

            // --- Vertical extents ---
            // Bottom stays at 0 (corner arcs are inside the profile, not below baseline)
            double minY = 0.0 - 0.0;            // keep as 0, you can add tiny margin if desired
                                                // Top is height; tolerance outline climbs by ~tol; the extra top pair centers shift up by tol, plus their radius
            double topWithTol = h + tol;
            double topWithPair = hasTopPair ? (h + tol + rTop) : topWithTol;
            double maxY = Math.Max(topWithTol, topWithPair);

            // Final bounds
            double minX = -halfX;
            double maxX = halfX;

            // Avoid zero-size bounds
            const double EPS = 1e-6;
            if (maxX - minX < EPS) { minX -= 1; maxX += 1; }
            if (maxY - minY < EPS) { minY -= 1; maxY += 1; }

            return (minX, maxX, minY, maxY);
        }

        // EnforceCornerCoupling disables the extra top corner radius feature when a main corner radius/chamfer is set.
        private void EnforceCornerCoupling()
        {
            // If there’s a main corner radius/chamfer value, disable the top-pair “cutout radius” feature.
            if (double.TryParse(connectorRadiusChamfer?.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var rc) && rc > 0)
            {
                if (connectorCornerCutoutRadius != null) connectorCornerCutoutRadius.IsChecked = false; // the checkbox
                if (connectorCornerCutoutRadiusValue != null) connectorCornerCutoutRadiusValue.Text = "0";
                if (connectorCornerCutoutRadius != null) connectorCornerCutoutRadius.IsEnabled = false;
                if (connectorCornerCutoutRadiusValue != null) connectorCornerCutoutRadiusValue.IsEnabled = false;
            }
            else
            {
                if (connectorCornerCutoutRadius != null) connectorCornerCutoutRadius.IsEnabled = true;
                if (connectorCornerCutoutRadiusValue != null) connectorCornerCutoutRadiusValue.IsEnabled = true;
            }
        }

        // DrawConnector renders a scaled 2D representation of the connector profile (plus tolerance outlines) onto the Graphics.
        private void DrawConnector(Graphics g, ConnectorModel connector)
        {
            try
            {
                bool showTolerance = connectorShowTolerance.IsChecked == true;

                int canvasWidth = picDrawing.ClientSize.Width;
                int canvasHeight = picDrawing.ClientSize.Height;

                // Compute robust bounds (mm)
                var (minXmm, maxXmm, minYmm, maxYmm) = ComputeBoundsMm(connector, showTolerance);

                // Derive scale (px/mm) with a 10% margin
                double modelW = maxXmm - minXmm;
                double modelH = maxYmm - minYmm;
                float scale = (float)(0.9 * Math.Min(canvasWidth / Math.Max(1.0, modelW),
                                                     canvasHeight / Math.Max(1.0, modelH)));

                // Convert model (mm) → pixels
                Func<double, float> Xpx = xmm => (float)((xmm - (minXmm + maxXmm) * 0.5) * scale + canvasWidth * 0.5);
                Func<double, float> Ypx = ymm => (float)(((maxYmm + minYmm) * 0.5 - ymm) * scale + canvasHeight * 0.5);

                // Convenience locals in pixels for main profile (your existing code uses these)
                float width1 = (float)connector.Width1 * scale;
                float width2 = (float)connector.Width2 * scale;
                float height = (float)connector.Height * scale;
                float radius = (float)connector.Radius * scale;
                float cornerCutoutRadius = connector.HasCornerCutout ? (float)connector.CornerCutoutRadius * scale : 0f;

                // Rebuild base anchor positions from model values using Xpx/Ypx
                // Bottom line: y = 0; Top line: y = connector.Height
                float leftX = Xpx(-connector.Width1 * 0.5);
                float rightX = Xpx(connector.Width1 * 0.5);
                float topX = Xpx(-connector.Width2 * 0.5);
                float bottomY = Ypx(0.0);
                //float topY = Ypx(connector.Height);
                // solid top (H - endRelief)
                float topY = Ypx(Math.Max(0.0, connector.Height - Math.Max(0.0, connector.EndRelief)));


                // Corner points in pixel space (same names as your code)
                float p0X = leftX; float p0Y = bottomY;
                float p1X = topX; float p1Y = topY;
                float p2X = Xpx(connector.Width2 * 0.5); float p2Y = topY;
                float p3X = rightX; float p3Y = bottomY;

                float p01X = leftX; float p01Y = bottomY;
                float p02X = leftX; float p02Y = bottomY;
                float p11X = topX; float p11Y = topY;
                float p12X = topX; float p12Y = topY;
                float p21X = Xpx(connector.Width2 * 0.5); float p21Y = topY;
                float p22X = Xpx(connector.Width2 * 0.5); float p22Y = topY;
                float p31X = rightX; float p31Y = bottomY;
                float p32X = rightX; float p32Y = bottomY;

                using var pen = new System.Drawing.Pen(System.Drawing.Color.Black, 2);
                using var pen1 = new System.Drawing.Pen(System.Drawing.Color.Gray, 2);
                using var penRed = new System.Drawing.Pen(System.Drawing.Color.Red, 2);

                float dX = 0;
                float dY = 0;
                const double EPS = 1e-9;
                double denom = 0.5 * width2 - 0.5 * width1;
                double alpha = Math.Abs(denom) < EPS ? (Math.PI * 0.5) : Math.Atan(height / denom);
                double alphaDegree = alpha * 180.0 / Math.PI;

                if (connector.HasCornerCutout && cornerCutoutRadius > 0)
                {
                    dX = Math.Abs((float)(cornerCutoutRadius * Math.Cos(alpha)));
                    dY = -Math.Abs((float)(cornerCutoutRadius * Math.Sin(alpha)));
                    if (width1 > width2) dX = -dX;

                    p01X += -cornerCutoutRadius;
                    p02X += -dX;
                    p02Y += dY;
                    p32X += cornerCutoutRadius;
                    p31Y += dY;
                    p31X += dX;

                    //if (lblAngle != null) lblAngle.Text = alphaDegree.ToString("0.###");

                    float drawAngle = (float)(360 - alphaDegree);
                    if (alphaDegree < 0) drawAngle = (float)(360 - (180 + alphaDegree));

                    g.DrawArc(pen, p0X - cornerCutoutRadius, p0Y - cornerCutoutRadius, 2 * cornerCutoutRadius, 2 * cornerCutoutRadius, 180, -drawAngle);
                    g.DrawArc(pen, p3X - cornerCutoutRadius, p3Y - cornerCutoutRadius, 2 * cornerCutoutRadius, 2 * cornerCutoutRadius, 0, drawAngle);
                }

                // Base lines
                g.DrawLine(pen1, 0, p01Y, p01X, p01Y);
                g.DrawLine(pen1, p32X, p32Y, canvasWidth, p32Y);

                if (radius > 0)
                {
                    float dX2 = radius;

                    if (connector.HasRounding)
                    {
                        if (width1 <= width2) dX2 = (float)(radius / Math.Tan(alpha / 2));
                        else
                        {
                            double beta = Math.PI + alpha;
                            dX2 = (float)(radius * Math.Tan(0.5 * (Math.PI - beta)));
                        }
                    }

                    if (width1 <= width2)
                    {
                        float dX1 = (float)(dX2 * Math.Sin(Math.PI * 0.5 - alpha));
                        float dY1 = (float)(dX2 * Math.Cos(Math.PI * 0.5 - alpha));

                        p12X += dX2;
                        p11X += dX1;
                        p11Y += dY1;

                        p21X += -dX2;
                        p22X += -dX1;
                        p22Y += dY1;
                    }
                    else
                    {
                        double beta = Math.PI + alpha;
                        float dX1 = (float)(dX2 * Math.Sin(beta - Math.PI * 0.5));
                        float dY1 = (float)(dX2 * Math.Cos(beta - Math.PI * 0.5));

                        p12X += dX2;
                        p11X += -dX1;
                        p11Y += dY1;

                        p21X += -dX2;
                        p22X += dX1;
                        p22Y += dY1;
                    }

                    if (connector.HasRounding)
                    {
                        if (width1 <= width2)
                        {
                            g.DrawArc(pen, p12X - radius, p1Y, 2 * radius, 2 * radius, 270, -(180 - (float)alphaDegree));
                            g.DrawArc(pen, p21X - radius, p2Y, 2 * radius, 2 * radius, 270, 180 - (float)alphaDegree);
                        }
                        else
                        {
                            g.DrawArc(pen, p12X - radius, p1Y, 2 * radius, 2 * radius, 270, (float)alphaDegree);
                            g.DrawArc(pen, p21X - radius, p2Y, 2 * radius, 2 * radius, 270, -(float)alphaDegree);
                        }
                    }
                    else
                    {
                        g.DrawLine(pen, p11X, p11Y, p12X, p12Y);
                        g.DrawLine(pen, p21X, p21Y, p22X, p22Y);
                    }
                }
                else
                {
                    g.DrawLine(pen1, 0, bottomY, leftX - cornerCutoutRadius, bottomY);
                    g.DrawLine(pen1, rightX + cornerCutoutRadius, bottomY, canvasWidth, bottomY);

                    g.DrawLine(pen, leftX - dX, bottomY + dY, topX, topY);
                    g.DrawLine(pen, rightX + dX, bottomY + dY, topX + width2, topY);
                }

                // top and sides
                g.DrawLine(pen, p02X, p02Y, p11X, p11Y);
                g.DrawLine(pen, p12X, p12Y, p21X, p21Y);
                g.DrawLine(pen, p22X, p22Y, p31X, p31Y);

                if (showTolerance)
                {
                    // pens
                    using var pen2 = new Pen(Color.Gray, 2) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash };

                    // px values
                    float tolPx = (float)Math.Max(0.0, connector.Tolerance) * scale;
                    float endReliefPx = (float)Math.Max(0.0, connector.EndRelief) * scale;

                    // keep your alpha/alphaDegree computed above
                    // alpha: slope angle; alphaDegree: alpha in degrees

                    // Guard the side/top offset math
                    denom = Math.Max(1e-9, (0.5 * width2 - 0.5 * width1));
                    double a = Math.Atan(height / denom) / 2.0;
                    double sinA = Math.Max(1e-9, Math.Sin(a));
                    double hyp = tolPx / sinA;
                    float delta = Math.Abs((float)(hyp * Math.Cos(a)));

                    // --- End-relief awareness ---
                    // Solid connector is drawn up to (Height - EndRelief).
                    // Dashed (tolerance) frame is drawn up to (Height + Tolerance).
                    // We only change the dashed Y reference here; the solid top is drawn elsewhere.
                    float solidTopY = Ypx(Math.Max(0.0, connector.Height - Math.Max(0.0, connector.EndRelief)));
                    float dashedTopY = Ypx(connector.Height + Math.Max(0.0, connector.Tolerance));

                    // Dashed-frame anchor corners (offset by tolerance). Use dashedTopY for top.
                    //float _topX1 = topX - delta, _topY1 = dashedTopY - delta;
                    //float _topX2 = topX + width2 + delta, _topY2 = dashedTopY - delta;
                    float _topX1 = topX - delta, _topY1 = dashedTopY;
                    float _topX2 = topX + width2 + delta, _topY2 = dashedTopY;

                    //float _bottomX1 = leftX - delta, _bottomY1 = bottomY - delta;
                    //float _bottomX2 = leftX + width1 + delta, _bottomY2 = _bottomY1;
                    float _bottomX1 = leftX - delta, _bottomY1 = bottomY;
                    float _bottomX2 = leftX + width1 + delta, _bottomY2 = bottomY;

                    // Always draw the bottom dashed segments
                    g.DrawLine(pen2, 0, _bottomY1, _bottomX1, _bottomY1);
                    g.DrawLine(pen2, _bottomX2, _bottomY2, canvasWidth, _bottomY2);

                    // Helpers
                    float Len(float x, float y) => (float)Math.Sqrt(x * x + y * y);
                    void Normalize(ref float x, ref float y) { float L = Len(x, y); if (L > 1e-6f) { x /= L; y /= L; } }
                    float Deg(float vx, float vy) => (float)(Math.Atan2(-vy, vx) * 180.0 / Math.PI); // GDI+: +X=0°, CW+
                    float SweepCW(float start, float end) { float s = end - start; while (s <= 0) s += 360f; return s; }

                    // Slanted side directions in the dashed frame
                    float rDirX = _bottomX2 - _topX2, rDirY = _bottomY2 - _topY2; Normalize(ref rDirX, ref rDirY);
                    float lDirX = _bottomX1 - _topX1, lDirY = _bottomY1 - _topY1; Normalize(ref lDirX, ref lDirY);

                    // Keep ONLY the optional top-pair circles in dashed view
                    bool showTopPair = connector.RadiusInCutOut && connector.RadiusInCutOut_Radius > 0;

                    if (showTopPair)
                    {
                        // Draw sides up to tangency and the two top arcs at the dashed top
                        float rTopPx = (float)(connector.RadiusInCutOut_Radius * scale);

                        // side tangency points (distance rTopPx from top corners along the sides)
                        float rSideTanX = _topX2 + rDirX * rTopPx, rSideTanY = _topY2 + rDirY * rTopPx;
                        float lSideTanX = _topX1 + lDirX * rTopPx, lSideTanY = _topY1 + lDirY * rTopPx;

                        // top tangency points (along the top dashed edge)
                        float rTopTanX = _topX2 - rTopPx, rTopTanY = _topY2;
                        float lTopTanX = _topX1 + rTopPx, lTopTanY = _topY1;

                        // slanted dashed up to tangency
                        g.DrawLine(pen2, _bottomX2, _bottomY2, rSideTanX, rSideTanY);
                        g.DrawLine(pen2, _bottomX1, _bottomY1, lSideTanX, lSideTanY);

                        // top dashed between tangencies
                        g.DrawLine(pen2, lTopTanX, lTopTanY, rTopTanX, rTopTanY);

                        // arcs at top corners (centered on the dashed top corners)
                        float lStart = Deg(lDirX, lDirY) + 180f; // start on the side tangent
                        float lEnd = 0f;                       // end on the top
                        float lSweep = SweepCW(lStart, lEnd);

                        float rStart = 180f;
                        float rEnd = Deg(rDirX, rDirY) + 180f;
                        float rSweep = SweepCW(rStart, rEnd);

                        g.DrawArc(pen2, _topX1 - rTopPx, _topY1 - rTopPx, 2 * rTopPx, 2 * rTopPx, lStart, lSweep);
                        g.DrawArc(pen2, _topX2 - rTopPx, _topY2 - rTopPx, 2 * rTopPx, 2 * rTopPx, rStart, rSweep);
                        return; // done with dashed overlay for this case
                    }

                    // Default dashed outline: simple trapezoid (no dashed chamfer/radius)
                    g.DrawLine(pen2, _bottomX1, _bottomY1, _topX1, _topY1); // left side
                    g.DrawLine(pen2, _topX1, _topY1, _topX2, _topY2);       // top
                    g.DrawLine(pen2, _topX2, _topY2, _bottomX2, _bottomY2); // right side
                }
            }
            catch (Exception ex)
            {
            }
        }

        // InvalidateDrawing requests a repaint of the PictureBox to show the latest connector preview.
        private void InvalidateDrawing()
        {
            try { picDrawing?.Invalidate(); }
            catch (Exception ex) { }
        }

        // CheckSelectedEdgeWidth ensures the connector bottom width does not exceed the available length on the selected edge.
        private bool CheckSelectedEdgeWidth()
        {
            var win = Window.ActiveWindow;
            var ctx = win?.ActiveContext;
            if (win == null || ctx?.Selection == null || ctx.Selection.Count == 0)
                return true; // nothing selected → don't block

            var iEdge = ctx.Selection.OfType<IDesignEdge>().FirstOrDefault();
            if (iEdge == null)
                return true; // not an edge → don't block

            var c = ConnectorModel.CreateConnector(this);
            if (c == null)
                return true; // invalid inputs handled elsewhere

            var edge = iEdge.Master;

            // Get endpoints in model units (meters)
            var p0 = edge.Shape.StartPoint;
            var p1 = edge.Shape.EndPoint;

            double availableWidthMm;

            // 1) Straight line: full edge length
            var line = edge.Shape.Geometry as Line;
            if (line != null)
            {
                double lenM = (p1 - p0).Magnitude;
                availableWidthMm = lenM * 1000.0;
            }
            else
            {
                // 2) Circle/arc: use radius + endpoints to compute arc length
                var circle = edge.Shape.GetGeometry<Circle>();
                if (circle != null)
                {
                    var center = circle.Axis.Origin;
                    double R = circle.Radius;

                    var v0 = p0 - center;
                    var v1 = p1 - center;

                    const double eps = 1e-9;
                    double chordM = (p1 - p0).Magnitude;

                    double theta; // radians along the shorter arc
                    if (chordM < eps || v0.Magnitude < eps || v1.Magnitude < eps)
                    {
                        // Degenerate endpoints → treat as full circle
                        theta = 2.0 * Math.PI;
                    }
                    else
                    {
                        double dot = Vector.Dot(v0, v1) / (v0.Magnitude * v1.Magnitude);
                        // Clamp numerical noise
                        if (dot > 1.0) dot = 1.0;
                        if (dot < -1.0) dot = -1.0;

                        theta = Math.Acos(dot);
                        if (theta < eps && chordM > eps)
                        {
                            theta = 2.0 * Math.PI;
                        }
                    }

                    double arcLenM = R * theta;
                    availableWidthMm = arcLenM * 1000.0;
                }
                else
                {
                    // 3) Other curve types: fallback to chord length between endpoints
                    double chordM = (p1 - p0).Magnitude;
                    availableWidthMm = chordM * 1000.0;
                }
            }
            return (availableWidthMm - c.Width1) >= 0.0;
        }

        // MapTool remaps a tool body from the owner's master space into a target occurrence's master space via world space.
        private static Body MapTool(Body toolInOwnerMaster, IDesignBody ownerOcc, IDesignBody targetOcc)
        {
            var mapped = toolInOwnerMaster.Copy();
            mapped.Transform(ownerOcc.TransformToMaster.Inverse);
            mapped.Transform(targetOcc.TransformToMaster);
            return mapped;
        }
    }
}
