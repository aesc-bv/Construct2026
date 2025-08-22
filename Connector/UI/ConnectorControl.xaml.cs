// SpaceClaim APIs
using AESCConstruct25.FrameGenerator.Utilities;
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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Threading;
using Application = SpaceClaim.Api.V242.Application;
using CheckBox = System.Windows.Controls.CheckBox;
using Component = SpaceClaim.Api.V242.Component;
using ConnectorModel = global::Connector;
using Frame = SpaceClaim.Api.V242.Geometry.Frame;
using KeyEventHandler = System.Windows.Input.KeyEventHandler;
using Line = SpaceClaim.Api.V242.Geometry.Line;
using MessageBox = System.Windows.MessageBox;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using RadioButton = System.Windows.Controls.RadioButton;
using Settings = AESCConstruct25.Properties.Settings;
using TextBox = System.Windows.Controls.TextBox;
using UserControl = System.Windows.Controls.UserControl;
using WF = System.Windows.Forms;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct25.UI
{
    public partial class ConnectorControl : UserControl
    {
        private WF.PictureBox picDrawing;
        private readonly TimeSpan UiDebounceInterval = TimeSpan.FromMilliseconds(150);
        private DispatcherTimer _uiDebounce;
        private class ConnectorPresetRecord
        {
            public string Name { get; set; } = "";
            public double? Height { get; set; }
            public double? Width1 { get; set; }
            public double? Width2 { get; set; }
            public double? Tolerance { get; set; }
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
            public bool? Straight { get; set; }

            // "Radius" or "Chamfer"
            public string CornerStyle { get; set; } = "";
        }


        private readonly List<ConnectorPresetRecord> _presets = new List<ConnectorPresetRecord>();
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
                Logger.Log($"ConnectorControl ctor failed: {ex}");
                MessageBox.Show($"Failed to initialize ProfileSelectionControl:\n{ex.Message}",
                    "Initialization Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
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
                picDrawing.Resize += (_, __) => { try { InvalidateDrawing(); } catch (Exception ex) { Logger.Log($"PictureBox Resize invalidate failed: {ex}"); } };

                // host from XAML
                picDrawingHost.Child = picDrawing;

                // keep drawing responsive to WPF size changes too
                this.SizeChanged += (_, __) =>
                {
                    try { InvalidateDrawing(); }
                    catch (Exception ex) { Logger.Log($"SizeChanged invalidate failed: {ex}"); }
                };
            }
            catch (Exception ex)
            {
                Logger.Log($"InitializeDrawingHost failed: {ex}");
            }
        }

        private void PicDrawing_Paint(object sender, WF.PaintEventArgs e)
        {
            try
            {
                if (connector != null)
                    DrawConnector(e.Graphics, connector);
            }
            catch (Exception ex)
            {
                Logger.Log($"PicDrawing_Paint error: {ex}");
            }
        }

        private void LocalizeUI()
        {
            Localization.Language.LocalizeFrameworkElement(this);
        }

        // === Legacy field (now strongly-typed via alias) ===
        private ConnectorModel connector;

        private void drawConnector()
        {
            try
            {
                connector = ConnectorModel.CreateConnector(this);
                InvalidateDrawing();
            }
            catch (Exception ex)
            {
                Logger.Log($"drawConnector failed: {ex}");
            }
        }

        private void btnCreate_connector(object sender, RoutedEventArgs e) => createConnector();

        // === Legacy Enter/validate handlers adapted to WPF control names ===
        // Map txtTubeLockHeight -> connectorHeight (the TextBox in your XAML)
        private void txtDouble_Validated(object sender, EventArgs e)
        {
            ValidateAndDrawConnector();
        }
        private void txtDouble_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                ValidateAndDrawConnector();
                e.Handled = true; // swallow Enter
            }
        }
        private void ValidateAndDrawConnector()
        {
            if (double.TryParse(connectorHeight?.Text, out _))
                drawConnector();
            else
            {
                MessageBox.Show("Please enter a valid double value.", "Input", MessageBoxButton.OK, MessageBoxImage.Information);
                connectorHeight?.Focus();
            }
        }

        // === Optional: call drawConnector on load (keeps legacy behavior) ===
        private void ConnectorControl_Loaded(object sender, RoutedEventArgs e)
        {
            try {
                WireUiChangeHandlers();
                LoadConnectorPresets();
                drawConnector(); 
            }
            catch (Exception ex) { Logger.Log($"ConnectorControl_Loaded failed: {ex}"); }
        }

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
                Logger.Log($"LoadConnectorPresets failed: {ex}");
                MessageBox.Show($"Failed to load connector presets:\n{ex.Message}",
                    "Connector Presets", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


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
                Logger.Log($"ConnectorPresetCombo_SelectionChanged failed: {ex}");
            }
        }

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

            // Do NOT touch Location or any other controls here
        }

        // Helpers for CSV parsing
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

        private static string[] SplitCsv(string line, char delimiter)
        {
            // simple split (fields expected to be plain values)
            return line.Split(delimiter).Select(s => s.Trim()).ToArray();
        }

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

        private static double? GetDouble(Dictionary<string, int> headers, string[] cols, string key)
        {
            if (!headers.TryGetValue(key, out var i) || i < 0 || i >= cols.Length) return null;
            if (double.TryParse(cols[i], NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) return v;
            return null;
        }

        private static bool? GetBool(Dictionary<string, int> headers, string[] cols, string key)
        {
            if (!headers.TryGetValue(key, out var i) || i < 0 || i >= cols.Length) return null;
            var s = (cols[i] ?? "").Trim().ToLowerInvariant();
            if (bool.TryParse(s, out var b)) return b;
            if (s == "1" || s == "y" || s == "yes") return true;
            if (s == "0" || s == "n" || s == "no") return false;
            return null;
        }

        private static IDesignBody createIndependentDesignBody(IDesignBody iDesignBody)
        {
            DesignBody designBody = iDesignBody.Master;

            Part part = designBody.Parent;
            IPart iPart = iDesignBody.Parent;
            Matrix transformToMaster = iDesignBody.TransformToMaster;
            string newName = part.DisplayName + "_";

            Part partPart = part.Document.MainPart;
            try { partPart = (Part)iPart.Parent.Parent.Master; }
            catch { }

            if (iDesignBody.Parent.Parent != null) // body is in component
            {
                Component compPart = Component.Create(partPart, Part.Create(part.Document, newName));
                DesignBody db = DesignBody.Create(compPart.Template, designBody.Name, designBody.Shape.Copy());
                compPart.Transform(transformToMaster.Inverse);
                db.SetColor(null, designBody.GetColor(null));
                db.Layer = designBody.Layer;

                if (iDesignBody.Parent.Parent != null)
                    iDesignBody.Parent.Parent.Delete();

                foreach (IDesignBody idb in compPart.Document.MainPart.GetDescendants<IDesignBody>())
                {
                    if (idb.Parent.Master.DisplayName == newName)
                    {
                        return idb;
                    }
                }
            }
            else
            {
                return iDesignBody;
            }
            return null;
        }

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

        private void createConnector()
        {
            try
            {
                Window activeWindow = Window.ActiveWindow;
                if (activeWindow == null)
                    return;

                InteractionContext context = activeWindow.ActiveContext;
                Document doc = activeWindow.Document;
                Part mainPart = doc.MainPart;

                DesignEdge desEdge = null;

                if (context.Selection.Count == 0)
                {
                    //SC.reportStatus("Please select an edge");
                    return;
                }

                if (context.Selection.Count > 1)
                {
                    //SC.reportStatus("Select a single edge");
                    return;
                }




                List<DesignBody> selDBodyList = new List<DesignBody> { };
                List<Point> selPointList = new List<Point> { };
                List<Part> selParts = new List<Part> { };
                List<DesignEdge> selEdges = new List<DesignEdge>();
                List<Part> suspendSheetMetalParts = new List<Part> { };
                List<IDesignBody> selIDBodyList = new List<IDesignBody> { };
                List<IDesignEdge> selIDesignEdges = new List<IDesignEdge> { };

                #region Checking the selection
                foreach (IDocObject sel in context.Selection)
                {
                    var de = sel as DesignEdge;
                    if (de != null)
                        desEdge = de;


                    var ide = sel as IDesignEdge;
                    if (ide != null)
                    {
                        desEdge = ide.Master;
                    }

                    if (desEdge == null)
                    {
                        //SC.reportStatus("Please select an edge");
                        return;
                    }

                    bool correctCurve = checkDesignEdge(ide, connector.ClickPosition, out Point selPoint);

                    if (!correctCurve)
                    {
                        //SC.reportStatus("Selection geometry is not supported. Please select edges only.");
                        return;
                    }

                    // Check if it fits on the line
                    if (!checkFitsLine(ide, connector))
                    {
                        //SC.reportStatus("The connector does not fully fit on the selected line");
                        return;
                    }

                    // Check if selected body in component or single body
                    Part part = null;
                    if (mainPart.GetDescendants<IDesignBody>().Count > 1)
                    {
                        part = desEdge.Parent.Parent;
                        if (mainPart == part)
                        {
                            //SC.reportStatus("Selected body is not within a component.");
                            continue;
                        }
                    }
                    else
                        part = mainPart;

                    selPointList.Add(selPoint);


                    selParts.Add(part);
                    selEdges.Add(desEdge);
                    selDBodyList.Add(desEdge.Parent);
                    selIDesignEdges.Add(ide);
                    if (ide != null)
                        selIDBodyList.Add(ide.Parent);

                    //if (selDBodyList.Count != selIDBodyList.Count)
                    //    SC.reportStatus("Nr DB != Nr IDB");


                    //WriteBlock.ExecuteTask("connector", () =>
                    //{
                    //    DatumPoint.Create(mainPart, "selPoint", selPoint);
                    //});

                    // CHeck sheet metal
                    bool suspendSheetMetal = false;
                    if (part.SheetMetal != null || part.IsSheetMetalSuspended)
                        suspendSheetMetal = true;


                    if (suspendSheetMetal)
                        suspendSheetMetalParts.Add(part);
                }

                #endregion

                if (selEdges.Count == 0)
                {
                    //SC.reportStatus("No connectors to be created");
                    return;
                }

                // Get parameters
                double width1 = 0.001 * connector.Width1;
                double width2 = 0.001 * connector.Width2;
                double tolerance = 0.001 * connector.Tolerance;
                double height = 0.001 * connector.Height;
                double radius = 0.001 * connector.Radius;
                bool hasRounding = connector.HasRounding;
                bool hasCornerCutout = connector.HasCornerCutout;
                double cornerCutoutRadius = 0.001 * connector.CornerCutoutRadius;
                bool dynamicHeight = connector.DynamicHeight;
                bool ClickPosition = connector.ClickPosition;

                bool CylinderStraightCut = false; // To add in interface if needed


                // Iterate through all selected edges
                for (int i = 0; i < selEdges.Count; i++)
                {
                    Point pCenter = selPointList[i];
                    DesignBody designBody = selDBodyList[i];
                    IDesignBody iDesignBody = selIDBodyList[i];
                    IDesignEdge ide = selIDesignEdges[i];

                    getFacesFromSelection(ide, out DesignFace bigFace, out DesignFace smallFace);
                    if (bigFace == null || smallFace == null)
                        continue;

                    var bigFacePlane = bigFace.Shape.Geometry as Plane;
                    var bigFaceCylinder = bigFace.Shape.Geometry as Cylinder;

                    bool isPlane = bigFacePlane != null;
                    bool isCylinder = bigFaceCylinder != null;

                    if (!isPlane && !isCylinder)
                    {
                        //SC.reportStatus("Select a planar or cylindrical surface");
                        continue;
                    }

                    DesignFace oppositeFace = getOppositeFace(smallFace, bigFace, isPlane, out double thickness, out Direction dirY2);


                    if (oppositeFace == null)
                    {
                        //SC.reportStatus("Could not find opposite face. Ensure the connector is made between planar or cylindrical faces");
                        continue;
                    }


                    List<IDesignBody> nrBodies = getIDesignBodiesFromBody(mainPart, designBody.Shape);
                    bool allBodies = true; // Default to modifying only the selected body
                    // Check if there are more than one body
                    if (nrBodies.Count > 1)
                    {
                        // Prompt the user with a message box
                        //DialogResult result = MessageBox.Show($"The selected file exists multiple times in the design ({nrBodies.Count}). Do you want to modify all identical parts? If not, only the selected part is affected.",
                        //                                      "Modify Bodies", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        // If user clicks 'Yes', set allBodies to true
                        //if (result == DialogResult.No)
                        //{
                        //    allBodies = false;
                        //}
                    }


                    // check faces (small face must be planar)


                    //Point pX1, pX2;
                    Body connectorBody = null;
                    Body cutBody = null;
                    Body collisionBody = null;
                    List<Body> cutBodiesSource = new List<Body>();

                    WriteBlock.ExecuteTask("connector", () =>
                    {
                        if (isPlane)
                        {
                            // if only the selected body needs modification, make it distinct
                            if (!allBodies)
                            {
                                iDesignBody = createIndependentDesignBody(iDesignBody);
                                designBody = iDesignBody.Master;
                                getFacesFromSelection(ide, out bigFace, out smallFace);
                            }

                            // get Directions

                            Direction dirX = desEdge.Shape.ProjectPoint(pCenter).Derivative.Direction;
                            Direction dirY = getNormalDirectionPlanarFace(bigFace.Shape);
                            Direction dirZ = getNormalDirectionPlanarFace(smallFace.Shape);

                            // Check the max height
                            if (dynamicHeight)
                            {
                                double dynheigth = connector.GetDynamicHeigth(designBody.Parent, dirX, dirY, dirZ, pCenter, height, thickness);
                                if (dynheigth > 0)
                                    height = Math.Min(height, dynheigth);

                                return;
                            }

                            connector.CreateGeometry(designBody.Parent, dirX, dirY, dirZ, pCenter, height, thickness, out connectorBody, out cutBodiesSource, out cutBody, out collisionBody, false);
                            //Collision body is used to check for collisions, if there is a collision, the cutBody is subtracted. cutBodiesSource is used for the Corner Cutouts

                        }
                        else
                        {
                            // Check if there is an inner/outer cylindrical face with same axis
                            var (innerFace, outerFace, thickness1, outerRadius, axis) = CylInfo.GetCoaxialCylPair(iDesignBody, bigFace);


                            if (innerFace == null || outerFace == null)
                            {

                                Application.ReportStatus($"No inner/outer Face found", StatusMessageType.Information, null);
                                return;
                            }
                            if (outerRadius * 2 < width1)
                            {
                                Application.ReportStatus($"Too wide, width should be less than: {outerRadius * 2000} mm", StatusMessageType.Information, null);
                                return;
                            }

                            double distSelectionPoint2Axis = (axis.ProjectPoint(pCenter).Point - pCenter).Magnitude;
                            bool selectedInnerFace = outerRadius - distSelectionPoint2Axis > 1e-5;
                            DesignFace selectedFace = selectedInnerFace ? innerFace : outerFace;

                            // Derive all points for drawing the inner and outer profile
                            Point pAxis = axis.ProjectPoint(pCenter).Point;
                            Direction dirPoint2Axis = (pAxis - pCenter).Direction;
                            Direction dirZ = axis.Direction;
                            // Check Correct Direction;
                            Point pTest = pCenter + 0.00001 * dirZ;
                            if (selectedFace.Shape.ContainsPoint(pTest))
                                dirZ = -dirZ;

                            Direction dirX = Direction.Cross(dirPoint2Axis, dirZ);
                            double maxWidth = Math.Max(width1, width2);
                            double distWidth = Math.Sqrt(outerRadius * outerRadius - (0.5 * maxWidth) * (0.5 * maxWidth));
                            Point pWidth = pAxis - distWidth * dirPoint2Axis;
                            Point pWidth_A = pAxis - distWidth * dirPoint2Axis + 0.5 * maxWidth * dirX;
                            Point pWidth_B = pAxis - distWidth * dirPoint2Axis - 0.5 * maxWidth * dirX;

                            Direction dir_A = CylinderStraightCut ? dirPoint2Axis : (pAxis - pWidth_A).Direction;
                            Direction dir_B = CylinderStraightCut ? dirPoint2Axis : (pAxis - pWidth_B).Direction;
                            Point pInner_A = pWidth_A + thickness1 * dir_A;
                            Point pInner_B = pWidth_B + thickness1 * dir_B;

                            Point pInnerMid = pInner_A - 0.5 * (pInner_A - pInner_B).Magnitude * dirX;
                            Point p_InnerFace = pAxis - (outerRadius - thickness1) * dirPoint2Axis;
                            Point p_OuterFace = pAxis - (outerRadius) * dirPoint2Axis;

                            double alpha = Math.Acos(distWidth / outerRadius);
                            double distOuter = outerRadius / Math.Cos(alpha);

                            Point pOuter_A = pAxis - distOuter * dir_A;
                            Point pOuter_B = pAxis - distOuter * dir_B;


                            //// Check the inner and outer edges closest to the connecter on the inner and outer face
                            var (success, innerEdge, outerEdge) = CylInfo.GetEdges(innerFace, p_InnerFace, outerFace, p_OuterFace);
                            // Derive maximal distance to edges, checking from all bottom corners of the part.
                            var (successInnerA, p_InnerEdge_A) = CylInfo.GetClosestPoint(innerEdge, pInner_A, dirZ);
                            var (successInnerB, p_InnerEdge_B) = CylInfo.GetClosestPoint(innerEdge, pInner_B, dirZ);
                            var (successOuterA, p_OuterEdge_A) = CylInfo.GetClosestPoint(outerEdge, pWidth_A, dirZ);
                            var (successOuterB, p_OuterEdge_B) = CylInfo.GetClosestPoint(outerEdge, pWidth_B, dirZ);

                            double maxDistanceBottom = 0;
                            if (successInnerA && (pInner_A - p_InnerEdge_A).Direction == dirZ)
                                maxDistanceBottom = Math.Max(maxDistanceBottom, (pInner_A - p_InnerEdge_A).Magnitude);
                            if (successInnerB && (pInner_B - p_InnerEdge_B).Direction == dirZ)
                                maxDistanceBottom = Math.Max(maxDistanceBottom, (pInner_B - p_InnerEdge_B).Magnitude);
                            if (successOuterA && (pWidth_A - p_OuterEdge_A).Direction == dirZ)
                                maxDistanceBottom = Math.Max(maxDistanceBottom, (pWidth_A - p_OuterEdge_A).Magnitude);
                            if (successOuterB && (pWidth_B - p_OuterEdge_B).Direction == dirZ)
                                maxDistanceBottom = Math.Max(maxDistanceBottom, (pWidth_B - p_OuterEdge_B).Magnitude);


                            // Create planes
                            Plane planeInner = Plane.Create(Frame.Create(pInnerMid, dirX, dirZ));
                            Plane planeOuter = Plane.Create(Frame.Create(p_OuterFace, dirX, dirZ));

                            // Check difference in width
                            double WidthDifferenceInner = CylinderStraightCut ? 0 : (pWidth_A - pWidth_B).Magnitude - (pInner_A - pInner_B).Magnitude;
                            double WidthDifferenceOuter = CylinderStraightCut ? 0 : (pWidth_A - pWidth_B).Magnitude - (pOuter_A - pOuter_B).Magnitude;


                            // For debugging, draw all points and planes

                            if (false)
                            {
                                WriteBlock.ExecuteTask("connector", () =>
                                {
                                    DatumPoint.Create(iDesignBody.Parent, "pCenter", pCenter);
                                    DatumPoint.Create(iDesignBody.Parent, "pTest", pTest);
                                    DatumPoint.Create(iDesignBody.Parent, "pAxis", pAxis);
                                    DatumPoint.Create(iDesignBody.Parent, "pWidth", pWidth);
                                    DatumPoint.Create(iDesignBody.Parent, "pWidth_A", pWidth_A);
                                    DatumPoint.Create(iDesignBody.Parent, "pWidth_B", pWidth_B);
                                    DatumPoint.Create(iDesignBody.Parent, "pOuter_B", pOuter_B);
                                    DatumPoint.Create(iDesignBody.Parent, "pOuter_A", pOuter_A);
                                    DatumPoint.Create(iDesignBody.Parent, "pInner_A", pInner_A);
                                    DatumPoint.Create(iDesignBody.Parent, "pInner_B", pInner_B);
                                    DatumPoint.Create(iDesignBody.Parent, "pInnerMid", pInnerMid);
                                    DatumPoint.Create(iDesignBody.Parent, "p_InnerFace", p_InnerFace);
                                    DatumPoint.Create(iDesignBody.Parent, "p_OuterFace", p_OuterFace);
                                    DatumPoint.Create(iDesignBody.Parent, "p_InnerEdge_A", p_InnerEdge_A);
                                    DatumPoint.Create(iDesignBody.Parent, "p_InnerEdge_B", p_InnerEdge_B);
                                    DatumPoint.Create(iDesignBody.Parent, "p_OuterEdge_A", p_OuterEdge_A);
                                    DatumPoint.Create(iDesignBody.Parent, "p_OuterEdge_B", p_OuterEdge_B);

                                    DatumPlane.Create(iDesignBody.Parent, "planeInner", planeInner);
                                    DatumPlane.Create(iDesignBody.Parent, "planeOuter", planeOuter);

                                });

                            }

                            var boundary_Outer = connector.CreateBoundary(dirX, dirPoint2Axis, dirZ, p_OuterFace, WidthDifferenceOuter, maxDistanceBottom);
                            var boundary_Inner = connector.CreateBoundary(dirX, dirPoint2Axis, dirZ, pInnerMid, WidthDifferenceInner, maxDistanceBottom);

                            if (false)
                            {
                                foreach (ITrimmedCurve curve in boundary_Outer)
                                    DesignCurve.Create(iDesignBody.Parent, curve);
                                foreach (ITrimmedCurve curve in boundary_Inner)
                                    DesignCurve.Create(iDesignBody.Parent, curve);

                            }

                            connectorBody = connector.CreateLoft(boundary_Outer, planeOuter, boundary_Inner, planeInner);

                            // Remove boundaries;
                            Plane circlePlane = Plane.Create(Frame.Create(axis.Origin - 10 * axis.Direction, axis.Direction));
                            Body cylinder = Body.ExtrudeProfile(new CircleProfile(circlePlane, outerRadius - thickness1), 20);
                            Body cylinder1 = Body.ExtrudeProfile(new CircleProfile(circlePlane, outerRadius), 20);
                            Body cylinder2 = Body.ExtrudeProfile(new CircleProfile(circlePlane, outerRadius + 1), 20);
                            cylinder2.Subtract(cylinder1);
                            connectorBody.Subtract(cylinder);
                            connectorBody.Subtract(cylinder2);


                            collisionBody = connectorBody.Copy();
                            cutBody = connectorBody.Copy();
                            cutBody.OffsetFaces(cutBody.Faces, connector.Tolerance * 0.001);

                            if (false)
                            {

                                DesignBody.Create(iDesignBody.Parent.Master, "connectorBody", connectorBody.Copy());
                                DesignBody.Create(iDesignBody.Parent.Master, "collisionBody", collisionBody.Copy());
                                DesignBody.Create(iDesignBody.Parent.Master, "cutBody", cutBody.Copy());
                            }

                        }
                        //Collision body is used to check for collisions, if there is a collision, the cutBody is subtracted. cutBodiesSource is used for the Corner Cutouts
                        // Add connector to selected Master designbody
                        DesignBody desBodyMaster = iDesignBody.Master;
                        DesignBody.Create(desBodyMaster.Parent, "connectorBody", connectorBody);
                        desBodyMaster.Shape.Unite(connectorBody);

                        if (cutBodiesSource.Count > 0)
                        {
                            foreach (Body cb in cutBodiesSource)
                            {
                                DesignBody.Create(desBodyMaster.Parent, "cut", cb);
                                desBodyMaster.Shape.Subtract(cb);
                            }
                        }

                        DesignBody.Create(desBodyMaster.Parent, "collisionBody", collisionBody);
                        DesignBody.Create(desBodyMaster.Parent, "cutBody", cutBody);


                        List<IDesignBody> _listIDB = mainPart.GetDescendants<IDesignBody>().ToList();
                        List<IDesignBody> listIDBCollisionBody = new List<IDesignBody> { };
                        List<IDesignBody> listIDBCutBody = new List<IDesignBody> { };
                        foreach (IDesignBody idb in _listIDB)
                        {
                            if (idb.Master.Shape == collisionBody)
                            {
                                listIDBCollisionBody.Add(idb);
                            }
                            if (idb.Master.Shape == cutBody)
                            {
                                listIDBCutBody.Add(idb);
                            }
                        }

                        // Subtract listIDBCutBody from _listIDB
                        _listIDB = _listIDB.Except(listIDBCollisionBody).ToList();
                        _listIDB = _listIDB.Except(listIDBCutBody).ToList();

                        foreach (IDesignBody idb in _listIDB)
                        {
                            if (idb.Master == desBodyMaster)
                                continue;

                            int j = 0;
                            foreach (IDesignBody idbCollision in listIDBCollisionBody)
                            {
                                try
                                {
                                    if (idb.Shape.GetCollision(idbCollision.Shape) == Collision.Intersect)
                                    {
                                        Body _cutBody = listIDBCutBody[j].Master.Shape.Copy();
                                        _cutBody.Transform(idbCollision.TransformToMaster.Inverse);
                                        _cutBody.Transform(idb.TransformToMaster);

                                        DesignBody.Create(idb.Master.Parent, "_cutBody", _cutBody);
                                        try
                                        {
                                            idb.Master.Shape.Subtract(_cutBody);
                                        }
                                        catch { }
                                    }
                                }
                                catch { }

                                j++;
                            }
                        }

                        foreach (IDesignBody idb in listIDBCollisionBody)
                        {
                            if (!idb.IsDeleted)
                                idb.Delete();
                        }
                        foreach (IDesignBody idb in listIDBCutBody)
                        {
                            if (!idb.IsDeleted)
                                idb.Delete();
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private static List<IDesignBody> getIDesignBodiesFromBody(Part part, Body body)
        {
            List<IDesignBody> list = new List<IDesignBody> { };
            foreach (IDesignBody idb in part.GetDescendants<IDesignBody>())
            {
                if (idb.Master.Shape == body)
                    list.Add(idb);
            }
            return list;
        }

        static Direction getNormalDirectionPlanarFace(Face face)
        {
            var plane = face.GetGeometry<Plane>();
            return plane != null ? GetPlaneNormal(plane, face.IsReversed) : Direction.Zero;
        }

        static Direction GetPlaneNormal(Plane plane, bool reversed)
        {
            Direction planeNormal = plane.Frame.DirZ;
            return reversed ? -planeNormal : planeNormal;
        }

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

        private bool checkFitsLine(IDesignEdge iDesEdge, ConnectorModel connector)
        {
            DesignEdge desEdge = iDesEdge.Master;
            Point midPoint = Point.Origin;

            var line = desEdge.Shape.Geometry as Line;
            if (line == null)
                return true;

            if (connector.ClickPosition)
            {
                Point selPoint = (Point)Window.ActiveWindow.ActiveContext.GetSelectionPoint(iDesEdge);
                midPoint = iDesEdge.Shape.ProjectPoint(selPoint).Point;
            }
            else
            {
                Point selPoint = Point.Origin;
                double paramStart = desEdge.Shape.ProjectPoint(desEdge.Shape.StartPoint).Param;
                double paramEnd = desEdge.Shape.ProjectPoint(desEdge.Shape.EndPoint).Param;
                midPoint = line.Evaluate(paramStart + 0.5 * (paramEnd - paramStart)).Point;
            }
            double minDist = Math.Min((desEdge.Shape.StartPoint - midPoint).Magnitude, (desEdge.Shape.EndPoint - midPoint).Magnitude);
            //Application.ReportStatus($"minDist: {minDist}", StatusMessageType.Information, null);
            return (minDist - 0.001 * connector.Width1 * 0.5) > 0;

        }

        private bool checkDesignEdge(IDesignEdge iDesEdge, bool clickPosition, out Point midPoint)
        {
            DesignEdge desEdge = iDesEdge.Master;

            midPoint = Point.Origin;
            // Check if line or circle or ellipse or nurbscurve

            var test = desEdge.Shape.Geometry as Line;
            var test1 = desEdge.Shape.Geometry as Circle;
            var test2 = desEdge.Shape.Geometry as NurbsCurve;
            var test3 = desEdge.Shape.Geometry as Ellipse;
            var test4 = desEdge.Shape.Geometry as ProceduralCurve;

            if (clickPosition)
            {
                midPoint = (Point)Window.ActiveWindow.ActiveContext.GetSelectionPoint(iDesEdge);
                Point af = iDesEdge.Shape.ProjectPoint(midPoint).Point;
                midPoint = iDesEdge.Shape.ProjectPoint(midPoint).Point;

                //WriteBlock.ExecuteTask("connector", () =>
                //{
                //    DatumPoint.Create(iDesEdge.Parent.Parent, "selPoint", af);
                //});

                //midPoint = midPoint + 1 * (iDesEdge.TransformToMaster).Translation;
            }
            else
            {
                Point selPoint = Point.Origin;
                double paramStart = desEdge.Shape.ProjectPoint(desEdge.Shape.StartPoint).Param;
                double paramEnd = desEdge.Shape.ProjectPoint(desEdge.Shape.EndPoint).Param;

                if (false)
                {
                    WriteBlock.ExecuteTask("connector", () =>
                    {
                        DatumPoint.Create(desEdge.Parent.Parent, "StartPoint", desEdge.Shape.StartPoint);
                        DatumPoint.Create(desEdge.Parent.Parent, "EndPoint", desEdge.Shape.EndPoint);
                    });

                }

                if (test != null)
                {
                    midPoint = test.Evaluate(paramStart + 0.5 * (paramEnd - paramStart)).Point;

                }
                else if (test1 != null)
                {
                    double paramMid = paramStart + 0.5 * (paramEnd - paramStart);
                    midPoint = test1.Evaluate(paramMid).Point;
                }
                else if (test2 != null)
                {
                    double paramMid = paramStart + 0.5 * (paramEnd - paramStart);
                    midPoint = test2.Evaluate(paramMid).Point;
                }
                else if (test3 != null)
                {
                    double paramMid = paramStart == paramEnd ? 0.5 * paramEnd : paramStart + 0.5 * (paramEnd - paramStart);
                    midPoint = test3.Evaluate(paramMid).Point;
                }
                else if (test4 != null)
                    midPoint = test4.Evaluate(0.5).Point;
                /* TODO add location paramater support */
            }



            return !(test == null && test1 == null && test2 == null && test3 == null && test4 == null);

        }

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
                connectorStraight
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

        private void OnTextChanged(object sender, TextChangedEventArgs e) => DebouncedRedraw();
        private void OnRoutedChanged(object sender, RoutedEventArgs e) => DebouncedRedraw();
        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e) => DebouncedRedraw();

        private void OnTextBoxKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                DebouncedRedraw();
            }
        }

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
                };
            }
            _uiDebounce.Stop();
            _uiDebounce.Start();
        }

        private void DrawConnector(Graphics g, ConnectorModel connector)
        {
            try
            {
                bool showTolerance = connectorShowTolerance.IsChecked == true;

                // NOTE: PictureBox client size (avoid Width/Height properties because of borders) > TODO implement tolerance height/width based on height calculated from chamfer etc.
                int canvasWidth = picDrawing.ClientSize.Width;
                int canvasHeight = picDrawing.ClientSize.Height;

                double heightTotal = connector.Height
                    + (connector.HasCornerCutout ? 2 * connector.CornerCutoutRadius : 0)
                    + (showTolerance ? connector.Tolerance : 0);

                float scale = (float)(0.9 * Math.Min(
                    canvasWidth / (float)((connector.HasCornerCutout ? connector.CornerCutoutRadius * 4 : 0) + Math.Max(connector.Width1, connector.Width2)) + (showTolerance ? connector.Tolerance : 0),
                    canvasHeight / (float)heightTotal));

                float width1 = (float)connector.Width1 * scale;
                float width2 = (float)connector.Width2 * scale;
                float height = (float)connector.Height * scale;
                float radius = (float)connector.Radius * scale;
                float cornerCutoutRadius = connector.HasCornerCutout ? (float)connector.CornerCutoutRadius * scale : 0;

                float leftX = (canvasWidth - width1) / 2f;
                float rightX = leftX + width1;
                float topX = (canvasWidth - width2) / 2f;
                float bottomY = canvasHeight / 2f + height / 2f;
                float topY = bottomY - height;

                float p0X = leftX;
                float p0Y = bottomY;
                float p1X = topX;
                float p1Y = topY;
                float p2X = topX + width2;
                float p2Y = topY;
                float p3X = leftX + width1;
                float p3Y = bottomY;

                float p01X = leftX;
                float p01Y = bottomY;
                float p02X = leftX;
                float p02Y = bottomY;

                float p11X = topX;
                float p11Y = topY;
                float p12X = topX;
                float p12Y = topY;

                float p21X = topX + width2;
                float p21Y = topY;
                float p22X = topX + width2;
                float p22Y = topY;

                float p31X = leftX + width1;
                float p31Y = bottomY;
                float p32X = leftX + width1;
                float p32Y = bottomY;

                using var pen = new System.Drawing.Pen(System.Drawing.Color.Black, 2);
                using var pen1 = new System.Drawing.Pen(System.Drawing.Color.Gray, 2);
                using var penRed = new System.Drawing.Pen(System.Drawing.Color.Red, 2);

                float dX = 0;
                float dY = 0;
                double alpha = Math.Atan(height / (0.5 * width2 - 0.5 * width1));
                double alphaDegree = alpha / Math.PI * 180;

                if (connector.HasCornerCutout)
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
                    float tol = (float)connector.Tolerance * scale;
                    double a = Math.Atan(height / (0.5 * width2 - 0.5 * width1)) / 2;
                    double hypotenuse = tol / Math.Sin(a);
                    float delta = Math.Abs((float)(hypotenuse * Math.Cos(a)));

                    using var pen2 = new Pen(Color.Gray, 2)
                    {
                        DashStyle = System.Drawing.Drawing2D.DashStyle.Dash
                    };

                    float _topX1 = topX - delta;
                    float _topY1 = topY - delta;
                    float _topX2 = topX + width2 + delta;
                    float _topY2 = topY - delta;

                    float _bottomX1 = leftX - delta;
                    float _bottomY1 = bottomY - delta;
                    float _bottomX2 = leftX + width1 + delta;
                    float _bottomY2 = _bottomY1;

                    g.DrawLine(pen2, 0, _bottomY1, _bottomX1, _bottomY1);
                    g.DrawLine(pen2, _bottomX2, _bottomY2, canvasWidth, _bottomY2);
                    g.DrawLine(pen2, _bottomX1, _bottomY1, _topX1, _topY1);
                    g.DrawLine(pen2, _bottomX2, _bottomY2, _topX2, _topY2);
                    g.DrawLine(pen2, _topX1, _topY1, _topX2, _topY2);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"DrawConnector error: {ex}");
            }
        }
        private void InvalidateDrawing()
        {
            try { picDrawing?.Invalidate(); }
            catch (Exception ex) { Logger.Log($"InvalidateDrawing error: {ex}"); }
        }
    }
}