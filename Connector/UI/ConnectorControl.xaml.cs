// SpaceClaim APIs
using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Display;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.Integration;
using System.Windows.Input;
using System.Windows.Threading;
// Aliases matching your legacy code
using Application = SpaceClaim.Api.V242.Application;
using Component = SpaceClaim.Api.V242.Component;
// IMPORTANT: your Connector class is in the global namespace
using ConnectorModel = global::Connector;
using Frame = SpaceClaim.Api.V242.Geometry.Frame;
using Line = SpaceClaim.Api.V242.Geometry.Line;
using MessageBox = System.Windows.MessageBox;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using WF = System.Windows.Forms;
using Window = SpaceClaim.Api.V242.Window;
using System.Windows.Threading;


namespace AESCConstruct25.UI
{
    public partial class ConnectorControl : UserControl
    {
        private WF.PictureBox picDrawing;
        private readonly TimeSpan UiDebounceInterval = TimeSpan.FromMilliseconds(150);
        private DispatcherTimer _uiDebounce;
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

        // === Legacy "drawConnector" adapted: no WinForms PictureBox ===
        private void drawConnector()
        {
            try
            {
                // try to build from UI; if it fails, use defaults so we still draw
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
                drawConnector(); 
            }
            catch (Exception ex) { Logger.Log($"ConnectorControl_Loaded failed: {ex}"); }
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
                        idb.Parent.Master.Name = part.DisplayName;
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
            Logger.Log("createconnector");
            try
            {

                var updated = ConnectorModel.CreateConnector(this);
                if (updated == null)
                {
                    MessageBox.Show("Please enter valid numeric values in all connector fields.",
                        "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                connector = updated;

                var activeWindow = Window.ActiveWindow;
                if (activeWindow == null)
                    return;

                Logger.Log("createconnector1");
                InteractionContext context = activeWindow.ActiveContext;
                Document doc = activeWindow.Document;
                Part mainPart = doc.MainPart;
                Logger.Log("createconnector2");

                DesignEdge desEdge = null;

                if (context.Selection.Count == 0)
                {
                    //SC.reportStatus("Please select an edge");
                    return;
                }
                Logger.Log("createconnector3");

                if (context.Selection.Count > 1)
                {
                    //SC.reportStatus("Select a single edge");
                    return;
                }

                Logger.Log("createconnector4");

                List<DesignBody> selDBodyList = new List<DesignBody> { };
                List<Point> selPointList = new List<Point> { };
                List<Part> selParts = new List<Part> { };
                List<DesignEdge> selEdges = new List<DesignEdge>();
                List<Part> suspendSheetMetalParts = new List<Part> { };
                List<IDesignBody> selIDBodyList = new List<IDesignBody> { };
                List<IDesignEdge> selIDesignEdges = new List<IDesignEdge> { };

                Logger.Log("createconnector5");
                #region Checking the selection
                foreach (IDocObject sel in context.Selection)
                {
                    Logger.Log("createconnector foreach");
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

                    Logger.Log("createconnector5.1");
                    bool correctCurve = checkDesignEdge(ide, connector.ClickPosition, out Point selPoint);

                    Logger.Log("createconnector5.2");
                    if (!correctCurve)
                    {
                        //SC.reportStatus("Selection geometry is not supported. Please select edges only.");
                        return;
                    }

                    Logger.Log("createconnector5.3");
                    // Check if it fits on the line
                    if (!checkFitsLine(ide, connector))
                    {
                        //SC.reportStatus("The connector does not fully fit on the selected line");
                        return;
                    }

                    Logger.Log("createconnector5.4");
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

                    Logger.Log("createconnector5.5");
                    selPointList.Add(selPoint);

                    selParts.Add(part);
                    selEdges.Add(desEdge);
                    selDBodyList.Add(desEdge.Parent);
                    selIDesignEdges.Add(ide);
                    if (ide != null)
                        selIDBodyList.Add(ide.Parent);

                    Logger.Log("createconnector5.6");

                    bool suspendSheetMetal = false;
                    if (part.SheetMetal != null || part.IsSheetMetalSuspended)
                        suspendSheetMetal = true;

                    Logger.Log("createconnector5.7");

                    if (suspendSheetMetal)
                        suspendSheetMetalParts.Add(part);
                    Logger.Log("createconnector5.8");
                }

                Logger.Log("createconnector6");
                #endregion

                if (selEdges.Count == 0)
                {
                    //SC.reportStatus("No connectors to be created");
                    return;
                }
                Logger.Log("createconnector7");

                // Convert & proceed with the freshly rebuilt 'connector'
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

                Logger.Log("createconnector8");

                // Iterate through all selected edges
                for (int i = 0; i < selEdges.Count; i++)
                {
                    Logger.Log("createconnector for");
                    Point pCenter = selPointList[i];
                    DesignBody designBody = selDBodyList[i];
                    IDesignBody iDesignBody = selIDBodyList[i];
                    IDesignEdge ide = selIDesignEdges[i];

                    Logger.Log("createconnector 8.1");
                    getFacesFromSelection(ide, out DesignFace bigFace, out DesignFace smallFace);
                    if (bigFace == null || smallFace == null)
                        continue;

                    Logger.Log("createconnector 8.2");
                    var bigFacePlane = bigFace.Shape.Geometry as Plane;
                    var bigFaceCylinder = bigFace.Shape.Geometry as Cylinder;

                    bool isPlane = bigFacePlane != null;
                    bool isCylinder = bigFaceCylinder != null;

                    Logger.Log("createconnector 8.3");
                    if (!isPlane && !isCylinder)
                    {
                        //SC.reportStatus("Select a planar or cylindrical surface");
                        continue;
                    }

                    Logger.Log("createconnector 8.4");
                    DesignFace oppositeFace = getOppositeFace(smallFace, bigFace, isPlane, out double thickness, out Direction dirY2);

                    Logger.Log("createconnector 8.5");

                    if (oppositeFace == null)
                    {
                        //SC.reportStatus("Could not find opposite face. Ensure the connector is made between planar or cylindrical faces");
                        continue;
                    }

                    Logger.Log("createconnector 8.6");

                    List<IDesignBody> nrBodies = getIDesignBodiesFromBody(mainPart, designBody.Shape);
                    bool allBodies = true; // Default to modifying only the selected body
                    // Check if there are more than one body
                    if (nrBodies.Count > 1)
                    {
                        var result = MessageBox.Show(
                            $"There are more than one bodies ({nrBodies.Count}) which have geometry added for the connector. Do you want to modify all identical bodies",
                            "Modify Bodies",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question);

                        if (result == MessageBoxResult.No)
                        {
                            allBodies = false;
                        }
                    }
                    Logger.Log("createconnector 8.7");


                    // check faces (small face must be planar)


                    Point pX1, pX2;
                    Body connectorBody = null;
                    Body cutBody = null;
                    Body collisionBody = null;
                    List<Body> cutBodiesSource = new List<Body>();

                    WriteBlock.ExecuteTask("connector", () =>
                    {
                        Logger.Log("createconnector WriteBlock");
                        if (isPlane)
                        {
                            // if only the selected body needs modification, make it distinct
                            if (!allBodies)
                            {
                                iDesignBody = createIndependentDesignBody(iDesignBody);
                                designBody = iDesignBody.Master;
                                getFacesFromSelection(ide, out bigFace, out smallFace);
                            }
                            Logger.Log("createconnector 8.1.1");

                            // get Directions
                            Direction dirX = desEdge.Shape.ProjectPoint(pCenter).Derivative.Direction;
                            Direction dirY = getNormalDirectionPlanarFace(bigFace.Shape);
                            Direction dirZ = getNormalDirectionPlanarFace(smallFace.Shape);

                            Logger.Log("createconnector 8.1.2");
                            Logger.Log($"createconnector dynamicHeight{dynamicHeight}");
                            // Check the max height
                            if (dynamicHeight)
                            {
                                double dynheigth = connector.GetDynamicHeigth(designBody.Parent, dirX, dirY, dirZ, pCenter, height, thickness);
                                Logger.Log("createconnector 8.1.2.1");
                                if (dynheigth > 0)
                                    height = Math.Min(height, dynheigth);

                                Logger.Log("createconnector 8.1.2.2");
                                return;
                            }

                            Logger.Log("createconnector 8.1.2.3");
                            connector.CreateGeometry(designBody.Parent, dirX, dirY, dirZ, pCenter, height, thickness, out connectorBody, out cutBodiesSource, out cutBody, out collisionBody, false);
                            //Collision body is used to check for collisions, if there is a collision, the cutBody is subtracted. cutBodiesSource is used for the Corner Cutouts

                            Logger.Log("createconnector 8.1.3");
                        }
                        else
                        {
                            Logger.Log("createconnector 8.2.1");
                            // Check if there is an inner/outer cylindrical face with same axis
                            var (innerFace, outerFace, thickness1, outerRadius, axis) = CylInfo.GetCoaxialCylPair(iDesignBody, bigFace);

                            Logger.Log("createconnector 8.2.2");

                            if (innerFace == null || outerFace == null)
                            {

                                Application.ReportStatus($"No inner/outer Face found", StatusMessageType.Information, null);
                                return;
                            }

                            Logger.Log("createconnector 8.2.3");
                            // Derive all points for drawing the inner and outer profile
                            Point pAxis = axis.ProjectPoint(pCenter).Point;
                            Direction dirPoint2Axis = (pAxis - pCenter).Direction;
                            Direction dirZ = axis.Direction;
                            // Check Correct Direction;
                            Point pTest = pCenter + 0.00001 * dirZ;
                            if (outerFace.Shape.ContainsPoint(pTest))
                                dirZ = -dirZ;

                            Logger.Log("createconnector 8.2.4");
                            Direction dirX = Direction.Cross(dirPoint2Axis, dirZ);
                            double maxWidth = Math.Max(width1, width2);
                            double distWidth = Math.Sqrt(outerRadius * outerRadius - (0.5 * maxWidth) * (0.5 * maxWidth));
                            Point pWidth = pAxis - distWidth * dirPoint2Axis;
                            Point pWidth_A = pAxis - distWidth * dirPoint2Axis + 0.5 * maxWidth * dirX;
                            Point pWidth_B = pAxis - distWidth * dirPoint2Axis - 0.5 * maxWidth * dirX;

                            Direction dir_A = (pAxis - pWidth_A).Direction;
                            Direction dir_B = (pAxis - pWidth_B).Direction;
                            Point pInner_A = pWidth_A + thickness1 * dir_A;
                            Point pInner_B = pWidth_B + thickness1 * dir_B;

                            Point pInnerMid = pInner_A - 0.5 * (pInner_A - pInner_B).Magnitude * dirX;
                            Point p_InnerFace = pAxis - (outerRadius - thickness1) * dirPoint2Axis;
                            Point p_OuterFace = pAxis - (outerRadius) * dirPoint2Axis;

                            double alpha = Math.Acos(distWidth / outerRadius);
                            double distOuter = outerRadius / Math.Cos(alpha);

                            Point pOuter_A = pAxis - distOuter * dir_A;
                            Point pOuter_B = pAxis - distOuter * dir_B;

                            Logger.Log("createconnector 8.2.5");

                            //// Check the inner and outer edges closest to the connecter on the inner and outer face
                            var (success, innerEdge, outerEdge) = CylInfo.GetEdges(innerFace, p_InnerFace, outerFace, p_OuterFace);
                            // Derive maximal distance to edges, checking from all bottom corners of the part.
                            var (successInnerA, p_InnerEdge_A) = CylInfo.GetClosestPoint(innerEdge, pInner_A, dirZ);
                            var (successInnerB, p_InnerEdge_B) = CylInfo.GetClosestPoint(innerEdge, pInner_B, dirZ);
                            var (successOuterA, p_OuterEdge_A) = CylInfo.GetClosestPoint(outerEdge, pWidth_A, dirZ);
                            var (successOuterB, p_OuterEdge_B) = CylInfo.GetClosestPoint(outerEdge, pWidth_B, dirZ);

                            Logger.Log("createconnector 8.2.6");
                            double maxDistanceBottom = 0;
                            if (successInnerA && (pInner_A - p_InnerEdge_A).Direction == dirZ)
                                maxDistanceBottom = Math.Max(maxDistanceBottom, (pInner_A - p_InnerEdge_A).Magnitude);
                            if (successInnerB && (pInner_B - p_InnerEdge_B).Direction == dirZ)
                                maxDistanceBottom = Math.Max(maxDistanceBottom, (pInner_B - p_InnerEdge_B).Magnitude);
                            if (successOuterA && (pWidth_A - p_OuterEdge_A).Direction == dirZ)
                                maxDistanceBottom = Math.Max(maxDistanceBottom, (pWidth_A - p_OuterEdge_A).Magnitude);
                            if (successOuterB && (pWidth_B - p_OuterEdge_B).Direction == dirZ)
                                maxDistanceBottom = Math.Max(maxDistanceBottom, (pWidth_B - p_OuterEdge_B).Magnitude);


                            Logger.Log("createconnector 8.2.7");
                            // Create planes
                            Plane planeInner = Plane.Create(Frame.Create(pInnerMid, dirX, dirZ));
                            Plane planeOuter = Plane.Create(Frame.Create(pCenter, dirX, dirZ));

                            Logger.Log("createconnector 8.2.8");
                            // Check difference in width
                            double WidthDifferenceInner = (pWidth_A - pWidth_B).Magnitude - (pInner_A - pInner_B).Magnitude;
                            double WidthDifferenceOuter = (pWidth_A - pWidth_B).Magnitude - (pOuter_A - pOuter_B).Magnitude;

                            Logger.Log("createconnector 8.2.9");

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

                            var boundary_Outer = connector.CreateBoundary(dirX, dirPoint2Axis, dirZ, pCenter, WidthDifferenceOuter, maxDistanceBottom);
                            var boundary_Inner = connector.CreateBoundary(dirX, dirPoint2Axis, dirZ, pInnerMid, WidthDifferenceInner, maxDistanceBottom);

                            Logger.Log("createconnector 8.2.10");
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

                            Logger.Log("createconnector 8.2.11");

                            collisionBody = connectorBody.Copy();
                            cutBody = connectorBody.Copy();
                            cutBody.OffsetFaces(cutBody.Faces, connector.Tolerance * 0.001);

                            Logger.Log("createconnector 8.2.12");
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

                        Logger.Log("createconnector 8.3");
                        if (cutBodiesSource.Count > 0)
                        {
                            foreach (Body cb in cutBodiesSource)
                            {
                                DesignBody.Create(desBodyMaster.Parent, "cut", cb);
                                desBodyMaster.Shape.Subtract(cb);
                            }
                        }
                        Logger.Log("createconnector 8.4");

                        DesignBody.Create(desBodyMaster.Parent, "collisionBody", collisionBody);
                        DesignBody.Create(desBodyMaster.Parent, "cutBody", cutBody);

                        Logger.Log("createconnector 8.5");

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
                        Logger.Log("createconnector 8.6");

                        // Subtract listIDBCutBody from _listIDB
                        _listIDB = _listIDB.Except(listIDBCollisionBody).ToList();
                        _listIDB = _listIDB.Except(listIDBCutBody).ToList();

                        Logger.Log("createconnector 8.7");
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

                        Logger.Log("createconnector 8.8");

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

                        Logger.Log("createconnector 8.9");


                    });




                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private static Matrix getPathToRoot(IDesignBody idb)
        {
            Matrix returnMatrix = Matrix.Identity;

            IPart parent = idb.Parent;
            IInstance test = parent.Parent;
            IDocObject test2 = parent.Parent.Parent;
            int a = 1;
            //while (parent != null)
            //{
            //    returnMatrix = returnMatrix * parent.TransformToMaster;
            //    parent = parent.Parent.Parent;
            //}

            return returnMatrix;

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
        private static List<Matrix> getMatrixFromBody(Part part, Body body)
        {
            List<Matrix> list = new List<Matrix> { };
            foreach (IDesignBody idb in part.GetDescendants<IDesignBody>())
            {
                if (idb.Master.Shape == body)
                    list.Add(idb.TransformToMaster);
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
        private static double DotProduct(Point point, Direction direction)
        {
            return point.X * direction.X + point.Y * direction.Y + point.Z * direction.Z;
        }

        public static List<double> FindLowestHighestAndFurthestPoint(Direction direction, List<Point> points)
        {
            if (points == null || points.Count == 0)
                throw new ArgumentException("Points list cannot be null or empty");

            Point lowestPoint = points[0];
            Point highestPoint = points[0];
            double minProjection = DotProduct(lowestPoint, direction);
            double maxProjection = minProjection;
            List<double> result = new List<double>();

            foreach (var point in points)
            {
                double projection = DotProduct(point, direction);
                result.Add(projection);
                if (projection < minProjection)
                {
                    minProjection = projection;
                    lowestPoint = point;
                }
                if (projection > maxProjection)
                {
                    maxProjection = projection;
                    highestPoint = point;
                }
            }

            double maxDistance = maxProjection - minProjection;

            List<double> result1 = new List<double>();
            for (int i = 0; i < result.Count; i++)
                result1.Add(maxProjection - result[i]);



            return result1;
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
            Application.ReportStatus($"minDist: {minDist}", StatusMessageType.Information, null);
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

                //if (true)
                //{
                //    WriteBlock.ExecuteTask("connector", () =>
                //    {
                //        DatumPoint.Create(desEdge.Parent.Parent, "StartPoint", desEdge.Shape.StartPoint);
                //        DatumPoint.Create(desEdge.Parent.Parent, "EndPoint", desEdge.Shape.EndPoint);
                //    });

                //}

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

            }



            return !(test == null && test1 == null && test2 == null && test3 == null && test4 == null);

        }

        //private void createConnector1()
        //{
        //    try
        //    {
        //        var activeWindow = Window.ActiveWindow;
        //        if (activeWindow == null)
        //            return;

        //        InteractionContext context = activeWindow.ActiveContext;
        //        Document doc = activeWindow.Document;
        //        Part mainPart = doc.MainPart;

        //        DesignEdge selectedEdge = null;

        //        DesignEdge desEdge = null;

        //        if (context.Selection.Count == 0)
        //        {
        //            //SC.reportStatus("Please select an edge");
        //            return;
        //        }

        //        List<DesignBody> selDBodyList = new List<DesignBody> { };
        //        List<Point> selPointList = new List<Point> { };
        //        List<Part> selParts = new List<Part> { };
        //        List<DesignEdge> selEdges = new List<DesignEdge>();
        //        List<Part> suspendSheetMetalParts = new List<Part> { };

        //        #region Checking the selection
        //        foreach (IDocObject sel in context.Selection)
        //        {
        //            var de = sel as DesignEdge;
        //            if (de != null)
        //                desEdge = de;

        //            try
        //            {
        //                var de1 = sel.Master as DesignEdge;
        //                if (de1 != null)
        //                    desEdge = de1;
        //            }
        //            catch { }

        //            if (desEdge == null)
        //            {
        //                //SC.reportStatus("Please select an edge");
        //                continue;
        //            }

        //            // Check if line or circle or ellipse or nurbscurve
        //            var test = desEdge.Shape.Geometry as Line;
        //            var test1 = desEdge.Shape.Geometry as Circle;
        //            var test2 = desEdge.Shape.Geometry as NurbsCurve;
        //            var test3 = desEdge.Shape.Geometry as Ellipse;

        //            Point selPoint = Point.Origin;
        //            if (test != null)
        //                selPoint = test.Evaluate(0.5 * desEdge.Shape.Length).Point;
        //            else if (test1 != null)
        //                selPoint = test1.Evaluate(0.5).Point;
        //            else if (test2 != null)
        //                selPoint = test2.Evaluate(0.5).Point;
        //            else if (test3 != null)
        //                selPoint = test3.Evaluate(0.5).Point;


        //            if (test == null && test1 == null && test2 == null && test3 == null)
        //            {
        //                //SC.reportStatus("Selection geometry is not supported. Please select edges only.");
        //                continue;
        //            }

        //            // Check if in component or single body
        //            Part part = null;
        //            if (mainPart.GetDescendants<IDesignBody>().Count > 1)
        //            {
        //                part = de.Parent.Parent;
        //                if (mainPart == part)
        //                {
        //                    //SC.reportStatus("Selected body is not within a component.");
        //                    continue;
        //                }
        //            }
        //            else
        //                part = mainPart;

        //            selPointList.Add(selPoint);


        //            selParts.Add(part);
        //            selEdges.Add(de);
        //            selDBodyList.Add(de.Parent);


        //            //WriteBlock.ExecuteTask("connector", () =>
        //            //{
        //            //    DatumPoint.Create(mainPart, "selPoint", selPoint);
        //            //});
        //            // CHeck sheet metal
        //            bool suspendSheetMetal = false;
        //            if (part.SheetMetal != null || part.IsSheetMetalSuspended)
        //                suspendSheetMetal = true;


        //            if (suspendSheetMetal)
        //                suspendSheetMetalParts.Add(part);
        //        }

        //        #endregion

        //        if (selEdges.Count == 0)
        //        {
        //            //SC.reportStatus("No connectors to be created");
        //            return;
        //        }

        //        // Get parameters
        //        double width1 = 0.001 * connector.Width1;
        //        double width2 = 0.001 * connector.Width2;
        //        double tolerance = 0.001 * connector.Tolerance;
        //        double height = 0.001 * connector.Height;
        //        double radius = 0.001 * connector.Radius;
        //        bool hasRounding = connector.HasRounding;
        //        bool hasCornerCutout = connector.HasCornerCutout;
        //        double cornerCutoutRadius = 0.001 * connector.CornerCutoutRadius;
        //        bool dynamicHeight = connector.DynamicHeight;


        //        for (int i = 0; i < selEdges.Count; i++)
        //        {
        //            Point pCenter = selPointList[i];
        //            // Determine biggest face
        //            DesignFace bigFace = null;
        //            DesignFace smallFace = null;

        //            DesignFace face0 = desEdge.Faces.ElementAt(0);
        //            DesignFace face1 = desEdge.Faces.ElementAt(1);

        //            if (desEdge.Faces.Count != 2)
        //                continue;

        //            if (face0.Area > face1.Area)
        //            {
        //                bigFace = face0;
        //                smallFace = face1;
        //            }
        //            else if (face0.Area < face1.Area)
        //            {
        //                bigFace = face1;
        //                smallFace = face0;
        //            }
        //            else
        //                return;

        //            // check faces (small face must be planar)

        //            // get Directions
        //            Direction dirZ = smallFace.Shape.ProjectPoint(pCenter).Normal;
        //            Direction dirY = bigFace.Shape.ProjectPoint(pCenter).Normal;
        //            Direction dirX = Direction.Cross(dirY, dirZ);


        //            // Additional values
        //            double tol2 = 0.0001;
        //            double alpha = Math.Atan(height / (0.5 * width2 - 0.5 * width1));
        //            double alphaDegree = alpha / Math.PI * 180;
        //            double dX = 0;
        //            double dY = 0;

        //            if (connector.HasCornerCutout)
        //            {
        //                dX = Math.Abs((cornerCutoutRadius * Math.Cos(alpha)));
        //                dY = -Math.Abs((cornerCutoutRadius * Math.Sin(alpha)));

        //                if (width1 > width2)
        //                {
        //                    dX = -dX;
        //                }
        //            }

        //            // Define points

        //            // outer contours without roundings/cutouts
        //            Point B0 = pCenter - (0.5 * width1) * dirX;
        //            Point B3 = pCenter + (0.5 * width2) * dirX;
        //            Point B1 = pCenter - (0.5 * width2) * dirX + height * dirZ;
        //            Point B2 = pCenter + (0.5 * width1) * dirX + height * dirZ;

        //            Point p0 = pCenter - (0.5 * width1 + tol2) * dirX;
        //            Point p1 = pCenter - (0.5 * width1) * dirX;
        //            Point p9 = p0 - tol2 * dirZ;
        //            Point p7 = pCenter + (0.5 * width1 + tol2) * dirX;
        //            Point p6 = pCenter + (0.5 * width1) * dirX;
        //            Point p8 = p7 - tol2 * dirZ;

        //            Point p2 = B1;
        //            Point p3 = B1;
        //            Point p4 = B2;
        //            Point p5 = B2;

        //            if (radius > 0)
        //            {
        //                double dX2 = radius;

        //                if (width1 <= width2)
        //                {
        //                    if (connector.HasRounding)
        //                        dX2 = (radius / Math.Tan(alpha / 2));

        //                    double dX1 = (dX2 * Math.Sin(Math.PI * 0.5 - alpha));
        //                    double dY1 = (dX2 * Math.Cos(Math.PI * 0.5 - alpha));

        //                    p2 = B1 + dX1 * dirX - dY1 * dirZ;
        //                    p3 = B1 + dX2 * dirX;
        //                    p4 = B2 - dX2 * dirX;
        //                    p5 = B2 - dX1 * dirX - dY1 * dirZ;
        //                }
        //                else
        //                {
        //                    double beta = Math.PI + alpha;
        //                    if (connector.HasRounding)
        //                        dX2 = (radius * Math.Tan(0.5 * (Math.PI - beta)));

        //                    double dX1 = ((dX2 * Math.Sin(beta - Math.PI * 0.5)));
        //                    double dY1 = ((dX2 * Math.Cos(beta - Math.PI * 0.5)));

        //                    p2 = B1 + dX1 * dirX - dY1 * dirZ;
        //                    p3 = B1 + dX2 * dirX;
        //                    p4 = B2 - dX2 * dirX;
        //                    p5 = B2 - dX1 * dirX - dY1 * dirZ;
        //                }
        //            }


        //            Frame frame = Frame.Create(pCenter, dirX, dirZ);
        //            Plane plane = Plane.Create(frame);

        //            // create curves for the addBody
        //            List<ITrimmedCurve> boundaryAddBody = new List<ITrimmedCurve>
        //            {
        //                CurveSegment.Create(p0, p1),
        //                CurveSegment.Create(p1, p2),
        //                CurveSegment.Create(p2, p3),
        //                CurveSegment.Create(p3, p4),
        //                CurveSegment.Create(p4, p5),
        //                CurveSegment.Create(p5, p6),
        //                CurveSegment.Create(p6, p7),
        //                CurveSegment.Create(p7, p8),
        //                CurveSegment.Create(p8, p9),
        //                CurveSegment.Create(p9, p0),
        //            };
        //            int direction = 1;
        //            if (plane.Frame.DirZ != dirZ)
        //                direction = -1;

        //            WriteBlock.ExecuteTask("connector", () =>
        //            {

        //                DesignCurve.Create(mainPart, CurveSegment.Create(B0, B1));
        //                DesignCurve.Create(mainPart, CurveSegment.Create(B1, B2));
        //                DesignCurve.Create(mainPart, CurveSegment.Create(B2, B3));
        //                DesignCurve.Create(mainPart, CurveSegment.Create(B3, B0));
        //                //    DatumPoint.Create(mainPart, "B0", B0);
        //                //    DatumPoint.Create(mainPart, "B1", B1);
        //                //    DatumPoint.Create(mainPart, "B2", B2);
        //                //    DatumPoint.Create(mainPart, "B3", B3);
        //                DatumPoint.Create(mainPart, "p0", p0);
        //                DatumPoint.Create(mainPart, "p1", p1);
        //                DatumPoint.Create(mainPart, "p3", p3);
        //                DatumPoint.Create(mainPart, "p2", p2);
        //                DatumPoint.Create(mainPart, "p4", p4);
        //                DatumPoint.Create(mainPart, "p5", p5);
        //                DatumPoint.Create(mainPart, "p6", p6);
        //                DatumPoint.Create(mainPart, "p7", p7);
        //                DatumPoint.Create(mainPart, "p8", p8);
        //                DatumPoint.Create(mainPart, "p9", p9);


        //                try
        //                {
        //                    Body addBody = Body.ExtrudeProfile(new Profile(plane, boundaryAddBody), direction * height);
        //                    DesignBody.Create(mainPart, "addBody", addBody);
        //                }
        //                catch
        //                {
        //                    foreach (ITrimmedCurve itc in boundaryAddBody)
        //                    {
        //                        try
        //                        {
        //                            DesignCurve dc = DesignCurve.Create(mainPart, itc);
        //                            //dc.SetColor(null, Color.Red);
        //                        }
        //                        catch { }
        //                    }
        //                }


        //            });

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //}

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
                connectorUseCustom,
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


        private void OnUiChanged(object sender, RoutedEventArgs e)
        {
            DebouncedRedraw();
        }

        private void OnTextChanged(object sender, TextChangedEventArgs e) => DebouncedRedraw();
        private void OnRoutedChanged(object sender, RoutedEventArgs e) => DebouncedRedraw();
        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e) => DebouncedRedraw();

        private void OnTextBoxKeyDown(object sender, KeyEventArgs e)
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


        private void picDrawing_Paint(object sender, WF.PaintEventArgs e)
        {
            if (connector != null)
            {
                DrawConnector(e.Graphics, connector);
            }
        }

        private void DrawConnector(System.Drawing.Graphics g, ConnectorModel connector)
        {
            try
            {
                bool showTolerance = connectorShowTolerance.IsChecked == true;

                // NOTE: PictureBox client size (avoid Width/Height properties because of borders)
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

                    using var pen2 = new System.Drawing.Pen(System.Drawing.Color.Gray, 2)
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

        private void chk_CheckedChanged(object sender, EventArgs e)
        {
            drawConnector();
        }

        private void FormConnector_Load(object sender, EventArgs e)
        {

            drawConnector();
        }
    }
}
