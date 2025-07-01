using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpaceClaim.Api.V242.Modeler;
using Document = SpaceClaim.Api.V242.Document;
using System.Windows.Forms;
using Application = SpaceClaim.Api.V242.Application;

namespace AESCConstruct25.Fastener.Module
{
    internal class FastenerModule
    {
        //private static Part GetWasherPart(Document doc, string name, Washer washer, out Component component)
        //{
        //    // Try to find an existing “Fasteners” part
        //    Part part = doc.Parts
        //        .FirstOrDefault(p => p.DisplayName == name);

        //    // Create Fasteners Part if not existing
        //    Part fastenersPart = GetFastenersPart(doc);
        //    if (part is null)
        //    {

        //        // if not found, create it and hook it up
        //        part = Part.Create(doc, name);
        //        component = Component.Create(fastenersPart, part);

        //        // Create body
        //        Body body = createFasteners.createWasher(washer.d1 * 0.001, washer.d2 * 0.001, washer.s * 0.001);

        //        DesignBody dbBolt = DesignBody.Create(part, name, body);
        //        CustomPartProperty.Create(part, "AESC_Construct", true);
        //    }
        //    else
        //    {
        //        // Create a copy of the original component
        //        component = Component.Create(fastenersPart, part);
        //    }
        //    return part;
        //}

        //public static List<PlacementData> GetPlacementDataSelection(ICollection<IDocObject> selection, ICollection<IDocObject> secondSelection)
        //{
        //    // 1) Make sure we actually have something selected
        //    if (selection == null || selection.Count == 0)
        //    {
        //        Application.ReportStatus("No objects selected.", StatusMessageType.Information, null);
        //        return null;
        //    }

        //    if (selection.Count > 1 && secondSelection.Count > 0)
        //    {
        //        Application.ReportStatus("Please select a single edge when selecting faces.", StatusMessageType.Information, null);
        //        return null;
        //    }

        //    List<IDesignFace> planarFaces = new List<IDesignFace>();
        //    if (secondSelection.Count > 0)
        //    {
        //        foreach (var dobj in secondSelection)
        //        {
        //            var selface = dobj as IDesignFace;
        //            if (selface == null)
        //            {
        //                Application.ReportStatus("Second selection must be planar face(s).", StatusMessageType.Information, null);
        //                return null;
        //            }
        //            var planarFace = selface.Shape.GetGeometry<Plane>();
        //            if (planarFace == null)
        //            {
        //                Application.ReportStatus("Second selection must be planar face(s).", StatusMessageType.Information, null);
        //                return null;
        //            }
        //            planarFaces.Add(selface);
        //        }
        //    }


        //    var placements = new List<PlacementData>();

        //    foreach (var dobj in selection)
        //    {
        //        // expect the selection to be a DesignEdge whose geometry is a Circle
        //        var edge = dobj as IDesignEdge;
        //        if (edge == null)
        //        {
        //            Application.ReportStatus("Selection contains non‐edge objects.", StatusMessageType.Information, null);
        //            return null;
        //        }

        //        // try to get a Circle from the edge
        //        var circle = edge.Shape.GetGeometry<Circle>();
        //        if (circle == null)
        //        {
        //            Application.ReportStatus("All selected edges must be circles.", StatusMessageType.Information, null);
        //            return null;
        //        }

        //        if (planarFaces.Count > 0)
        //        {
        //            double radius = circle.Radius;
        //            placements.AddRange(CreatePlacementDataForCirclesInPlanes(planarFaces, radius));
        //        }
        //        else
        //        {
        //            placements.Add(CreatePlacementData(circle, edge));
        //        }

        //    }

        //    return placements;
        //}

        //private static List<PlacementData> CreatePlacementDataForCirclesInPlanes(List<IDesignFace> planarFaces, double radius)
        //{
        //    List<PlacementData> returnList = new List<PlacementData>();
        //    foreach (IDesignFace idf in planarFaces)
        //    {
        //        foreach (IDesignEdge ide in idf.Edges)
        //        {
        //            var circle = ide.Shape.GetGeometry<Circle>();
        //            if (circle == null)
        //            {
        //                continue;
        //            }
        //            if (Math.Abs(radius - circle.Radius) < 1e-9)
        //                returnList.Add(CreatePlacementData(circle, ide));
        //        }
        //    }
        //    return returnList;
        //}

        //private static PlacementData CreatePlacementData(Circle circle, IDesignEdge edge)
        //{
        //    // 2) origin is simply the circle's axis origin
        //    var origin = circle.Axis.Origin;

        //    // Check if there is no other designface in the center of the circle
        //    Part mainPart = edge.Root as Part;
        //    foreach (IDesignBody idb in mainPart.GetDescendants<IDesignBody>())
        //        if (idb.Shape.ContainsPoint(origin))
        //        {
        //            Application.ReportStatus("There is an object in the hole, no fastener created.", StatusMessageType.Information, null);
        //            return null;
        //        }

        //    // 3) direction: look for an adjacent cylindrical face
        //    Direction direction = circle.Axis.Direction; // fallback
        //    double depth = 0;
        //    foreach (IDesignFace face in edge.Faces)
        //    {
        //        var cyl = face.Master.Shape.GetGeometry<Cylinder>();

        //        if (cyl != null)
        //        {
        //            direction = face.TransformToMaster.Inverse * cyl.Axis.Direction;
        //            Point test1 = origin + 0.01 * direction;
        //            Point test2 = origin - 0.01 * direction;
        //            Point projectPoint1 = face.Shape.ProjectPoint(test1).Point;
        //            Point projectPoint2 = face.Shape.ProjectPoint(test2).Point;
        //            double dist1 = (projectPoint1 - test1).Magnitude;
        //            double dist2 = (projectPoint2 - test2).Magnitude;
        //            //WriteBlock.ExecuteTask("UpdateComponentNames", () =>
        //            //{
        //            //    DatumPoint.Create(face.Parent.Parent, "origin", origin);
        //            //    DatumPoint.Create(face.Parent.Parent, "test1", test1);
        //            //    DatumPoint.Create(face.Parent.Parent, "test2", test2);
        //            //    DatumPoint.Create(face.Parent.Parent, "projectPoint1", projectPoint1);
        //            //    DatumPoint.Create(face.Parent.Parent, "projectPoint2", projectPoint2);
        //            //});
        //            direction = dist1 > dist2 ? direction : -direction;

        //            var cyl2 = face.Shape.Geometry as Cylinder;
        //            Matrix mat = Matrix.CreateMapping(Frame.Create(cyl2.Axis.Origin, cyl2.Axis.Direction)).Inverse;
        //            Box boundingBox = face.Shape.GetBoundingBox(mat, true);
        //            depth = boundingBox.Size.Z;

        //            break;

        //        }
        //    }

        //    return new PlacementData
        //    {
        //        Circle = circle,
        //        Origin = origin,
        //        Direction = direction,
        //        Depth = depth
        //    };
        //}


        //private void CreateFasteners()
        //{
        //    // Check selection
        //    Window win = Window.ActiveWindow;
        //    Document doc = win.Document;
        //    Part mainPart = doc.MainPart;

        //    if (!CheckSelectedCircle(win))
        //        return;

        //    ICollection<IDocObject> selection = win.ActiveContext.Selection;
        //    ICollection<IDocObject> secSelection = win.ActiveContext.SecondarySelection;


        //    if (selection.Count > 1 && secSelection.Count > 0)
        //    {
        //        Application.ReportStatus("Please select only one circle when selecting faces", StatusMessageType.Information, null);
        //        return;
        //    }

        //    // To do, get list of center points and directions from the selections
        //    List<PlacementData> placementDatas = GetPlacementDataSelection(selection, secSelection);
        //    if (placementDatas == null)
        //    {
        //        Application.ReportStatus("Could not find correct placement(s)", StatusMessageType.Information, null);
        //        return;
        //    }
        //    ReadUIControls(); // inlezen van data vanuit UI

        //    try
        //    {   //Execute a Command 
        //        WriteBlock.ExecuteTask("AESCConstruct.FastenersCreate",
        //        delegate
        //        {
        //            Bolt _bolt = listBolt.First();

        //            foreach (Bolt bolt in listBolt)
        //            {
        //                if (bolt.type == boltType && bolt.size == boltSize)
        //                {
        //                    _bolt = bolt;
        //                }
        //            }

        //            foreach (Bolt bolt in listBolt)
        //            {
        //                if (bolt.type == boltType && bolt.size == boltSize)
        //                {
        //                    _bolt = bolt;

        //                }
        //            }

        //            Washer _washerBottom = listWasherBottom.First();
        //            Washer _washerTop = listWasherTop.First();

        //            foreach (Washer washer in listWasherBottom)
        //            {
        //                if (washer.type == washerBottomType && washer.size == washerBottomSize)
        //                {
        //                    _washerBottom = washer;
        //                }
        //            }
        //            foreach (Washer washer in listWasherTop)
        //            {
        //                if (washer.type == washerTopType && washer.size == washerTopSize)
        //                {
        //                    _washerTop = washer;
        //                }
        //            }
        //            Nut _nut = listNut.First();
        //            foreach (Nut nut in listNut)
        //            {
        //                if (nut.type == nutType && nut.size == nutSize)
        //                {
        //                    _nut = nut;
        //                }
        //            }

        //            string customPartPath = "none";
        //            if (File.Exists(_addinPathProfiles + "\\Fasteners\\Custom\\" + customFile))
        //                customPartPath = _addinPathProfiles + "\\Fasteners\\Custom\\" + customFile;

        //            // TODO:
        //            // - handle custom part
        //            // - Lock parts

        //            ///// Create Geometry     
        //            string boltName = boltType.Split(' ')[0] + " " + boltSize + " x " + (parBoltL).ToString();
        //            string washerTopName = washerTopType.Split(' ')[0] + " " + washerTopSize;
        //            string washerBottomName = washerBottomType.Split(' ')[0] + " " + washerBottomType;
        //            string nutName = nutType.Split(' ')[0] + " " + nutSize;

        //            foreach (PlacementData PD in placementDatas)
        //            {
        //                if (PD == null)
        //                    continue;
        //                Part boltPart = GetBoltPart(doc, boltName, boltType, _bolt, parBoltL, out Component componentBolt);
        //                Matrix matrixMapping = Matrix.CreateMapping(Frame.Create(PD.Origin, PD.Direction));
        //                Matrix matrixBolt = matrixMapping;

        //                if (chkFastenersIncludeWasherTop.Checked)
        //                {
        //                    Part washerPart = GetWasherPart(doc, washerTopName, _washerTop, out Component componentWasherTop);
        //                    matrixBolt = matrixMapping * Matrix.CreateTranslation(Vector.Create(0, 0, _washerTop.s * 0.001));
        //                    componentWasherTop.Transform(matrixMapping);
        //                }
        //                componentBolt.Transform(matrixBolt);
        //                double displacementZ = -PD.Depth;

        //                if (chkFastenersIncludeWasherBottom.Checked)
        //                {
        //                    Part washerPart = GetWasherPart(doc, washerBottomName, _washerBottom, out Component componentWasherBottom);
        //                    displacementZ += -0.001 * _washerBottom.s;
        //                    componentWasherBottom.Transform(matrixMapping * Matrix.CreateTranslation(Vector.Create(0, 0, displacementZ)));
        //                }

        //                if (chkFastenersIncludeNut.Checked)
        //                {
        //                    Part nutPart = GetNutPart(doc, nutName, nutType, _nut, out Component componentNut);
        //                    displacementZ += -0.001 * _nut.h;
        //                    componentNut.Transform(matrixMapping * Matrix.CreateTranslation(Vector.Create(0, 0, displacementZ)));
        //                }
        //            }
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //}

        //private static Part GetFastenersPart(Document doc)
        //{
        //    // Try to find an existing “Fasteners” part
        //    var fastenersPart = doc.Parts
        //        .FirstOrDefault(p => p.DisplayName == "Fasteners");

        //    if (fastenersPart is null)
        //    {
        //        // if not found, create it and hook it up
        //        fastenersPart = Part.Create(doc, "Fasteners");
        //        Component.Create(doc.MainPart, fastenersPart);

        //        CustomPartProperty.Create(fastenersPart, "AESC_Construct", true);

        //    }
        //    return fastenersPart;
        //}

        //private static Part GetBoltPart(Document doc, string name, string boltType, Bolt _bolt, double parBoltL, out Component component)
        //{
        //    // Try to find an existing “Fasteners” part
        //    Part boltPart = doc.Parts
        //        .FirstOrDefault(p => p.DisplayName == name);

        //    // Create Fasteners Part if not existing
        //    Part fastenersPart = GetFastenersPart(doc);
        //    if (boltPart is null)
        //    {

        //        // if not found, create it and hook it up
        //        boltPart = Part.Create(doc, name);
        //        component = Component.Create(fastenersPart, boltPart);

        //        // Create body
        //        Body bodyBolt = createFasteners.Create_Bolt(boltType, _bolt, parBoltL);
        //        DesignBody dbBolt = DesignBody.Create(boltPart, name, bodyBolt);
        //        CustomPartProperty.Create(boltPart, "AESC_Construct", true);
        //    }
        //    else
        //    {
        //        // Create a copy of the original component
        //        component = Component.Create(fastenersPart, boltPart);
        //    }
        //    return boltPart;
        //}

        //private static Part GetNutPart(Document doc, string name, string type, Nut nut, out Component component)
        //{
        //    // Try to find an existing “Fasteners” part
        //    Part part = doc.Parts
        //        .FirstOrDefault(p => p.DisplayName == name);

        //    // Create Fasteners Part if not existing
        //    Part fastenersPart = GetFastenersPart(doc);
        //    if (part is null)
        //    {

        //        // if not found, create it and hook it up
        //        part = Part.Create(doc, name);
        //        component = Component.Create(fastenersPart, part);

        //        // Create body
        //        Body body = createFasteners.Create_Nut(boltType, nut);
        //        DesignBody designBody = DesignBody.Create(part, name, body);
        //        CustomPartProperty.Create(part, "AESC_Construct", true);
        //    }
        //    else
        //    {
        //        // Create a copy of the original component
        //        component = Component.Create(fastenersPart, part);
        //    }
        //    return part;
        //}

        //private bool CheckSelectedCircle(Window window, bool singleSelection = false)
        //{
        //    var selection = window.ActiveContext.Selection;

        //    // no selection or too many
        //    if (selection == null || selection.Count == 0 ||
        //        (singleSelection && selection.Count != 1))
        //    {
        //        Application.ReportStatus("Please select a circle.", StatusMessageType.Information, null);
        //        return false;
        //    }

        //    // helper to grab a Circle from either a curve or an edge
        //    Circle GetCircle(IDocObject obj)
        //    {
        //        var curve = obj as IDesignCurve;
        //        if (curve != null)
        //            return curve.Shape.GetGeometry<Circle>();

        //        var edge = obj.Master as DesignEdge;
        //        if (edge != null)
        //            return edge.Shape.GetGeometry<Circle>();

        //        return null;
        //    }

        //    // ensure every selected item is actually a circle
        //    foreach (var obj in selection)
        //    {
        //        if (GetCircle(obj) == null)
        //        {
        //            Application.ReportStatus("Please select a circle.", StatusMessageType.Information, null);
        //            return false;
        //        }
        //    }

        //    return true;
        //}


        //private double GetSizeCircle(Window window, out double depthMM)
        //{
        //    double radiusMM = 0;

        //    depthMM = 0;

        //    InteractionContext interContext = window.ActiveContext;
        //    var selection = interContext.SingleSelection;

        //    if (selection != null)
        //    {
        //        var desCurve = selection as IDesignCurve;
        //        var desEdge = selection.Master as DesignEdge;
        //        if (desCurve != null)
        //        {
        //            string desCurveType = desCurve.Shape.Geometry.GetType().Name;

        //            if (desCurveType == "Circle")
        //            {
        //                var circle = desCurve.Shape.GetGeometry<Circle>();
        //                radiusMM = 1000 * circle.Radius;
        //            }
        //        }
        //        else if (desEdge != null)
        //        {
        //            string desEdgeType = desEdge.Shape.Geometry.GetType().Name;

        //            if (desEdgeType == "Circle")
        //            {
        //                var circle = desEdge.Shape.GetGeometry<Circle>();

        //                radiusMM = 1000 * circle.Radius;
        //                foreach (DesignFace df in desEdge.Faces)
        //                {
        //                    var cyl = df.Shape.Geometry as Cylinder;
        //                    if (cyl != null)
        //                    {
        //                        Matrix mat = Matrix.CreateMapping(Frame.Create(cyl.Axis.Origin, cyl.Axis.Direction)).Inverse;
        //                        Box boundingBox = df.Shape.GetBoundingBox(mat, true);
        //                        depthMM = boundingBox.Size.Z * 1000;
        //                    }

        //                }
        //            }
        //        }
        //        else
        //        {
        //            Application.ReportStatus("Please select a circle.", StatusMessageType.Information, null);
        //        }

        //    }

        //    return radiusMM;
        //}


        //private void CheckSize()
        //{

        //    Window window = Window.ActiveWindow;
        //    Document doc = window.Document;
        //    Part rootPart = doc.MainPart;

        //    if (!CheckSelectedCircle(window, true))
        //        return;

        //    double radiusMM = GetSizeCircle(window, out double depthMM);

        //    if (radiusMM == 0)
        //        return;

        //    //// Apply filter to the Sizes of all comboboxes
        //    try
        //    {
        //        List<Bolt> listBolt = File.ReadAllLines(_addinPathProfiles + "\\Fasteners\\" + "Bolt.csv").Skip(1)
        //        .Select(v => Bolt.FromCsv(v))
        //        .ToList();

        //        cboFastenersFastenersize.Items.Clear();
        //        foreach (Bolt bolt in listBolt)
        //        {
        //            if (bolt.type == cboFastenersBoltType.SelectedItem.ToString())
        //            {
        //                if (bolt.d < 2 * radiusMM)
        //                {
        //                    if (!cboFastenersFastenersize.Items.Cast<string>().Contains(bolt.size))
        //                    {
        //                        cboFastenersFastenersize.Items.Add(bolt.size);
        //                    }
        //                }
        //            }
        //        }
        //        cboFastenersFastenersize.DisplayMember = "size";
        //        cboFastenersFastenersize.SelectedIndex = cboFastenersFastenersize.Items.Count - 1;
        //    }
        //    catch
        //    {
        //        MessageBox.Show("Unable to find/read the file Filter: " + _addinPathProfiles + "\\Fasteners\\" + "Bolt.csv", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }

        //    try
        //    {
        //        List<Washer> listWasherTop = File.ReadAllLines(_addinPathProfiles + "\\Fasteners\\" + "Washer.csv").Skip(1)
        //  .Select(v => Washer.FromCsv(v))
        //  .ToList();

        //        cboFastenersWasherTopSize.Items.Clear();
        //        foreach (Washer washer in listWasherTop)
        //        {
        //            if (washer.type == cboFastenersWasherTopType.SelectedItem.ToString())
        //            {
        //                if (Convert.ToDouble(washer.size.TrimStart('M')) < 2 * radiusMM)
        //                {
        //                    if (!cboFastenersWasherTopSize.Items.Cast<string>().Contains(washer.size))
        //                    {
        //                        cboFastenersWasherTopSize.Items.Add(washer.size);
        //                    }
        //                }
        //            }
        //        }
        //        cboFastenersWasherTopSize.DisplayMember = "size";
        //        cboFastenersWasherTopSize.SelectedIndex = cboFastenersWasherTopSize.Items.Count - 1;
        //    }
        //    catch
        //    {
        //        MessageBox.Show("Unable to find/read the file Filter: " + _addinPathProfiles + "\\Fasteners\\" + "Washer.csv", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }

        //    try
        //    {
        //        List<Washer> listWasherBottom = File.ReadAllLines(_addinPathProfiles + "\\Fasteners\\" + "Washer.csv").Skip(1)
        //  .Select(v => Washer.FromCsv(v))
        //  .ToList();

        //        cboFastenersWasherBottomSize.Items.Clear();
        //        foreach (Washer washer in listWasherBottom)
        //        {
        //            if (washer.type == cboFastenersWasherBottomType.SelectedItem.ToString())
        //            {
        //                if (Convert.ToDouble(washer.size.TrimStart('M')) < 2 * radiusMM)
        //                {
        //                    if (!cboFastenersWasherBottomSize.Items.Cast<string>().Contains(washer.size))
        //                    {
        //                        cboFastenersWasherBottomSize.Items.Add(washer.size);
        //                    }
        //                }
        //            }
        //        }
        //        cboFastenersWasherBottomSize.DisplayMember = "size";
        //        cboFastenersWasherBottomSize.SelectedIndex = cboFastenersWasherBottomSize.Items.Count - 1;
        //    }
        //    catch
        //    {
        //        MessageBox.Show("Unable to find/read the file Filter Washer Bottom: " + _addinPathProfiles + "\\Fasteners\\" + "Washer.csv", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }



        //    try
        //    {
        //        List<Nut> listNut = File.ReadAllLines(_addinPathProfiles + "\\Fasteners\\" + "Nut.csv").Skip(1)
        //  .Select(v => Nut.FromCsv(v))
        //  .ToList();


        //        cboFastenersNutSize.Items.Clear();
        //        foreach (Nut nut in listNut)
        //        {
        //            if (nut.type == cboFastenersNutType.SelectedItem.ToString())
        //            {
        //                if (Convert.ToDouble(nut.size.TrimStart('M')) < 2 * radiusMM)
        //                {
        //                    if (!cboFastenersNutSize.Items.Cast<string>().Contains(nut.size))
        //                    {
        //                        cboFastenersNutSize.Items.Add(nut.size);
        //                    }
        //                }
        //            }
        //        }
        //        cboFastenersNutSize.DisplayMember = "size";
        //        cboFastenersNutSize.SelectedIndex = cboFastenersNutSize.Items.Count - 1;
        //    }
        //    catch
        //    {
        //        MessageBox.Show("Unable to find/read the file Filter: " + _addinPathProfiles + "\\Fasteners\\" + "Nut.csv", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }
        //}


    }
}
