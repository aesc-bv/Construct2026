using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Application = SpaceClaim.Api.V242.Application;
using Body = SpaceClaim.Api.V242.Modeler.Body;
using Document = SpaceClaim.Api.V242.Document;
using Frame = SpaceClaim.Api.V242.Geometry.Frame;
using Settings = AESCConstruct25.Properties.Settings;

namespace AESCConstruct25.Fastener.Module
{
    internal class FastenerModule
    {
        // backing lists
        private readonly List<Bolt> _bolts;
        private readonly List<Washer> _washers;
        private readonly List<Nut> _nuts;

        public FastenerModule()
        {
            var basePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
               "AESCConstruct", "Fasteners");
            //string csvPath = Settings.Default.profiles;
            //_bolts = File.ReadAllLines(Path.Combine(basePath, "Bolt.csv"))
            //             .Skip(1).Select(Bolt.FromCsv).ToList();

            //_washers = File.ReadAllLines(Path.Combine(basePath, "Washer.csv"))
            //                .Skip(1).Select(Washer.FromCsv).ToList();

            //_nuts = File.ReadAllLines(Path.Combine(basePath, "Nut.csv"))
            //             .Skip(1).Select(Nut.FromCsv).ToList();
            _bolts = File.ReadAllLines(Settings.Default.Bolt)
                         .Skip(1).Select(Bolt.FromCsv).ToList();

            _washers = File.ReadAllLines(Settings.Default.Washer)
                        .Skip(1).Select(Washer.FromCsv).ToList();

            _nuts = File.ReadAllLines(Settings.Default.Nut)
                         .Skip(1).Select(Nut.FromCsv).ToList();
        }

        // top-level lists for your “Type” ComboBoxes
        //public IEnumerable<string> BoltTypes => _bolts.Select(b => b.type).Distinct();
        //public IEnumerable<string> WasherTypes => _washers.Select(w => w.type).Distinct();
        //public IEnumerable<string> NutTypes => _nuts.Select(n => n.type).Distinct();
        public IEnumerable<string> BoltNames => _bolts.Select(b => b.Name).Distinct();
        public IEnumerable<string> WasherNames => _washers.Select(w => w.Name).Distinct();
        public IEnumerable<string> NutNames => _nuts.Select(n => n.Name).Distinct();

        public string GetBoltTypeByName(string name) =>
        _bolts.FirstOrDefault(b => b.Name == name)?.type
        ?? throw new InvalidOperationException($"Unknown bolt name '{name}'");

        public string GetWasherTopTypeByName(string name) =>
        _washers.FirstOrDefault(w => w.Name == name)?.type
        ?? throw new InvalidOperationException($"Unknown washer name '{name}'");

        public string GetNutTypeByName(string name) =>
        _nuts.FirstOrDefault(n => n.Name == name)?.type
        ?? throw new InvalidOperationException($"Unknown nut name '{name}'");

        // called when the user picks a type, to fill the “Size” list
        //public IEnumerable<string> BoltSizesFor(string type) =>
        //    _bolts.Where(b => b.type == type).Select(b => b.size).Distinct();

        //public IEnumerable<string> WasherSizesFor(string type) =>
        //    _washers.Where(w => w.type == type).Select(w => w.size).Distinct();

        //public IEnumerable<string> NutSizesFor(string type) =>
        //    _nuts.Where(n => n.type == type).Select(n => n.size).Distinct();

        public IEnumerable<string> BoltSizesFor(string selectedName)
        {
            // find all underlying types for that name
            var types = _bolts.Where(b => b.Name == selectedName)
                              .Select(b => b.type)
                              .Distinct();
            // aggregate sizes across all matching types
            return _bolts.Where(b => types.Contains(b.type))
                         .Select(b => b.size)
                         .Distinct();
        }
        public IEnumerable<string> WasherSizesFor(string selectedName)
        {
            var types = _washers.Where(w => w.Name == selectedName).Select(w => w.type).Distinct();
            return _washers.Where(w => types.Contains(w.type)).Select(w => w.size).Distinct();
        }
        public IEnumerable<string> NutSizesFor(string selectedName)
        {
            var types = _nuts.Where(n => n.Name == selectedName).Select(n => n.type).Distinct();
            return _nuts.Where(n => types.Contains(n.type)).Select(n => n.size).Distinct();
        }

        double parBoltL = 20;
        double parDistance = 20;
        static string boltType, boltSize, washerTopType, washerTopSize, washerBottomType, washerBottomSize, nutType, nutSize, customFile;
        private bool includeWasherTop, includeWasherBottom, includeNut, useCustom;

        static string _addinPathProfiles = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\AESCConstruct";
        static string _addinPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\SpaceClaim\\Addins\\AESCConstruct";

        List<Bolt> listBolt;
        List<Washer> listWasherTop;
        List<Washer> listWasherBottom;
        List<Nut> listNut;



        public void SetBoltType(string t) => boltType = t;
        public void SetBoltSize(string s) => boltSize = s;
        public void SetBoltLength(string len) => parBoltL = double.Parse(len);
        public void SetIncludeWasherTop(bool inc) => includeWasherTop = inc;
        public void SetWasherTopType(string t) => washerTopType = t;
        public void SetWasherTopSize(string s) => washerTopSize = s;
        public void SetIncludeWasherBottom(bool inc) => includeWasherBottom = inc;
        public void SetWasherBottomType(string t) => washerBottomType = t;
        public void SetWasherBottomSize(string s) => washerBottomSize = s;
        public void SetIncludeNut(bool inc) => includeNut = inc;
        public void SetNutType(string t) => nutType = t;
        public void SetNutSize(string s) => nutSize = s;
        public void SetCustomFile(string f) => customFile = f;
        public void SetUseCustomPart(bool u) => useCustom = u;
        public void SetDistance(int d) => parDistance = d;

        private bool _lockDistance;
        public void SetLockDistance(bool l) => _lockDistance = l;

        private bool _overwriteDistance;
        public void SetOverwriteDistance(bool od) => _overwriteDistance = od;


        private Part GetWasherPart(Document doc, string name, Washer washer, out Component component)
        {
            // Logger.Log($"[GetWasherPart] START name={name}, d1={washer.d1}, d2={washer.d2}, s={washer.s}");

            // Always create new part
            Part fastenersPart = GetFastenersPart(doc);
            string uniqueName = $"{name}_{Guid.NewGuid():N}";
            // Logger.Log($"[GetWasherPart] Creating new Part '{uniqueName}' under 'Fasteners'");
            Part part = Part.Create(doc, uniqueName);
            component = Component.Create(fastenersPart, part);

            // Create geometry
            Body body;
            try
            {
                body = createFasteners.createWasher(washer.d1 * 0.001, washer.d2 * 0.001, washer.s * 0.001);
            }
            catch (Exception)
            {
                // Logger.Log($"[GetWasherPart] ERROR in createWasher: {ex}");
                throw;
            }
            // Logger.Log($"[GetWasherPart] Body created; calling DesignBody.Create");

            DesignBody db = DesignBody.Create(part, name, body);
            db.IsLocked = _lockDistance;
            // Logger.Log($"[GetWasherPart] DesignBody created, attaching property");
            CustomPartProperty.Create(part, "AESC_Construct", true);

            // Logger.Log($"[GetWasherPart] COMPLETE for '{uniqueName}'");
            return part;
        }

        public static List<PlacementData> GetPlacementDataSelection(ICollection<IDocObject> selection, ICollection<IDocObject> secondSelection)
        {
            // 1) Make sure we actually have something selected
            if (selection == null || selection.Count == 0)
            {
                Application.ReportStatus("No objects selected.", StatusMessageType.Information, null);
                return null;
            }

            if (selection.Count > 1 && secondSelection.Count > 0)
            {
                Application.ReportStatus("Please select a single edge when selecting faces.", StatusMessageType.Information, null);
                return null;
            }

            List<IDesignFace> planarFaces = new List<IDesignFace>();
            if (secondSelection.Count > 0)
            {
                foreach (var dobj in secondSelection)
                {
                    var selface = dobj as IDesignFace;
                    if (selface == null)
                    {
                        Application.ReportStatus("Second selection must be planar face(s).", StatusMessageType.Information, null);
                        return null;
                    }
                    var planarFace = selface.Shape.GetGeometry<Plane>();
                    if (planarFace == null)
                    {
                        Application.ReportStatus("Second selection must be planar face(s).", StatusMessageType.Information, null);
                        return null;
                    }
                    planarFaces.Add(selface);
                }
            }


            var placements = new List<PlacementData>();

            foreach (var dobj in selection)
            {
                // expect the selection to be a DesignEdge whose geometry is a Circle
                var edge = dobj as IDesignEdge;
                if (edge == null)
                {
                    Application.ReportStatus("Selection contains non‐edge objects.", StatusMessageType.Information, null);
                    return null;
                }

                // try to get a Circle from the edge
                var circle = edge.Shape.GetGeometry<Circle>();
                if (circle == null)
                {
                    Application.ReportStatus("All selected edges must be circles.", StatusMessageType.Information, null);
                    return null;
                }

                if (planarFaces.Count > 0)
                {
                    double radius = circle.Radius;
                    placements.AddRange(CreatePlacementDataForCirclesInPlanes(planarFaces, radius));
                }
                else
                {
                    placements.Add(CreatePlacementData(circle, edge));
                }

            }

            return placements;
        }

        private static List<PlacementData> CreatePlacementDataForCirclesInPlanes(List<IDesignFace> planarFaces, double radius)
        {
            List<PlacementData> returnList = new List<PlacementData>();
            foreach (IDesignFace idf in planarFaces)
            {
                foreach (IDesignEdge ide in idf.Edges)
                {
                    var circle = ide.Shape.GetGeometry<Circle>();
                    if (circle == null)
                    {
                        continue;
                    }
                    if (Math.Abs(radius - circle.Radius) < 1e-9)
                        returnList.Add(CreatePlacementData(circle, ide));
                }
            }
            return returnList;
        }

        private static PlacementData CreatePlacementData(Circle circle, IDesignEdge edge)
        {
            // 2) origin is simply the circle's axis origin
            var origin = circle.Axis.Origin;

            // Check if there is no other designface in the center of the circle
            Part mainPart = edge.Root as Part;
            foreach (IDesignBody idb in mainPart.GetDescendants<IDesignBody>())
                if (idb.Shape.ContainsPoint(origin))
                {
                    Application.ReportStatus("There is an object in the hole, no fastener created.", StatusMessageType.Information, null);
                    return null;
                }

            // 3) direction: look for an adjacent cylindrical face
            Direction direction = circle.Axis.Direction; // fallback
            double depth = 0;
            foreach (IDesignFace face in edge.Faces)
            {
                var cyl = face.Master.Shape.GetGeometry<Cylinder>();

                if (cyl != null)
                {
                    direction = face.TransformToMaster.Inverse * cyl.Axis.Direction;
                    Point test1 = origin + 0.01 * direction;
                    Point test2 = origin - 0.01 * direction;
                    Point projectPoint1 = face.Shape.ProjectPoint(test1).Point;
                    Point projectPoint2 = face.Shape.ProjectPoint(test2).Point;
                    double dist1 = (projectPoint1 - test1).Magnitude;
                    double dist2 = (projectPoint2 - test2).Magnitude;
                    //WriteBlock.ExecuteTask("UpdateComponentNames", () =>
                    //{
                    //    DatumPoint.Create(face.Parent.Parent, "origin", origin);
                    //    DatumPoint.Create(face.Parent.Parent, "test1", test1);
                    //    DatumPoint.Create(face.Parent.Parent, "test2", test2);
                    //    DatumPoint.Create(face.Parent.Parent, "projectPoint1", projectPoint1);
                    //    DatumPoint.Create(face.Parent.Parent, "projectPoint2", projectPoint2);
                    //});
                    direction = dist1 > dist2 ? direction : -direction;

                    var cyl2 = face.Shape.Geometry as Cylinder;
                    Matrix mat = Matrix.CreateMapping(SpaceClaim.Api.V242.Geometry.Frame.Create(cyl2.Axis.Origin, cyl2.Axis.Direction)).Inverse;
                    Box boundingBox = face.Shape.GetBoundingBox(mat, true);
                    depth = boundingBox.Size.Z;

                    break;

                }
            }

            return new PlacementData
            {
                Circle = circle,
                Origin = origin,
                Direction = direction,
                Depth = depth
            };
        }

        public void CreateFasteners()
        {
            // Check selection
            Window win = Window.ActiveWindow;
            Document doc = win.Document;
            Part mainPart = doc.MainPart;

            if (!CheckSelectedCircle(win))
                return;

            ICollection<IDocObject> selection = win.ActiveContext.Selection;
            ICollection<IDocObject> secSelection = win.ActiveContext.SecondarySelection;

            listBolt = _bolts.ToList();
            listWasherTop = _washers.ToList();
            listWasherBottom = _washers.ToList();
            listNut = _nuts.ToList();

            if (selection.Count > 1 && secSelection.Count > 0)
            {
                Application.ReportStatus("Please select only one circle when selecting faces", StatusMessageType.Information, null);
                return;
            }

            // To do, get list of center points and directions from the selections
            List<PlacementData> placementDatas = GetPlacementDataSelection(selection, secSelection);
            if (placementDatas == null)
            {
                Application.ReportStatus("Could not find correct placement(s)", StatusMessageType.Information, null);
                return;
            }

            try
            {   //Execute a Command 
                WriteBlock.ExecuteTask("AESCConstruct.FastenersCreate",
                delegate
                {
                    // Logger.Log($"[FastenerModule] CreateFasteners() starting; includeTop={includeWasherTop}, includeBottom={includeWasherBottom}, includeNut={includeNut}");
                    // Logger.Log($"[FastenerModule] boltType={boltType}, boltSize={boltSize}, washerTopType={washerTopType}, washerTopSize={washerTopSize}, washerBottomType={washerBottomType}, washerBottomSize={washerBottomSize}");

                    Bolt _bolt = listBolt.First();

                    foreach (Bolt bolt in listBolt)
                    {
                        if (bolt.type == boltType && bolt.size == boltSize)
                        {
                            _bolt = bolt;
                        }
                    }

                    foreach (Bolt bolt in listBolt)
                    {
                        if (bolt.type == boltType && bolt.size == boltSize)
                        {
                            _bolt = bolt;

                        }
                    }

                    Washer _washerBottom = listWasherBottom.First();
                    Washer _washerTop = listWasherTop.First();

                    foreach (Washer washer in listWasherBottom)
                    {
                        if (washer.type == washerBottomType && washer.size == washerBottomSize)
                        {
                            _washerBottom = washer;
                        }
                    }
                    foreach (Washer washer in listWasherTop)
                    {
                        if (washer.type == washerTopType && washer.size == washerTopSize)
                        {
                            _washerTop = washer;
                        }
                    }
                    Nut _nut = listNut.First();
                    foreach (Nut nut in listNut)
                    {
                        if (nut.type == nutType && nut.size == nutSize)
                        {
                            _nut = nut;
                        }
                    }

                    string customPartPath = "none";
                    if (File.Exists(_addinPathProfiles + "\\Fasteners\\Custom\\" + customFile))
                        customPartPath = _addinPathProfiles + "\\Fasteners\\Custom\\" + customFile;

                    // TODO:
                    // - handle custom part
                    // - Lock parts

                    ///// Create Geometry     
                    string boltName = boltType.Split(' ')[0] + " " + boltSize + " x " + (parBoltL).ToString();
                    string washerTopName = washerTopType.Split(' ')[0] + " " + washerTopSize;
                    string washerBottomName = washerBottomType.Split(' ')[0] + " " + washerBottomSize;
                    string nutName = nutType.Split(' ')[0] + " " + nutSize;

                    foreach (PlacementData PD in placementDatas)
                    {
                        if (PD == null)
                            continue;

                        Matrix matrixMapping = Matrix.CreateMapping(SpaceClaim.Api.V242.Geometry.Frame.Create(PD.Origin, PD.Direction));
                        Matrix matrixBolt = matrixMapping;
                        if (_overwriteDistance)
                        {
                            // parDistance is in mm; convert to m (×0.001)
                            matrixBolt = matrixMapping
                                * Matrix.CreateTranslation(Vector.Create(0, 0, parDistance * 0.001));
                        }

                        if (this.includeWasherTop)
                        {
                            // Logger.Log($"[CreateFasteners] Inserting TOP washer 'name' at {PD.Origin} dir={PD.Direction}");
                            Part washerPart = GetWasherPart(doc, washerTopName, _washerTop, out Component componentWasherTop);
                            // Logger.Log($"[CreateFasteners] Transforming TOP washer component...");
                            matrixBolt = matrixMapping * Matrix.CreateTranslation(Vector.Create(0, 0, _washerTop.s * 0.001));
                            if (_overwriteDistance)
                            {
                                // parDistance is in mm; convert to m (×0.001)
                                matrixBolt = matrixMapping
                                    * Matrix.CreateTranslation(Vector.Create(0, 0, (parDistance + _washerTop.s) * 0.001));
                            }
                            // Logger.Log($"[CreateFasteners] Transform COMPLETE for TOP washer");
                            componentWasherTop.Transform(matrixMapping);
                        }
                        if (useCustom && !string.IsNullOrEmpty(customFile))
                        {
                            // Logger.Log($"[CreateFasteners] Inserting CUSTOM part: {customFile}");

                            // 1) Build the full path
                            string customDir = Path.Combine(
                                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                                "AESCConstruct", "Fasteners", "Custom"
                            );
                            string path = Path.Combine(customDir, customFile);
                            if (!File.Exists(path))
                            {
                                Application.ReportStatus($"Custom file not found: {path}",
                                                         StatusMessageType.Warning, null);
                                continue;
                            }

                            // 2) Branch on extension
                            string ext = Path.GetExtension(path).ToLowerInvariant();
                            Document customDoc;
                            if (ext == ".scdoc" || ext == ".scdocx")
                            {
                                // Native SpaceClaim document
                                // Logger.Log($"[CreateFasteners] Loading SpaceClaim doc via Document.Load: {path}");
                                customDoc = Document.Load(path);
                            }
                            else if (ext == ".stp" || ext == ".step")
                            {
                                // STEP import
                                // Logger.Log($"[CreateFasteners] Importing STEP via Document.Open + ImportOptions: {path}");
                                var opts = ImportOptions.Create();
                                customDoc = Document.Open(path, opts);
                            }
                            else
                            {
                                Application.ReportStatus($"Unsupported format: {ext}",
                                                         StatusMessageType.Warning, null);
                                continue;
                            }

                            // 3) Bring bodies into your Fasteners part
                            Part fastenersPart = GetFastenersPart(doc);
                            foreach (var db in customDoc.MainPart.GetDescendants<DesignBody>())
                            {
                                Component comp = Component.Create(fastenersPart, db.Parent as Part);
                                comp.Transform(matrixBolt);
                            }

                            continue;
                        }

                        if (!useCustom)
                        {
                            Part boltPart = GetBoltPart(doc, boltName, boltType, _bolt, parBoltL, out Component componentBolt);
                            componentBolt.Transform(matrixBolt);
                        }

                        double displacementZ = -PD.Depth;

                        if (this.includeWasherBottom)
                        {
                            // Logger.Log($"[CreateFasteners] Inserting BOTTOM washer 'name'");
                            Part washerPart = GetWasherPart(doc, washerBottomName, _washerBottom, out Component componentWasherBottom);
                            // Logger.Log($"[CreateFasteners] Transforming BOTTOM washer component...");
                            displacementZ += -0.001 * _washerBottom.s;
                            componentWasherBottom.Transform(matrixMapping * Matrix.CreateTranslation(Vector.Create(0, 0, displacementZ)));
                            // Logger.Log($"[CreateFasteners] Transform COMPLETE for BOTTOM washer");
                        }

                        if (this.includeNut)
                        {
                            Part nutPart = GetNutPart(doc, nutName, nutType, _nut, out Component componentNut);
                            displacementZ += -0.001 * _nut.h;
                            componentNut.Transform(matrixMapping * Matrix.CreateTranslation(Vector.Create(0, 0, displacementZ)));
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public static Part GetFastenersPart(Document doc)
        {
            // Try to find an existing “Fasteners” part
            var fastenersPart = doc.Parts
                .FirstOrDefault(p => p.DisplayName == "Fasteners");

            if (fastenersPart is null)
            {
                // if not found, create it and hook it up
                fastenersPart = Part.Create(doc, "Fasteners");
                Component.Create(doc.MainPart, fastenersPart);

                CustomPartProperty.Create(fastenersPart, "AESC_Construct", true);

            }
            return fastenersPart;
        }

        private Part GetBoltPart(Document doc, string name, string boltType, Bolt _bolt, double parBoltL, out Component component)
        {
            // Try to find an existing “Fasteners” part
            Part boltPart = doc.Parts
                .FirstOrDefault(p => p.DisplayName == name);


            // Logger.Log("getboltpart");
            // Create Fasteners Part if not existing
            Part fastenersPart = GetFastenersPart(doc);
            if (boltPart is null)
            {

                // if not found, create it and hook it up
                boltPart = Part.Create(doc, name);
                component = Component.Create(fastenersPart, boltPart);

                // Create body
                Body bodyBolt = createFasteners.Create_Bolt(boltType, _bolt, parBoltL);
                DesignBody dbBolt = DesignBody.Create(boltPart, name, bodyBolt);
                dbBolt.IsLocked = _lockDistance;
                CustomPartProperty.Create(boltPart, "AESC_Construct", true);
            }
            else
            {
                // Create a copy of the original component
                component = Component.Create(fastenersPart, boltPart);
            }
            return boltPart;
        }

        private Part GetNutPart(Document doc, string name, string type, Nut nut, out Component component)
        {
            // Try to find an existing “Fasteners” part
            Part part = doc.Parts
                .FirstOrDefault(p => p.DisplayName == name);

            // Create Fasteners Part if not existing
            Part fastenersPart = GetFastenersPart(doc);
            if (part is null)
            {

                // if not found, create it and hook it up
                part = Part.Create(doc, name);
                component = Component.Create(fastenersPart, part);

                // Create body
                Body body = createFasteners.Create_Nut(boltType, nut);
                DesignBody designBody = DesignBody.Create(part, name, body);
                designBody.IsLocked = _lockDistance;
                CustomPartProperty.Create(part, "AESC_Construct", true);
            }
            else
            {
                // Create a copy of the original component
                component = Component.Create(fastenersPart, part);
            }
            return part;
        }

        public static bool CheckSelectedCircle(Window window, bool singleSelection = false)
        {
            var selection = window.ActiveContext.Selection;

            // no selection or too many
            if (selection == null || selection.Count == 0 ||
                (singleSelection && selection.Count != 1))
            {
                Application.ReportStatus("Please select a circle.", StatusMessageType.Information, null);
                return false;
            }

            // helper to grab a Circle from either a curve or an edge
            Circle GetCircle(IDocObject obj)
            {
                var curve = obj as IDesignCurve;
                if (curve != null)
                    return curve.Shape.GetGeometry<Circle>();

                var edge = obj.Master as DesignEdge;
                if (edge != null)
                    return edge.Shape.GetGeometry<Circle>();

                return null;
            }

            // ensure every selected item is actually a circle
            foreach (var obj in selection)
            {
                if (GetCircle(obj) == null)
                {
                    Application.ReportStatus("Please select a circle.", StatusMessageType.Information, null);
                    return false;
                }
            }

            return true;
        }


        public static double GetSizeCircle(Window window, out double depthMM)
        {
            double radiusMM = 0;

            depthMM = 0;

            InteractionContext interContext = window.ActiveContext;
            var selection = interContext.SingleSelection;

            if (selection != null)
            {
                var desCurve = selection as IDesignCurve;
                var desEdge = selection.Master as DesignEdge;
                if (desCurve != null)
                {
                    string desCurveType = desCurve.Shape.Geometry.GetType().Name;

                    if (desCurveType == "Circle")
                    {
                        var circle = desCurve.Shape.GetGeometry<Circle>();
                        radiusMM = 1000 * circle.Radius;
                    }
                }
                else if (desEdge != null)
                {
                    string desEdgeType = desEdge.Shape.Geometry.GetType().Name;

                    if (desEdgeType == "Circle")
                    {
                        var circle = desEdge.Shape.GetGeometry<Circle>();

                        radiusMM = 1000 * circle.Radius;
                        foreach (DesignFace df in desEdge.Faces)
                        {
                            var cyl = df.Shape.Geometry as Cylinder;
                            if (cyl != null)
                            {
                                Matrix mat = Matrix.CreateMapping(Frame.Create(cyl.Axis.Origin, cyl.Axis.Direction)).Inverse;
                                Box boundingBox = df.Shape.GetBoundingBox(mat, true);
                                depthMM = boundingBox.Size.Z * 1000;
                            }

                        }
                    }
                }
                else
                {
                    Application.ReportStatus("Please select a circle.", StatusMessageType.Information, null);
                }

            }

            return radiusMM;
        }


        //public static void CheckSize()
        //{

        //    Window window = Window.ActiveWindow;
        //    Document doc = window.Document;
        //    Part rootPart = doc.MainPart;

        //    if (!CheckSelectedCircle(window, true))
        //        return;

        //    double radiusMM = GetSizeCircle(window, out double depthMM);
        //   // Logger.Log($"radiusMM - {radiusMM}");

        //    if (radiusMM == 0)
        //        return;

        //    //// Apply filter to the Sizes of all comboboxes

        //}
    }
}
