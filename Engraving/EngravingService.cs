using AESCConstruct2026.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Unsupported;  // for ScriptEnvironment
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using Application = SpaceClaim.Api.V242.Application;
using Point = SpaceClaim.Api.V242.Geometry.Point;

namespace AESCConstruct2026.UIMain
{

    public static class EngravingService
    {
        private static int version = 0;
        private static void RunPythonScriptFromString(string logicalName, string scriptBody)
        {
            var tempRoot = Path.Combine(Path.GetTempPath(), "AESCConstruct2026", "Scripts");
            Directory.CreateDirectory(tempRoot);

            var scriptPath = Path.Combine(tempRoot, logicalName);

            // Write Python to disk exactly as given
            File.WriteAllText(scriptPath, scriptBody, Encoding.UTF8);

            // Simplest overload: just run the file
            bool ok = Application.RunScript(scriptPath);

            Logger.Log($"RunScript('{logicalName}') returned {ok}");
        }
        private static int GetHostApiVersion()
        {
            var env = ScriptEnvironment.GetOrCreate(false);
            ScriptEnvironment.ActiveEnvironment = env;
            // Logger.Log($"API = {env.ApiVersion}");
            if(env.ApiVersion > version)
            {
                version = env.ApiVersion;
            }
            return version;
        }


        /// <summary>
        /// Mirrors the old btnMarkeer_ImprintToMarking_Click Python logic.
        /// </summary>
        public static void ImprintToEngravingAndExport()
        {
            Logger.Log($"Detected Host API Version: {GetHostApiVersion()}");
            var api = GetHostApiVersion();
            var script = @$"
# Python Script, API Version = V{api}
from SpaceClaim.Api.V{api} import Layer
part = GetRootPart()
window = GetActiveWindow()
doc = part.Document


def length2points(point1, point2):
    return math.sqrt((point1.X - point2.X)**2 + (point1.Y - point2.Y)**2 + (point1.Z - point2.Z)**2)

############# Check where to save the DXF
dialog = SaveFileDialog()
dialog.Filter = ""DXF|*.dxf""
dialog.Title = ""Please select a file to save the DXF.""
result = dialog.Show()
if result == True:

    ########### Get marking Edges (adjacent faces with the same normal)
    MarkEdges = []
    for b in GetRootPart().Bodies:
        for e in b.Edges:
            if e.Faces.Count == 1:
                MarkEdges.append(e)
            if e.Faces.Count > 1:
                _faceNormal = e.Faces[0].GetFaceNormal(0.5, 0.5)
                error = 0
                for f in e.Faces:
                    if not (_faceNormal == f.GetFaceNormal(0.5, 0.5) and _faceNormal == f.GetFaceNormal(0.25, 0.25)):
                        error = 1
                if error == 0:
                    MarkEdges.append(e)

    ################## CREATE NEW LAYER MARKINGS ##############
    try:
        _newLayer = Layer.Create(doc, ""Engravings"", Color.Brown)
    except:
        a = 1

    lay = doc.GetLayer(""Engravings"")
    window.ActiveLayer = lay

    ################## Copy to layer ""Engravings"" #################
    nrCurves = part.Curves.Count
    for e in MarkEdges:
        result = Copy.ToClipboard(Selection.Create(e))
        result = Paste.FromClipboard()

    window.ActiveLayer = doc.DefaultLayer

    ################### EXPORT TO TEMP ####################
    fileName = r""C:\Windows\Temp\_TempDXF.dxf""

    # Save File
    options = ExportOptions.Create()
    DocumentSave.Execute(fileName, options)

    ################### Remove Lines from Engraving Layer ####################
    for i in range(nrCurves, part.Curves.Count):
        part.Curves[nrCurves].Delete()

    ##################### IMPORT DXF #####################
    importOptions = ImportOptions.Create()
    DocumentOpen.Execute(fileName, importOptions)
    part = GetRootPart()
    window = GetActiveWindow()
    doc = part.Document

    curves = part.DatumPlanes[0].Curves

    ## Check which curve is attached to layer ""Markings""
    lay = doc.GetLayer(""Engravings"")
    curvesMarkings = []
    otherCurves = []
    for c in curves:
        if c.Layer == lay:
            curvesMarkings.append(c)
        else:
            otherCurves.append(c)

    deleteList = []
    for cm in curvesMarkings:
        for oc in otherCurves:
            if cm.Shape.StartPoint == oc.Shape.StartPoint and cm.Shape.EndPoint == oc.Shape.EndPoint:
                deleteList.append(oc)

    for d in deleteList:
        d.Delete()

    ################## Save File #####################3
    fileName = dialog.FileName
    options = ExportOptions.Create()
    DocumentSave.Execute(fileName, options)

    Window.ActiveWindow.Close()
";

            //if (GetHostApiVersion() < 252)
            //{
                var env = ScriptEnvironment.GetOrCreate(false);
                ScriptEnvironment.ActiveEnvironment = env;
                env.ApiVersion = api;
                var success = env.RunCommand(script);
            //}
            //else
            //{
            //    RunPythonScriptFromString("ImprintToEngravingAndExport.py", script);
            //}
        }


        /// <summary>
        /// Mirrors the old btnMarkeer_ImprintBody_Click Python logic.
        /// </summary>
        public static void ImprintBody()
        {
            Logger.Log($"Detected Host API Version: {GetHostApiVersion()}");
            var api = GetHostApiVersion();
            var script = @$"
# Python Script, API Version = V{api}
from System.Collections.Generic import List
from SpaceClaim.Api.V{api} import Application, StatusMessageType
from SpaceClaim.Api.V{api}.Geometry import Point

win = Window.ActiveWindow
ctx = win.ActiveContext
sel = ctx.Selection

if sel.Count < 2:
    MessageBox.Show(""Select at least two solid bodies (source first, then one or more targets)."", ""Warning"", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    raise SystemExit(0)

firstSel = sel[0]

def is_body(o):
    return isinstance(o, IDesignBody) and o.Master is not None

# body-to-body imprint: imprint edges of each touching face of body2 onto body1
if is_body(firstSel):
    # validate selection
    for s in sel:
        if not is_body(s):
            MessageBox.Show(""Select solid bodies only."", ""Warning"", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            raise SystemExit(0)
        if s.Root == s.Parent:
            MessageBox.Show(""Body not in a component."", ""Warning"", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            raise SystemExit(0)

    body1 = sel[0]
    totalPairs = 0

    # tolerance on opposite normals (cosine close to 1)
    angDotTol = 0.9999

    for i in range(1, sel.Count):
        body2 = sel[i]

        pairs = []
        for f1 in body1.Faces:
            n1 = f1.GetFaceNormal(0.5, 0.5)
            for f2 in body2.Faces:
                n2 = f2.GetFaceNormal(0.5, 0.5)
                dot = n1.X * (-n2.X) + n1.Y * (-n2.Y) + n1.Z * (-n2.Z)
                if dot < angDotTol:
                    continue
                # midpoint of f2 must lie on f1 (coincident faces)
                mp2 = f2.Master.EvalMid().Point
                if not f1.Shape.ContainsPoint(mp2):
                    continue
                pairs.append((f1, f2))

        for (fp1, fp2) in pairs:
            selSet = Selection.Empty()
            for e in fp2.Edges:
                selSet = selSet + Selection.Create(e)
            selSet = selSet + Selection.Create(fp1)

            pts = List[Point]()  # optional disambiguation points (unused)
            FixImprint.FixSpecific(selSet, pts)
            totalPairs += 1

    if totalPairs == 0:
        MessageBox.Show(""No touching faces were found to imprint. Ensure bodies actually touch (coincident faces)."", ""Information"", MessageBoxButtons.OK, MessageBoxIcon.Information)
    else:
        Application.ReportStatus(""Imprint finished"", StatusMessageType.Information, None)

# curve-to-face imprint: curves first, face last
elif isinstance(firstSel, IDesignCurve) and isinstance(sel[-1], IDesignFace):
    curves = [c for c in sel[0:sel.Count-1] if isinstance(c, IDesignCurve)]
    face   = sel[-1]
    selSet = Selection.Empty()
    for c in curves:
        selSet = selSet + Selection.Create(c)
    selSet = selSet + Selection.Create(face)
    pts = List[Point]()
    FixImprint.FixSpecific(selSet, pts)
    Application.ReportStatus(""Imprint finished"", StatusMessageType.Information, None)

else:
    MessageBox.Show(""Unsupported selection."", ""Warning"", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    raise SystemExit(0)
";
            //if (GetHostApiVersion() < 252)
            //{
                var env = ScriptEnvironment.GetOrCreate(false);
                ScriptEnvironment.ActiveEnvironment = env;
                env.ApiVersion = api;
                var success = env.RunCommand(script);
            //}
            //else
            //{
            //    RunPythonScriptFromString("ImprintBody.py", script);
            //}
        }
        public static void AddNote(
                        double parLetterSize,
                        bool parEngraving,
                        bool parCustomText,
                        string parText,
                        bool parCutOut,
                        string parFont,
                        bool parBodyName,
                        bool parLocationCenter)
        {

            Logger.Log($"Detected Host API Version: {GetHostApiVersion()}");
            //public static void AddNote(
            //            double letterSize,
            //            bool engraving,
            //            bool customText,
            //            string text,
            //            bool cutout,
            //            string fontName,
            //            bool useBodyName,
            //            bool centerLocation)
            //        {

            // Parameters with defaults matching the Python fallback
            //double parLetterSize = 3.0;
            //bool parEngraving = false;
            //bool parCustomText = true;
            //string parText = "Test11";
            //bool parCutOut = true;
            //string parFont = "Arial Black";
            //bool parBodyName = false;
            //bool parLocationCenter = true;
            bool parLocationBottomRight = false;


            var win = Window.ActiveWindow;
            if (win == null)
            {
                Application.ReportStatus("No active window found.", StatusMessageType.Information, null);
                return;
            }

            var context = win.ActiveContext;
            if (context == null)
            {
                Application.ReportStatus("No active context found.", StatusMessageType.Information, null);
                return;
            }

            var selection = context.Selection;
            if (selection == null || selection.Count == 0)
            {
                Application.ReportStatus("Please select one or more faces.", StatusMessageType.Information, null);
                return;
            }

            Part part = win.Document.MainPart;
            if (part == null)
            {
                Application.ReportStatus("No root part found.", StatusMessageType.Information, null);
                return;
            }

            var window = Window.ActiveWindow;
            var doc = part.Document;
            var layActive = window.ActiveLayer;

            // Collect selected faces and the pick points used in the Python script
            List<IDesignFace> selectedFaces = new List<IDesignFace>();
            foreach (var sel in selection)
            {
                IDesignFace face = sel as IDesignFace;
                if (face != null)
                    selectedFaces.Add(face);
            }

            if (selectedFaces.Count == 0)
            {
                Application.ReportStatus("Please select one or more faces.", StatusMessageType.Information, null);
                return;
            }

            foreach (var desFace in selectedFaces)
            {
                if (desFace == null)
                {
                    Application.ReportStatus("Please select a single face.", StatusMessageType.Information, null);
                    continue;
                }

                Plane plane = desFace.Shape.Geometry as Plane;
                if (plane == null)
                {
                    Application.ReportStatus("Selected face is not planar.", StatusMessageType.Information, null);
                    continue;
                }

                ////Point selPoint1 = (Point)win.ActiveContext.GetSelectionPoint(selection.First());
                // Retrieve the pick point for this selection item
                Point selPoint1;
                try
                {
                    selPoint1 = (Point)context.GetSelectionPoint(desFace);
                }
                catch
                {
                    // Fallback to face mid if no explicit pick point is available
                    selPoint1 = desFace.Master.Shape.GetBoundingBox(Matrix.Identity, true).Center;
                }

                Part parentPart = desFace.Parent.Parent.Master;

                // Ensure point is on plane
                Point selPoint = plane.ProjectPoint(selPoint1).Point;


                Direction dirX = plane.Frame.DirX;
                Direction dirY = plane.Frame.DirY;
                Direction dirZ = desFace.Shape.IsReversed ? -desFace.Shape.ProjectPoint(selPoint).Normal : desFace.Shape.ProjectPoint(selPoint).Normal;


                WriteBlock.ExecuteTask("Engraving", () =>
                {
                    // Choose note text per logic
                    string noteText;
                    if (parCustomText)
                        noteText = parText;
                    else if (parBodyName)
                        noteText = desFace.Parent.Master.Name;
                    else
                        noteText = parentPart.DisplayName;


                    // Optional repositioning to the face center or bottom right
                    DatumPoint dp = null;
                    if (parLocationCenter)
                    {
                        selPoint1 = desFace.Master.Shape.GetBoundingBox(Matrix.Identity, true).Center;
                        selPoint = desFace.Shape.ProjectPoint(selPoint1).Point;
                        // dp = DatumPoint.Create(part, "dp", selPoint);
                    }
                    else if (parLocationBottomRight)
                    {
                        //dp = DatumPoint.Create(part, "dp", selPoint1);
                    }


                    DatumPlane datumPlane = DatumPlane.Create(parentPart, "dpAtPoint", Plane.Create(Frame.Create(selPoint, dirX, dirY)));

                    Direction _dirZ = datumPlane.Shape.Geometry.Frame.DirZ;
                    if (_dirZ != dirZ)
                    {
                        datumPlane.Delete();
                        datumPlane = DatumPlane.Create(parentPart, "dpAtPoint", Plane.Create(Frame.Create(selPoint, dirY, dirX)));
                    }

                    // Create the final annotation plane at the pick point with the face normal
                    if (datumPlane == null)
                    {
                        Application.ReportStatus("Could not create the target datum plane.", StatusMessageType.Information, null);
                        return;
                    }

                    // Create the note on the datum plane
                    Note note = Note.Create(datumPlane, PointUV.Origin, LocationPoint.Center, 0.001 * parLetterSize, noteText);
                    note.SetFontName(parFont);

                    //// Debug
                    //dp = DatumPoint.Create(part, "selPoint", selPoint);


                    // Engraving path, ensure layer exists and apply note to sheet metal
                    if (parEngraving)
                    {
                        var lay = doc.GetLayer("Engravings");
                        if (lay == null)
                        {
                            lay = Layer.Create(doc, "Engravings", Color.Brown);
                        }

                        //window.ActiveLayer = lay;
                        note.Layer = lay;

                        try
                        {
                            var sm = parentPart.SheetMetal;
                            sm.TryApplyNote(note, Lettering.Engraved, out DesignFace substrate);
                        }
                        catch
                        {
                            Application.ReportStatus(
                                "Engraving failed. Ensure that the selected face belongs to a sheet metal solid body.",
                                StatusMessageType.Information, null);
                        }
                    }

                    // Cutout path
                    if (parCutOut)
                    {
                        try
                        {
                            var sm = parentPart.SheetMetal;
                            sm.TryApplyNote(note, Lettering.Cutout, out DesignFace substrate);
                        }
                        catch
                        {
                            Application.ReportStatus(
                                "Cutout failed. Ensure that the selected face belongs to a sheet metal solid body.",
                                StatusMessageType.Information, null);
                        }
                    }

                    // cleanup
                    if (!datumPlane.IsDeleted)
                        datumPlane.Delete();
                });
            }
        }
    }
}

