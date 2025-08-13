using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Unsupported;  // for ScriptEnvironment
using System;

namespace AESCConstruct25.UIMain
{
    public static class EngravingService
    {
        /// <summary>
        /// Mirrors the old Markeer “Add Note/Engraving/CutOut” Python logic under V242.
        /// </summary>
        public static void AddNote(
            double letterSize,
            bool engraving,
            bool customText,
            string text,
            bool cutout,
            string fontName,
            bool useBodyName,
            bool centerLocation)
        {
            try
            {   //Execute a Command 
                WriteBlock.ExecuteTask("AddNote",
                delegate
                {

                    var dataList = string.Format(
                        "dataList = [\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\"]",
                        letterSize.ToString(),
                        engraving.ToString(),
                        customText.ToString(),
                        text.Replace("\"", "\\\""),
                        cutout.ToString(),
                        fontName,
                        useBodyName.ToString(),
                        centerLocation.ToString()
                    );
                    // Logger.Log(dataList);


                    string script = dataList + "\n" + @"
# Python Script, API Version = V242
##########
from SpaceClaim.Api.V242.Geometry import PointUV
from SpaceClaim.Api.V242 import LocationPoint
from SpaceClaim.Api.V242 import Note
from SpaceClaim.Api.V242 import Layer

from SpaceClaim.Api.V242 import Lettering

try:
    parLetterSize= float(dataList[0]);
    parEngraving = dataList[1];
    parCustomText =  dataList[2];
    parText =  dataList[3];
    parCutOut =  dataList[4];
    parFont = dataList[5];
    parBodyName = dataList[6];
    parLocationCenter = dataList[7];

except:
    ########### Parameters
    parLetterSize = 3
    parEngraving = ""False""
    parCutOut = ""False""
    parCustomText = ""True""
    parText = ""Test11""
    parFont = ""Arial Black""
    parBodyName = ""False""
    parLocationCenter = ""True""
    #parLocationBottomRight = ""False""

#######################
parLocationBottomRight = ""False""
win = Window.ActiveWindow
context = win.ActiveContext;
multiSelection = context.Selection ## Check how to select two faces
error = 0;
part  = GetRootPart()
window = GetActiveWindow()
doc = part.Document
layActive = window.ActiveLayer

selPointList = List[Point]()
for selection in multiSelection:
    selPointList.Add( context.GetSelectionPoint(selection))

for i in range(0,multiSelection.Count):
    selection = multiSelection[i]
    if isinstance(selection, IDesignFace) == False:
        MessageBox.Show(""Please select a single face"", ""Warning"" , MessageBoxButtons.OK, MessageBoxIcon.Exclamation )
        error = error + 1;
        raise SystemExit(0)
    
    if selection.Root.GetName() == selection.Parent.Parent.GetName():
        raise MessageBox.Show(""Selected body is not within a component."", ""Warning"" , MessageBoxButtons.OK, MessageBoxIcon.Exclamation )
        raise SystemExit(0)

    selPoint = selPointList[i]

    desBody = selection.Parent
    comp = desBody.Parent

    #select component
    for c in part.GetAllComponents():
        if c.GetName() == comp.GetName():
            compSelected = c

    result = ComponentHelper.SetActive(Selection.Create(compSelected), None)

    #selPointD = DatumPoint.Create(GetRootPart(), ""test"", selPoint)

    # Referentievlak maken
    desFace = selection
    test = desFace.Shape
    #_frame = Frame.Create(selPoint, Direction.DirX, Direction.DirZ)
    #_datumplane = DatumPlane.Create(GetRootPart( ), ""test"", Plane.Create(_frame))

    # Referentievlak maken
    selection = Selection.Create(desFace)
    result = DatumPlaneCreator.Create(selection, True, None)

    plane = result.CreatedPlanes[0]

    dirX = plane.AnnotationPlane.Frame.DirX
    dirY = plane.AnnotationPlane.Frame.DirY
    dirZ = plane.AnnotationPlane.Frame.DirZ
    plane.Delete()

    #print dirX
    #print dirY
    #print dirZ 

    if True: # dirX.X <0:
         frame = Frame.Create(selPoint, -dirX,-dirY)     
         plane2 = Plane.Create(frame)
         result = DatumPlaneCreator.Create(selPoint, dirZ)
         plane = result.CreatedPlanes[0]
      
    if parCustomText == ""True"":
        _noteText = parText
    elif parBodyName == ""True"": 
        _noteText = str(desBody.GetName())
    else: 
        _noteText = str(compSelected.GetName())
        
    note = Note.Create( plane,PointUV.Origin,LocationPoint.Center,0.001*parLetterSize, _noteText)
    note.SetFontName(parFont)
    
    if parLocationCenter == ""True"":
        #_NoteCenterLocation = desFace.MidPoint().Point
        _NoteCenterLocation = desFace.Master.EvalMid().Point
        _dp = DatumPoint.Create(part, ""dp"", _NoteCenterLocation)
        
    if parLocationBottomRight == ""True"":
        _NoteCenterLocation = desFace.GetFacePoint(0.8,0.1)
        _dp = DatumPoint.Create(part, ""dp"", _NoteCenterLocation)
    
    if parLocationCenter == ""True"" or parLocationBottomRight == ""True"":
        selection =  Selection.Create(note)
        upToSelection =  Selection.Create(_dp)
        anchorPoint =  Move.GetAnchorPoint(selection)
        options = MoveOptions()
        result = Move.UpTo(selection, upToSelection, anchorPoint, options, None)
        
        _dp.Delete()
    
    if parEngraving == ""True"":                
        lay = doc.GetLayer(""Engravings"")
        ################## CREATE NEW LAYER MARKINGS ##############
        if lay == None:
            lay = Layer.Create(doc, ""Engravings"", Color.Brown)
            
        window.ActiveLayer = lay
        note.Master.Layer = lay
            
        try:
            sm = comp.SheetMetal
            sm.TryApplyNote(note, Lettering.Engraved )
            plane.Delete()
        except:
            MessageBox.Show(""Engraving Failed. Ensure that the selected face belongs to a sheet metal solid body."", ""Warning"" , MessageBoxButtons.OK, MessageBoxIcon.Exclamation )
            plane.Delete()
            result = ComponentHelper.SetRootActive(None)

    if parCutOut == ""True"":
        try:
            sm = comp.SheetMetal
            sm.TryApplyNote(note, Lettering.Cutout )
            plane.Delete()
        except:
            MessageBox.Show(""Engraving Failed. Ensure that the selected face belongs to a sheet metal solid body."", ""Warning"" , MessageBoxButtons.OK, MessageBoxIcon.Exclamation )
            plane.Delete()
            result = ComponentHelper.SetRootActive(None)
        
    window.ActiveLayer = layActive   

result = ComponentHelper.SetRootActive(None)
";
                    var env = ScriptEnvironment.GetOrCreate(true);
                    var success = env.RunCommand(script);
                    // log the result
                    // Logger.Log($"AddNote script returned: {success}");
                });
            }
            catch (Exception)
            {
                // Logger.Log(ex.ToString());
            }

        }

        /// <summary>
        /// Mirrors the old btnMarkeer_ImprintToMarking_Click Python logic.
        /// </summary>
        public static void ImprintToEngravingAndExport()
        {
            // Logger.Log("ImprintToEngravingAndExport");
            var script = @"
# Python Script, API Version = V242
from SpaceClaim.Api.V242 import Layer
part = GetRootPart()
window = GetActiveWindow()
doc = part.Document


def length2points(point1,point2):
return math.sqrt( (point1.X-point2.X)**2 + (point1.Y-point2.Y)**2 + (point1.Z-point2.Z)**2)

############# Check where to save the DXF
dialog = SaveFileDialog()
dialog.Filter = ""DXF|*.dxf"";
dialog.Title = ""Please select a file to save the DXF."";
result = dialog.Show()
if result == True:

########### Get marking Edges (adjacent faces with the same normal)
MarkEdges = []
for b in GetRootPart().Bodies:
    for e in b.Edges:
        if e.Faces.Count == 1:
            MarkEdges.append(e)
        if e.Faces.Count >1:
            _faceNormal = e.Faces[0].GetFaceNormal(0.5, 0.5)
            error = 0
            for f in e.Faces:
                if not (_faceNormal == f.GetFaceNormal(0.5, 0.5) and _faceNormal == f.GetFaceNormal(0.25, 0.25))  :
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

################## Copy to layer ""Engravings""#################
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
    
    
################### Remove Lines from Engriving Layer ####################
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
            
deleteList= []
for cm in curvesMarkings:
    for oc in otherCurves:
        if cm.Shape.StartPoint == oc.Shape.StartPoint and  cm.Shape.EndPoint == oc.Shape.EndPoint:
            deleteList.append(oc)
        
for d in deleteList:
    d.Delete()

      

################## Save File #####################3

fileName = dialog.FileName
options = ExportOptions.Create()
DocumentSave.Execute(fileName, options)

Window.ActiveWindow.Close()
";
            var env = ScriptEnvironment.GetOrCreate(true);
            // Logger.Log($"API = {env.ApiVersion}");
            env.ApiVersion = 242;
            // Logger.Log($"API = {env.ApiVersion}");
            var success = env.RunCommand(script);
        }

        /// <summary>
        /// Mirrors the old btnMarkeer_ImprintBody_Click Python logic.
        /// </summary>
        public static void ImprintBody()
        {
            var script = @"
# Python Script, API Version = V242
from System.Collections.Generic import List
from SpaceClaim.Api.V242 import Application, StatusMessageType
from SpaceClaim.Api.V242.Geometry import Point

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
            var env = ScriptEnvironment.GetOrCreate(true);
            env.ApiVersion = 242;
            var success = env.RunCommand(script);
        }
    }
}
/*
# Python Script, API Version = V242
##########
from SpaceClaim.Api.V242.Geometry import PointUV
from SpaceClaim.Api.V242 import LocationPoint
from SpaceClaim.Api.V242 import Note
from SpaceClaim.Api.V242 import Layer

from SpaceClaim.Api.V242 import Lettering

MessageBox.Show(""test"")

try:
    parLetterSize= float(dataList[0]);
    parEngraving = dataList[1];
    parCustomText =  dataList[2];
    parText =  dataList[3];
    parCutOut =  dataList[4];
    parFont = dataList[5];
    parBodyName = dataList[6];
    parLocationCenter = dataList[7];
    
    

except:
    ########### Parameters
    parLetterSize = 3
    parEngraving = ""False""
    parCutOut = ""False""
    parCustomText = ""True""
    parText = ""Test11""
    parFont = ""Arial Black""
    parBodyName = ""False""
    parLocationCenter = ""True""
    #parLocationBottomRight = ""False""
    


#######################
parLocationBottomRight = ""False""
win = Window.ActiveWindow
context = win.ActiveContext;
multiSelection = context.Selection ## Check how to select two faces
error = 0;
part  = GetRootPart()
window = GetActiveWindow()
doc = part.Document
layActive = window.ActiveLayer

selPointList = List[Point]()
for selection in multiSelection:
    selPointList.Add( context.GetSelectionPoint(selection))
   

for i in range(0,multiSelection.Count):
    selection = multiSelection[i]
    if isinstance(selection, IDesignFace) == False:
        MessageBox.Show(""Please select a single face"", ""Warning"" , MessageBoxButtons.OK, MessageBoxIcon.Exclamation )
        error = error + 1;
        raise SystemExit(0)
    
    if selection.Root.GetName() == selection.Parent.Parent.GetName():
        raise MessageBox.Show(""Selected body is not within a component."", ""Warning"" , MessageBoxButtons.OK, MessageBoxIcon.Exclamation )
        raise SystemExit(0)

    selPoint = selPointList[i]

    desBody = selection.Parent
    comp = desBody.Parent

    #select component
    for c in part.GetAllComponents():
        if c.GetName() == comp.GetName():
            compSelected = c

    
    
    result = ComponentHelper.SetActive(Selection.Create(compSelected), None)

    #selPointD = DatumPoint.Create(GetRootPart(), ""test"", selPoint)

    # Referentievlak maken
    desFace = selection
    test = desFace.Shape
    #_frame = Frame.Create(selPoint, Direction.DirX, Direction.DirZ)
    #_datumplane = DatumPlane.Create(GetRootPart( ), ""test"", Plane.Create(_frame))


    # Referentievlak maken
    selection = Selection.Create(desFace)
    result = DatumPlaneCreator.Create(selection, True, None)

    plane = result.CreatedPlanes[0]

    dirX = plane.AnnotationPlane.Frame.DirX
    dirY = plane.AnnotationPlane.Frame.DirY
    dirZ = plane.AnnotationPlane.Frame.DirZ
    plane.Delete()

    #print dirX
    #print dirY
    #print dirZ 



    if True: # dirX.X <0:
         frame = Frame.Create(selPoint, -dirX,-dirY)     
         plane2 = Plane.Create(frame)
         result = DatumPlaneCreator.Create(selPoint, dirZ)
         plane = result.CreatedPlanes[0]
      
    if parCustomText == ""True"":
        _noteText = parText
    elif parBodyName == ""True"": 
        _noteText = str(desBody.GetName())
    else: 
        _noteText = str(compSelected.GetName())
    
    
    
        
    note = Note.Create( plane,PointUV.Origin,LocationPoint.Center,0.001*parLetterSize, _noteText)
    note.SetFontName(parFont)
    
    
    
    
    if parLocationCenter == ""True"":
        #_NoteCenterLocation = desFace.MidPoint().Point
        _NoteCenterLocation = desFace.Master.EvalMid().Point
        _dp = DatumPoint.Create(part, ""dp"", _NoteCenterLocation)
        
    if parLocationBottomRight == ""True"":
        _NoteCenterLocation = desFace.GetFacePoint(0.8,0.1)
        _dp = DatumPoint.Create(part, ""dp"", _NoteCenterLocation)
    
    if parLocationCenter == ""True"" or parLocationBottomRight == ""True"":
        

        selection =  Selection.Create(note)
        upToSelection =  Selection.Create(_dp)
        anchorPoint =  Move.GetAnchorPoint(selection)
        options = MoveOptions()
        result = Move.UpTo(selection, upToSelection, anchorPoint, options, None)
        
        _dp.Delete()
    
    if parEngraving == ""True"":
        
                            
        lay = doc.GetLayer(""Engravings"")
        ################## CREATE NEW LAYER MARKINGS ##############
        if lay == None:
            lay = Layer.Create(doc, ""Engravings"", Color.Brown)
            
        window.ActiveLayer = lay
        note.Master.Layer = lay
            
        try:
            sm = comp.SheetMetal
            sm.TryApplyNote(note, Lettering.Engraved )
            plane.Delete()
        except:
            MessageBox.Show(""Engraving Failed. Ensure that the selected face belongs to a sheet metal solid body."", ""Warning"" , MessageBoxButtons.OK, MessageBoxIcon.Exclamation )
            plane.Delete()
            result = ComponentHelper.SetRootActive(None)

    if parCutOut == ""True"":
        try:
            sm = comp.SheetMetal
            sm.TryApplyNote(note, Lettering.Cutout )
            plane.Delete()
        except:
            MessageBox.Show(""Engraving Failed. Ensure that the selected face belongs to a sheet metal solid body."", ""Warning"" , MessageBoxButtons.OK, MessageBoxIcon.Exclamation )
            plane.Delete()
            result = ComponentHelper.SetRootActive(None)
    
    
        
    window.ActiveLayer = layActive
        

result = ComponentHelper.SetRootActive(None)
    

";
*/