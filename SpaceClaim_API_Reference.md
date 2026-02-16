# SpaceClaim API V242 Reference

Comprehensive reference for the SpaceClaim API used by AESCConstruct2026. The API assembly and documentation live at `C:\Program Files\ANSYS Inc\v251\scdm\SpaceClaim.Api.V242\`.

## SDK Files

| File | Description |
|------|-------------|
| `SpaceClaim.Api.V242.dll` | Main API assembly |
| `SpaceClaim.Api.V242.xml` | XML documentation (~28,500 lines) |
| `SpaceClaimCustomUI.V242.xsd` | Ribbon XML schema definition |
| `SpaceClaimCustomUI.V242.xsx` | Schema support file |
| `API_Class_Library.chm` | Class library help |
| `API_Combined_Class_Library.chm` | Combined class library help |
| `API_Scripting_Class_Library.chm` | Scripting class library help |
| `SpaceClaim_API.chm` | General API help |
| `Developers Guide.pdf` | Developer guide |
| `Building Sample Add-Ins.pdf` | Sample add-in guide |
| `Style Guide.pdf` | UI style guide |
| `ScriptingFoundation.dll` | Scripting foundation assembly |

## Namespaces Overview (~652 types)

| Namespace | Types | Purpose |
|-----------|-------|---------|
| `SpaceClaim.Api.V242.Extensibility` | 11 | Add-in interfaces and base classes |
| `SpaceClaim.Api.V242` (root) | ~520 | Core API: Application, Document, Part, DesignBody, Command, Tool, etc. |
| `SpaceClaim.Api.V242.Geometry` | ~109 | Point, Vector, Matrix, curves, surfaces, profiles |
| `SpaceClaim.Api.V242.Modeler` | ~70 | Geometry kernel: Body, Face, Edge, booleans, tessellation |
| `SpaceClaim.Api.V242.Display` | 18 | Graphics primitives for rendering |
| `SpaceClaim.Api.V242.Analysis` | 10 | Mesh generation: BodyMesh, FaceElement, VolumeElement |
| `SpaceClaim.Api.V242.Unsupported` | 6 | Internal/experimental features |
| `SpaceClaim.Api.V242.Unsupported.Internal` | 2 | Internal resources |
| `SpaceClaim.Api.V242.Unsupported.RuledCutting` | 1 | Ruled cutting |

---

## Extensibility

### IExtensibility
Required interface for all add-ins.

- `Connect(object connection) → bool` — Called at startup; `connection` is the `Application` instance. Return `true` to load.
- `Disconnect() → void` — Called at shutdown for cleanup.

### IRibbonExtensibility
Customizes the Ribbon UI.

- `GetCustomUI() → string` — Returns XML string (SpaceClaimCustomUI.xsd schema) describing ribbon tabs/groups/buttons.
- `GetRibbonLabel(string id, out string text, out string hint) → bool` — Dynamic text/tooltips for ribbon elements.

### ICommandExtensibility
Creates/modifies commands during startup.

- `Initialize() → void` — Register commands via `Command.Create()`.

### AddIn (base class)
All add-ins derive from this.

- `OnSetDefaultTool() → bool` — Override to set custom default tool when Escape pressed.
- `ExecuteWindowsFormsCode(ThreadStart task) → void` — Execute WinForms code in STA.
- `GetInstance<T>() → T` (static) — Retrieve current add-in instance.

### CommandCapsule
Alternative to raw `Command.Create()` — encapsulates a command with its tool and lifecycle.

---

## Application / Document / Window

### Application
Core API entry point.

**Properties:**
- `MainWindow → Window` — For modal dialogs
- `ActiveDocument → Document` — Currently active document
- `ActiveWindow → Window` — Currently active window
- `Documents → ICollection<Document>` — All open documents
- `Version → ApplicationVersion` — SpaceClaim version info
- `UndoSteps, RedoSteps → int` — Undo/redo counts
- `IsVisible, IsRibbonMinimized → bool` — UI state
- `TitleBarPrefix, TitleBarSuffix → string` — Window title customization

**Methods:**
- `Undo(int steps)`, `Redo(int steps)` — Undo/redo
- `PurgeUndo(int steps)` — Clear undo history
- `Exit()` — Exit application
- `BringToFront()` — Bring to front
- `ReportStatus(string text, StatusMessageType type, Task task)` — Status bar messages
- `CheckLicense(string feature) → bool` — License check
- `AddPanelContent(Command, Control, Panel)` — Replace panel content
- `ShowPanel(Panel, bool)` — Show/hide panels
- `AddPropertyDisplay(PropertyDisplay)` — Add property panel
- `AddSelectionHandler<T>(SelectionHandler<T>)` — Custom selection behavior
- `AddCommandFilter(NativeCommand, CommandFilter)` — Filter native commands
- `AddFileHandler(FileOpenHandler/FileSaveHandler)` — Custom file formats
- `RunScript(string path)` — Execute .scscript or .py
- `ExecuteOnMainThread(Task task)` — Async on main thread

**Events:** `ExitProposed`, `ExitAgreed`, `ContextMenuOpening`, `RadialMenuOpening`, `SessionChanged`, `SessionRolled`, `SessionRolledIncremental`

### Document
Represents a SpaceClaim document.

**Properties:**
- `MainPart → Part` — Primary part
- `Parts → ICollection<Part>` — All parts
- `Symbols → ICollection<Part>` — Design symbols (XY plane sketches)
- `DrawingSheets → ICollection<DrawingSheet>` — Drawing sheets
- `Layers → ICollection<Layer>` — Layer management
- `DefaultLayer → Layer` — Default layer
- `Materials → DocumentMaterialDictionary` — Material library
- `Units → Units` — Unit system
- `Path → string` — Full file path
- `IsModified → bool` — Needs saving
- `IsComplete → bool` — Fully loaded (not lightweight)

**Methods:**
- `Create() → Document` (static) — New document
- `Open(string path, ImportOptions) → Document` (static) — Open/import file
- `Load(string path) → Document` (static) — Load without windows
- `Save()`, `SaveAs(string path)`, `SaveAsSnapshot(string path)` — Save
- `DeleteObjects<T>(ICollection<T>)` — Batch delete
- `InternalizeParts(ICollection<Part>, bool deep)` — Copy external parts
- `ExternalizeParts(ICollection<Part>, string dir)` — Save parts to files
- `GetLayer(string name) → Layer` — Get layer by name
- `GetReferencedDocuments(bool recursive) → ICollection<Document>`

**Events:** `DocumentAdded`, `DocumentRemoved`, `DocumentOpened`, `DocumentClosed`, `DocumentSaving`, `DocumentSaved`, `DocumentChanged`, `CustomPropertiesChanged`, `DocumentCompleted`

### Window
Graphics window displaying a Part or DrawingSheet.

**Properties:**
- `Scene → IDocObject` — Current part or drawing sheet
- `Document → Document` — Parent document
- `ActiveWindow → Window` (static) — Currently active window
- `AllWindows → ICollection<Window>` (static) — All windows
- `ActiveContext → InteractionContext` — Interaction context for selection
- `SceneContext → InteractionContext` — Scene-space context
- `SelectionFilter → SelectionFilterType` — Selection type filtering
- `HomeProjection → Matrix` — Default view orientation
- `Projection → Matrix` — Current view matrix
- `Split → WindowSplit` — Window split arrangement

**Methods:**
- `Create(Part, bool show) → Window` (static) — Create for part
- `Create(DrawingSheet) → Window` (static) — Create for sheet
- `CreateEmbedded(Part, EmbeddedWindowHandler) → Window` (static) — Embedded window
- `Copy() → Window` — Duplicate window
- `GetWindows(Document/Part/DrawingSheet) → ICollection<Window>` (static)
- `GetContext(IDocObject) → InteractionContext` — Context for object
- `ZoomExtents()` — Fit scene
- `ZoomSelection()` — Fit selection
- `SetProjection(Matrix, bool animate, bool lockAxes)` — Set view by matrix
- `SetProjection(Frame, double size)` — Set view by plane/size
- `SetTool(Tool)` — Set active tool
- `Close()` — Close window
- `Export(WindowExportFormat, string path)` — Export view image

**Events:** `SelectionChanged`, `PreselectionChanged`, `WindowOpened`, `WindowClosed`, `ActiveWindowChanged` (static), `ActiveToolChanged`, `InteractionModeChanged`

---

## Design Objects

### Part
A piece part or assembly containing components and design bodies.

**Properties:**
- `Components → ICollection<Component>` — Sub-assemblies
- `Bodies → ICollection<DesignBody>` — Design bodies (solids/surfaces)
- `Curves → ICollection<DesignCurve>` — Design curves
- `CoordinateSystems → ICollection<CoordinateSystem>` — Reference frames
- `DatumPlanes, DatumLines, DatumPoints → ICollection` — Reference geometry
- `DatumFeatures → ICollection<DatumFeature>` — Combined features
- `Beams → ICollection<Beam>` — Structural beams
- `Meshes → ICollection<DesignMesh>` — Design meshes
- `SpotWeldJoints, FilletWelds, Bolts → ICollection` — Fasteners/welds
- `CustomProperties → CustomPartPropertyDictionary` — Part properties
- `IsEmpty → bool` — No design bodies
- `SheetMetal → SheetMetalAspect` — Sheet metal properties
- `AppearanceStates → ICollection<AppearanceState>` — Visual states

**Methods:**
- `Create(Document, string name) → Part` (static)
- `Export(PartExportFormat, string path, bool overwrite, ExportOptions)`
- `ConvertToSheetMetal() → SheetMetalAspect`
- `CreateBeamProfile(Document, string name, Profile) → Part` (static)
- `GetDesignCurvesInPlane(Plane) → ICollection<DesignCurve>`

### Component
Assembly instance referencing a Part template.

**Properties:**
- `Template → Part` — Referenced part master
- `Parent → Part` — Parent part master
- `Content → Part` — Component's bodies/curves
- `Components → ICollection<Component>` — Child components
- `Placement → Matrix` — Transform matrix (rotation/translation only)
- `Name → string` — Component name

**Methods:**
- `Create(Part parent, Part template) → Component` (static)
- `CreateFromFile(Part parent, string path, ImportOptions) → Component` (static)
- `ReplaceFromFile(string path)` — Replace template
- `Transform(Matrix)` — Apply transformation

### DesignBody
Geometry with design intent (solid or surface).

**Properties:**
- `Parent → Part` — Parent part
- `Faces → ICollection<DesignFace>` — Design faces
- `Edges → ICollection<DesignEdge>` — Design edges
- `Shape → Modeler.Body` — Underlying modeler body
- `Material → Material` — Body material
- `SurfaceMaterial → SurfaceMaterial` — Surface appearance
- `MassProperties → MassProperties` — Volume, mass, centroid
- `Layer → Layer` — Layer assignment
- `Style, FinishStyle, RenderingStyle` — Display styles
- `IsLocked → bool` — Prevent modification
- `CanSuppress, IsSuppressed → bool` — Suppression state

**Methods:**
- `Create(Part, string name, Modeler.Body) → DesignBody` (static) — Create from modeler body
- `GetDesignFace(Modeler.Face) → DesignFace` — Get design face wrapper
- `GetDesignEdge(Modeler.Edge) → DesignEdge` — Get design edge wrapper
- `GetDesignBody(Modeler.Body) → DesignBody` (static) — Get design body for modeler body
- `IdentifyHoles(IdentifyHoleOptions) → ICollection<Hole>` — Find holes
- `Save(BodySaveFormat, string path)` — Save body to file
- `GetTessellation(ICollection<Modeler.Face>) → Modeler.Mesh` — Faceted display
- `Transform(Matrix)` — Apply transformation
- `Scale(Frame, double x, double y, double z)` — Non-uniform scaling

### DesignFace
A face on a design body with surface properties.

**Properties:**
- `Parent → DesignBody` — Parent body
- `Edges → ICollection<DesignEdge>` — Boundary edges
- `AdjacentFaces → ICollection<DesignFace>` — Connected faces
- `Shape → Modeler.Face` — Underlying modeler face
- `Area, Perimeter → double` — Geometric properties
- `SurfaceMaterial → SurfaceMaterial` — Appearance

### DesignEdge
An edge on a face of a design body.

**Properties:**
- `Parent → DesignBody` — Parent body
- `Faces → ICollection<DesignFace>` — Adjacent faces
- `Shape → Modeler.Edge` — Underlying modeler edge

### DesignCurve
Construction or sketch curve not part of a body.

**Properties:**
- `Parent → IDesignCurveParent` — Parent part or datum plane
- `Shape → ITrimmedCurve` — Underlying curve geometry
- `IsConstruction → bool` — Sketch vs construction
- `Layer → Layer`, `IsLocked → bool`

**Methods:**
- `Create(IDesignCurveParent, ITrimmedCurve) → DesignCurve` (static)
- `Copy() → DesignCurve` — Duplicate
- `GetLineWeight/SetLineWeight`, `GetLineStyle/SetLineStyle`, `GetColor/SetColor`
- `Transform(Matrix)`

### Other Design Objects
- `CoordinateSystem` — Local coordinate system with `Frame`
- `DatumPlane` / `IDatumPlane` — Reference plane
- `DatumLine` / `IDatumLine` — Reference line
- `DatumPoint` / `IDatumPoint` — Reference point
- `Beam` / `IBeam` — Structural beam with `BeamType`
- `Bolt` / `IBolt` — Bolt with `BoltHeadShape`, `BoltProperties`
- `SpotWeldJoint` / `ISpotWeldJoint` — Spot weld
- `FilletWeld` / `IFilletWeld` — Fillet weld
- `Hole` / `IHole` — Hole feature with `HoleCreationInfo`

---

## DocObject Base

Base class for all design objects.

**Properties:**
- `Document → Document` — Parent document
- `Parent → IDocObject` — Parent object
- `Root → IDocObject` — Root ancestor
- `Moniker → Moniker` — Unique identifier
- `Name → string` — Object name (if `IHasName`)
- `IsDeleted → bool`
- `NumberAttributes → IDictionary<string, double>`
- `TextAttributes → IDictionary<string, string>`

**Methods (hierarchy):**
- `GetAncestor<T>() → T`
- `GetChildren<T>() → ICollection<T>`
- `GetDescendants<T>() → ICollection<T>`

**Methods (occurrences):**
- `GetOccurrence(IList<Instance> path) → IDocObject`
- `GetOccurrence(Instance) → IDocObject`
- `GetOccurrence(IDocObject context) → IDocObject`

**Methods (custom attributes):**
- `SetNumberAttribute(string name, double value)` / `TryGetNumberAttribute(string name, out double) → bool` / `RemoveNumberAttribute(string name)`
- `SetTextAttribute(string name, string value)` / `TryGetTextAttribute(string name, out string) → bool` / `RemoveTextAttribute(string name)`

**Methods (lifecycle):**
- `Delete()`

> Attributes should be prefixed with add-in name (e.g., `"MyAddIn.PropertyName"`) to avoid collisions.

---

## Modeler (Geometry Kernel)

### Body
Underlying geometry kernel body.

**Properties:**
- `Faces → ICollection<Face>` — All faces
- `Edges → ICollection<Edge>` — All edges
- `Vertices → ICollection<Vertex>` — All vertices
- `Shells → ICollection<Shell>` — All shells
- `IsClosed → bool` — Solid vs open
- `IsManifold → bool`
- `PieceCount → int` — Disjoint pieces

**Boolean operations:**
- `Unite(ICollection<Body> tools)` — Union
- `Subtract(ICollection<Body> tools)` — Difference
- `Intersect(ICollection<Body> tools)` — Intersection
- `Fuse(ICollection<Body>, bool removeTJoins, Tracker)` — Topological fusion
- `Imprint(Body tool)` — Imprint intersection edges
- `ImprintCurves(ICollection<ITrimmedCurve>, ICollection<Face>)` — Edges from curves

**Other methods:**
- `Copy() → Body` — Deep copy
- `Copy(IDictionary<Face,Face>, IDictionary<Edge,Edge>) → Body` — Copy with mapping
- `CopyFaces(ICollection<Face>) → Body` — Copy specific faces
- `SeparatePieces() → ICollection<Body>` — Separate disjoint pieces
- `SeparateNonManifold() → ICollection<Body>`
- `CombinePieces(ICollection<Body>)` — Combine disjoint bodies
- `GetIntersections(Body) → BodyIntersection` — Intersection geometry
- `GetTessellation(ICollection<Face>, TessellationOptions) → Mesh` — Faceted mesh
- `RoundEdges(ICollection<KeyValuePair<Edge, EdgeRound>>)` — Fillet/round
- `GetNormal(Point, out Direction) → bool` — Surface normal
- `Save(BodySaveFormat, string path)`

**Static creation methods:**
- `CreatePlanarBody(Plane, ICollection<ITrimmedCurve>) → Body` — Flat profile
- `CreateBox(Frame, double width, double depth, double height) → Body`
- `CreateSphere(Point center, double radius) → Body`
- `CreateCylinder(Frame, double radius, double height) → Body`
- `CreateCone(Frame, double radius, double height, double angle) → Body`
- `CreateTorus(Frame, double majorRadius, double minorRadius) → Body`

### Face
Surface boundary with loops and edges.

**Properties:**
- `Edges → ICollection<Edge>` — Boundary edges
- `Loops → ICollection<Loop>` — Edge loops (outer + inner)
- `AdjacentFaces → ICollection<Face>` — Connected faces
- `Shell → Shell` — Parent shell
- `Area, Perimeter → double`
- `Geometry → Surface` — Plane, Cylinder, Cone, Sphere, Torus, NurbsSurface
- `BoxUV → BoxUV` — Parameter space bounds
- `IsReversed → bool`

**Methods:**
- `GetAdjacentFace(Edge) → Face`
- `ProjectPoint(Point) → PointUV`
- `ContainsParam(PointUV) → bool`
- `IntersectCurve(ITrimmedCurve) → ICollection<CurveParam>`
- `IntersectCurves(IList<ITrimmedCurve>) → ICollection<IntPoint<CurveParam, PointUV>>`
- `ImprintCurves(ICollection<ITrimmedCurve>)`
- `GetClosestSeparation(ITrimmedGeometry) → Separation`
- `GetBoundingBox(Matrix, bool precise) → Box`

### Edge
Boundary curve between vertices.

**Properties:**
- `Faces → ICollection<Face>` — Adjacent faces (1 or 2)
- `Fins → ICollection<Fin>` — Face-edge associations
- `StartVertex, EndVertex → Vertex` — Endpoints
- `StartPoint, EndPoint → Point` — Point positions
- `Shell → Shell`
- `Length → double`
- `Bounds → Interval` — Parameter interval
- `Geometry → Curve` — Line, Circle, Ellipse, NurbsCurve
- `IsSmooth → bool` — Tangent continuous
- `IsConcave → bool` — Concave vs convex
- `IsReversed → bool`
- `Precision → double` — Tolerance

**Methods:**
- `ProjectPoint(Point) → CurveParam`
- `IntersectCurve(ITrimmedCurve) → ICollection<CurveParam>`
- `Offset(Plane, double distance) → ITrimmedCurve`
- `OffsetChain(ICollection<Edge>, Plane, double, OffsetCornerType) → ICollection<ITrimmedCurve>` (static)
- `GetPolyline(PolylineOptions) → ICollection<Point>`
- `GetClosestSeparation(ITrimmedGeometry) → Separation`

### Other Modeler Types
- `Vertex` — Point on body
- `Loop` — Closed sequence of edges bounding a face
- `Shell` — Connected set of faces
- `Fin` — Face-edge association (half-edge)
- `Mesh`, `MeshTopology`, `MeshVertex`, `MeshEdge`, `MeshFace` — Mesh representations
- `BodyRendering`, `FaceRendering`, `EdgeRendering` — Display
- `FaceTessellation`, `Facet` — Tessellation
- `TransformedBody` — Transformed body reference
- `EdgeRound`, `FixedRadiusRound`, `VariableRadiusRound`, `RadiusPoint` — Rounding
- `TessellationOptions` — Tessellation control
- `BodySaveFormat` — Save format enum

---

## Geometry

### Point
3D point coordinates.

- `X, Y, Z → double`
- `Origin → Point` (static) — (0, 0, 0)
- `Create(double x, double y, double z) → Point` (static)
- `Distance(Point) → double`
- `Plus(Vector) → Point`, `Minus(Point) → Vector`, `Minus(Vector) → Point`

### Vector
Displacement vector.

- `X, Y, Z → double`
- `Magnitude → double`
- `Create(double x, double y, double z) → Vector` (static)
- `Dot(Vector, Vector) → double` (static), `Cross(Vector, Vector) → Vector` (static)
- Operators: `+`, `-`, `*`, `/`

### Direction
Unit direction vector.

- `IsZero → bool` — Indeterminate
- `UnitVector → Vector`
- `ArbitraryPerpendicular → Direction`
- `DirX, DirY, DirZ → Direction` (static) — Unit axes
- `Create(double x, double y, double z) → Direction` (static)
- `Create(Vector) → Direction` (static)
- `Cross(Direction, Direction) → Direction` (static)

### Frame
Local coordinate system (origin + XYZ axes).

- `Origin → Point`, `DirX, DirY, DirZ → Direction`
- `Create(Point, Direction dirX, Direction dirY) → Frame` (static)
- `Create(Point, Direction dirZ) → Frame` (static) — Auto X,Y
- `World → Frame` (static) — World frame at origin

### Matrix
4x4 transformation matrix (pre-multiplication).

- `IsIdentity, HasTranslation, HasScale, HasRotation, IsMirror → bool`
- `Translation → Vector`, `Scale → double`, `Rotation → Matrix`
- `CreateTranslation(Vector) → Matrix` (static)
- `CreateScale(double) → Matrix` (static)
- `CreateScale(double, Point origin) → Matrix` (static)
- `CreateRotation(Line axis, double angle) → Matrix` (static)
- `CreateMapping(Frame) → Matrix` (static) — Transform to frame
- `Decompose(out Vector, out Matrix scale, out Matrix rotation)`
- Operators: `*`, `Inverse`

### Curves
- `Curve` — Base curve type
- `Line` — Infinite line (origin + direction)
- `Circle` — Circle (frame + radius)
- `Ellipse` — Ellipse (frame + major/minor radius)
- `Helix` — Helical curve
- `NurbsCurve` — NURBS with `NurbsData`, `Knot`, `ControlPoint`
- `ProceduralCurve` — Procedural definition
- `PointCurve` — Degenerate curve at a point
- `LineSegment` — Bounded line segment

### ITrimmedCurve
Bounded curve segment with endpoints.

- `Geometry → Curve`, `StartPoint, EndPoint → Point`, `Length → double`
- `Bounds → Interval`, `IsReversed → bool`
- `Evaluate(double param) → CurveEvaluation`
- `ProjectPoint(Point) → CurveParam`
- `IntersectCurve(ITrimmedCurve) → ICollection<IntPoint<CurveParam, CurveParam>>`
- `Offset(Plane, double) → ITrimmedCurve`
- `OffsetChain(ICollection<ITrimmedCurve>, Plane, double, OffsetCornerType) → ICollection<ITrimmedCurve>` (static)
- `GetPolyline(PolylineOptions) → ICollection<Point>`
- `IsCoincident(ITrimmedCurve) → bool`
- `ProjectToPlane(Plane) → ITrimmedCurve`

### Surfaces
- `Surface` — Base surface type
- `Plane` — Infinite plane
- `Cylinder` — Cylindrical surface
- `Cone` — Conical surface
- `Sphere` — Spherical surface
- `Torus` — Toroidal surface
- `NurbsSurface` — NURBS surface
- `ProceduralSurface` — Procedural definition

### Profiles
- `Profile` / `IProfile` — Base profile
- `CircleProfile` / `ICircle` — Circle profile
- `RectangleProfile` / `IRectangle` — Rectangle profile
- `SquareProfile` / `ISquare` — Square profile
- `OblongProfile` / `IOblong` — Oblong profile
- `PolygonProfile` / `IPolygon` — Polygon profile
- `RegularPolygonProfile` / `IRegularPolygon` — Regular polygon
- `ArrowProfile` / `IArrow` — Arrow profile

### 2D Geometry
`PointUV`, `DirectionUV`, `VectorUV`, `PlacementUV`, `BoxUV`, `CurveEvaluationUV`, `NurbsCurveUV`

### Other
`Interval`, `Accuracy`, `Parameterization`, `CurveEvaluation`, `CurveParam`, `Collision`, `Separation`, `SelectFragmentResult`

---

## Tools & Interaction

### Tool (base class)
Base for interaction tools. Override mouse/keyboard methods to handle user input.

**Constructors:**
- `Tool()` — All interaction modes
- `Tool(InteractionMode mode)` — Single mode
- `Tool(InteractionMode mode1, InteractionMode mode2)` — Two modes

**Properties:**
- `Window → Window`, `InteractionContext → InteractionContext`
- `StatusText → string` — Status bar text
- `Cursor → Cursor` — Custom cursor
- `Rendering → Graphic` — Temporary graphics
- `OverlayRendering → Graphic` — Pixel-space graphics
- `IsEnabled → bool`, `IsDragging → bool`
- `OptionsXml → string` — Options panel layout XML
- `SelectionTypes → ICollection<Type>` — Filter selectable objects
- `PickboxUpselectBodies → bool` — Auto up-select faces to bodies
- `PickboxIncludedObjectTypes → PickboxObjectType` — Override UI filter
- `Readouts → ICollection<Readout>` — Readout collection
- `Indicators → ICollection<Indicator>` — Point/face/edge/curve indicators
- `Handles → ICollection<Graphic>` — Directional/axis handles

**Lifecycle:**
- `OnEnable(bool enable)` — Activation/deactivation
- `Close()` — Close tool

**Mouse events:**
- `OnDragEvent(Point, Line) → bool` — Triggers OnDragStart/Move/End
- `OnDragStart(Point, Line)`, `OnDragMove(Point, Line)`, `OnDragEnd(Point, Line)`, `OnDragCancel()`
- `OnClickEvent(Point, Line) → bool` — Triggers OnClickStart/Move/End
- `OnClickStart(Point, Line)`, `OnClickMove(Point, Line)`, `OnClickEnd(Point, Line)`, `OnClickCancel()`
- `OnMouseDown(Point, Line, MouseButtons) → bool`, `OnMouseUp(...)`, `OnMouseMove(Point, Line) → bool`
- `OnMouseWheel(Point, int delta) → bool`
- `OnDoubleClick() → bool`, `OnTripleClick() → bool`

**Selection:**
- `OnPickbox(ICollection<IDocObject>) → bool` — Pickbox (click) selection
- `AdjustSelection(IDocObject) → IDocObject` — Custom filtering

**Readout creation:**
- `CreateOffsetReadout(Curve, bool showDistance) → Readout`
- `CreateRadiusReadout(Point center) → Readout`
- `CreateDiameterReadout(Point, bool showRadius) → Readout`
- `CreateLengthReadout(Point, Direction, bool showDistance, bool showOffset) → Readout`
- `CreateAngleReadout(Point/Frame, ...) → Readout`

**Indicator/handle creation:**
- `CreateIndicator(Point/ICollection) → Indicator`
- `CreateDirectionHandle(Color, Point, Direction) → Graphic`
- `CreateAxisHandle(Color, Point, Direction) → Graphic`

---

## Selection System

### SelectionFilterType (flags enum)
`DesignBody`, `DesignFace`, `DesignEdge`, `DesignCurve`, `DatumPlane`, `DatumLine`, `DatumPoint`, `Component`, `CoordinateSystem`, `Beam`, `Mesh`, etc.

### InteractionContext
Scene interaction state (selection, preselection, coordinate mapping).

**Properties:**
- `Window → Window`, `Context → IDocObject`
- `TransformToContext, TransformToScene → Matrix`
- `ActivePart → Part`, `ActiveDatum → DatumPlane`
- `VisibleBodies → ICollection<DesignBody>`
- `Selection → ICollection<IDocObject>` — Currently selected
- `SingleSelection → IDocObject` — Single selection or null
- `SecondarySelection → ICollection<IDocObject>`
- `Preselection → IDocObject` — Currently highlighted
- `PreselectionPoint → Point` — Hit point

**Mapping methods:**
- `MapToContext<T>(IDocObject) → T` — Scene → context
- `MapToContext<T>(ICollection<T>) → ICollection<T>`
- `MapToScene<T>(IDocObject) → T` — Context → scene
- `MapToScene<T>(ICollection<T>) → ICollection<T>`

**Selection methods:**
- `GetSelection<T>() → ICollection<T>` — Get typed selection
- `GetSelectionPoint(IDocObject) → Point` — Hit point for selection
- `SelectById(string[])`, `SecondarySelectById(string[])`
- `SelectByGroup(string[])`

---

## Transaction System

### WriteBlock
Transaction wrapper for modifications (undo system).

- `IsAvailable → bool` (static) — Can enter write block without interrupting interaction
- `IsActive → bool` (static) — Currently in a write block
- `IsInterrupted → bool` (static) — Operation interrupted by user
- `ExecuteTask(string text, Task task)` (static) — Execute within write block with undo step
- `AppendTask(Task task)` (static) — Execute and append to previous undo step

> All design modifications must occur within a `WriteBlock`. Throwing `CommandException` rolls back the entire write block. Do not modify design during Tool drag/click interactions.

---

## Panels & UI

### Panel (enum)
Built-in panels: `Properties`, `Structure`, `Groups`, `Layers`, `Selection`, `Options`, etc.

### PanelTab
- `PanelTab.Create(Command, Control, DockLocation, int width, bool isLocked)` — Create docked panel tab

### Application Panel Methods
- `Application.AddPanelContent(Command, Control, Panel)` — Add to existing panel
- `Application.ShowPanel(Panel, bool)` — Show/hide panel

### DockLocation
`Left`, `Right`, `Top`, `Bottom`

---

## Custom Objects

- `CustomObject` / `ICustomObject` — Object that persists in the document with custom data
- `CustomWrapper<T>` — Typed wrapper for custom objects
- `LightweightCustomWrapper<T>` — Lightweight (no body) custom wrapper
- `LightweightCustomObject` — Lightweight custom object base
- `ICustomObjectParent` / `ICustomObjectParentMaster` — Parent interfaces

---

## Import / Export

### ImportOptions
- `Completeness → ImportCompleteness` — How much to import
- `CleanUpBodies, StitchSurfaces, ImportNames → bool`

### ExportOptions
- `CleanUpBodies, ExportNames → bool`

### Part Export
- `Part.Export(PartExportFormat, string path, bool overwrite, ExportOptions)`

### Document Open
- `Document.Open(string path, ImportOptions) → Document` (static)

### Format-Specific Options
`AcisExportOptions`, `AcisImportOptions`, `CatiaExportOptions`, `CatiaImportOptions`, `StepExportOptions`, `ParasolidExportOptions`, `StlExportOptions` (`StlExportFormat`, `StlFileGranularity`), `StlImportOptions` (`StlImportType`), `ObjExportOptions`, `ObjImportOptions`, `FluentExportOptions`, `FmdExportOptions`, `WorkbenchExportOptions`, `WorkbenchImportOptions`

---

## Ribbon XML Schema

Based on `SpaceClaimCustomUI.V242.xsd`. Full structure:

```xml
<customUI xmlns="http://schemas.spaceclaim.com/customui">
  <ribbon startFromScratch="false">
    <tabGroups>
      <tabGroup id="..." label="..." command="..." />
    </tabGroups>
    <tabs>
      <tab id="..." label="..." command="..." tabGroup="..." insertBefore="...">
        <group id="..." label="..." command="..." layoutOrientation="horizontal|vertical"
               itemSpacing="3" showOptionsButton="false" verticalAlign="middle"
               insertBefore="..." overwriteExisting="false">
          <!-- controls -->
        </group>
      </tab>
    </tabs>
    <menu>
      <!-- application menu buttons/labels -->
    </menu>
  </ribbon>
</customUI>
```

### Control Types

| Control | Key Attributes |
|---------|----------------|
| `<button>` | `id`, `command`, `label`, `size` (small/large), `style` (default/imageandtext/textonlyalways), `position`, `isGroup` |
| `<label>` | `id`, `text`, `width`, `align` (near/center/far), `command` |
| `<textBox>` | `id`, `command`, `width`, `height`, `multiline` |
| `<spinBox>` | `id`, `command`, `width`, `minimumValue`, `maximumValue`, `increment`, `decimalPlaces` |
| `<slider>` | `id`, `command`, `label`, `width`, `minimumValue`, `maximumValue`, `increment`, `orientation` |
| `<checkBox>` | `id`, `command`, `text` |
| `<radioButton>` | `id`, `command`, `text` |
| `<comboBox>` | `id`, `command`, `type` (default/font), `width`, `textEditable`; children: `<item label="...">` |
| `<spacer>` | `id`, `width` (default 3) |
| `<separator>` | `id`, `width` (default 1) |
| `<container>` | `id`, `layoutOrientation`, `itemSpacing`, `verticalAlign`, `horizontalAlign`, `resizeItemsToFit` |
| `<galleryContainer>` | `id`, `width`, `height`, `minimumWidth`, `minimumHeight`, `popup` |

### ST_CommandAttributes (space-separated)
`formattext`, `striphtml`, `noimage`, `notext`, `nokeytips`

### ST_Position (application menu placement)
`new`, `newButton`, `newStart`, `newEnd`, `open`, `openButton`, `save`, `saveButton`, `saveStart`, `saveEnd`, `saveAs`, `saveAsButton`, `saveAsStart`, `saveAsEnd`, `share`, `shareButton`, `print`, `printButton`, `close`, `closeButton`

---

## Complete Type Listing by Namespace

### SpaceClaim.Api.V242.Extensibility (11)
`AddIn`, `CustomHelper`, `LightweightCustomHelper`, `CommandCapsule`, `IVersionedScriptExtensibility`, `IMultiVersionedScriptExtensibility`, `IScriptExtensibility`, `IScriptedCommand`, `IExtensibility`, `ICommandExtensibility`, `IRibbonExtensibility`

### SpaceClaim.Api.V242 (root, ~520)

**Command:** `NativeCommand`, `CommandFilter`, `Command`, `CommandData<T>`, `ControlState`, `ComboBoxState`, `GalleryState`, `SliderState`, `SpinBoxState`, `ExecutionContext`, `CommandException`

**Document/Window:** `Application`, `Document`, `Window`, `Session`, `StartupOptions`, `ApplicationVersion`, `ApplicationWindowState`, `InteractionContext`, `InteractionMode`, `Camera`, `ViewProjection`

**Transactions:** `WriteBlock`, `WriteBlockOptions`, `AutoKeepAliveBlock`, `ManualKeepAliveBlock`, `KeepAliveObject`, `IKeepAliveObject`

**Parts/Components:** `Part`, `IPart`, `Component`, `IComponent`, `PartType`, `PartAspect`, `PartExportFormat`

**Design Objects:** `DesignBody`, `DesignFace`, `DesignEdge`, `DesignCurve`, `IDesignCurve`, `IDesignCurveGroup`, `DesignCurveGroupGeneral`, `DesignCurveGroup`, `DesignBodyAspect`, `MidSurfaceAspect`, `VolumeExtractionAspect`

**Reference Geometry:** `DatumPlane`, `IDatumPlane`, `DatumLine`, `IDatumLine`, `DatumPoint`, `IDatumPoint`, `DatumFeature`, `IDatumFeature`, `DatumAxis`, `IDesignAxis`, `CoordinateSystem`, `ICoordinateSystem`, `CoordinateAxis`, `ICoordinateAxis`, `AxisType`

**Structural:** `Beam`, `IBeam`, `BeamType`, `Bolt`, `IBolt`, `BoltHeadShape`, `BoltProperties`, `SpotWeldJoint`, `ISpotWeldJoint`, `FilletWeld`, `IFilletWeld`, `IWeld`, `WeldIntermittentParameters`, `WeldParameters`, `Gusset`, `IGusset`, `GussetStyle`, `GussetSite`, `UnknownGusset`, `RegularGusset`, `FlatGusset`, `RoundGusset`, `SectionProperties`, `SectionAnchor`, `ConnectionTable`, `WireCurveConnection`, `WireCurvePoint`, `PointType`

**Sheet Metal:** `SheetMetalOptions`, `SheetMetalAspect`, `Bend`, `IBend`, `BendSpecification`, `IHasBendRadius`, `IHasBendAngle`, `AxialBend`, `IAxialBend`, `ConicalBend`, `IConicalBend`, `CylindricalBend`, `ICylindricalBend`, `JoggleBend`, `IJoggleBend`, `JoggleAngles`, `HemBend`, `IHemBend`, `HemStyle`, `ClosedHem`, `OpenHem`, `TeardropHem`, `RolledHem`, `BendStep`, `BendOptions`, `SheetMetalBendHandler`, `SheetMetalForm`, `ISheetMetalForm`, `SheetMetalFormHandler`, `SheetMetalFormPreview`, `SheetMetalFeature`, `ISheetMetalFeature`, `SheetMetalCorner`, `ISheetMetalCorner`, `Bead`, `IBead`, `FlatPatternAspect`, `FlatPatternFormState`, `CustomFormOptions`

**Sheet Metal Forms:** `BridgeForm`, `IBridgeForm`, `CardGuideForm`, `ICardGuideForm`, `BossForm`, `IBossForm`, `DimpleForm`, `IDimpleForm`, `RoundedLouverForm`, `IRoundedLouverForm`, `LouverForm`, `ILouverForm`, `ThreadPunchForm`, `IThreadPunchForm`, `RaisedCountersinkForm`, `IRaisedCountersinkForm`, `CountersinkForm`, `ICountersinkForm`, `CustomForm`, `ICustomForm`, `ExtrusionForm`, `IExtrusionForm`, `CupForm`, `ICupForm`, `KnockoutForm`, `IKnockoutForm`, `CutoutForm`, `ICutoutForm`, `CutoutFormPattern`, `CutoutFacePattern`, `HeavyweightFormPattern`, `LightweightFormPattern`, `HeavyweightFacePattern`, `LightweightFacePattern`

**Sheet Metal Reliefs:** `CornerRelief`, `CornerReliefPlacement`, `CornerReliefPosition`, `DiametricRelief`, `CircularRelief`, `LaserEdgeRelief`, `LaserSymmetricRelief`, `OblongRelief`, `SmoothRelief`, `SquareRelief`, `TriangularRelief`, `DiagonalRelief`, `RectangularRelief`, `CrossBreak`, `Lettering`

**Holes/Features:** `Hole`, `IHole`, `HoleCreationInfo`, `HoleDepthMeasurementType`, `HoleFit`, `IdentifyHoleOptions`, `CurvePoint`, `ICurvePoint`, `FacePoint`, `IFacePoint`, `FacePointType`, `Marker`, `IMarker`, `Section`, `ISection`, `SectionCurve`, `ISectionCurve`

**Meshes:** `DesignMesh`, `IDesignMesh`, `IDesignMeshLoop`, `IDesignMeshItem`, `IDesignMeshRegion`, `IDesignMeshTopology`, `DesignMeshStyle`

**Annotations:** `Note`, `INote`, `NoteCurvature`, `Callout`, `CountCallout`, `Dimension`, `IDimension`, `DimensionCallout`, `DatumFeatureCallout`, `DatumFeatureSymbol`, `IDatumFeatureSymbol`, `DatumReference`, `GeometricToleranceCallout`, `GeometricCharacteristic`, `MaterialCondition`, `FeatureControlFrame`, `ToleranceFormat`, `SurfaceFinish`, `ISurfaceFinish`, `SurfaceFinishType`, `SurfaceFinishParameterType`, `SurfaceFinishParameter`, `Symbol`, `ISymbol`, `SymbolInsert`, `ISymbolInsert`, `SymbolSize`, `SymbolSizeType`, `Table`, `ITable`, `TableCell`, `TableColumn`, `TableRow`, `TableCellChangedEventArgs`, `Image`, `IImage`, `ImageAttachment`, `AttachToPlane`, `AttachToFace`, `ImageLock`, `Barcode`, `IBarcode`, `BarcodeCodePage`, `QRCodeErrorCorrectionLevel`, `DataMatrixBarcodeSize`, `TextInfo`, `IHasText`, `ITextFormattingAttributes`, `IRichText`, `IBlockAnnotation`, `HorizontalAlignment`, `VerticalAlignment`, `TextFit`, `AnnotationSpace`

**Drawing:** `DrawingSheet`, `DrawingSheetContents`, `DrawingSheetBatch`, `DrawingView`, `IDrawingView`, `DrawingViewAspect`, `DrawingViewBoundary`, `DrawingViewSection`, `DrawingViewSectionType`, `DrawingViewStyle`, `AlignedViewAspect`, `DerivedViewAspect`, `DetailViewAspect`, `DrawingViewRendering`, `DrawingViewColorRendering`, `DrawingViewScale`, `DecimalViewScale`, `FractionalViewScale`, `DrawingViewDisplayName`, `LongDisplayName`, `ShortDisplayName`, `CustomDisplayName`, `RenderingBatch`, `ColorRenderingBatch`, `DrawingWindowOptions`, `DrawingSheetWindowExportFormat`

**Materials:** `Material`, `DocumentMaterial`, `LibraryMaterial`, `IHasMaterial`, `DocumentMaterialDictionary`, `LibraryMaterialDictionary`, `MaterialProperty`, `MaterialPropertyDictionary`, `MaterialPropertyId`, `SurfaceMaterial`, `KeyShotSurfaceMaterial`, `TextureSurfaceMaterial`, `AppearanceState`, `IAppearanceState`, `DefaultVisibility`, `IAppearanceContext`, `BodyStyle`, `BodyFinishStyle`, `BodyRenderingStyle`, `LineStyle`, `LineWeight`, `LineWeightType`

**Layers/Groups:** `Layer`, `ILayer`, `IHasLayer`, `Group`, `IGroup`, `Instance`, `IInstance`

**Properties:** `CoreProperties`, `CustomProperty`, `CustomPropertyDictionary`, `CustomPartProperty`, `CustomPartPropertyDictionary`, `PropertyDisplay`, `IHasPropertyGetter`, `IHasPropertySetter`, `ReadOnlyPropertyDisplay`, `SimplePropertyDisplay`, `AdvancedPropertyDisplay`, `IPropertyEditorContext`

**Tools:** `Tool`, `ToolOperation`, `ToolGuidePosition`, `ToolGuideLayout`, `MoveToolHandler`, `BarcodeToolHandler`, `SheetMetalFormHandler`, `Readout`, `ReadoutField`, `Indicator`, `Spacing`, `ValueConverter`, `DoubleConverter`, `LengthConverter`, `AngleConverter`, `ValueDefinition`, `ValueType`

**Selection:** `SelectionResult`, `SelectionHandler<T>`, `SelectionOrder`, `SelectionFilterType`, `PickboxObjectType`

**Units:** `Units`, `UnitsSystem`, `UnitsSystemType`, `IHasUnits`, `MetricUnits`, `ImperialUnits`, `MetricLengthUnit`, `MetricMassUnit`, `ImperialLengthUnit`, `ImperialMassUnit`, `AngleUnit`, `MeasurementUnit`, `INumericPresentation`, `LengthScale`, `MassProperties`, `IHasMassProperties`

**Import/Export:** `ImportOptions`, `ImportData`, `ImportCompleteness`, `MixedImportResolution`, `ExportOptions`, `PartExportFormat`, `WindowExportFormat`, `PartWindowExportFormat`, `AcisExportOptions`, `AcisImportOptions`, `CatiaExportOptions`, `CatiaImportOptions`, `StepExportOptions`, `ParasolidExportOptions`, `StlExportOptions`, `StlExportFormat`, `StlFileGranularity`, `StlImportOptions`, `StlImportType`, `ObjExportOptions`, `ObjImportOptions`, `FluentExportOptions`, `FmdExportOptions`, `WorkbenchExportOptions`, `WorkbenchImportOptions`, `AnalysisOptions`, `AnalysisType`, `AnalysisAspect`, `CADImportParameterType`, `ReplayParameter`, `FileLoadResolver`, `FileOpenHandler`, `FileSaveHandler`

**Window Management:** `WindowSplit`, `LeftRightSplit`, `TopBottomSplit`, `FourWaySplit`, `WindowTab`, `WindowTabHandler`, `EmbeddedWindowHandler`, `DesignWindowOptions`, `WorldOrientation`, `DockLocation`

**Panels/UI:** `Panel`, `PanelTab`, `Tab`, `OptionsPage`, `OptionsPageSection`, `GridSettings`, `LayoutGrid`, `ILayoutGrid`, `LayoutGridDisplay`

**Animation:** `Animation`, `Animator`, `AnimationCompletedEventArgs`, `AnimationResult`, `VideoCapture`, `PixelDepth`

**Context Menus:** `ContextMenu`, `ContextMenuEventArgs`, `ContextMenuId`, `ContextSubMenu`, `RadialMenuEventArgs`

**Custom Objects:** `CustomObject`, `ICustomObject`, `ICustomObjectParent`, `ICustomObjectParentMaster`, `CustomWrapper<T>`, `LightweightCustomWrapper<T>`, `LightweightCustomObject`

**Sketching:** `SketchConstraint`, `ISketchConstraint`, `SketchConstraintType`, `Blocking`

**Base/Misc:** `Base`, `DocObject`, `IDocObject`, `Bounds<T>`, `LocationPoint`, `Moniker<T>`, `IHasName`, `IHasColor`, `IHasShape`, `IHasVisibility`, `ITransformable`, `IDeletable`, `IReplaceable`, `IHasSuppressState`, `UpdateState`, `UpdateFrequency`, `RollDirection`, `RollEventType`, `Aspect<T>`, `IAspect<T>`, `EnclosureAspect`, `ScriptObjectHelper`, `ScriptResult`, `ProgressTracker`, `CancelHandler<T>`, `CancelStatus`, `StatusMessageType`, `ModifyLockedDocumentBlock`, `SaveDocumentOriginator`, `DocumentChangedEventArgs`, `PropertiesChangedEventArgs`, `SessionChangedEventArgs`, `SessionRolledEventArgs`, `ActiveToolChangedEventArgs`, `DisplayImage`

### SpaceClaim.Api.V242.Analysis (10)
`IRegionMeshInfo`, `IRegionMeshInfo<T>`, `RegionMeshInfoBase<T>`, `BodyMesh`, `FaceElement`, `EdgeElement`, `MeshNode`, `PartMesh`, `AssemblyMesh`, `VolumeElement`

### SpaceClaim.Api.V242.Geometry (~109)

**Basic:** `Point`, `Vector`, `Direction`, `Frame`, `IHasFrame`, `Matrix`, `Interval`, `Accuracy`, `CircularSense`, `FitMethod`, `ParamForm`, `Parameterization`

**2D:** `PointUV`, `DirectionUV`, `VectorUV`, `PlacementUV`, `BoxUV`, `UV<T>`, `ControlPointUV`, `CurveEvaluationUV`, `NurbsCurveUV`

**Curves:** `Curve`, `ICurveShape`, `ICurveShape<T>`, `IHasCurveShape`, `IHasCurveShape<T>`, `CurveEvaluation`, `CurveParam`, `Line`, `IHasAxis`, `LineSegment`, `Circle`, `ICircle`, `Ellipse`, `Helix`, `NurbsCurve`, `NurbsData`, `Knot`, `ControlPoint`, `ProceduralCurve`, `PointCurve`, `Polygon`, `IPolygon`, `Polyline`, `OffsetCornerType`

**Trimmed:** `ITrimmedCurve`, `ITrimmedCurve<T>`, `IHasTrimmedCurve`, `IHasTrimmedCurve<T>`, `ITrimmedSurface`, `ITrimmedSurface<T>`, `IHasTrimmedSurface`, `IHasTrimmedSurface<T>`, `ITrimmedSpace`, `IHasTrimmedSpace`

**Surfaces:** `Surface`, `ISurfaceShape`, `ISurfaceShape<T>`, `IHasSurfaceShape`, `IHasSurfaceShape<T>`, `Plane`, `Cylinder`, `Cone`, `Sphere`, `Torus`, `NurbsSurface`, `ProceduralSurface`

**Profiles:** `Profile`, `IProfile`, `CircleProfile`, `RectangleProfile`, `IRectangle`, `SquareProfile`, `ISquare`, `OblongProfile`, `IOblong`, `PolygonProfile`, `RegularPolygonProfile`, `IRegularPolygon`, `IHasRotationalSymmetry`, `ArrowProfile`, `IArrow`, `IHasPlanarPlacement`, `IStar`

**Space:** `Space`, `ISpaceShape`, `Collision`, `Separation`, `IntPoint<T1,T2>`, `IBounded`, `ISpatial`, `IShape`, `SelectFragmentResult`

### SpaceClaim.Api.V242.Modeler (~70)
**Body/Topology:** `Body`, `Face`, `Edge`, `Vertex`, `Loop`, `Shell`, `Fin`, `Topology`, `TransformedBody`

**Mesh:** `Mesh`, `MeshTopology`, `MeshVertex`, `MeshEdge`, `MeshFace`

**Rendering:** `BodyRendering`, `EdgeRendering`, `FaceRendering`, `FaceTessellation`, `Facet`

**Operations:** `BooleanException`, `BooleanException.BooleanFailureType`, `BodyIntersection`, `ProjectionExtent`, `ProjectionOptions`, `TessellationOptions`, `EdgeRound`, `FixedRadiusRound`, `VariableRadiusRound`, `RadiusPoint`, `Tracker`, `LoftTracker`, `BodySaveFormat`

### SpaceClaim.Api.V242.Display (18)
`Primitive`, `PointPrimitive`, `LinePrimitive`, `PolylinePrimitive`, `PolypointPrimitive`, `CurvePrimitive`, `TextPrimitive`, `NotePrimitive`, `ArrowPrimitive`, `MeshEdgeDisplay`, `Polygon`, `PolygonColored`, `PositionNormal`, `PositionColored`, `PositionNormalTextured`, `Graphic`, `GraphicStyle`, `TextPadding`, `DisplayRange`

### SpaceClaim.Api.V242.Unsupported (6)
`SweepOutcome`, `SweepOptions`, `ModelerCreateTransaction`, `ModelerOperations`, `PenStroke`, `IPenStroke`

### SpaceClaim.Api.V242.Unsupported.Internal (2)
`SpaceClaimResources`, `BitmapExtensions`

### SpaceClaim.Api.V242.Unsupported.RuledCutting (1)
`CutLine`
