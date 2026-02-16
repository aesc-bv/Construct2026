using AESCConstruct2026.FrameGenerator.Utilities;
using AESCConstruct2026.UI;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

public class Connector
{
    public double Width1 { get; set; }        // TubeLockWidth
    public double Height { get; set; }        // TubeLockHeight
    public double Tolerance { get; set; }     // TubeLockTolerance
    public double EndRelief { get; set; }
    public double Width2 { get; set; }        // TubeLockWidth2
    public double Radius { get; set; }        // TubeLockRadius
    public bool OneSide { get; set; }         // chkTubeLockOneSide.Checked
    public double Location { get; set; }      // TubeLockLocation
    public bool ClickPosition { get; set; }   // chkTubeLockClickPosition.Checked
    public bool HasRounding { get; set; }     // radTubeLockRounding.Checked
    public bool DynamicHeight { get; set; }   // chkTubeLockDynamicHeight.Checked
    public bool RoundCutout { get; set; }     // chkTubeLockRoundCutout.Checked
    public int PatternQty { get; set; }       // TubeLockPattern
    public bool HasPattern { get; set; }      // chkTubeLockPattern.Checked
    public bool HasCornerCutout { get; set; }
    public double CornerCutoutRadius { get; set; }
    public bool RadiusInCutOut { get; set; } 
    public double RadiusInCutOut_Radius { get; set; } 
    public bool ConnectorStraight { get; set; }

    // Constructor
    public Connector(
        double width1,
        double height,
        double tolerance,
        double endRelief,
        double width2,
        double radius,
        bool oneSide,
        double location,
        bool clickPosition,
        bool rounding,
        bool dynamicHeight,
        bool roundCutout,
        int patternQty,
        bool hasPattern,
        bool hasCornerCutout,
        double cornerCutoutRadius,
        bool radiusInCutOut,
        double radiusInCutOut_Radius,
        bool connectorStraight)
    {
        Width1 = width1;
        Height = height;
        Tolerance = tolerance;
        EndRelief = endRelief;
        Width2 = width2;
        Radius = radius;
        OneSide = oneSide;
        Location = location;
        ClickPosition = clickPosition;
        HasRounding = rounding;
        DynamicHeight = dynamicHeight;
        RoundCutout = roundCutout;
        PatternQty = patternQty;
        HasPattern = hasPattern;
        HasCornerCutout = hasCornerCutout;
        CornerCutoutRadius = cornerCutoutRadius;
        RadiusInCutOut = radiusInCutOut;
        RadiusInCutOut_Radius = radiusInCutOut_Radius;
        ConnectorStraight = connectorStraight;
    }

    // Simplified method to create a Connector instance using the FormConnector
    public static Connector CreateConnector(ConnectorControl form)
    {
        // correlation id to tie UI parse → geometry creation → unite/propagation
        string rid = Guid.NewGuid().ToString("N").Substring(0, 8);
        var sw = Stopwatch.StartNew();

        try
        {
            double ParseWithTrace(string label, string s)
            {
                string raw = s ?? "<null>";

                // Accept current culture, invariant, and comma→dot fallback
                if (!string.IsNullOrWhiteSpace(s))
                {
                    if (double.TryParse(s, NumberStyles.Float, CultureInfo.CurrentCulture, out var v1))
                    {
                        return v1;
                    }
                    if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v2))
                    {
                        return v2;
                    }
                    var norm = s.Replace(',', '.');
                    if (double.TryParse(norm, NumberStyles.Float, CultureInfo.InvariantCulture, out var v3))
                    {
                        return v3;
                    }
                }

                throw new FormatException($"Invalid number for {label}: \"{raw}\"");
            }

            // --- parse textboxes ---
            double width1 = ParseWithTrace(nameof(width1), form.connectorWidth1.Text);
            double height = ParseWithTrace(nameof(height), form.connectorHeight.Text);
            double tolerance = ParseWithTrace(nameof(tolerance), form.connectorTolerance.Text);
            double width2 = ParseWithTrace(nameof(width2), form.connectorWidth2.Text);
            double radius = ParseWithTrace(nameof(radius), form.connectorRadiusChamfer.Text);
            double location = ParseWithTrace(nameof(location), form.connectorLocation.Text);
            double endRelief = ParseWithTrace(nameof(endRelief), form.connectorSpacing.Text);
            double cornerCutoutRadius = ParseWithTrace(nameof(cornerCutoutRadius), form.connectorCornerCutoutValue.Text);
            double radiusInCutOut_Radius = ParseWithTrace(nameof(radiusInCutOut_Radius), form.connectorCornerCutoutRadiusValue.Text);

            // --- booleans/radios ---
            bool clickPosition = form.connectorClickLocation.IsChecked == true;
            bool dynamicHeight = form.connectorDynamicHeight.IsChecked == true;
            bool hasCornerCutout = form.connectorCornerCutout.IsChecked == true;
            bool radiusInCutOut = form.connectorCornerCutoutRadius.IsChecked == true;
            bool rounding = form.connectorRadius.IsChecked == true;
            bool chamfering = form.connectorChamfer.IsChecked == true;
            bool oneSide = false;
            bool roundCutout = false;
            bool hasPattern = form.connectorPattern.IsChecked == true;
            int patternQty = (int)ParseWithTrace(nameof(patternQty), form.connectorPatternValue.Text);
            bool connectorStraight = false;

            var result = new Connector(
                width1, height, tolerance, endRelief, width2, radius,
                oneSide, location, clickPosition, rounding, dynamicHeight,
                roundCutout, patternQty, hasPattern,
                hasCornerCutout, cornerCutoutRadius,
                radiusInCutOut, radiusInCutOut_Radius, connectorStraight
            );

            sw.Stop();
            return result;
        }
        catch (Exception ex)
        {
            sw.Stop();
            return null;
        }
    }

    public double GetDynamicHeigth(
        Part part,
        Direction dirX,
        Direction dirY,
        Direction dirZ,
        Point center,
        double height,
        double thickness)
    {
        double DynamicHeigth = 0;
        var boundary = new List<ITrimmedCurve>();

        // Convert parameters into meters
        double halfWidth1 = (Width1 / 2) / 1000;
        double halfWidth2 = (Width2 / 2) / 1000;
        Point midPoint = center + halfWidth1 * dirX - thickness / 2 * dirY;

        // Calculate initial corner points
        Point p1 = center - halfWidth1 * dirX;  // Bottom left
        Point p6 = center + halfWidth1 * dirX;  // Bottom right
        Point p01 = center - halfWidth2 * dirX; // Top left
        Point p06 = center + halfWidth2 * dirX; // Top right

        Point p2 = center - halfWidth2 * dirX + height * dirZ;
        Point p5 = center + halfWidth2 * dirX + height * dirZ;

        boundary.Add(CurveSegment.Create(p1, p2));
        boundary.Add(CurveSegment.Create(p2, p5));
        boundary.Add(CurveSegment.Create(p5, p6));
        boundary.Add(CurveSegment.Create(p6, p1));

        Body body = Body.ExtrudeProfile(
            new Profile(Plane.Create(Frame.Create(center, dirX, dirZ)), boundary),
            thickness);

        // check direction of thicknening
        int sign1 = 1;
        if (body.ContainsPoint(center + (0.99 * thickness) * dirY))
        {
            sign1 = -1;
            body = Body.ExtrudeProfile(
                new Profile(Plane.Create(Frame.Create(center, dirX, dirZ)), boundary),
                sign1 * thickness);
        }

        List<IDesignBody> _listIDB = part.Document.MainPart.GetDescendants<IDesignBody>().ToList();
        DesignBody.Create(part.Document.MainPart, "checkDynHeightBody", body.Copy());

        return height;
    }

    public List<ITrimmedCurve> CreateBoundary(
        Direction dirX,
        Direction dirY,
        Direction dirZ,
        Point center,
        double widthDiff,
        double bottomHeigth,
        double dynHeightVal)
    {
        var boundary = new List<ITrimmedCurve>();

        // -------- inputs (mm → m conversions inline too) --------
        Point pCenter = center - bottomHeigth * dirZ;

        double halfWidth1 = ((Width1 - widthDiff * 1000.0) * 0.5) / 1000.0; // m
        double halfWidth2 = ((Width2 - widthDiff * 1000.0) * 0.5) / 1000.0; // m
        double r = (Radius) / 1000.0;                                       // m
        double h = DynamicHeight ? dynHeightVal : (Height / 1000.0) + bottomHeigth; // m

        // Bottom and top references (X only at top)
        Point p1 = pCenter - halfWidth1 * dirX;        // bottom-left
        Point p6 = pCenter + halfWidth1 * dirX;        // bottom-right
        Point pTopLRef = pCenter - halfWidth2 * dirX;  // top-left ref (no Z)
        Point pTopRRef = pCenter + halfWidth2 * dirX;  // top-right ref (no Z)

        Point p2; // side tangency (left)
        Point p5; // side tangency (right)

        // -------- rectangle / pure trapezoid path --------
        if (r <= 0.0)
        {
            p2 = pTopLRef + h * dirZ;
            p5 = pTopRRef + h * dirZ;

            boundary.Add(CurveSegment.Create(p1, p2));
            boundary.Add(CurveSegment.Create(p2, p5));
            boundary.Add(CurveSegment.Create(p5, p6));
            boundary.Add(CurveSegment.Create(p6, p1));
            return boundary;
        }

        // -------- fillet/chamfer path --------
        const double EPSM = 1e-9;
        double dxAbs = Math.Abs(halfWidth2 - halfWidth1);
        double sideLen = Math.Sqrt(dxAbs * dxAbs + h * h);
        double alpha = (dxAbs < EPSM) ? (Math.PI * 0.5) : Math.Atan(h / dxAbs);


        // distance from TOP corner down along side to arc/chamfer start
        double dist;
        if (HasRounding)
        {
            double gamma = (Math.PI - alpha) * 0.5;
            double tg = Math.Tan(gamma);
            dist = (Math.Abs(tg) < 1e-12) ? r : (r / tg);
        }
        else
        {
            dist = r;
        }

        // clamp dist to guarantee top span and avoid overshoot
        double distBefore = dist;
        dist = Math.Max(0.0, Math.Min(dist, sideLen - 1e-12));
        dist = Math.Min(dist, Math.Max(0.0, halfWidth2 - 1e-12));

        // top inner (tangency) points p3/p4
        Point p3 = pTopLRef + h * dirZ + dist * dirX; // left top inner
        Point p4 = pTopRRef + h * dirZ - dist * dirX; // right top inner

        // from bottom corners to side tangency p2/p5
        double rem = Math.Max(0.0, sideLen - dist);
        double cosA = (sideLen < EPSM) ? 0.0 : (dxAbs / sideLen);
        double sinA = (sideLen < EPSM) ? 1.0 : (h / sideLen);
        double stepX = rem * cosA; // horizontal
        double stepZ = rem * sinA; // vertical

        bool topWider = (halfWidth2 >= halfWidth1);
        p2 = topWider ? (p1 - stepX * dirX + stepZ * dirZ)
                      : (p1 + stepX * dirX + stepZ * dirZ);
        p5 = topWider ? (p6 + stepX * dirX + stepZ * dirZ)
                      : (p6 - stepX * dirX + stepZ * dirZ);

        // tiny degen guards: if any consecutive points are ~identical, nudge a hair on Z
        double Near(Point a, Point b) => (a - b).Magnitude;

        if (Near(p1, p2) < 1e-12)
        {
            p2 = p2 + 1e-9 * dirZ;
        }
        if (Near(p2, p3) < 1e-12)
        {
            p3 = p3 + 1e-9 * dirZ;
        }
        if (Near(p3, p4) < 1e-12)
        {
            p4 = p4 + 1e-9 * dirZ;
        }
        if (Near(p4, p5) < 1e-12)
        {
            p5 = p5 + 1e-9 * dirZ;
        }
        if (Near(p5, p6) < 1e-12)
        {
            p6 = p6 + 1e-9 * dirZ;
        }
        if (Near(p6, p1) < 1e-12)
        {
            p1 = p1 + 1e-9 * dirZ;
        }

        // build loop: p1→p2→(corner)→p3→p4→(corner)→p5→p6→p1
        boundary.Add(CurveSegment.Create(p1, p2));

        if (HasRounding)
        {
            // LEFT arc p2→p3 (center below top edge by r along -dirZ)
            Point cL = p3 - r * dirZ;
            var aL1 = CurveSegment.CreateArc(cL, p2, p3, -dirY);
            var aL2 = CurveSegment.CreateArc(cL, p2, p3, dirY);
            // choose the one consistent with top plane (smaller distance to top ref)
            var aL = (aL1.ProjectPoint(pTopLRef).Point - pTopLRef).Magnitude >
                     (aL2.ProjectPoint(pTopLRef).Point - pTopLRef).Magnitude ? aL1 : aL2;
            boundary.Add(aL);

            // top span
            boundary.Add(CurveSegment.Create(p3, p4));

            // RIGHT arc p4→p5
            Point cR = p4 - r * dirZ;
            var aR1 = CurveSegment.CreateArc(cR, p4, p5, dirY);
            var aR2 = CurveSegment.CreateArc(cR, p4, p5, -dirY);
            var aR = (aR1.ProjectPoint(pTopRRef).Point - pTopRRef).Magnitude >
                     (aR2.ProjectPoint(pTopRRef).Point - pTopRRef).Magnitude ? aR1 : aR2;
            boundary.Add(aR);
        }
        else
        {
            // chamfers
            boundary.Add(CurveSegment.Create(p2, p3));
            boundary.Add(CurveSegment.Create(p3, p4));
            boundary.Add(CurveSegment.Create(p4, p5));
        }

        boundary.Add(CurveSegment.Create(p5, p6)); //Logger.Log("[CreateBoundary] seg p5→p6");
        boundary.Add(CurveSegment.Create(p6, p1)); //Logger.Log("[CreateBoundary] seg p6→p1");

        //Logger.Log($"[CreateBoundary] EXIT segments={boundary.Count}");
        return boundary;
    }

    public Body CreateLoft(
        List<ITrimmedCurve> bound1,
        Plane plane1,
        List<ITrimmedCurve> bound2,
        Plane plane2)
    {
        Body returnBody = null;
        Body loft = null;

        var profiles = new List<ICollection<ITrimmedCurve>>
        {
            bound1,
            bound2
        };

        loft = Body.LoftProfiles(profiles, periodic: false, ruled: false);
        Body cap0 = Body.CreatePlanarBody(plane2, bound2);
        Body cap1 = Body.CreatePlanarBody(plane1, bound1);

        var tracker = Tracker.Create();

        // Stitch modifies 'loft' in place. Do NOT include 'loft' in the tool list.
        loft.Stitch(new List<Body> { cap0, cap1 }, 1e-6, tracker);
        loft.KeepAlive(true);

        returnBody = loft;
        return returnBody;
    }

    public void CreateGeometry(
        Part part,
        Direction dirX,
        Direction dirY,
        Direction dirZ,
        Point center,
        double height,
        double adjustedHeight,
        double thickness,
        out Body connector,
        out List<Body> cutBodiesSource,
        out Body cutBody,
        out Body collisionBody,
        bool drawBodies = true,
        double? cutHeightOverride = null,
        bool rectangularCut = false,
        bool allowCornerFeatures = true) // << NEW optional flag
    {
        // ========= DEBUG UTILITIES =========
        const bool DEBUG = false; // flip to false to silence logs/visuals (keeps drawBodies behavior)

        void DrawPoint(Part prt, string name, Point p, double r = 0.00075)
        {
            try
            {
                if (!DEBUG) return;
                var plane = Plane.Create(Frame.Create(p, dirX, dirZ)); // ⟂ dirY
                var cyl = Body.ExtrudeProfile(new CircleProfile(plane, r), 2 * r);
                var cyn = Body.ExtrudeProfile(new CircleProfile(plane, r), -2 * r);
                cyl.Unite(cyn);
                DesignBody.Create(prt, $"DBG_Point_{name}", cyl);
            }
            catch
            {
            }
        }

        void DrawLine(Part prt, string name, Point a, Point b)
        {
            try
            {
                if (!DEBUG) return;
                var seg = CurveSegment.Create(a, b);
                DesignCurve.Create(prt, seg);
                // Also a very thin rod to visualize in shaded mode
                var pa = Plane.Create(Frame.Create(a, dirX, dirZ));
                var pb = Plane.Create(Frame.Create(b, dirX, dirZ));
                var rodA = Body.ExtrudeProfile(new CircleProfile(pa, 0.0002), 0.0004);
                var rodB = Body.ExtrudeProfile(new CircleProfile(pb, 0.0002), -0.0004);
                rodA.Unite(rodB);
                DesignBody.Create(prt, $"DBG_Line_{name}_A", rodA);
            }
            catch
            {
            }
        }

        Body DrawWireBox(Part prt, string name, IList<Point> loop)
        {
            try
            {
                if (!DEBUG) return null;
                var plane = Plane.Create(Frame.Create(loop[0], dirX, dirZ));
                var prof = new Profile(plane, new List<ITrimmedCurve>
                {
                    CurveSegment.Create(loop[0], loop[1]),
                    CurveSegment.Create(loop[1], loop[2]),
                    CurveSegment.Create(loop[2], loop[3]),
                    CurveSegment.Create(loop[3], loop[0])
                });
                // wafer-thin extrude so it’s visible in shaded display
                var wafer = Body.ExtrudeProfile(prof, 0.0004);
                if (wafer.ContainsPoint(loop[0] + 0.00039 * dirY))
                    wafer = Body.ExtrudeProfile(prof, -0.0004);
                DesignBody.Create(prt, $"DBG_Wire_{name}", wafer.Copy());
                return wafer;
            }
            catch
            {
                return null;
            }
        }

        void TryDrawBody(Part prt, string name, Body b)
        {
            try
            {
                if (DEBUG && b != null)
                    DesignBody.Create(prt, $"DBG_{name}", b.Copy());
            }
            catch
            {
            }
        }

        // Initialize outs
        connector = null;
        cutBody = null;
        collisionBody = null;
        cutBodiesSource = new List<Body>();

        double cutHeight = cutHeightOverride ?? height;
        double connHeight = Math.Max(0.0, adjustedHeight);

        // --- Build profile boundary (unchanged base logic) ---
        var boundary = new List<ITrimmedCurve>();

        double halfWidth1 = Width1 * 0.001 * 0.5;
        double halfWidth2 = Width2 * 0.001 * 0.5;
        double chamferOrRadius = Radius / 1000.0;
        double tol = 0.001 * Tolerance; 

        // Base corners (at bottom)
        Point p1 = center - halfWidth1 * dirX;  // bottom-left
        Point p6 = center + halfWidth1 * dirX;  // bottom-right
        Point p01 = center - halfWidth2 * dirX; // top-left (X ref)
        Point p06 = center + halfWidth2 * dirX; // top-right (X ref)

        DrawPoint(part, "p1_bottomLeft", p1);
        DrawPoint(part, "p6_bottomRight", p6);

        if (chamferOrRadius == 0.0)
        {
            Point p2 = center - halfWidth2 * dirX + connHeight * dirZ;
            Point p5 = center + halfWidth2 * dirX + connHeight * dirZ;

            boundary.Add(CurveSegment.Create(p1, p2));
            boundary.Add(CurveSegment.Create(p2, p5));
            boundary.Add(CurveSegment.Create(p5, p6));
            boundary.Add(CurveSegment.Create(p6, p1));

            DrawPoint(part, "p2_topLeft_noFillet", p2);
            DrawPoint(part, "p5_topRight_noFillet", p5);
            //DBG($"No fillet. p2={p2}, p5={p5}");
        }
        else
        {
            // --- Fillet OR Chamfer build via offset–offset ---
            bool useRadius = this.HasRounding;

            // Top corners of the trapezoid at the connector height (in XZ plane)
            Point TL = center - halfWidth2 * dirX + connHeight * dirZ;  // top-left corner
            Point TR = center + halfWidth2 * dirX + connHeight * dirZ;  // top-right corner
            DrawPoint(part, "TL_unfilleted", TL);
            DrawPoint(part, "TR_unfilleted", TR);

            // Side directions (XZ plane)
            Direction sL = (TL - p1).Direction;
            Direction sR = (TR - p6).Direction;
            Direction topDir = dirX;

            Direction InwardNormal(Point sideBase, Direction sideDir)
            {
                Direction nA = Direction.Cross(dirY, sideDir);
                Direction nB = -nA;
                Point interiorProbe = center + 0.001 * dirZ;
                double dA = Vector.Dot((interiorProbe - sideBase), nA.ToVector());
                double dB = Vector.Dot((interiorProbe - sideBase), nB.ToVector());
                return (dA >= dB) ? nA : nB;
            }

            Point IntersectLinesXZ(Point a0, Direction ad, Point b0, Direction bd)
            {
                Vector ax = ad.ToVector(), bx = bd.ToVector();
                Vector ex = dirX.ToVector(), ez = dirZ.ToVector();
                double a11 = Vector.Dot(ax, ex), a12 = -Vector.Dot(bx, ex);
                double a21 = Vector.Dot(ax, ez), a22 = -Vector.Dot(bx, ez);
                Vector rhs = b0 - a0;
                double b1 = Vector.Dot(rhs, ex), b2 = Vector.Dot(rhs, ez);
                double det = a11 * a22 - a12 * a21;
                if (Math.Abs(det) < 1e-12)
                {
                    double t = (Math.Abs(a11) > 1e-12) ? (b1 / a11) : 0.0;
                    return a0 + t * ad.ToVector();
                }
                double t1 = (b1 * a22 - b2 * a12) / det;
                return a0 + t1 * ad.ToVector();
            }

            Point ProjectToLine(Point a, Point l0, Direction ld)
            {
                Vector v = a - l0;
                double t = Vector.Dot(v, ld.ToVector());
                return l0 + t * ld.ToVector();
            }

            // --- CORNER HANDLING ---
            // shared variables for chamfer connection
            Direction nL = InwardNormal(p1, sL);
            Direction nR = InwardNormal(p6, sR);

            if (useRadius)
            {
                // --- LEFT FILLET ---
                Point topOff_L0 = TL - chamferOrRadius * dirZ;
                Point sideOff_L0 = p1 + chamferOrRadius * nL;
                Point cL = IntersectLinesXZ(topOff_L0, topDir, sideOff_L0, sL);
                Point p3 = cL + chamferOrRadius * dirZ;
                Point p2 = ProjectToLine(cL, p1, sL);
                boundary.Add(CurveSegment.Create(p1, p2));

                var aL_ccw = CurveSegment.CreateArc(cL, p2, p3, dirY);
                var aL_cw = CurveSegment.CreateArc(cL, p2, p3, -dirY);
                Point midProbeL = p2 + 0.5 * (p3 - p2) + 1e-6 * dirZ;
                var arcL = ((aL_ccw.ProjectPoint(midProbeL).Point - midProbeL).Magnitude <
                            (aL_cw.ProjectPoint(midProbeL).Point - midProbeL).Magnitude) ? aL_ccw : aL_cw;
                boundary.Add(arcL);
                TL = p3;

                // --- RIGHT FILLET ---
                Point topOff_R0 = TR - chamferOrRadius * dirZ;
                Point sideOff_R0 = p6 + chamferOrRadius * nR;
                Point cR = IntersectLinesXZ(topOff_R0, topDir, sideOff_R0, sR);
                Point p4 = cR + chamferOrRadius * dirZ;
                Point p5 = ProjectToLine(cR, p6, sR);

                boundary.Add(CurveSegment.Create(TL, p4));
                var aR_ccw = CurveSegment.CreateArc(cR, p4, p5, dirY);
                var aR_cw = CurveSegment.CreateArc(cR, p4, p5, -dirY);
                Point midProbeR = p4 + 0.5 * (p5 - p4) + 1e-6 * dirZ;
                var arcR = ((aR_ccw.ProjectPoint(midProbeR).Point - midProbeR).Magnitude <
                            (aR_cw.ProjectPoint(midProbeR).Point - midProbeR).Magnitude) ? aR_ccw : aR_cw;
                boundary.Add(arcR);
                boundary.Add(CurveSegment.Create(p5, p6));
                boundary.Add(CurveSegment.Create(p6, p1));
            }
            else
            {
                // --- CHAMFER BASED ON LEG LENGTHS (equal legs from chamfer length c) ---
                double c = chamferOrRadius; // given chamfer length (hypotenuse)
                double a = c / Math.Sqrt(2.0); // both legs equal (a=b=c/√2)
                double b = a;

                // Left side chamfer
                Point p2 = p1 + (connHeight - a) * dirZ;                     // vertical leg (down from full height)
                Point p3 = center - (halfWidth2 * dirX) + b * dirX + connHeight * dirZ; // top leg (inward)

                // Right side chamfer
                Point p5 = p6 + (connHeight - a) * dirZ;                     // vertical leg (down)
                Point p4 = center + (halfWidth2 * dirX) - b * dirX + connHeight * dirZ; // top leg (inward)

                // Build loop
                boundary.Add(CurveSegment.Create(p1, p2)); // left side
                boundary.Add(CurveSegment.Create(p2, p3)); // left chamfer
                boundary.Add(CurveSegment.Create(p3, p4)); // top edge
                boundary.Add(CurveSegment.Create(p4, p5)); // right chamfer
                boundary.Add(CurveSegment.Create(p5, p6)); // right side
                boundary.Add(CurveSegment.Create(p6, p1)); // bottom close

                DrawPoint(part, "p2_chamfer_L", p2);
                DrawPoint(part, "p3_chamfer_L", p3);
                DrawPoint(part, "p4_chamfer_R", p4);
                DrawPoint(part, "p5_chamfer_R", p5);
            }
        }

        // Optional visual wire of the boundary
        if (DEBUG)
        {
            foreach (var curve in boundary)
            {
                try { DesignCurve.Create(part, curve); }
                catch { }
            }
            DatumPoint.Create(part, "DBG_center", center);
        }

        // --- Build main connector as a thin extrusion (sheet direction = dirY) ---
        Body main = Body.ExtrudeProfile(
            new Profile(Plane.Create(Frame.Create(center, dirX, dirZ)), boundary),
            thickness);

        int sign1 = 1;
        if (main.ContainsPoint(center + (0.99 * thickness) * dirY))
        {
            sign1 = -1;
            main = Body.ExtrudeProfile(
                new Profile(Plane.Create(Frame.Create(center, dirX, dirZ)), boundary),
                sign1 * thickness);
        }

        TryDrawBody(part, "Connector_profile_extrude_raw", main);
        if (drawBodies)
            DesignBody.Create(part, "Connector", main.Copy());
        connector = main;

        // === Rectangle for collision/cut ===

        // Top corners from widths and cutHeight (no tolerance baked into X/Z)
        Point pTopL = center - halfWidth2 * dirX + cutHeight * dirZ;
        Point pTopR = center + halfWidth2 * dirX + cutHeight * dirZ;

        DrawPoint(part, "pTopL_true", pTopL);
        DrawPoint(part, "pTopR_true", pTopR);
        DrawLine(part, "rectTop", pTopL, pTopR);
        DrawLine(part, "rectLeft", p1, pTopL);
        DrawLine(part, "rectRight", pTopR, p6);
        DrawLine(part, "rectBottom", p6, p1);

        var rectBoundary = new List<ITrimmedCurve>
        {
            CurveSegment.Create(p1,     pTopL),
            CurveSegment.Create(pTopL,  pTopR),
            CurveSegment.Create(pTopR,  p6),
            CurveSegment.Create(p6,     p1)
        };
        var rectProfilePlane = Plane.Create(Frame.Create(center, dirX, dirZ));
        Body rect = Body.ExtrudeProfile(new Profile(rectProfilePlane, rectBoundary), thickness);

        if (rect.ContainsPoint(center + (0.99 * thickness) * dirY))
        {
            rect = Body.ExtrudeProfile(new Profile(rectProfilePlane, rectBoundary), -thickness);
        }

        // Collision: plain rectangle (no tolerance), top at cutHeight (usedHeight)
        collisionBody = rect.Copy();
        TryDrawBody(part, "CollisionRect", collisionBody);

        if (rectangularCut)
        {
            // Explicit rectangular cutter:
            // - bottom locked to baseline
            // - sides widened by +tol
            // - top lifted to connectorHeight + connectorTolerance (mm→m)
            //   unless dynamic height is active (then use cutHeightOverride or height)

            double widest = Math.Max(halfWidth1, halfWidth2);
            double halfW1_g = widest + tol;
            double halfW2_g = widest + tol;

            // connectorHeight and connectorTolerance are in mm
            double topZ_g;
            if (DynamicHeight)
            {
                // use already-calculated cutHeight (converted to meters)
                topZ_g = cutHeight;
            }
            else
            {
                // use connectorHeight + connectorTolerance (convert mm → m)
                topZ_g = (Height + Tolerance) * 0.001;
            }

            Point rp1 = center - halfW1_g * dirX;
            Point rp6 = center + halfW1_g * dirX;
            Point rpTopL = center - halfW2_g * dirX + topZ_g * dirZ;
            Point rpTopR = center + halfW2_g * dirX + topZ_g * dirZ;

            var rLoop = new List<ITrimmedCurve>
            {
                CurveSegment.Create(rp1,    rpTopL),
                CurveSegment.Create(rpTopL, rpTopR),
                CurveSegment.Create(rpTopR, rp6),
                CurveSegment.Create(rp6,    rp1)
            };

            var rProfPlane = Plane.Create(Frame.Create(center, dirX, dirZ));
            cutBody = Body.ExtrudeProfile(new Profile(rProfPlane, rLoop), thickness);

            if (cutBody.ContainsPoint(center + (0.99 * thickness) * dirY))
                cutBody = Body.ExtrudeProfile(new Profile(rProfPlane, rLoop), -thickness);

            TryDrawBody(part, "CutRect_rectangular_match_height", cutBody.Copy());
        }
        else
        {
            // Non-rectangular: inflate by face offset (keeps geometry relationship to connector)
            cutBody = rect.Copy();
            try
            {
                double grow = Math.Max(1e-6, tol);
                cutBody.OffsetFaces(cutBody.Faces, grow);
            }
            catch
            {
            }
            TryDrawBody(part, "CutRect_afterOffset", cutBody.Copy());
        }

        // === Cylindrical cutters (TOP pair) — keep tolerance in +dirY ONLY ===
        if (allowCornerFeatures && RadiusInCutOut && RadiusInCutOut_Radius > 0)
        {
            double r = RadiusInCutOut_Radius * 0.001;
            double depth = thickness + 2 * tol;

            // Anchor to true rect top corners, shift +tol in +Y only
            Point cTopR = pTopR + tol * dirY;
            Point cTopL = pTopL + tol * dirY;

            DrawPoint(part, "cTopR_cylCenter", cTopR);
            DrawPoint(part, "cTopL_cylCenter", cTopL);

            Plane pr = Plane.Create(Frame.Create(cTopR, dirX, dirZ)); // ⟂ dirY
            Plane pl = Plane.Create(Frame.Create(cTopL, dirX, dirZ));

            Body r_pos = Body.ExtrudeProfile(new CircleProfile(pr, r), depth);
            Body r_neg = Body.ExtrudeProfile(new CircleProfile(pr, r), -depth);
            r_pos.Unite(r_neg);
            Body l_pos = Body.ExtrudeProfile(new CircleProfile(pl, r), depth);
            Body l_neg = Body.ExtrudeProfile(new CircleProfile(pl, r), -depth);
            l_pos.Unite(l_neg);

            TryDrawBody(part, "TopCyl_R_raw", r_pos.Copy());
            TryDrawBody(part, "TopCyl_L_raw", l_pos.Copy());

            // Unite cylinders into cutBody so propagation subtracts them:
            try { cutBody.Unite(r_pos.Copy()); }
            catch { }
            try { cutBody.Unite(l_pos.Copy()); }
            catch { }

            // Also drop visual combined top pair for inspection
            if (DEBUG)
            {
                try
                {
                    var combo = r_pos.Copy();
                    combo.Unite(l_pos.Copy());
                    DesignBody.Create(part, "DBG_TopPair_Unified", combo);
                }
                catch
                {
                }
            }
        }

        // === Bottom pair (CornerCutout) — owner cuts via cutBodiesSource ===
        if (allowCornerFeatures && HasCornerCutout && CornerCutoutRadius > 0)
        {
            double rCC = CornerCutoutRadius * 0.001;
            double depth = Math.Abs(sign1) * (thickness + 2 * tol);

            Plane circlePlane1 = Plane.Create(Frame.Create(p1 + tol * dirY, dirX, dirZ));
            Plane circlePlane2 = Plane.Create(Frame.Create(p6 + tol * dirY, dirX, dirZ));

            Body c1_pos = Body.ExtrudeProfile(new CircleProfile(circlePlane1, rCC), depth);
            Body c1_neg = Body.ExtrudeProfile(new CircleProfile(circlePlane1, rCC), -depth);
            c1_pos.Unite(c1_neg);
            Body c2_pos = Body.ExtrudeProfile(new CircleProfile(circlePlane2, rCC), depth);
            Body c2_neg = Body.ExtrudeProfile(new CircleProfile(circlePlane2, rCC), -depth);
            c2_pos.Unite(c2_neg);

            cutBodiesSource.Add(c1_pos.Copy());
            cutBodiesSource.Add(c2_pos.Copy());

            TryDrawBody(part, "CornerCutout_L", c1_pos.Copy());
            TryDrawBody(part, "CornerCutout_R", c2_pos.Copy());
        }

        // === Final sanity visuals ===
        TryDrawBody(part, "CutBody_final", cutBody.Copy());
        TryDrawBody(part, "CollisionBody_final", collisionBody.Copy());

        // Wireframe of the rectangle used for cut/collision
        var wire = DrawWireBox(part, "RectLoop", new[] { p1, pTopL, pTopR, p6 });
    }
}
