using AESCConstruct25.FrameGenerator.Utilities;
using AESCConstruct25.UI;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

public class Connector
{
    public double Width1 { get; set; }        // TubeLockWidth
    public double Height { get; set; }        // TubeLockHeight
    public double Tolerance { get; set; }     // TubeLockTolerance
    public double Width2 { get; set; }        // TubeLockWidth2
    public double Radius { get; set; }        // TubeLockRadius
    public bool OneSide { get; set; }         // chkTubeLockOneSide.Checked
    public double Location { get; set; }      // TubeLockLocation
    public bool ClickPosition { get; set; }   // chkTubeLockClickPosition.Checked
    public bool HasRounding { get; set; }        // radTubeLockRounding.Checked
    public bool DynamicHeight { get; set; }   // chkTubeLockDynamicHeight.Checked
    public bool RoundCutout { get; set; }     // chkTubeLockRoundCutout.Checked
    public int PatternQty { get; set; }       // TubeLockPattern
    public bool HasPattern { get; set; }      // chkTubeLockPattern.Checked
    public bool HasCornerCutout { get; set; } // new property
    public double CornerCutoutRadius { get; set; } // new property
    public bool RadiusInCutOut { get; set; } // new property
    public double RadiusInCutOut_Radius { get; set; } // new property

    // Constructor
    public Connector(double width1, double height, double tolerance, double width2, double radius, bool oneSide, double location, bool clickPosition, bool rounding, bool dynamicHeight, bool roundCutout, int patternQty, bool hasPattern, bool hasCornerCutout, double cornerCutoutRadius, bool radiusInCutOut, double radiusInCutOut_Radius)
    {
        Width1 = width1;
        Height = height;
        Tolerance = tolerance;
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
    }

    // Simplified method to create a Connector instance using the FormConnector
    public static Connector CreateConnector(ConnectorControl form)
    {
        try
        {
            double ReadDouble(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) throw new FormatException("One or more numeric fields are empty.");
                // Try current culture first (e.g., nl-NL uses comma), then invariant (dot)
                if (double.TryParse(s, NumberStyles.Float, CultureInfo.CurrentCulture, out var v))
                    return v;
                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v))
                    return v;
                // Last-resort normalize commas/dots
                var norm = s.Replace(',', '.');
                if (double.TryParse(norm, NumberStyles.Float, CultureInfo.InvariantCulture, out v))
                    return v;

                throw new FormatException($"Invalid number: \"{s}\"");
            }

            // TextBoxes
            double width1 = ReadDouble(form.connectorWidth1.Text);
            double height = ReadDouble(form.connectorHeight.Text);
            double tolerance = ReadDouble(form.connectorTolerance.Text);
            double width2 = ReadDouble(form.connectorWidth2.Text);
            double radius = ReadDouble(form.connectorRadiusChamfer.Text);
            double location = ReadDouble(form.connectorLocation.Text);

            // CheckBoxes / RadioButtons
            bool oneSide = false;
            bool clickPosition = form.connectorClickLocation.IsChecked == true;
            bool rounding = form.connectorRadius.IsChecked == true;
            bool dynamicHeight = form.connectorDynamicHeight.IsChecked == true;
            bool roundCutout = false;
            bool hasPattern = false;
            bool hasCornerCutout = form.connectorCornerCutout.IsChecked == true;
            bool radiusInCutOut = form.connectorCornerCutoutRadius.IsChecked == true;

            // Corner cutout / optional radius
            double cornerCutoutRadius = ReadDouble(form.connectorCornerCutoutValue.Text);
            double radiusInCutOut_Radius = ReadDouble(form.connectorCornerCutoutRadiusValue.Text);

            int patternQty = 1;

            //Logger.Log(width1.ToString(CultureInfo.CurrentCulture));
            //Logger.Log(height.ToString(CultureInfo.CurrentCulture));
            //Logger.Log(tolerance.ToString(CultureInfo.CurrentCulture));
            //Logger.Log(width2.ToString(CultureInfo.CurrentCulture));
            //Logger.Log(radius.ToString(CultureInfo.CurrentCulture));
            //Logger.Log(oneSide.ToString());
            //Logger.Log(location.ToString(CultureInfo.CurrentCulture));
            //Logger.Log(clickPosition.ToString());
            //Logger.Log(rounding.ToString());
            //Logger.Log(dynamicHeight.ToString());
            //Logger.Log(roundCutout.ToString());
            //Logger.Log(patternQty.ToString(CultureInfo.InvariantCulture));
            //Logger.Log(hasPattern.ToString());
            //Logger.Log(hasCornerCutout.ToString());
            //Logger.Log(cornerCutoutRadius.ToString(CultureInfo.CurrentCulture));
            //Logger.Log(radiusInCutOut.ToString());
            //Logger.Log(radiusInCutOut_Radius.ToString(CultureInfo.CurrentCulture));

            return new Connector(
                width1, height, tolerance, width2, radius,
                oneSide, location, clickPosition, rounding, dynamicHeight,
                roundCutout, patternQty, hasPattern,
                hasCornerCutout, cornerCutoutRadius,
                radiusInCutOut, radiusInCutOut_Radius
            );
        }
        catch (Exception ex)
        {
            //Logger.Log($"CreateConnector parse failed: {ex}");
            return null;
        }
    }


    public double GetDynamicHeigth(Part part, Direction dirX, Direction dirY, Direction dirZ, Point center, double height, double thickness)
    {
        double DynamicHeigth = 0;
        var boundary = new List<ITrimmedCurve>();

        // Convert parameters into meters
        double halfWidth1 = (Width1 / 2) / 1000;
        double halfWidth2 = (Width2 / 2) / 1000;
        Point midPoint = center + halfWidth1 * dirX - thickness / 2 * dirY;

        // Calculate initial corner points
        Point p1 = center - halfWidth1 * dirX; // Bottom left
        Point p6 = center + halfWidth1 * dirX; // Bottom right
        Point p01 = center - halfWidth2 * dirX; // Top left
        Point p06 = center + halfWidth2 * dirX; // Top right


        Point p2 = center - halfWidth2 * dirX + height * dirZ;
        Point p5 = center + halfWidth2 * dirX + height * dirZ;

        boundary.Add(CurveSegment.Create(p1, p2));
        boundary.Add(CurveSegment.Create(p2, p5));
        boundary.Add(CurveSegment.Create(p5, p6));
        boundary.Add(CurveSegment.Create(p6, p1));

        Body body = Body.ExtrudeProfile(new Profile(Plane.Create(Frame.Create(center, dirX, dirZ)), boundary), thickness);

        // check direction of thicknening
        int sign1 = 1;
        if (body.ContainsPoint(center + (0.99 * thickness) * dirY))
        {
            sign1 = -1;
            body = Body.ExtrudeProfile(new Profile(Plane.Create(Frame.Create(center, dirX, dirZ)), boundary), sign1 * thickness);
        }


        List<IDesignBody> _listIDB = part.Document.MainPart.GetDescendants<IDesignBody>().ToList();
        DesignBody.Create(part.Document.MainPart, "checkDynHeightBody", body.Copy());

        IDesignBody idbDynHeightBody = null;
        foreach (IDesignBody idb in _listIDB)
        {
            if (idb.Master.Shape == body)
            {
                idbDynHeightBody = idb;
            }
        }

        // ik krijg de juiste hoogte niet terug, kijk naar dynHeight.scdoc

        return height;

        Box bbox = Box.Empty;
        Matrix mat = Matrix.CreateMapping(Frame.Create(center, dirX, dirY));
        int i = 0;
        foreach (IDesignBody idb in part.Document.MainPart.GetDescendants<IDesignBody>())
        {
            Body _body2 = idb.Master.Shape.Copy();
            _body2.Transform(idb.TransformToMaster);
            Body _body3 = idb.Master.Shape.Copy();
            _body3.Transform(idb.TransformToMaster.Inverse);
            Body _body4 = idb.Master.Shape.Copy();
            DesignBody.Create(part, "body_" + i.ToString(), _body2);
            DesignBody.Create(part, "bodya_" + i.ToString(), _body3);
            DesignBody.Create(part, "bodyb_" + i.ToString(), _body4);
            if (idb.Shape.GetCollision(body) == Collision.Intersect)
            {
                Body body1 = body.Copy();
                Body body2 = idb.Master.Shape.Copy();
                body2.Transform(idb.TransformToMaster.Inverse);
                var a = body1.GetIntersections(body2);

                foreach (BodyIntersection bi in a)
                {
                    var segment = bi.Segment;
                    double length = segment.Length;
                    if (length != 0)
                        continue;
                    Vector vec = (segment.StartPoint - midPoint);
                    double distZ = Vector.Dot(vec, 1.0*dirZ);
                    DynamicHeigth = Math.Max(DynamicHeigth, distZ);

                    //DesignCurve.Create(part, bi.Segment);
                    bbox = bbox | bi.Segment.GetBoundingBox(Matrix.Identity);
                }
            }
        }
       
        return DynamicHeigth;

    }

    public List<ITrimmedCurve> CreateBoundary(Direction dirX, Direction dirY, Direction dirZ, Point center, double widthDiff, double bottomHeigth)
    {
        var boundary = new List<ITrimmedCurve>();
        Point pCenter = center - bottomHeigth * dirZ;

        // Convert parameters into meters
        double halfWidth1 = ((Width1 - widthDiff * 1000) / 2) / 1000;
        double halfWidth2 = ((Width2 - widthDiff * 1000) / 2) / 1000;
        double chamferOrRadius = (Radius / 1000); // Convert radius or chamfer length into meters
        double height = Height / 1000 + bottomHeigth;
        double tol = 0.001 * Tolerance;

        // Calculate initial corner points
        Point p1 = pCenter - halfWidth1 * dirX; // Bottom left
        Point p6 = pCenter + halfWidth1 * dirX; // Bottom right
        Point p01 = pCenter - halfWidth2 * dirX; // Top left
        Point p06 = pCenter + halfWidth2 * dirX; // Top right



        if (chamferOrRadius == 0) // no chamfer/radius
        {
            Point p2 = pCenter - halfWidth2 * dirX + height * dirZ;
            Point p5 = pCenter + halfWidth2 * dirX + height * dirZ;

            boundary.Add(CurveSegment.Create(p1, p2));
            boundary.Add(CurveSegment.Create(p2, p5));
            boundary.Add(CurveSegment.Create(p5, p6));
            boundary.Add(CurveSegment.Create(p6, p1));
        }
        else
        {
            // Calculating angles for adjustment
            double alpha = Math.Atan(height / (halfWidth1 - halfWidth2));
            double alphaDegree = alpha / Math.PI * 180;


            double lengthSide = Math.Sqrt((halfWidth2 - halfWidth1) * (halfWidth2 - halfWidth1) + height * height);
            double dist;
            if (HasRounding)
            {
                double gamma = (Math.PI - alpha) / 2;
                if (halfWidth2 > halfWidth1)
                {
                    gamma = -alpha / 2;
                }
                double gammaDegree = gamma / Math.PI * 180;
                dist = Math.Abs(chamferOrRadius / Math.Tan(gamma));

            }
            else // chamfer
            {
                dist = chamferOrRadius;
            }

            Point p3 = pCenter + (dist - halfWidth2) * dirX + height * dirZ;
            Point p4 = pCenter + (halfWidth2 - dist) * dirX + height * dirZ;

            double dX = (lengthSide - dist) * Math.Cos(alpha);
            double dY = Math.Abs((lengthSide - dist) * Math.Sin(alpha));


            Point p2 = p1 - (halfWidth2 > halfWidth1 ? 1 : -1) * dX * dirX + dY * dirZ;
            Point p5 = p6 + (halfWidth2 > halfWidth1 ? 1 : -1) * dX * dirX + dY * dirZ;


            boundary.Add(CurveSegment.Create(p1, p2));
            if (HasRounding)
            {
                Point pCenter3 = p3 - chamferOrRadius * dirZ;

                ITrimmedCurve itc1 = CurveSegment.CreateArc(pCenter3, p3, p2, -dirY);
                ITrimmedCurve itc2 = CurveSegment.CreateArc(pCenter3, p3, p2, dirY);

                ITrimmedCurve selectedCurve =
                    (itc1.ProjectPoint(p01).Point - p01).Magnitude > (itc2.ProjectPoint(p01).Point - p01).Magnitude
                    ? itc1
                    : itc2;

                boundary.Add(selectedCurve);

                boundary.Add(CurveSegment.Create(p3, p4));
                Point pCenter4 = p4 - chamferOrRadius * dirZ;

                itc1 = CurveSegment.CreateArc(pCenter4, p4, p5, dirY);
                itc2 = CurveSegment.CreateArc(pCenter4, p4, p5, -dirY);

                selectedCurve =
                    (itc1.ProjectPoint(p06).Point - p06).Magnitude > (itc2.ProjectPoint(p06).Point - p06).Magnitude
                    ? itc1
                    : itc2;

                boundary.Add(selectedCurve);

            }
            else // chamfer
            {
                boundary.Add(CurveSegment.Create(p2, p3));
                boundary.Add(CurveSegment.Create(p3, p4));
                boundary.Add(CurveSegment.Create(p4, p5));
            }
            boundary.Add(CurveSegment.Create(p5, p6));
            boundary.Add(CurveSegment.Create(p6, p1));

        }

        return boundary;
    }

    public Body CreateLoft(List<ITrimmedCurve> bound1, Plane plane1, List<ITrimmedCurve> bound2, Plane plane2)
    {
        Body returnBody = null;
        Body loft = null;

        var profiles = new List<ICollection<ITrimmedCurve>> {
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
        return returnBody ;
    }

    public void CreateGeometry(Part part, Direction dirX, Direction dirY, Direction dirZ, Point center, double height, double thickness, out Body connector, out List<Body> cutBodiesSource, out Body cutBody, out Body collisionBody, bool drawBodies = false)
    {
        cutBodiesSource = new List<Body> { };
        connector = null;
        cutBody = null;
        collisionBody = null;

        //Logger.Log("CreateGeometry 1");
        //Logger.Log($"CreateGeometry Width1 {Width1}");
        //Logger.Log($"CreateGeometry Width2 {Width2}");
        //Logger.Log($"CreateGeometry Radius {Radius}");
        //Logger.Log($"CreateGeometry Tolerance {Tolerance}");
        //Logger.Log($"CreateGeometry Height {Height}");

        var boundary = new List<ITrimmedCurve>();
        // Convert parameters into meters
        double halfWidth1 = (Width1 / 2) / 1000;
        double halfWidth2 = (Width2 / 2) / 1000;
        double chamferOrRadius = (Radius / 1000); // Convert radius or chamfer length into meters
        double tol = 0.001 * Tolerance;

        //Logger.Log("CreateGeometry 2");
        // Calculate initial corner points
        Point p1 = center - halfWidth1 * dirX; // Bottom left
        Point p6 = center + halfWidth1 * dirX; // Bottom right
        Point p01 = center - halfWidth2 * dirX; // Top left
        Point p06 = center + halfWidth2 * dirX; // Top right


        //Logger.Log("CreateGeometry 3");

        if (chamferOrRadius == 0) // no chamfer/radius
        {
            Point p2 = center - halfWidth2 * dirX + height * dirZ;
            Point p5 = center + halfWidth2 * dirX + height * dirZ;

            boundary.Add(CurveSegment.Create(p1, p2));
            boundary.Add(CurveSegment.Create(p2, p5));
            boundary.Add(CurveSegment.Create(p5, p6));
            boundary.Add(CurveSegment.Create(p6, p1));

            //Logger.Log("CreateGeometry 4");
        }
        else
        {
            // Calculating angles for adjustment
            double alpha = Math.Atan(height / (halfWidth1 -halfWidth2));
            double alphaDegree = alpha / Math.PI * 180;


            double lengthSide = Math.Sqrt((halfWidth2 - halfWidth1) * (halfWidth2 - halfWidth1) + height * height);
            double dist;
            //Logger.Log($"CreateGeometry HasRounding {HasRounding}");
            if (HasRounding)
            {
                double gamma = (Math.PI - alpha) / 2;
                if (halfWidth2 > halfWidth1)
                {
                    gamma = -alpha/2;
                }
                double gammaDegree = gamma / Math.PI * 180;
                dist = Math.Abs(chamferOrRadius / Math.Tan(gamma));

            }
            else // chamfer
            {
                dist = chamferOrRadius;
            }

            //Logger.Log("CreateGeometry 5");

            Point p3 = center + (dist - halfWidth2) * dirX + height * dirZ;
            Point p4 = center + (halfWidth2 - dist) * dirX + height * dirZ;

            double dX = (lengthSide - dist) * Math.Cos(alpha);
            double dY = Math.Abs((lengthSide - dist) * Math.Sin(alpha));


            Point p2 = p1 - (halfWidth2 > halfWidth1 ? 1 : -1) * dX * dirX + dY * dirZ;
            Point p5 = p6 + (halfWidth2 > halfWidth1 ? 1 : -1) * dX * dirX + dY * dirZ;

            //Logger.Log("CreateGeometry 6");

            boundary.Add(CurveSegment.Create(p1, p2));
            if (HasRounding)
            {
                //Logger.Log("CreateGeometry 7");
                Point pCenter3 = p3 - chamferOrRadius * dirZ;

                ITrimmedCurve itc1 = CurveSegment.CreateArc(pCenter3, p3, p2, -dirY);
                ITrimmedCurve itc2 = CurveSegment.CreateArc(pCenter3, p3, p2, dirY);

                ITrimmedCurve selectedCurve =
                    (itc1.ProjectPoint(p01).Point - p01).Magnitude > (itc2.ProjectPoint(p01).Point - p01).Magnitude
                    ? itc1
                    : itc2;

                //Logger.Log("CreateGeometry 8");
                boundary.Add(selectedCurve);

                boundary.Add(CurveSegment.Create(p3, p4));
                Point pCenter4 = p4 - chamferOrRadius * dirZ;

                itc1 = CurveSegment.CreateArc(pCenter4, p4, p5, dirY);
                itc2 = CurveSegment.CreateArc(pCenter4, p4, p5, -dirY);

                //Logger.Log("CreateGeometry 9");
                selectedCurve =
                    (itc1.ProjectPoint(p06).Point - p06).Magnitude > (itc2.ProjectPoint(p06).Point - p06).Magnitude
                    ? itc1
                    : itc2;

                //Logger.Log("CreateGeometry 10");
                boundary.Add(selectedCurve);
                //Logger.Log("CreateGeometry 11");

            }
            else // chamfer
            {
                //Logger.Log("CreateGeometry 12");
                boundary.Add(CurveSegment.Create(p2, p3));
                boundary.Add(CurveSegment.Create(p3, p4));
                boundary.Add(CurveSegment.Create(p4, p5));
                //Logger.Log("CreateGeometry 13");
            }
            boundary.Add(CurveSegment.Create(p5, p6));
            boundary.Add(CurveSegment.Create(p6, p1));
            //Logger.Log("CreateGeometry 14");

        }

        //Logger.Log("CreateGeometry 15");
        // Draw the design curves in the SpaceClaim environment
        if (false)
        {
            foreach (ITrimmedCurve curve in boundary)
            {
                DesignCurve.Create(part, curve);
            }
            DatumPoint.Create(part, "center",center);
        }
        //Logger.Log($"CreateGeometry center {center}");
        //Logger.Log($"CreateGeometry dirX {dirX}");
        //Logger.Log($"CreateGeometry dirZ {dirZ}");
        //Logger.Log($"CreateGeometry thickness {thickness}");
        //Logger.Log($"CreateGeometry boundary {boundary}");
        foreach(ITrimmedCurve curve in boundary)
        {
            //Logger.Log($"CreateGeometry curve {curve}");
        }
        Frame frame = Frame.Create(center, dirX, dirZ);
        //Logger.Log($"CreateGeometry 15.1 {frame}");
        Plane planep = Plane.Create(frame);
        //Logger.Log($"CreateGeometry 15.2 {planep}");
        Profile profile = new Profile(planep, boundary);
        //Logger.Log($"CreateGeometry 15.3 {profile}");

        Body body = Body.ExtrudeProfile(profile, thickness);
        //Logger.Log("CreateGeometry 15.4");

        //Logger.Log("CreateGeometry 16");
        // check direction of thicknening
        int sign1 = 1;
        if (body.ContainsPoint(center + (0.99 * thickness) * dirY))
        {
            sign1 = -1;
            body = Body.ExtrudeProfile(new Profile(Plane.Create(Frame.Create(center, dirX, dirZ)), boundary), sign1 * thickness);
        }

        //Logger.Log("CreateGeometry 17");

        if (drawBodies)
            DesignBody.Create(part,"Connector", body.Copy());
        connector = body;

        //Logger.Log("CreateGeometry 18");
        // Create box body for offset
        var boundaryBox = new List<ITrimmedCurve>();
        double maxWidth = Math.Max(halfWidth1, halfWidth2);
        boundaryBox.Add(CurveSegment.Create(center - maxWidth * dirX, center - maxWidth * dirX + height * dirZ));
        boundaryBox.Add(CurveSegment.Create(center - maxWidth * dirX + height * dirZ, center + maxWidth * dirX + height * dirZ));
        boundaryBox.Add(CurveSegment.Create(center + maxWidth * dirX + height * dirZ, center + maxWidth * dirX));
        boundaryBox.Add(CurveSegment.Create(center - maxWidth * dirX, center + maxWidth * dirX));

        //Logger.Log("CreateGeometry 19");
        Plane plane = Plane.Create(Frame.Create(center, dirX, dirZ));
        Body boxBody = Body.ExtrudeProfile(new Profile(plane, boundaryBox), sign1 * thickness);

        collisionBody = boxBody.Copy();

        //Logger.Log("CreateGeometry 20");

        boxBody.OffsetFaces(boxBody.Faces, tol);
        if (RadiusInCutOut)
        {
            Plane circlePlane = Plane.Create(Frame.Create(center  + (maxWidth + tol) * dirX + tol * dirY - tol  *dirZ, dirX,dirY));
            Body cylinder = Body.ExtrudeProfile(new CircleProfile(circlePlane, RadiusInCutOut_Radius*0.001), height + 2 * tol);
            int sign = boxBody.GetCollision(cylinder) == Collision.Intersect ? 1 : -1;

            cylinder = Body.ExtrudeProfile(new CircleProfile(circlePlane, RadiusInCutOut_Radius * 0.001), sign * (height + 2 * tol));

            //DesignBody.Create(part, "boxBody 0", boxBody.Copy());
            boxBody.Unite(cylinder.Copy());
            //DesignBody.Create(part, "cylinder 1", cylinder.Copy());
            cylinder.Transform(Matrix.CreateTranslation(-(thickness + 2  * tol) * dirY));
            boxBody.Unite(cylinder.Copy());
            //DesignBody.Create(part, "cylinder 2", cylinder.Copy());
            cylinder.Transform(Matrix.CreateTranslation(-2 * ( maxWidth + tol) * dirX));
            boxBody.Unite(cylinder.Copy());
            //DesignBody.Create(part, "cylinder 3", cylinder.Copy());
            cylinder.Transform(Matrix.CreateTranslation((thickness + 2 * tol) * dirY));
            boxBody.Unite(cylinder.Copy());
            //DesignBody.Create(part, "cylinder 4", cylinder.Copy());
        }
        //Logger.Log("CreateGeometry 21");
        cutBody = boxBody;

        if (drawBodies)
            DesignBody.Create(part, "boxBody", boxBody.Copy());

        Body CCO1 = null;
        Body CCO2 = null;

        //Logger.Log("CreateGeometry 22");
        if (HasCornerCutout)
        {
            Plane circlePlane1 = Plane.Create(Frame.Create(p1 + tol * dirY, dirX, dirZ));
            CCO1 = Body.ExtrudeProfile(new CircleProfile(circlePlane1, CornerCutoutRadius * 0.001), sign1 * (thickness + 2 * tol));
            if (drawBodies)
                DesignBody.Create(part, "CCO1", CCO1.Copy());
            Plane circlePlane2 = Plane.Create(Frame.Create(p6 + tol * dirY, dirX, dirZ));
            CCO2 = Body.ExtrudeProfile(new CircleProfile(circlePlane2, CornerCutoutRadius * 0.001), sign1 * (thickness + 2 * tol));
            if (drawBodies)
                DesignBody.Create(part, "CCO2", CCO2.Copy());

            cutBodiesSource.Add(CCO1);
            cutBodiesSource.Add(CCO2);
        }

        //Logger.Log("CreateGeometry 23");
    }
}