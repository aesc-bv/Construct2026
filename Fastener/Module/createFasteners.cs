using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace AESCConstruct25.Fastener.Module
{

    class createFasteners
    {

        public static Body Create_Bolt(string boltType, Bolt _bolt, double parBoltL = 0)
        {
            Body bodyBolt = null;
            if (boltType.StartsWith("DIN931") || boltType.StartsWith("DIN933"))
            {
                bodyBolt = createFasteners.createBolt(_bolt.d * 0.001, _bolt.c * 0.001, _bolt.k * 0.001, _bolt.s * 0.001, parBoltL * 0.001);
            }
            else if (boltType.StartsWith("DIN84") || boltType.StartsWith("DIN85"))
            {
                bodyBolt = createFasteners.createBolt84(_bolt.d * 0.001, _bolt.c * 0.001, _bolt.k * 0.001, _bolt.s * 0.001, parBoltL * 0.001, _bolt.t * 0.001, boltType.StartsWith("DIN85"));
            }
            else if (boltType.StartsWith("DIN912") || boltType.Contains("ISO7380"))
            {
                bodyBolt = createFasteners.createBolt912(_bolt.d * 0.001, _bolt.c * 0.001, _bolt.k * 0.001, _bolt.s * 0.001, parBoltL * 0.001, _bolt.t * 0.001, boltType.StartsWith("DIN912"));
            }
            else if (boltType.StartsWith("DIN7991"))
            {
                bodyBolt = createFasteners.createBolt7991(_bolt.d * 0.001, _bolt.c * 0.001, _bolt.k * 0.001, _bolt.s * 0.001, parBoltL * 0.001, _bolt.t * 0.001);
            }
            else
            {
                bodyBolt = createFasteners.createBolt(_bolt.d * 0.001, _bolt.c * 0.001, _bolt.k * 0.001, _bolt.s * 0.001, parBoltL * 0.001);
            }
            return bodyBolt;
        }


        public static Body Create_Nut(string type, Nut _nut)
        {
            Body body = null;
            if (type.StartsWith("DIN982"))
            {
                body = createFasteners.createNut982(_nut.d * 0.001, _nut.s * 0.001, _nut.e * 0.001, _nut.h * 0.001, _nut.k * 0.001);
            }
            else
            {
                body = createFasteners.createNut(_nut.d * 0.001, _nut.s * 0.001, _nut.e * 0.001, _nut.h * 0.001);
            }
            return body;
        }

        private static Body createBolt(double d, double c, double k, double s, double l) //DIN931 or DIN933
        {
            Part mainPart = Window.ActiveWindow.Document.MainPart;
            Point point = Point.Origin;
            Direction dirX = Direction.DirX;
            Direction dirY = Direction.DirY;
            Direction dirZ = Direction.DirZ;
            Body body = null;



            body = Body.ExtrudeProfile(new CircleProfile(Plane.PlaneXY, 0.5 * d), -l);

            Body cyl2 = Body.ExtrudeProfile(new CircleProfile(Plane.PlaneXY, 0.99 * s / 2), c);

            body.Unite(cyl2);

            // Calculate the length of each side of the hexagon based on the pitch
            double sideLength = s / Math.Sqrt(3);

            // Calculate the corners of the hexagon
            Point p0 = point + 0.5 * s * dirY + 0.5 * sideLength * dirX;
            Point p1 = point + 0.5 * s * dirY - 0.5 * sideLength * dirX;
            Point p2 = point - 0.5 * s * dirY + 0.5 * sideLength * dirX;
            Point p3 = point - 0.5 * s * dirY - 0.5 * sideLength * dirX;
            Point p4 = point + (point - p3).Magnitude * dirX;
            Point p5 = point - (point - p3).Magnitude * dirX;


            // Create a list to store the boundary segments
            List<ITrimmedCurve> boundary = new List<ITrimmedCurve> { };


            // Create and store the boundary segments
            boundary.Add(CurveSegment.Create(p1, p0));
            boundary.Add(CurveSegment.Create(p0, p4));
            boundary.Add(CurveSegment.Create(p4, p2));
            boundary.Add(CurveSegment.Create(p2, p3));
            boundary.Add(CurveSegment.Create(p3, p5));
            boundary.Add(CurveSegment.Create(p5, p1));

            //foreach (ITrimmedCurve itc in boundary)
            //    DesignCurve.Create(mainPart, itc);

            //DatumPoint.Create(mainPart, "point", point);
            //DatumPoint.Create(mainPart, "p0", p0);
            //DatumPoint.Create(mainPart, "p1", p1);
            //DatumPoint.Create(mainPart, "p2", p2);
            //DatumPoint.Create(mainPart, "p3", p3);
            //DatumPoint.Create(mainPart, "p4", p4);
            //DatumPoint.Create(mainPart, "p5", p5);

            Body bodyHead = Body.ExtrudeProfile(new Profile(Plane.PlaneXY, boundary), k - c);
            bodyHead.Transform(Matrix.CreateTranslation(c * dirZ));


            body.Unite(bodyHead);

            // round edges head
            Body bodyRound = createRoundBody(point, k, s, dirX, dirY, dirZ);
            //DesignBody.Create(mainPart, "bodyRound", bodyRound);
            body.Subtract(bodyRound);

            // round edges bottom
            List<Point> listPoints = new List<Point> { };
            listPoints.Add(point - l * dirZ + 0.5 * d * dirX);
            AddRounds(body, listPoints, 0.001 * 0.5);

            return body;

        }
        public static void AddRounds(Body body, List<Point> listPoints, double radius)
        {
            Debug.Assert(body != null, "body != null");

            var edgeRounds = new Dictionary<Edge, EdgeRound>();
            foreach (Edge edge in body.Edges)
            {
                foreach (Point point in listPoints)
                {
                    if (edge.ContainsPoint(point))
                    {
                        Interval bounds = edge.Bounds;
                        if (!edgeRounds.ContainsKey(edge))
                        {
                            edgeRounds[edge] = new VariableRadiusRound(new[] {
                        new RadiusPoint(bounds[0.5], radius),
                        new RadiusPoint(bounds[0.5], radius)
                    });
                        }
                        // Break the loop if a matching point is found, 
                        // assuming you don't want to add more rounds for the same edge.
                        break;
                    }
                }
            }
            body.RoundEdges(edgeRounds);
        }

        private static Body createBolt84(double d, double c, double k, double s, double l, double t, bool is85) //DIN84 or DIN85
        {
            Part mainPart = Window.ActiveWindow.Document.MainPart;
            Point point = Point.Origin;
            Direction dirX = Direction.DirX;
            Direction dirY = Direction.DirY;
            Direction dirZ = Direction.DirZ;
            Body body = null;



            body = Body.ExtrudeProfile(new CircleProfile(Plane.PlaneXY, 0.5 * d), -l);



            // Calculate the corners of the head
            Point p0 = point + 0.5 * s * dirX;
            Point p3 = point - 0.5 * s * dirX;
            Point p1 = p0 + k * dirZ - (k * Math.Tan(5.0 / 180 * Math.PI)) * dirX;
            Point p2 = point + k * dirZ;

            // Create a list to store the boundary segments
            List<ITrimmedCurve> boundary = new List<ITrimmedCurve> { };


            // Create and store the boundary segments
            boundary.Add(CurveSegment.Create(p0, p1));
            boundary.Add(CurveSegment.Create(p1, p2));
            boundary.Add(CurveSegment.Create(p2, point));
            boundary.Add(CurveSegment.Create(point, p0));

            Plane planeRoundZX = Plane.Create(Frame.Create(point, dirZ, dirX));
            Plane planeRoundXY = Plane.Create(Frame.Create(point, dirX, dirY));

            Profile profile_round = new Profile(planeRoundZX, boundary);
            Body bodyTop = Body.SweepProfile(profile_round, new CircleProfile(planeRoundXY, 0.01).Boundary);
            body.Unite(bodyTop);


            // round edges
            double rounding = is85 ? s / 4 : 0.001 * 0.5;
            List<Point> listPoints = new List<Point> { };
            listPoints.Add(p1);
            AddRounds(body, listPoints, rounding);

            listPoints = new List<Point> { };
            listPoints.Add(point - l * dirZ + 0.5 * d * dirX);
            AddRounds(body, listPoints, 0.001 * 0.5);

            // remove slot

            // Curves removeBody            
            p0 = point + k * dirZ + 0.5 * c * dirX;
            p1 = p0 - c * dirX;
            p2 = p1 - t * dirZ;
            p3 = p2 + c * dirX;


            List<ITrimmedCurve> boundaryRemoveBody = new List<ITrimmedCurve>{
            CurveSegment.Create(p0, p1),
CurveSegment.Create(p1, p2),
CurveSegment.Create(p2, p3),
CurveSegment.Create(p3, p0)
            };

            Profile profile_remove = new Profile(planeRoundZX, boundaryRemoveBody);
            Body removeBody = Body.ExtrudeProfile(profile_remove, 2 * s);
            removeBody.Transform(Matrix.CreateTranslation(-s * dirY));
            body.Subtract(removeBody);

            return body;

        }


        private static Body createBolt7991(double d, double c, double k, double s, double l, double t) // ISO 7991
        {
            Part mainPart = Window.ActiveWindow.Document.MainPart;
            Point point = Point.Origin;
            Direction dirX = Direction.DirX;
            Direction dirY = Direction.DirY;
            Direction dirZ = Direction.DirZ;

            Body body = Body.ExtrudeProfile(new CircleProfile(Plane.PlaneXY, 0.5 * d), -l);

            // Calculate the corners of the head
            Point p0 = point + 0.5 * s * dirX;
            Point p2 = point - k * dirZ;
            Point p1 = p2 + (0.5 * d) * dirX;

            // Create a list to store the boundary segments
            List<ITrimmedCurve> boundary = new List<ITrimmedCurve> { };

            // Create and store the boundary segments
            boundary.Add(CurveSegment.Create(p0, p1));
            boundary.Add(CurveSegment.Create(p1, p2));
            boundary.Add(CurveSegment.Create(p2, point));
            boundary.Add(CurveSegment.Create(point, p0));


            Plane planeRoundZX = Plane.Create(Frame.Create(point, dirZ, dirX));
            Plane planeRoundXY = Plane.Create(Frame.Create(point, dirX, dirY));

            Profile profile_round = new Profile(planeRoundZX, boundary);
            Body bodyTop = Body.SweepProfile(profile_round, new CircleProfile(planeRoundXY, 0.01).Boundary);

            body.Unite(bodyTop);


            //// Create hexagon cutout
            // Calculate the length of each side of the hexagon based on the pitch
            double sideLength = c / Math.Sqrt(3);

            // Calculate the corners of the hexagon
            p0 = point + 0.5 * c * dirY + 0.5 * sideLength * dirX;
            p1 = point + 0.5 * c * dirY - 0.5 * sideLength * dirX;
            p2 = point - 0.5 * c * dirY + 0.5 * sideLength * dirX;
            Point p3 = point - 0.5 * c * dirY - 0.5 * sideLength * dirX;
            Point p4 = point + (point - p3).Magnitude * dirX;
            Point p5 = point - (point - p3).Magnitude * dirX;

            // Create a list to store the boundary segments
            boundary = new List<ITrimmedCurve> { };


            // Create and store the boundary segments
            boundary.Add(CurveSegment.Create(p1, p0));
            boundary.Add(CurveSegment.Create(p0, p4));
            boundary.Add(CurveSegment.Create(p4, p2));
            boundary.Add(CurveSegment.Create(p2, p3));
            boundary.Add(CurveSegment.Create(p3, p5));
            boundary.Add(CurveSegment.Create(p5, p1));


            Body bodyRemove = Body.ExtrudeProfile(new Profile(Plane.PlaneXY, boundary), -t);
            body.Subtract(bodyRemove);

            // round edges bottom and top
            List<Point> listPoints = new List<Point> { };
            listPoints.Add(point - l * dirZ + 0.5 * d * dirX);
            AddRounds(body, listPoints, 0.001 * 0.5);


            listPoints = new List<Point> { };
            listPoints.Add(point + 0.5 * s * dirX);
            AddRounds(body, listPoints, 0.001 * 0.2);

            return body;

        }
        private static Body createBolt912(double d, double c, double k, double s, double l, double t, bool is912) //DIN912
        {
            Part mainPart = Window.ActiveWindow.Document.MainPart;
            Point point = Point.Origin;
            Direction dirX = Direction.DirX;
            Direction dirY = Direction.DirY;
            Direction dirZ = Direction.DirZ;

            Body body = Body.ExtrudeProfile(new CircleProfile(Plane.PlaneXY, 0.5 * d), -l);
            Body bodyHead = Body.ExtrudeProfile(new CircleProfile(Plane.PlaneXY, 0.5 * s), k);
            body.Unite(bodyHead);

            //// Create hexagon cutout
            // Calculate the length of each side of the hexagon based on the pitch
            double sideLength = c / Math.Sqrt(3);

            // Calculate the corners of the hexagon
            Point p0 = point + 0.5 * c * dirY + 0.5 * sideLength * dirX;
            Point p1 = point + 0.5 * c * dirY - 0.5 * sideLength * dirX;
            Point p2 = point - 0.5 * c * dirY + 0.5 * sideLength * dirX;
            Point p3 = point - 0.5 * c * dirY - 0.5 * sideLength * dirX;
            Point p4 = point + (point - p3).Magnitude * dirX;
            Point p5 = point - (point - p3).Magnitude * dirX;

            // Create a list to store the boundary segments
            List<ITrimmedCurve> boundary = new List<ITrimmedCurve> { };


            // Create and store the boundary segments
            boundary.Add(CurveSegment.Create(p1, p0));
            boundary.Add(CurveSegment.Create(p0, p4));
            boundary.Add(CurveSegment.Create(p4, p2));
            boundary.Add(CurveSegment.Create(p2, p3));
            boundary.Add(CurveSegment.Create(p3, p5));
            boundary.Add(CurveSegment.Create(p5, p1));


            Body bodyRemove = Body.ExtrudeProfile(new Profile(Plane.PlaneXY, boundary), t - k);
            bodyRemove.Transform(Matrix.CreateTranslation(k * dirZ));

            //DesignBody.Create(mainPart, "bodyRemove", bodyRemove);
            body.Subtract(bodyRemove);

            // round edges bottom and top
            List<Point> listPoints = new List<Point> { };
            listPoints.Add(point - l * dirZ + 0.5 * d * dirX);
            AddRounds(body, listPoints, 0.001 * 0.5);


            listPoints = new List<Point> { };
            double rad = is912 ? 0.001 * 0.5 : s / 4;
            listPoints.Add(point + k * dirZ + 0.5 * s * dirX);
            AddRounds(body, listPoints, rad);
            return body;
        }
        public static Body createWasher(double d1, double d2, double s)
        {
            Body outer = Body.ExtrudeProfile(
                new CircleProfile(Plane.PlaneXY, 0.5 * d2),
                s
            );

            Body cutter = Body.ExtrudeProfile(
                new CircleProfile(Plane.PlaneXY, 0.5 * d1),
                s
            );

            Body cutterTool = cutter.Copy();

            var doc = Window.ActiveWindow.Document;
            var fasteners = FastenerModule.GetFastenersPart(doc);
            outer.Subtract(new[] { cutterTool });

            return outer;
        }

        public static Body createNut(double d, double s, double e, double h)
        {
            Point point = Point.Origin;
            Direction dirX = Direction.DirX;
            Direction dirY = Direction.DirY;
            Direction dirZ = Direction.DirZ;

            Body cyl2 = Body.ExtrudeProfile(new CircleProfile(Plane.PlaneXY, d / 2), h);


            // Calculate the length of each side of the hexagon based on the pitch
            double sideLength = s / Math.Sqrt(3);

            // Calculate the corners of the hexagon
            Point p0 = point + 0.5 * s * dirY + 0.5 * sideLength * dirX;
            Point p1 = point + 0.5 * s * dirY - 0.5 * sideLength * dirX;
            Point p2 = point - 0.5 * s * dirY + 0.5 * sideLength * dirX;
            Point p3 = point - 0.5 * s * dirY - 0.5 * sideLength * dirX;
            Point p4 = point + (point - p3).Magnitude * dirX;
            Point p5 = point - (point - p3).Magnitude * dirX;

            // Create a list to store the boundary segments
            List<ITrimmedCurve> boundary = new List<ITrimmedCurve> { };

            // Create and store the boundary segments
            boundary.Add(CurveSegment.Create(p1, p0));
            boundary.Add(CurveSegment.Create(p0, p4));
            boundary.Add(CurveSegment.Create(p4, p2));
            boundary.Add(CurveSegment.Create(p2, p3));
            boundary.Add(CurveSegment.Create(p3, p5));
            boundary.Add(CurveSegment.Create(p5, p1));

            //Part mainPart = Window.ActiveWindow.Document.MainPart;
            //foreach (ITrimmedCurve itc in boundary)
            //    DesignCurve.Create(mainPart, itc);

            //DatumPoint.Create(mainPart, "point", point);
            //DatumPoint.Create(mainPart, "p0", p0);
            //DatumPoint.Create(mainPart, "p1", p1);
            //DatumPoint.Create(mainPart, "p2", p2);
            //DatumPoint.Create(mainPart, "p3", p3);
            //DatumPoint.Create(mainPart, "p4", p4);
            //DatumPoint.Create(mainPart, "p5", p5);

            Body body = Body.ExtrudeProfile(new Profile(Plane.PlaneXY, boundary), h);
            body.Subtract(cyl2);

            Body bodyRound = createRoundBody(point, h, s, dirX, dirY, dirZ);
            body.Subtract(bodyRound);
            bodyRound = createRoundBody(point, 0, s, dirX, dirY, -dirZ);
            body.Subtract(bodyRound);

            return body;

        }


        public static Body createNut982(double d, double s, double e, double h, double k)
        {
            Point point = Point.Origin;
            Direction dirX = Direction.DirX;
            Direction dirY = Direction.DirY;
            Direction dirZ = Direction.DirZ;

            Body cyl2 = Body.ExtrudeProfile(new CircleProfile(Plane.PlaneXY, d / 2), h);


            // Calculate the length of each side of the hexagon based on the pitch
            double sideLength = s / Math.Sqrt(3);

            // Calculate the corners of the hexagon
            Point p0 = point + 0.5 * s * dirY + 0.5 * sideLength * dirX;
            Point p1 = point + 0.5 * s * dirY - 0.5 * sideLength * dirX;
            Point p2 = point - 0.5 * s * dirY + 0.5 * sideLength * dirX;
            Point p3 = point - 0.5 * s * dirY - 0.5 * sideLength * dirX;
            Point p4 = point + (point - p3).Magnitude * dirX;
            Point p5 = point - (point - p3).Magnitude * dirX;

            // Create a list to store the boundary segments
            List<ITrimmedCurve> boundary = new List<ITrimmedCurve> { };

            // Create and store the boundary segments
            boundary.Add(CurveSegment.Create(p1, p0));
            boundary.Add(CurveSegment.Create(p0, p4));
            boundary.Add(CurveSegment.Create(p4, p2));
            boundary.Add(CurveSegment.Create(p2, p3));
            boundary.Add(CurveSegment.Create(p3, p5));
            boundary.Add(CurveSegment.Create(p5, p1));


            Body body = Body.ExtrudeProfile(new Profile(Plane.PlaneXY, boundary), h);
            body.Subtract(cyl2);

            Body bodyRound = createRoundBody(point, h, s, dirX, dirY, dirZ);
            body.Subtract(bodyRound);
            bodyRound = createRoundBody(point, 0, s, dirX, dirY, -dirZ);
            body.Subtract(bodyRound);

            //add ring
            Body bodyRing = Body.ExtrudeProfile(new CircleProfile(Plane.PlaneXY, 0.5 * s), h - k);
            Body extrBody = Body.ExtrudeProfile(new CircleProfile(Plane.PlaneXY, 0.5 * d), h - k);
            bodyRing.Subtract(extrBody);
            body.Transform(Matrix.CreateTranslation((h - k) * dirZ));
            body.Unite(bodyRing);
            return body;

        }

        public static Body createRoundBody(Point point, double k, double s, Direction dirX, Direction dirY, Direction dirZ)
        {
            // add rounding top
            List<ITrimmedCurve> boundaryTop = new List<ITrimmedCurve> { };
            Point pTop1 = point + k * dirZ + 0.5 * s * dirX;
            Point pTop2 = point + (k - s) * dirZ + (0.5 * s + s * Math.Sqrt(3)) * dirX;
            Point pTop3 = point + k * dirZ + s * dirX;


            // Create a list to store the boundary segments
            boundaryTop.Add(CurveSegment.Create(pTop1, pTop2));
            boundaryTop.Add(CurveSegment.Create(pTop2, pTop3));
            boundaryTop.Add(CurveSegment.Create(pTop3, pTop1));
            Plane planeRoundZX = Plane.Create(Frame.Create(point, dirZ, dirX));
            Plane planeRoundXY = Plane.Create(Frame.Create(point, dirX, dirY));
            //Body bodyTop = Body.CreatePlanarBody(planeRound,boundaryTop);
            Profile profile_round = new Profile(planeRoundZX, boundaryTop);

            Body bodyRound = Body.SweepProfile(profile_round, new CircleProfile(planeRoundXY, 0.01).Boundary);


            return bodyRound;
        }
    }
}
