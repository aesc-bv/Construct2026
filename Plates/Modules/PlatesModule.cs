using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpaceClaim.Api.V242.Modeler;
using System.Windows.Forms;

namespace AESCConstruct25.Plates.Modules
{
    internal class PlatesModule
    {
        /// <summary>
        /// Public wrapper: call this from your UI, passing in everything from PlatesControl.
        /// </summary>
        public static void CreatePlateFromUI(
            string type, string sizeName, double angleDeg,
            double L1, double L2, int count1,
            double B1, double B2, int count2,
            double thickness, double filletRadius, double holeDiameter,
            bool insertPlateMid = false)
        {
            // you can store those parameters as fields if you want,
            // or just call the old createPart workflow (that reads from selection)
            // and then call a new internal method that actually builds the plate body.
            //
            // For example:
            createPart(type, sizeName, angleDeg,
                            L1, L2, count1,
                            B1, B2, count2,
                            thickness, filletRadius, holeDiameter,
                            insertPlateMid);
        }




        private static void createPart(
            string type, string name,
            double angleDeg,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam,
            bool insertPlateMid)
        {
            //// To do:
            // - Names of the part
            // - Placement when not adding to middle is not always correct
            // - 
            Window win = Window.ActiveWindow;
            Document doc = win.Document;
            Part mainPart = doc.MainPart;

            var selection = win.ActiveContext.Selection;


            foreach (var obj in selection)
            {
                DesignFace df = obj as DesignFace;
                Matrix matrix = Matrix.Identity;
                //if (df != null)
                //{

                //}

                IDesignFace idf = obj as IDesignFace;
                if (idf != null)
                {
                    df = idf.Master;
                    matrix = idf.TransformToMaster;
                }

                if (df == null)
                    continue;

                Plane plane = df.Shape.Geometry as Plane;
                if (plane == null)
                    continue;

                Point selPoint1 = (Point)win.ActiveContext.GetSelectionPoint(selection.First());
                Point selPoint = plane.ProjectPoint(selPoint1).Point;

                Direction dirZ = df.Shape.IsReversed ? -df.Shape.ProjectPoint(selPoint).Normal : df.Shape.ProjectPoint(selPoint).Normal;


                if (insertPlateMid)
                {
                    Box bb = df.Shape.GetBoundingBox(matrix.Inverse, true);
                    var size = bb.Size;
                    selPoint = bb.Center;
                }

                if (!matrix.IsIdentity)
                {
                    plane = plane.CreateTransformedCopy(matrix.Inverse);
                    dirZ = matrix.Inverse * dirZ;
                }


                createPlate(selPoint, plane, dirZ,
                            type, name, angleDeg,
                            L1, L2, Lnr,
                            B1, B2, Bnr,
                            T, Rad, Diam,
                            insertPlateMid);

            }

        }

        private static Part CheckIfPartExists(string type, double L1, double L2, int Lnr, double B1, double B2, int Bnr, double T, double Rad, double Diam)
        {
            Window win = Window.ActiveWindow;
            Document doc = win.Document;
            Part mainPart = doc.MainPart;

            foreach (Part p in doc.Parts)
            {
                if (CompareCustomProperties(p, type, L1, L2, Lnr, B1, B2, Bnr, T, Rad, Diam))
                {
                    return p;
                }
            }
            return null;
        }

        private static Body CreatePlateBody(string type, double L1, double L2, int Lnr, double B1, double B2, int Bnr, double T, double Rad, double Diam)
        {
            Window win = Window.ActiveWindow;
            Document doc = win.Document;
            Part mainPart = doc.MainPart;


            Plane plane = Plane.PlaneXY;
            Point point = Point.Origin;
            Direction dirX = plane.Frame.DirX;
            Direction dirY = plane.Frame.DirY;
            Direction dirZ = plane.Frame.DirZ;

            Body body = null;

            // Create Geometry
            if (type.Contains("support"))
            {
                List<ITrimmedCurve> boundary = new List<ITrimmedCurve> { };
                Plane holePlane = plane.CreateTransformedCopy(Matrix.CreateTranslation(point - plane.Frame.Origin));

                Point p0, p1, p2, p3, p4, p5;
                Point point0 = point + 0.5 * T * dirY;

                p0 = point0 - (0.5 * L1 - Rad) * dirX;
                p1 = p0 + Rad * dirZ - Rad * dirX;
                p2 = p1 + (B1 - Rad) * dirZ;
                p3 = p2 + L1 * dirX;
                p4 = p3 - (B1 - Rad) * dirZ;
                p5 = p4 - Rad * dirZ - Rad * dirX;

                boundary.Add(CurveSegment.Create(p0, p1));
                boundary.Add(CurveSegment.Create(p1, p2));
                boundary.Add(CurveSegment.Create(p2, p3));
                boundary.Add(CurveSegment.Create(p3, p4));
                boundary.Add(CurveSegment.Create(p4, p5));
                boundary.Add(CurveSegment.Create(p5, p0));

                body = Body.ExtrudeProfile(new Profile(Plane.Create(Frame.Create(point0, dirX, dirZ)), boundary), T);
                if (!body.ContainsPoint(point))
                    body.Transform(Matrix.CreateTranslation(T * -dirY));
            }
            else if (type.Contains("flange"))
            {

                Plane holePlane = plane.CreateTransformedCopy(Matrix.CreateTranslation(point - plane.Frame.Origin));
                body = Body.ExtrudeProfile(new CircleProfile(holePlane, 0.5 * L1), 1 * T);
                if (type == "Flat flange")
                {
                    Body cylinderCutOut = Body.ExtrudeProfile(new CircleProfile(holePlane, 0.5 * B1), 1.1 * T);
                    body.Subtract(cylinderCutOut);
                }

                Body cylinderHole = Body.ExtrudeProfile(new CircleProfile(holePlane, 0.5 * Diam), 1.1 * T);
                cylinderHole.Transform(Matrix.CreateTranslation(Vector.Create(L2 * 0.5, 0, 0)));
                int nr = Convert.ToInt32(Lnr);
                double angle = 2 * Math.PI / nr;

                for (int i = 0; i < nr; i++)
                {
                    cylinderHole.Transform(Matrix.CreateRotation(Line.Create(point, dirZ), angle));
                    body.Subtract(cylinderHole.Copy());
                }


            }
            else
            {
                body = createRoundedRectangle(L1, L2, Lnr, B1, B2, Bnr, T, Rad, Diam);
            }
            return body;
        }

        private static Body createRoundedRectangle(double length, double length2, int nrHolesLength, double width, double width2, int nrHolesWidth, double thickness, double radius, double holeDiameter)
        {
            Window win = Window.ActiveWindow;
            Document doc = win.Document;
            Part mainPart = doc.MainPart;


            Plane plane = Plane.PlaneXY;
            Point point = Point.Origin;
            Direction dirX = plane.Frame.DirX;
            Direction dirY = plane.Frame.DirY;
            Direction dirZ = plane.Frame.DirZ;

            List<ITrimmedCurve> boundary = new List<ITrimmedCurve> { };


            // Create base plate
            Point p0, p1, p2, p3, p4, p5, p6, p7;

            if (radius <= 0)
            {
                p0 = point - 0.5 * length * dirX - 0.5 * width * dirY;
                p1 = p0 + width * dirY;
                p2 = p1 + length * dirX;
                p3 = p2 - width * dirY;
                boundary.Add(CurveSegment.Create(p0, p1));
                boundary.Add(CurveSegment.Create(p1, p2));
                boundary.Add(CurveSegment.Create(p2, p3));
                boundary.Add(CurveSegment.Create(p3, p0));
            }
            else
            {
                p0 = point - 0.5 * (length - 2 * radius) * dirX - 0.5 * (width) * dirY;
                p1 = p0 + (length - 2 * radius) * dirX;
                p2 = p1 + radius * dirX + radius * dirY;
                ITrimmedCurve arc12 = CurveSegment.CreateArc(p1 + radius * dirY, p1, p2, dirZ);
                p3 = p2 + (width - 2 * radius) * dirY;
                p4 = p3 - radius * dirX + radius * dirY;
                ITrimmedCurve arc34 = CurveSegment.CreateArc(p4 - radius * dirY, p3, p4, dirZ);
                p5 = p4 - (length - 2 * radius) * dirX;
                p6 = p5 - radius * dirX - radius * dirY;
                ITrimmedCurve arc56 = CurveSegment.CreateArc(p6 + radius * dirX, p5, p6, dirZ);
                p7 = p6 - (width - 2 * radius) * dirY;
                ITrimmedCurve arc70 = CurveSegment.CreateArc(p0 + radius * dirY, p7, p0, dirZ);

                boundary.Add(CurveSegment.Create(p0, p1));
                boundary.Add(arc12);
                boundary.Add(CurveSegment.Create(p2, p3));
                boundary.Add(arc34);
                boundary.Add(CurveSegment.Create(p4, p5));
                boundary.Add(arc56);
                boundary.Add(CurveSegment.Create(p6, p7));
                boundary.Add(arc70);




            }
            //foreach (ITrimmedCurve itc in boundary)
            //{
            //    DesignCurve.Create(mainPart, itc);
            //}

            Body body = Body.ExtrudeProfile(new Profile(plane, boundary), thickness);

            // Create holes
            Body cylinder = Body.ExtrudeProfile(new CircleProfile(plane, 0.5 * holeDiameter), 1.1 * thickness);
            for (int i = 0; i < nrHolesLength; i++)
            {
                double dX = -0.5 * length2 + (nrHolesLength > 1 ? i * length2 / (nrHolesLength - 1) : 0);
                for (int j = 0; j < nrHolesWidth; j++)
                {
                    double dY = -0.5 * width2 + (nrHolesWidth > 1 ? j * width2 / (nrHolesWidth - 1) : 0);
                    Body cylCopy = cylinder.Copy();
                    cylCopy.Transform(Matrix.CreateTranslation(Vector.Create(dX, dY, 0)));
                    try
                    {

                        body.Subtract(cylCopy);
                    }
                    catch
                    {
                        cylCopy.Dispose();
                    }
                }

            }
            cylinder.Dispose();




            return body;
        }

        private static void createPlate(Point point, Plane plane, Direction dirZ,
            string type, string name,
            double angleDeg,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam,
            bool insertPlateMid)
        {
            Window win = Window.ActiveWindow;
            Document doc = win.Document;
            Part mainPart = doc.MainPart;

            Direction dirX = plane.Frame.DirX;
            Direction dirY = plane.Frame.DirY;

            string partName = type + "_" + name; // TO DO

            WriteBlock.ExecuteTask("Create part", () =>
            {
                // Create Plates part
                Part partPlates = doc.Parts.FirstOrDefault(p => p.DisplayName == "Plates");
                if (partPlates == null)
                {
                    partPlates = Part.Create(doc, "Plates");
                    Component compPlates = Component.Create(mainPart, partPlates);
                }


                Part part = CheckIfPartExists(type, L1, L2, Lnr, B1, B2, Bnr, T, Rad, Diam);
                Component component = null;

                if (part != null)
                {
                    // Create a copy of the original component
                    component = Component.Create(partPlates, part);

                }
                else
                {
                    Body body = CreatePlateBody(type, L1, L2, Lnr, B1, B2, Bnr, T, Rad, Diam);

                    if (body == null)
                        return;


                    Part partBody = Part.Create(doc, partName);
                    DesignBody.Create(partBody, partName, body);
                    component = Component.Create(partPlates, partBody);


                    CreateCustomPartPropertiesPlate(partBody, type, L1, B1, T, L2, Lnr, B2, Bnr, Diam, Rad, angleDeg);
                }
                // Move component to correct position
                Frame frame = Frame.Create(point, dirX, dirY);
                if (frame.DirZ != dirZ)
                    frame = Frame.Create(point, -dirX, dirY);
                Matrix matrixMapping = Matrix.CreateMapping(frame);
                component.Transform(matrixMapping);

                if (angleDeg != 0)
                {
                    double angleRad = angleDeg / 180 * Math.PI;
                    component.Transform(Matrix.CreateRotation(Line.Create(point, dirZ), angleRad));
                }

            });
        }


        private static void CreateCustomPartPropertiesPlate(Part part, string type, double L1, double B1, double T, double L2, int Lnr, double B2, int Bnr, double Diam, double Rad, double Angle)
        {

            // Add custom component properties
            CustomPartProperty.Create(part, "AESC_Construct", "Plates");
            CustomPartProperty.Create(part, "Type", type);
            CustomPartProperty.Create(part, "L1", L1);
            CustomPartProperty.Create(part, "B1", B1);
            CustomPartProperty.Create(part, "T", T);
            CustomPartProperty.Create(part, "L2", L2);
            CustomPartProperty.Create(part, "Lnr", Lnr);
            CustomPartProperty.Create(part, "B2", B2);
            CustomPartProperty.Create(part, "Bnr", Bnr);
            CustomPartProperty.Create(part, "Diam", Diam);
            CustomPartProperty.Create(part, "Rad", Rad);
            CustomPartProperty.Create(part, "Angle", Angle);

        }
        private static bool CompareCustomProperties(
            Part part,
            string type,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam)
        {
            // Shorthand to the custom‐properties dictionary
            var props = part.CustomProperties;
            string name = part.DisplayName;

            // 1) Marker must exist and must == "Plates"
            if (!props.TryGetValue("AESC_Construct", out var markerProp) ||
                markerProp.Value.ToString() != "Plates")
                return false;

            // 2) Type must exist and match
            if (!props.TryGetValue("Type", out var typeProp) ||
                typeProp.Value.ToString() != type)
                return false;

            // 3) Numeric values to check
            var expected = new Dictionary<string, double>
            {
                ["L1"] = L1,
                ["L2"] = L2,
                ["Lnr"] = Lnr,      // integers treated as doubles
                ["B1"] = B1,
                ["B2"] = B2,
                ["Bnr"] = Bnr,
                ["T"] = T,
                ["Diam"] = Diam,
                ["Rad"] = Rad
            };

            const double tol = 1e-6;

            // 4) Ensure each key exists, parses to double, and is within tolerance
            bool allMatch = expected.All(kv =>
            {
                if (!props.TryGetValue(kv.Key, out var prop))
                    return false;
                if (!double.TryParse(prop.Value.ToString(), out var actual))
                    return false;
                return Math.Abs(actual - kv.Value) <= tol;
            });

            return allMatch;
        }


    }
}
