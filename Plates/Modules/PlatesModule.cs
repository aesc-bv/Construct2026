using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct25.Plates.Modules
{
    internal class PlatesModule
    {
        public static void CreatePlateFromUI(
            string type, string sizeName, double angleDeg,
            double L1, double L2, int count1,
            double B1, double B2, int count2,
            double thickness, double filletRadius, double holeDiameter,
            bool insertPlateMid = false)
        {
            try
            {
                createPart(type, sizeName, angleDeg,
                           L1, L2, count1,
                           B1, B2, count2,
                           thickness, filletRadius, holeDiameter,
                           insertPlateMid);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating plate: {ex.Message}", "Create Plate Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static void createPart(
            string type, string name,
            double angleDeg,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam,
            bool insertPlateMid)
        {
            var win = Window.ActiveWindow;
            if (win == null)
            {
                MessageBox.Show("No active window found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            var doc = win.Document;
            var selection = win.ActiveContext.Selection;

            if (selection.Count == 0)
            {
                MessageBox.Show("No selection found. Please select a face.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (var obj in selection)
            {
                try
                {
                    DesignFace df = obj as DesignFace;
                    var matrix = Matrix.Identity;

                    if (obj is IDesignFace idf)
                    {
                        df = idf.Master;
                        matrix = idf.TransformToMaster;
                    }
                    if (df == null)
                    {
                        MessageBox.Show("Selected object is not a valid face.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        continue;
                    }

                    var plane = df.Shape.Geometry as Plane;
                    if (plane == null)
                    {
                        MessageBox.Show("Selected face is not planar.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        continue;
                    }

                    var rawPt = (Point)win.ActiveContext.GetSelectionPoint(selection.First());
                    var selPoint = plane.ProjectPoint(rawPt).Point;

                    var dirZ = df.Shape.IsReversed
                        ? -df.Shape.ProjectPoint(selPoint).Normal
                        : df.Shape.ProjectPoint(selPoint).Normal;

                    if (insertPlateMid)
                    {
                        var bb = df.Shape.GetBoundingBox(matrix.Inverse, true);
                        selPoint = bb.Center;
                    }
                    else
                    {
                        var bb = df.Shape.GetBoundingBox(matrix.Inverse, true);
                        selPoint = Point.Create(selPoint.X, selPoint.Y, bb.Center.Z);
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
                catch (Exception ex)
                {
                    MessageBox.Show($"Error processing selection: {ex.Message}", "Selection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private static bool ValidateParams(
            string type,
            double L1, double B1, double T,
            double radius,
            double L2, int Lnr,
            double B2, int Bnr,
            double holeDiam)
        {
            if (L2 > L1)
            {
                MessageBox.Show($"Hole pitch X (L2={L2}) > plate L1={L1}", "Parameter Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            if (B2 > B1)
            {
                MessageBox.Show($"Hole pitch Y (B2={B2}) > plate B1={B1}", "Parameter Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (holeDiam > 0 && Lnr > 1)
            {
                var spacingX = L2 / (Lnr - 1);
                if (spacingX < holeDiam)
                {
                    MessageBox.Show($"Holes too close in X (spacing {spacingX} < diam {holeDiam})", "Parameter Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            if (holeDiam > 0 && Bnr > 1)
            {
                var spacingY = B2 / (Bnr - 1);
                if (spacingY < holeDiam)
                {
                    MessageBox.Show($"Holes too close in Y (spacing {spacingY} < diam {holeDiam})", "Parameter Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }

            return true;
        }

        private static Part CheckIfPartExists(
            string type, double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam)
        {
            var doc = Window.ActiveWindow.Document;
            foreach (var p in doc.Parts)
                if (CompareCustomProperties(p, type, L1, L2, Lnr, B1, B2, Bnr, T, Rad, Diam))
                    return p;
            return null;
        }

        private static Body CreatePlateBody(
            string type,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam
        )
        {
            if (!ValidateParams(type, L1, B1, T, Rad, L2, Lnr, B2, Bnr, Diam))
                return null;

            // base coordinate frame
            var plane = Plane.PlaneXY;
            var origin = Point.Origin;
            var dirX = plane.Frame.DirX;
            var dirY = plane.Frame.DirY;
            var dirZ = plane.Frame.DirZ;

            Body body = null;

            try
            {
                if (type.Contains("support"))
                {
                    // build the exact 6‐segment V18 contour
                    var segs = getBaseCapContour(origin, dirX, dirY, type, L1, B1, Rad);
                    if (segs.Count == 6 && T > 0)
                    {
                        var supportPlane = Plane.Create(Frame.Create(origin + 0.5 * T * dirY, dirX, dirZ));
                        body = Body.ExtrudeProfile(new Profile(supportPlane, segs), T);
                    }
                    else
                    {
                        MessageBox.Show("Invalid contour count or thickness for support plate.", "Plate Creation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return null;
                    }
                }
                else if (type.Contains("flange"))
                {
                    if (T <= 0)
                    {
                        MessageBox.Show("Flange thickness must be positive.", "Plate Creation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return null;
                    }

                    // outer disk
                    var holePlane = plane.CreateTransformedCopy(Matrix.CreateTranslation(origin - plane.Frame.Origin));
                    body = Body.ExtrudeProfile(new CircleProfile(holePlane, 0.5 * L1), T);

                    // flat‐flange inner cut
                    if (type == "Flat flange" && B1 > 0)
                    {
                        var cut = Body.ExtrudeProfile(new CircleProfile(holePlane, 0.5 * B1), 1.1 * T);
                        body.Subtract(cut);
                    }

                    // bolt holes (always on diameter L1, never in the middle)
                    if (Diam > 0 && Lnr > 0)
                    {
                        var circ = new FrameGenerator.Modules.Profiles.CircularProfile(Diam, 0.0, false, 0, 0);
                        var loop = circ.GetProfileCurves(holePlane).ToList();
                        var prof = new Profile(holePlane, loop);
                        var cutter = Body.ExtrudeProfile(prof, 1.1 * T);
                        double step = 2 * Math.PI / Lnr;
                        double radius = 0.5 * L2; // Always use L1 as the bolt circle diameter
                        for (int i = 0; i < Lnr; i++)
                        {
                            // Place each hole on the bolt circle (diameter L1)
                            double angle = i * step;
                            var x = radius * Math.Cos(angle);
                            var y = radius * Math.Sin(angle);
                            var copy = cutter.Copy();
                            copy.Transform(Matrix.CreateTranslation(Vector.Create(x, y, 0)));
                            body.Subtract(copy);
                        }
                        cutter.Dispose();
                    }
                }
                else if (type == "UNP")
                {
                    // rectangular UNP with top‐corner chamfers
                    body = createUnpRectangle(
                        width: L1,
                        height: B1,
                        thickness: T,
                        chamfer: Rad,
                        holeDiam: Diam,
                        holePitchX: L2, holeCountX: Lnr,
                        holePitchY: B2, holeCountY: Bnr
                    );
                }
                else
                {
                    // Logger.Log("[CreatePlateBody] Branch: rectangle");
                    body = createRoundedRectangle(L1, L2, Lnr, B1, B2, Bnr, T, Rad, Diam);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating plate body: {ex.Message}", "Plate Creation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            return body;
        }

        //— base/cap/UNP contour helper —
        /// <summary>
        /// Returns exactly six trimmed‐curves for “support” / base / cap shapes,
        /// ported from your V18 getContour logic (no rotation here).
        /// </summary>
        private static List<ITrimmedCurve> getBaseCapContour(
            Point p,
            Direction dX,
            Direction dY,
            string type,
            double L1,
            double B1,
            double R
        )
        {
            var dZ = Direction.Cross(dX, dY);
            var list = new List<ITrimmedCurve>();

            if (type.Contains("support") || type.Contains("base") || type.Contains("cap") || type == "UNP")
            {
                // Re‐use exactly the six segments from your V18 code:
                var p1 = p + (-0.5 * L1 + R) * dX + (0.5 * B1) * dY;
                var p2 = p1 + (L1 - 2 * R) * dX;
                var p5 = p1 - R * dX + R * dZ;
                var p5a = p5 + (B1 - R) * dZ;
                var p7 = p2 + R * dX + R * dZ;
                var p7a = p7 + (B1 - R) * dZ;

                list.Add(CurveSegment.Create(p1, p2));
                list.Add(CurveSegment.Create(p2, p7));
                list.Add(CurveSegment.Create(p7, p7a));
                list.Add(CurveSegment.Create(p7a, p5a));
                list.Add(CurveSegment.Create(p5a, p5));
                list.Add(CurveSegment.Create(p5, p1));
            }

            return list;
        }

        private static Body createRoundedRectangle(
            double length, double length2, int nrL,
            double width, double width2, int nrW,
            double thickness, double radius, double holeDiam
        )
        {
            // Logger.Log("[createRoundedRectangle] Enter");

            // build the 2D profile curves:
            var boundary = CreateRectangleCurves(
                Plane.PlaneXY,
                width: length,
                height: width,
                radius: radius
            );
            // Logger.Log($"[createRoundedRectangle] boundary.Count = {boundary.Count}");

            // extrude the plate
            var plane = Plane.PlaneXY;
            var body = Body.ExtrudeProfile(new Profile(plane, boundary), thickness);
            // Logger.Log("[createRoundedRectangle] extruded plate");

            // drill holes if requested
            if (holeDiam > 0 && nrL > 0 && nrW > 0)
            {
                // Logger.Log("[createRoundedRectangle] adding holes");
                var cutter = Body.ExtrudeProfile(
                    new CircleProfile(plane, holeDiam * 0.5),
                    1.1 * thickness
                );

                for (int i = 0; i < nrL; i++)
                    for (int j = 0; j < nrW; j++)
                    {
                        double dx = -0.5 * length2 + (nrL > 1 ? i * (length2 / (nrL - 1)) : 0);
                        double dy = -0.5 * width2 + (nrW > 1 ? j * (width2 / (nrW - 1)) : 0);

                        var copy = cutter.Copy();
                        copy.Transform(Matrix.CreateTranslation(Vector.Create(dx, dy, 0)));

                        try
                        {
                            body.Subtract(copy);
                            // Logger.Log($"[createRoundedRectangle] subtracted hole at ({dx:0.###}, {dy:0.###})");
                        }
                        catch
                        {
                            copy.Dispose();
                            // Logger.Log("[createRoundedRectangle] hole subtract failed, disposing cutter copy");
                        }
                    }
                cutter.Dispose();
            }

            // Logger.Log("[createRoundedRectangle] Exit");
            return body;
        }

        private static List<ITrimmedCurve> CreateRectangleCurves(
            Plane profilePlane,
            double width,
            double height,
            double radius
        )
        {
            var curves = new List<ITrimmedCurve>();
            var frame = profilePlane.Frame;
            var center = frame.Origin;
            var dx = frame.DirX;
            var dy = frame.DirY;
            var dz = frame.DirZ;

            // Logger.Log($"[CreateRectangleCurves] radius = {radius}");

            // corner‐centers
            var m1 = center + (-width / 2 + radius) * dx + (height / 2 - radius) * dy; // top‐left
            var m2 = center + (width / 2 - radius) * dx + (height / 2 - radius) * dy; // top‐right
            var m3 = center + (width / 2 - radius) * dx + (-height / 2 + radius) * dy; // bottom‐right
            var m4 = center + (-width / 2 + radius) * dx + (-height / 2 + radius) * dy; // bottom‐left

            // arc endpoints
            var p1 = center + (-width / 2 + radius) * dx + (height / 2) * dy;
            var p2 = center + (width / 2 - radius) * dx + (height / 2) * dy;
            var p3 = center + (width / 2) * dx + (height / 2 - radius) * dy;
            var p4 = center + (width / 2) * dx + (-height / 2 + radius) * dy;
            var p5 = center + (width / 2 - radius) * dx + (-height / 2) * dy;
            var p6 = center + (-width / 2 + radius) * dx + (-height / 2) * dy;
            var p7 = center + (-width / 2) * dx + (-height / 2 + radius) * dy;
            var p8 = center + (-width / 2) * dx + (height / 2 - radius) * dy;

            // straight edges & optional arcs
            curves.Add(CurveSegment.Create(p1, p2));                          // top edge
            if (radius > 0) curves.Add(CurveSegment.CreateArc(m2, p2, p3, -dz)); // top‐right
            curves.Add(CurveSegment.Create(p3, p4));                          // right edge
            if (radius > 0) curves.Add(CurveSegment.CreateArc(m3, p4, p5, -dz)); // bottom‐right
            curves.Add(CurveSegment.Create(p5, p6));                          // bottom edge
            if (radius > 0) curves.Add(CurveSegment.CreateArc(m4, p6, p7, -dz)); // bottom‐left
            curves.Add(CurveSegment.Create(p7, p8));                          // left edge
            if (radius > 0) curves.Add(CurveSegment.CreateArc(m1, p8, p1, -dz)); // top‐left

            // Logger.Log($"[CreateRectangleCurves] returned {curves.Count} curves");
            return curves;
        }

        private static Body createUnpRectangle(
            double width,
            double height,
            double thickness,
            double chamfer,
            double holeDiam,
            double holePitchX, int holeCountX,
            double holePitchY, int holeCountY
        )
        {
            var plane = Plane.PlaneXY;
            var o = Point.Origin;
            var dx = plane.Frame.DirX;
            var dy = plane.Frame.DirY;
            var dz = plane.Frame.DirZ;

            double halfW = width / 2;
            double halfH = height / 2;

            // For a 45-degree chamfer, a = b = chamfer * sin(45°) = chamfer / sqrt(2)
            double a = chamfer / Math.Sqrt(2);
            double b = chamfer / Math.Sqrt(2);

            // corner points, going CCW from bottom-left
            var p0 = o + (-halfW) * dx + (-halfH) * dy;                 // BL
            var p1 = o + (halfW) * dx + (-halfH) * dy;                  // BR
            var p2 = o + (halfW) * dx + (halfH - b) * dy;               // just below TR chamfer start
            var p3 = o + (halfW - a) * dx + (halfH) * dy;               // chamfer meet point
            var p4 = o + (-halfW + a) * dx + (halfH) * dy;              // just right of TL chamfer start
            var p5 = o + (-halfW) * dx + (halfH - b) * dy;              // TL chamfer meet

            // build the poly with two chamfers
            var boundary = new List<ITrimmedCurve> {
                CurveSegment.Create(p0, p1),   // bottom
                CurveSegment.Create(p1, p2),   // right vertical
                CurveSegment.Create(p2, p3),   // TR chamfer
                CurveSegment.Create(p3, p4),   // top between chamfers
                CurveSegment.Create(p4, p5),   // TL chamfer
                CurveSegment.Create(p5, p0),   // left vertical
            };

            // extrude
            var body = Body.ExtrudeProfile(new Profile(plane, boundary), thickness);

            // optional holes
            if (holeDiam > 0 && holeCountX > 0 && holeCountY > 0)
            {
                var cutter = Body.ExtrudeProfile(
                    new CircleProfile(plane, holeDiam * 0.5),
                    1.1 * thickness
                );
                for (int i = 0; i < holeCountX; i++)
                    for (int j = 0; j < holeCountY; j++)
                    {
                        double dx2 = -0.5 * holePitchX + (holeCountX > 1 ? i * (holePitchX / (holeCountX - 1)) : 0);
                        double dy2 = -0.5 * holePitchY + (holeCountY > 1 ? j * (holePitchY / (holeCountY - 1)) : 0);
                        var copy = cutter.Copy();
                        copy.Transform(Matrix.CreateTranslation(Vector.Create(dx2, dy2, 0)));
                        try
                        {
                            body.Subtract(copy);
                        }
                        catch
                        {
                            copy.Dispose();
                        }
                    }
                cutter.Dispose();
            }

            return body;
        }



        private static void createPlate(
            Point p, Plane pl, Direction dZ,
            string type, string name, double ang,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double R, double D,
            bool insertMid)
        {
            // Logger.Log($"[createPlate] Enter: placing '{type}_{name}'");
            var doc = Window.ActiveWindow.Document;
            var main = doc.MainPart;
            var dX = pl.Frame.DirX;
            var dY = pl.Frame.DirY;
            var pn = $"{type}_{name}";

            try
            {
                WriteBlock.ExecuteTask("Create part", () =>
                {
                    var plates = doc.Parts.FirstOrDefault(p => p.DisplayName == "Plates")
                                  ?? Part.Create(doc, "Plates")
                                     .Also(pp => Component.Create(main, pp));

                    var exist = CheckIfPartExists(type, L1, L2, Lnr, B1, B2, Bnr, T, R, D);
                    Component comp;
                    if (exist != null)
                    {
                        // Logger.Log($"[createPlate] Reusing '{exist.DisplayName}'");
                        comp = Component.Create(plates, exist);
                        comp.Transform(
                            Matrix.CreateTranslation(
                                Vector.Create(0.0, 0.0, 0.0)
                            )
                        );
                    }
                    else
                    {
                        // Logger.Log("[createPlate] Building new plate body");
                        var b = CreatePlateBody(type, L1, L2, Lnr, B1, B2, Bnr, T, R, D);
                        if (b == null)
                        {// Logger.Log("[createPlate] abort body==null");
                            return;
                        }
                        var pb = Part.Create(doc, pn);
                        DesignBody.Create(pb, pn, b);
                        comp = Component.Create(plates, pb);
                        CreateCustomPartPropertiesPlate(pb, type, L1, B1, T, L2, Lnr, B2, Bnr, D, R, ang);
                    }

                    var frame = Frame.Create(p, dX, dY);
                    if (frame.DirZ != dZ) frame = Frame.Create(p, -dX, dY);
                    comp.Transform(Matrix.CreateMapping(frame));
                    // Logger.Log("[createPlate] moved into place");

                    if (ang != 0)
                    {
                        comp.Transform(Matrix.CreateRotation(Line.Create(p, dZ), ang * Math.PI / 180));
                        // Logger.Log($"[createPlate] rotated {ang}°");
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating plate component: {ex.Message}", "Plate Placement Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static void CreateCustomPartPropertiesPlate(
            Part part, string type,
            double L1, double B1, double T,
            double L2, int Lnr,
            double B2, int Bnr,
            double Diam, double Rad,
            double Angle)
        {
            try
            {
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
            catch (Exception ex)
            {
                MessageBox.Show($"Error setting custom properties: {ex.Message}", "Property Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static bool CompareCustomProperties(
            Part part, string type,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam)
        {
            var props = part.CustomProperties;
            if (!props.TryGetValue("AESC_Construct", out var m) || m.Value.ToString() != "Plates")
                return false;
            if (!props.TryGetValue("Type", out var tp) || tp.Value.ToString() != type)
                return false;
            var exp = new Dictionary<string, double>
            {
                ["L1"] = L1,
                ["L2"] = L2,
                ["Lnr"] = Lnr,
                ["B1"] = B1,
                ["B2"] = B2,
                ["Bnr"] = Bnr,
                ["T"] = T,
                ["Diam"] = Diam,
                ["Rad"] = Rad
            };
            const double tol = 1e-6;
            foreach (var kv in exp)
            {
                if (!props.TryGetValue(kv.Key, out var pr))
                    return false;
                if (!double.TryParse(pr.Value.ToString(), out var a))
                    return false;
                if (Math.Abs(a - kv.Value) > tol)
                    return false;
            }
            return true;
        }
    }

    internal static class Extensions
    {
        public static T Also<T>(this T self, Action<T> act) { act(self); return self; }
    }
}