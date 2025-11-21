/*
 PlatesModule contains all geometry and part-creation logic for standard plates.
 It validates UI parameters, reuses or creates plate parts with custom properties,
 and inserts them as components on selected planar faces in the active SpaceClaim document.
*/

using AESCConstruct25.FrameGenerator.Utilities;
using DocumentFormat.OpenXml.Wordprocessing;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Application = SpaceClaim.Api.V242.Application;
using Body = SpaceClaim.Api.V242.Modeler.Body;
using Frame = SpaceClaim.Api.V242.Geometry.Frame;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct25.Plates.Modules
{
    internal class PlatesModule
    {
        // Entry point for the UI: validates and forwards plate parameters to createPart and reports errors to the user.
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
                Application.ReportStatus($"Error creating plate: {ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Creates plate components on the selected planar faces in the active window using the given dimensions and options.
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
                Application.ReportStatus("No active window found.", StatusMessageType.Error, null);
                return;
            }
            var doc = win.Document;
            var selection = win.ActiveContext.Selection;

            if (selection.Count == 0)
            {
                Application.ReportStatus("No selection found. Please select a face.", StatusMessageType.Warning, null);
                return;
            }

            WriteBlock.ExecuteTask("Create part", () =>
            {
                // 1) Expand the current selection into face jobs (face + its transform to master)
                var jobs = ExpandSelectionToPlanarFaces(Window.ActiveWindow.ActiveContext.Selection);

                if (jobs.Count == 0)
                {
                    Application.ReportStatus("No planar faces found in selection.", StatusMessageType.Warning, null);
                    return;
                }

                foreach (var (df, toMaster, src) in jobs)
                {
                    try
                    {
                        // --- master geometry & plane
                        var plane = df.Shape.Geometry as Plane;
                        if (plane == null) continue; // filtered already, but double-safeguard

                        // Try to get the user click for *this* source; fall back to face center
                        var bbM = df.Shape.GetBoundingBox(Matrix.Identity, true);
                        Point selPointWorld = bbM.Center;
                        try
                        {
                            if (src is IDocObject srcObj)
                            {
                                var sp = Window.ActiveWindow.ActiveContext.GetSelectionPoint(srcObj);
                                if (sp is Point pw) selPointWorld = pw;
                            }
                        }
                        catch { /* box/whole-body case -> keep center */ }

                        // Project to this face plane (MASTER)
                        var selOnPlane = plane.ProjectPoint(selPointWorld).Point;

                        // Choose placement point
                        Point pMaster = insertPlateMid
                            ? bbM.Center
                            : Point.Create(selOnPlane.X, selOnPlane.Y, bbM.Center.Z);

                        // Face normal (respect “reversed”)
                        var nMaster = df.Shape.ProjectPoint(selOnPlane).Normal;
                        if (df.Shape.IsReversed) nMaster = -nMaster;

                        // If this face came from an occurrence, move plane/normal/point into that local space
                        if (!toMaster.IsIdentity)
                        {
                            plane = plane.CreateTransformedCopy(toMaster.Inverse);
                            nMaster = toMaster.Inverse * nMaster;
                            pMaster = toMaster.Inverse * pMaster;
                        }

                        // Create the plate on this face
                        createPlate(
                            pMaster, plane, nMaster,
                            type, name, angleDeg,
                            L1, L2, Lnr,
                            B1, B2, Bnr,
                            T, Rad, Diam,
                            insertPlateMid
                        );
                    }
                    catch (Exception ex)
                    {
                        Application.ReportStatus($"Error on a selected face: {ex.Message}", StatusMessageType.Error, null);
                    }
                }
            });
        }

        // Expands the selection into a list of planar DesignFaces with their transform to master and source object.
        private static List<(DesignFace df, Matrix toMaster, object src)>
        ExpandSelectionToPlanarFaces(ICollection<IDocObject> sel)
        {
            var outList = new List<(DesignFace df, Matrix toMaster, object src)>();

            bool AlreadyAdded(DesignFace f, Matrix m) =>
                outList.Any(t => ReferenceEquals(t.df, f) && t.toMaster.Equals(m));

            foreach (var obj in sel)
            {
                // Occurrence face
                if (obj is IDesignFace idf)
                {
                    var df = idf.Master;
                    if (df?.Shape?.Geometry is Plane && !AlreadyAdded(df, idf.TransformToMaster))
                        outList.Add((df, idf.TransformToMaster, obj));
                    continue;
                }

                // Master face (rare, but allow)
                if (obj is DesignFace dfM)
                {
                    if (dfM?.Shape?.Geometry is Plane && !AlreadyAdded(dfM, Matrix.Identity))
                        outList.Add((dfM, Matrix.Identity, obj));
                    continue;
                }

                // Occurrence body → expand to its planar faces
                if (obj is IDesignBody idb)
                {
                    var master = idb.Master;
                    if (master?.Shape != null)
                    {
                        foreach (var f in master.Faces)
                            if (f?.Shape?.Geometry is Plane && !AlreadyAdded(f, idb.TransformToMaster))
                                outList.Add((f, idb.TransformToMaster, obj));
                    }
                    continue;
                }

                // Master body → expand to its planar faces
                if (obj is DesignBody db)
                {
                    foreach (var f in db.Faces)
                        if (f?.Shape?.Geometry is Plane && !AlreadyAdded(f, Matrix.Identity))
                            outList.Add((f, Matrix.Identity, obj));
                    continue;
                }

                // ignore edges, vertices, etc.
            }

            return outList;
        }

        // Validates plate dimensions and hole layout for basic geometric consistency and spacing constraints.
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
                Application.ReportStatus($"Hole pitch X (L2={L2}) > plate L1={L1}", StatusMessageType.Error, null);
                return false;
            }
            if (B2 > B1)
            {
                Application.ReportStatus($"Hole pitch Y (B2={B2}) > plate B1={B1}", StatusMessageType.Error, null);
                return false;
            }

            if (holeDiam > 0 && Lnr > 1)
            {
                var spacingX = L2 / (Lnr - 1);
                if (spacingX < holeDiam)
                {
                    Application.ReportStatus($"Holes too close in X (spacing {spacingX} < diam {holeDiam})", StatusMessageType.Error, null);
                    return false;
                }
            }
            if (holeDiam > 0 && Bnr > 1)
            {
                var spacingY = B2 / (Bnr - 1);
                if (spacingY < holeDiam)
                {
                    Application.ReportStatus($"Holes too close in Y (spacing {spacingY} < diam {holeDiam})", StatusMessageType.Error, null);
                    return false;
                }
            }

            return true;
        }

        // Checks the current document for an existing plate Part with matching custom properties and returns it if found.
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

        // Builds the solid plate Body geometry for the given type using dimensions and hole parameters.
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
                    // SUPPORT: same geometry as Base/Cap (rounded rectangle + holes) but UPRIGHT
                    if (T <= 0)
                    {
                        Application.ReportStatus("Support plate thickness must be positive.", StatusMessageType.Error, null);
                        return null;
                    }

                    var uprightPlane = Plane.Create(Frame.Create(origin + (0.5 * T * dirY) + (0.5 * B1 * dirZ), dirX, dirZ));
                    body = createRoundedRectangle(
                        uprightPlane,
                        length: L1, length2: L2, nrL: Lnr,
                        width: B1, width2: B2, nrW: Bnr,
                        thickness: T, radius: Rad, holeDiam: Diam
                    );
                }
                else if (type.Contains("flange"))
                {
                    if (T <= 0)
                    {
                        Application.ReportStatus("Flange thickness must be positive.", StatusMessageType.Error, null);
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
                    if (T <= 0)
                    {
                        Application.ReportStatus("UNP thickness must be positive.", StatusMessageType.Error, null);
                        return null;
                    }

                    var uprightPlane = Plane.Create(
                        Frame.Create(origin + (0.5 * T * dirY) + (0.5 * B1 * dirZ), dirX, dirZ)
                    );

                    // Make contour coplanar with uprightPlane (puts TL/TR at y = +0.5*T)
                    var pContour = origin + (0.5 * T - 0.5 * B1) * dirY;

                    var segs = getBaseCapContour(pContour, dirX, dirY, type, L1, B1, Rad);
                    if (segs == null || (segs.Count != 4 && segs.Count != 6))
                    {
                        Application.ReportStatus("UNP contour could not be built.", StatusMessageType.Error, null);
                        return null;
                    }

                    body = Body.ExtrudeProfile(new Profile(uprightPlane, segs), T);
                }

                else
                {
                    body = createRoundedRectangle(L1, L2, Lnr, B1, B2, Bnr, T, Rad, Diam);
                }
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Error creating plate body: {ex.Message}", StatusMessageType.Error, null);
                return null;
            }

            return body;
        }

        // Builds the base/cap contour (optionally chamfered) for UNP/support plates as a list of trimmed curves.
        private static List<ITrimmedCurve> getBaseCapContour(
            Point p, Direction dX, Direction dY,
            string type, double L1, double B1, double R)
        {
            var dZ = Direction.Cross(dX, dY);
            var list = new List<ITrimmedCurve>();

            // Corners on the "top" edge (z = 0 in this sketch)
            var TL = p + (-0.5 * L1) * dX + (0.5 * B1) * dY;
            var TR = p + (0.5 * L1) * dX + (0.5 * B1) * dY;

            // Bottom corners
            var BL = TL + B1 * dZ;
            var BR = TR + B1 * dZ;

            // If no chamfer: plain rectangle with exactly 4 edges
            if (R <= 1e-9)
            {
                list.Add(CurveSegment.Create(TL, TR)); // top
                list.Add(CurveSegment.Create(TR, BR)); // right
                list.Add(CurveSegment.Create(BR, BL)); // bottom
                list.Add(CurveSegment.Create(BL, TL)); // left
                return list;
            }

            // Chamfer hypotenuse R -> 45° legs
            double a = R / Math.Sqrt(2.0); // along X
            double b = R / Math.Sqrt(2.0); // along Z
            a = Math.Min(a, 0.5 * L1);
            b = Math.Min(b, B1);

            // Top-corner chamfers (opposite to your previous bottom-corner version)
            var TL_topRight = TL + a * dX;     // along top edge to the right
            var TL_down = TL + b * dZ;     // down from TL
            var TR_topLeft = TR - a * dX;     // along top edge to the left
            var TR_down = TR + b * dZ;     // down from TR

            // 6 straight segments CCW:
            list.Add(CurveSegment.Create(TL_topRight, TR_topLeft)); // top between chamfers
            list.Add(CurveSegment.Create(TR_topLeft, TR_down));    // top-right chamfer
            list.Add(CurveSegment.Create(TR_down, BR));         // right vertical
            list.Add(CurveSegment.Create(BR, BL));         // bottom
            list.Add(CurveSegment.Create(BL, TL_down));    // left vertical
            list.Add(CurveSegment.Create(TL_down, TL_topRight));// top-left chamfer

            return list;
        }

        // Creates a rounded rectangle plate body on an arbitrary plane, including a hole grid if requested.
        private static Body createRoundedRectangle(
            Plane profilePlane,
            double length, double length2, int nrL,
            double width, double width2, int nrW,
            double thickness, double radius, double holeDiam
        )
        {
            // Build the 2D profile on the provided plane
            var boundary = CreateRectangleCurves(
                profilePlane,
                width: length,
                height: width,
                radius: radius
            );

            // Extrude the plate along the plane normal
            var body = Body.ExtrudeProfile(new Profile(profilePlane, boundary), thickness);

            // Drill holes (if requested) in the SAME plane
            if (holeDiam > 0 && nrL > 0 && nrW > 0)
            {
                var cutter = Body.ExtrudeProfile(
                    new CircleProfile(profilePlane, holeDiam * 0.5),
                    1.1 * thickness
                );

                for (int i = 0; i < nrL; i++)
                    for (int j = 0; j < nrW; j++)
                    {
                        double dx = -0.5 * length2 + (nrL > 1 ? i * (length2 / (nrL - 1)) : 0);
                        double dy = -0.5 * width2 + (nrW > 1 ? j * (width2 / (nrW - 1)) : 0);

                        var copy = cutter.Copy();
                        copy.Transform(Matrix.CreateTranslation(Vector.Create(dx, dy, 0)));
                        try { body.Subtract(copy); }
                        catch { copy.Dispose(); }
                    }
                cutter.Dispose();
            }

            return body;
        }

        // Legacy convenience overload: creates a rounded rectangle on PlaneXY, keeping existing callers working.
        private static Body createRoundedRectangle(
            double length, double length2, int nrL,
            double width, double width2, int nrW,
            double thickness, double radius, double holeDiam
        )
        {
            return createRoundedRectangle(
                Plane.PlaneXY,
                length, length2, nrL,
                width, width2, nrW,
                thickness, radius, holeDiam
            );
        }

        // Builds the rectangular (optionally corner-rounded) 2D contour for a plate on a given plane.
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

            return curves;
        }

        // Creates or reuses a plate component for one face, positions and rotates it according to the face and angle.
        private static void createPlate(
            Point p, Plane pl, Direction dZ,
            string type, string name, double ang,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double R, double D,
            bool insertMid)
        {
            var doc = Window.ActiveWindow.Document;
            var main = doc.MainPart;
            var dX = pl.Frame.DirX;
            var dY = pl.Frame.DirY;
            var pn = $"{type}_{name}";

            try
            {
                //WriteBlock.ExecuteTask("Create part", () =>
                //{
                var plates = doc.Parts.FirstOrDefault(p => p.DisplayName == "Plates")
                              ?? Part.Create(doc, "Plates")
                                 .Also(pp => Component.Create(main, pp));

                var exist = CheckIfPartExists(type, L1, L2, Lnr, B1, B2, Bnr, T, R, D);
                Component comp;
                if (exist != null)
                {
                    comp = Component.Create(plates, exist);
                    comp.Transform(
                        Matrix.CreateTranslation(
                            Vector.Create(0.0, 0.0, 0.0)
                        )
                    );
                }
                else
                {
                    var b = CreatePlateBody(type, L1, L2, Lnr, B1, B2, Bnr, T, R, D);
                    if (b == null)
                    {
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

                if (type != "UNP" && !type.Contains("support"))
                {
                    if (ang != 0)
                        comp.Transform(Matrix.CreateRotation(Line.Create(p, dZ), ang * Math.PI / 180));
                }
                else
                {
                    comp.Transform(Matrix.CreateRotation(Line.Create(p, dZ), (ang + 90) * Math.PI / 180));
                }
                //});
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Error creating plate component: {ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Adds custom properties to a plate part so it can be identified and reused by geometry parameters.
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
                Application.ReportStatus($"Error setting custom properties: {ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Compares a part's custom properties against the requested plate parameters to determine a match.
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
        // Fluent helper that executes an action on an object and then returns the same instance.
        public static T Also<T>(this T self, Action<T> act) { act(self); return self; }
    }
}
