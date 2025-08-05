using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Linq;

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
            Logger.Log($"[CreatePlateFromUI] Enter: type={type}, sizeName={sizeName}, angleDeg={angleDeg}, " +
                       $"L1={L1},L2={L2},count1={count1},B1={B1},B2={B2},count2={count2},thickness={thickness},filletRadius={filletRadius},holeDiameter={holeDiameter},insertPlateMid={insertPlateMid}");
            createPart(type, sizeName, angleDeg,
                       L1, L2, count1,
                       B1, B2, count2,
                       thickness, filletRadius, holeDiameter,
                       insertPlateMid);
            Logger.Log("[CreatePlateFromUI] Exit");
        }

        private static void createPart(
            string type, string name,
            double angleDeg,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam,
            bool insertPlateMid)
        {
            Logger.Log($"[createPart] Enter: type={type}, name={name}");
            var win = Window.ActiveWindow;
            var doc = win.Document;

            var selection = win.ActiveContext.Selection;
            Logger.Log($"[createPart] Selection count = {selection.Count}");

            foreach (var obj in selection)
            {
                Logger.Log($"[createPart] Processing selection object of type {obj.GetType().Name}");
                DesignFace df = obj as DesignFace;
                var matrix = Matrix.Identity;

                if (obj is IDesignFace idf)
                {
                    df = idf.Master;
                    matrix = idf.TransformToMaster;
                    Logger.Log("[createPart] Transformed IDesignFace to master");
                }
                if (df == null)
                {
                    Logger.Log("[createPart] Skipping: not a DesignFace");
                    continue;
                }

                var plane = df.Shape.Geometry as Plane;
                if (plane == null)
                {
                    Logger.Log("[createPart] Skipping: face geometry not Plane");
                    continue;
                }

                var rawPt = (Point)win.ActiveContext.GetSelectionPoint(selection.First());
                var selPoint = plane.ProjectPoint(rawPt).Point;
                Logger.Log($"[createPart] selPoint = {selPoint}");

                var dirZ = df.Shape.IsReversed
                    ? -df.Shape.ProjectPoint(selPoint).Normal
                    : df.Shape.ProjectPoint(selPoint).Normal;

                if (insertPlateMid)
                {
                    var bb = df.Shape.GetBoundingBox(matrix.Inverse, true);
                    selPoint = bb.Center;
                    Logger.Log($"[createPart] insertPlateMid: selPoint set to bounding-box center {selPoint}");
                }
                if (!matrix.IsIdentity)
                {
                    plane = plane.CreateTransformedCopy(matrix.Inverse);
                    dirZ = matrix.Inverse * dirZ;
                    Logger.Log("[createPart] Applied inverse transform");
                }

                createPlate(selPoint, plane, dirZ,
                    type, name, angleDeg,
                    L1, L2, Lnr,
                    B1, B2, Bnr,
                    T, Rad, Diam,
                    insertPlateMid);
            }

            Logger.Log("[createPart] Exit");
        }

        private static Part CheckIfPartExists(
            string type, double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam)
        {
            Logger.Log($"[CheckIfPartExists] Enter: type={type}, L1={L1},L2={L2},Lnr={Lnr},B1={B1},B2={B2},Bnr={Bnr},T={T},Rad={Rad},Diam={Diam}");
            var doc = Window.ActiveWindow.Document;
            foreach (var p in doc.Parts)
                if (CompareCustomProperties(p, type, L1, L2, Lnr, B1, B2, Bnr, T, Rad, Diam))
                {
                    Logger.Log($"[CheckIfPartExists] Found existing part: {p.DisplayName}");
                    return p;
                }
            Logger.Log("[CheckIfPartExists] No matching part found");
            return null;
        }

        private static Body CreatePlateBody(
            string type,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam
        )
        {
            Logger.Log($"[CreatePlateBody] Enter: type={type}, L1={L1}, L2={L2}, Lnr={Lnr}, B1={B1}, B2={B2}, Bnr={Bnr}, T={T}, Rad={Rad}, Diam={Diam}");

            // base coordinate frame
            var plane = Plane.PlaneXY;
            var origin = Point.Origin;
            var dirX = plane.Frame.DirX;
            var dirY = plane.Frame.DirY;
            var dirZ = plane.Frame.DirZ;

            Body body = null;

            if (type.Contains("support"))
            {
                Logger.Log("[CreatePlateBody] Branch: support");
                // extract exactly 6 segments from the V18-style contour:
                var segs = getBaseCapContour(origin, dirX, dirY, type, L1, B1, Rad);
                if (segs.Count == 6 && T > 0)
                {
                    var supportPlane = Plane.Create(Frame.Create(origin + 0.5 * T * dirY, dirX, dirZ));
                    body = Body.ExtrudeProfile(new Profile(supportPlane, segs), T);
                    Logger.Log("[CreatePlateBody] support: extruded support profile");
                }
                else
                    Logger.Log("[CreatePlateBody] support: invalid contour or thickness");
            }
            else if (type.Contains("flange"))
            {
                Logger.Log("[CreatePlateBody] Branch: flange");
                if (T <= 0)
                {
                    Logger.Log("[CreatePlateBody] flange: T<=0, abort");
                    return null;
                }

                // outer disk
                var holePlane = plane.CreateTransformedCopy(Matrix.CreateTranslation(origin - plane.Frame.Origin));
                body = Body.ExtrudeProfile(new CircleProfile(holePlane, 0.5 * L1), T);
                Logger.Log("[CreatePlateBody] flange: extruded disk");

                // flat‐flange inner cut
                if (type == "Flat flange")
                {
                    if (B1 > 0)
                    {
                        var cut = Body.ExtrudeProfile(new CircleProfile(holePlane, 0.5 * B1), 1.1 * T);
                        body.Subtract(cut);
                        Logger.Log("[CreatePlateBody] flange: subtracted inner cut");
                    }
                    else
                        Logger.Log("[CreatePlateBody] flange: B1<=0, skipped inner cut");
                }

                // bolt holes
                if (Diam > 0 && Lnr > 0)
                {
                    Logger.Log("[CreatePlateBody] flange: creating holes");
                    var circ = new FrameGenerator.Modules.Profiles.CircularProfile(Diam, 0.0, false, 0, 0);
                    var loop = circ.GetProfileCurves(holePlane).ToList();
                    Logger.Log($"[CreatePlateBody] flange: hole loop has {loop.Count} curves");
                    var prof = new Profile(holePlane, loop);
                    var cutter = Body.ExtrudeProfile(prof, 1.1 * T);
                    double step = 2 * Math.PI / Lnr;
                    for (int i = 0; i < Lnr; i++)
                    {
                        var copy = cutter.Copy();
                        copy.Transform(Matrix.CreateRotation(Line.Create(origin, dirZ), step * i));
                        body.Subtract(copy);
                    }
                    cutter.Dispose();
                    Logger.Log("[CreatePlateBody] flange: all holes cut");
                }
                else
                    Logger.Log("[CreatePlateBody] flange: no holes to cut");
            }
            else
            {
                Logger.Log("[CreatePlateBody] Branch: rectangle");
                body = createRoundedRectangle(L1, L2, Lnr, B1, B2, Bnr, T, Rad, Diam);
                Logger.Log("[CreatePlateBody] rectangle: done");
            }

            Logger.Log("[CreatePlateBody] Exit");
            return body;
        }


        //— V18 support contour helper —
        private static List<ITrimmedCurve> getContourV18(
            Point p, Direction dX, Direction dY,
            string type, double L1, double B1, double R, double T)
        {
            var dZ = Direction.Cross(dX, dY);
            var list = new List<ITrimmedCurve>();
            if (type.Contains("support"))
            {
                var p1 = p + (-0.5 * L1 + R) * dX + (0.5 * T) * dY;
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
            double thickness, double radius, double holeDiam)
        {
            Logger.Log("[createRoundedRectangle] Enter");
            var plane = Plane.PlaneXY;
            var o = Point.Origin;
            var dX = plane.Frame.DirX;
            var dY = plane.Frame.DirY;
            var dZ = plane.Frame.DirZ;
            var boundary = new List<ITrimmedCurve>();

            Logger.Log("[createRoundedRectangle] 1");
            if (radius <= 0)
            {
                var p0 = o - 0.5 * length * dX - 0.5 * width * dY;
                var p1 = p0 + width * dY;
                var p2 = p1 + length * dX;
                var p3 = p2 - width * dY;
                boundary.Add(CurveSegment.Create(p0, p1));
                boundary.Add(CurveSegment.Create(p1, p2));
                boundary.Add(CurveSegment.Create(p2, p3));
                boundary.Add(CurveSegment.Create(p3, p0));
            }
            else
            {
                var p0 = o - 0.5 * (length - 2 * radius) * dX - 0.5 * width * dY;
                var p1 = p0 + (length - 2 * radius) * dX;
                var p2 = p1 + radius * (dX.ToVector() + dY.ToVector());
                var p3 = p2 + (width - 2 * radius) * dY;
                var p4 = p3 - radius * (dX.ToVector() - radius * dY);
                var p5 = p4 - (length - 2 * radius) * dX;
                var p6 = p5 - radius * (dX.ToVector() + dY.ToVector());
                var p7 = p6 - (width - 2 * radius) * dY;
                boundary.Add(CurveSegment.Create(p0, p1));
                boundary.Add(CurveSegment.CreateArc(p1 + radius * dY, p1, p2, dZ));
                boundary.Add(CurveSegment.Create(p2, p3));
                boundary.Add(CurveSegment.CreateArc(p3 - radius * dX, p3, p4, dZ));
                boundary.Add(CurveSegment.Create(p4, p5));
                boundary.Add(CurveSegment.CreateArc(p5 - radius * dY, p5, p6, dZ));
                boundary.Add(CurveSegment.Create(p6, p7));
                boundary.Add(CurveSegment.CreateArc(p7 + radius * dX, p7, p0, dZ));
            }
            Logger.Log("[createRoundedRectangle] 2");

            var body = Body.ExtrudeProfile(new Profile(plane, boundary), thickness);
            Logger.Log("[createRoundedRectangle] extruded plate");

            if (holeDiam > 0 && nrL > 0 && nrW > 0)
            {
                var cut = Body.ExtrudeProfile(new CircleProfile(plane, 0.5 * holeDiam), 1.1 * thickness);
                for (int i = 0; i < nrL; i++)
                    for (int j = 0; j < nrW; j++)
                    {
                        double dx = -0.5 * length2 + (nrL > 1 ? i * (length2 / (nrL - 1)) : 0);
                        double dy = -0.5 * width2 + (nrW > 1 ? j * (width2 / (nrW - 1)) : 0);
                        var c = cut.Copy();
                        c.Transform(Matrix.CreateTranslation(Vector.Create(dx, dy, 0)));
                        try { body.Subtract(c); } catch { c.Dispose(); }
                    }
                cut.Dispose();
                Logger.Log("[createRoundedRectangle] added holes");
            }

            Logger.Log("[createRoundedRectangle] Exit");
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
            Logger.Log($"[createPlate] Enter: placing '{type}_{name}'");
            var doc = Window.ActiveWindow.Document;
            var main = doc.MainPart;
            var dX = pl.Frame.DirX;
            var dY = pl.Frame.DirY;
            var pn = $"{type}_{name}";

            WriteBlock.ExecuteTask("Create part", () =>
            {
                var plates = doc.Parts.FirstOrDefault(p => p.DisplayName == "Plates")
                              ?? Part.Create(doc, "Plates")
                                 .Also(pp => Component.Create(main, pp));

                var exist = CheckIfPartExists(type, L1, L2, Lnr, B1, B2, Bnr, T, R, D);
                Component comp;
                if (exist != null)
                {
                    Logger.Log($"[createPlate] Reusing '{exist.DisplayName}'");
                    comp = Component.Create(plates, exist);
                }
                else
                {
                    Logger.Log("[createPlate] Building new plate body");
                    var b = CreatePlateBody(type, L1, L2, Lnr, B1, B2, Bnr, T, R, D);
                    if (b == null) { Logger.Log("[createPlate] abort body==null"); return; }
                    var pb = Part.Create(doc, pn);
                    DesignBody.Create(pb, pn, b);
                    comp = Component.Create(plates, pb);
                    CreateCustomPartPropertiesPlate(pb, type, L1, B1, T, L2, Lnr, B2, Bnr, D, R, ang);
                }

                var frame = Frame.Create(p, dX, dY);
                if (frame.DirZ != dZ) frame = Frame.Create(p, -dX, dY);
                comp.Transform(Matrix.CreateMapping(frame));
                Logger.Log("[createPlate] moved into place");

                if (ang != 0)
                {
                    comp.Transform(Matrix.CreateRotation(Line.Create(p, dZ), ang * Math.PI / 180));
                    Logger.Log($"[createPlate] rotated {ang}°");
                }
            });

            Logger.Log("[createPlate] Exit");
        }

        private static void CreateCustomPartPropertiesPlate(
            Part part, string type,
            double L1, double B1, double T,
            double L2, int Lnr,
            double B2, int Bnr,
            double Diam, double Rad,
            double Angle)
        {
            Logger.Log("[CreateCustomPartPropertiesPlate] start");
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
            Logger.Log("[CreateCustomPartPropertiesPlate] done");
        }

        private static bool CompareCustomProperties(
            Part part, string type,
            double L1, double L2, int Lnr,
            double B1, double B2, int Bnr,
            double T, double Rad, double Diam)
        {
            Logger.Log($"[CompareCustomProperties] Checking '{part.DisplayName}'");
            var props = part.CustomProperties;
            if (!props.TryGetValue("AESC_Construct", out var m) || m.Value.ToString() != "Plates")
            {
                Logger.Log("[CompareCustomProperties] missing marker"); return false;
            }
            if (!props.TryGetValue("Type", out var tp) || tp.Value.ToString() != type)
            {
                Logger.Log($"[CompareCustomProperties] type mismatch '{type}'"); return false;
            }
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
            bool ok = exp.All(kv =>
            {
                if (!props.TryGetValue(kv.Key, out var pr)) { Logger.Log($"[Compare] missing {kv.Key}"); return false; }
                if (!double.TryParse(pr.Value.ToString(), out var a)) { Logger.Log($"[Compare] parse fail {kv.Key}"); return false; }
                if (Math.Abs(a - kv.Value) > tol) { Logger.Log($"[Compare] {kv.Key}:{a}≠{kv.Value}"); return false; }
                return true;
            });
            Logger.Log($"[CompareCustomProperties] allMatch={ok}");
            return ok;
        }
    }

    internal static class Extensions
    {
        public static T Also<T>(this T self, Action<T> act) { act(self); return self; }
    }
}
