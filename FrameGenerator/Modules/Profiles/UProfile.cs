using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;
using AESCConstruct25.FrameGenerator.Utilities; // Logger.Log

namespace AESCConstruct25.FrameGenerator.Modules.Profiles
{
    public class UProfile : ProfileBase
    {
        private readonly double height;              // m
        private readonly double width;               // m
        private readonly double webThickness;        // m
        private readonly double flangeThickness;     // m
        private readonly double innerCornerRadius;   // m (r1)
        private readonly double outerCornerRadius;   // m (r2)
        private readonly double offsetX;             // m (unused here; kept for parity)
        private readonly double offsetY;             // m (unused here; kept for parity)
        private readonly bool isUPN;                 // UPN (true) vs UPE (false)

        public UProfile(
            double height,
            double width,
            double webThickness,
            double flangeThickness,
            double innerCornerRadius,
            double outerCornerRadius,
            double offsetX,
            double offsetY,
            bool isUPN
        )
        {
            this.height = height;
            this.width = width;
            this.webThickness = webThickness;
            this.flangeThickness = flangeThickness;
            this.innerCornerRadius = innerCornerRadius;
            this.outerCornerRadius = outerCornerRadius;
            this.offsetX = offsetX;
            this.offsetY = offsetY;
            this.isUPN = isUPN;

            try
            {
                Logger.Log("[UProfile::.ctor] args (m): " +
                           $"h={F(height)}, w={F(width)}, s(web)={F(webThickness)}, t(flange)={F(flangeThickness)}, " +
                           $"r1(inner)={F(innerCornerRadius)}, r2(outer)={F(outerCornerRadius)}, " +
                           $"offX={F(offsetX)}, offY={F(offsetY)}, isUPN={isUPN}");

                // Basic sanity checks with warnings (not throwing to allow caller to see logs)
                if (height <= 0 || width <= 0)
                    Logger.Log("[UProfile::.ctor][WARN] Non-positive height/width.");
                if (webThickness <= 0 || flangeThickness <= 0)
                    Logger.Log("[UProfile::.ctor][WARN] Non-positive web/flange thickness.");
                if (innerCornerRadius < 0 || outerCornerRadius < 0)
                    Logger.Log("[UProfile::.ctor][WARN] Negative radius provided.");
                if (innerCornerRadius > Math.Min(webThickness, height))
                    Logger.Log("[UProfile::.ctor][WARN] r1 larger than web/height envelope.");
                if (outerCornerRadius > flangeThickness)
                    Logger.Log("[UProfile::.ctor][WARN] r2 larger than flange thickness.");
                if (webThickness >= width / 2.0)
                    Logger.Log("[UProfile::.ctor][WARN] web >= half width; geometry may collapse.");
                if (flangeThickness >= height / 2.0)
                    Logger.Log("[UProfile::.ctor][WARN] flange >= half height; geometry may collapse.");
            }
            catch (Exception ex)
            {
                Logger.Log($"[UProfile::.ctor][ERR] Exception while validating args: {ex}");
            }
        }

        public override ICollection<ITrimmedCurve> GetProfileCurves(Plane profilePlane)
        {
            var curves = new List<ITrimmedCurve>();

            try
            {
                Logger.Log("[UProfile::GetProfileCurves] Begin");
                if (profilePlane == null)
                {
                    Logger.Log("[UProfile::GetProfileCurves][ERR] profilePlane is null.");
                    return curves;
                }

                Frame frame = profilePlane.Frame;
                Point center = frame.Origin;

                DumpFrame("Sketch Frame", frame);

                // Convenience factory to construct points in this profile plane
                Point PX(double sx, double sy) => center + sx * frame.DirX + sy * frame.DirY;

                // === Base (no-draft) key scalars (m) ===
                double xLeft = -width / 2.0;
                double xRight = width / 2.0;
                double yTop = height / 2.0;
                double yBot = -height / 2.0;

                Logger.Log($"[UProfile] extents: xLeft={F(xLeft)}, xRight={F(xRight)}, yTop={F(yTop)}, yBot={F(yBot)}");

                // Outer corner radii geometry
                double xOuterCX = xRight - outerCornerRadius;                    // outer arc center X (right)
                double yOuterTopCY = yTop - flangeThickness + outerCornerRadius; // outer top-right arc center Y
                double yOuterBotCY = yBot + flangeThickness - outerCornerRadius; // outer bottom-right arc center Y

                Logger.Log($"[UProfile] outer arcs: centerX={F(xOuterCX)}, topCY={F(yOuterTopCY)}, botCY={F(yOuterBotCY)}");

                // Inner flange (before draft) Y-levels
                double yInnerTop = yTop - flangeThickness;   // top inner flange Y (flat case)
                double yInnerBot = yBot + flangeThickness;   // bottom inner flange Y (flat case)

                // X of the inner flange line at the web side (r1 from web)
                double xInnerWeb = xLeft + webThickness + innerCornerRadius;
                // X of the outer side at the flange fillet start
                double xOuterFilletStart = xRight - outerCornerRadius;

                // Distance along X of the inner flange segment (from outer fillet start to web+r1)
                double dxFlange = xOuterFilletStart - xInnerWeb;
                if (dxFlange < 0)
                {
                    Logger.Log($"[UProfile][WARN] dxFlange < 0 (dx={F(dxFlange)}). Clamping to 0.");
                    dxFlange = 0;
                }

                Logger.Log($"[UProfile] flange run dx={F(dxFlange)} (xOuterFilletStart={F(xOuterFilletStart)}, xInnerWeb={F(xInnerWeb)})");

                // === UPN-only inner flange draft ===
                double slope = 0.0;
                if (isUPN)
                {
                    double h_mm = height * 1000.0;
                    slope = (h_mm <= 300.0) ? 0.08 : 0.05; // tan(theta)
                    Logger.Log($"[UProfile] isUPN=true → slope={slope} (h={F(height)} m, {h_mm:0.#} mm)");
                }
                else
                {
                    Logger.Log("[UProfile] isUPN=false → slope=0 (UPE, no inner flange draft).");
                }

                // Draft amount in Y is slope * run in X
                double dY = slope * dxFlange;
                Logger.Log($"[UProfile] draft dY={F(dY)}");

                // Final (possibly drafted) inner flange Y levels:
                // Top flange gets thinner towards outer edge → inner line LOWER near the web
                double yInnerTop_final = yInnerTop - dY;
                // Bottom flange gets thinner towards outer edge → inner line HIGHER near the web
                double yInnerBot_final = yInnerBot + dY;

                Logger.Log($"[UProfile] innerY (flat): top={F(yInnerTop)}, bot={F(yInnerBot)} → (final): top={F(yInnerTop_final)}, bot={F(yInnerBot_final)}");

                // === Construct key points ===
                // Outer rectangle top
                Point p1 = PX(xLeft, yTop);
                Point p2 = PX(xRight, yTop);

                // Outer top-right short vertical start/end (leads into outer corner arc)
                Point p3 = PX(xRight, yTop - flangeThickness + outerCornerRadius);
                Point p4 = PX(xRight - outerCornerRadius, yTop - flangeThickness);

                // Inner top flange straight (p4 -> p5) — drafted ONLY for UPN
                Point p5 = PX(xInnerWeb, yInnerTop_final);

                // Inner web vertical (top) end point
                Point p6 = PX(xLeft + webThickness, yInnerTop_final - innerCornerRadius); // TOP web tangent
                Point p7 = PX(xLeft + webThickness, yInnerBot_final + innerCornerRadius);

                // Inner bottom flange straight (p8 -> p9) — drafted ONLY for UPN
                Point p8 = PX(xInnerWeb, yInnerBot_final);
                Point p9 = PX(xRight - outerCornerRadius, yBot + flangeThickness);

                // Outer bottom-right short vertical start/end (leads into outer corner arc)
                Point p10 = PX(xRight, yBot + flangeThickness - outerCornerRadius);

                // Outer rectangle bottom
                Point p11 = PX(xRight, yBot);
                Point p12 = PX(xLeft, yBot);

                DumpPoint("p1", p1); DumpPoint("p2", p2);
                DumpPoint("p3", p3); DumpPoint("p4", p4);
                DumpPoint("p5", p5); DumpPoint("p6", p6);
                DumpPoint("p7", p7); DumpPoint("p8", p8);
                DumpPoint("p9", p9); DumpPoint("p10", p10);
                DumpPoint("p11", p11); DumpPoint("p12", p12);

                // === Arc centers ===
                // Outer arcs unchanged
                Point m1 = PX(xOuterCX, yOuterTopCY);
                Point m4 = PX(xOuterCX, yOuterBotCY);

                // Inner arcs follow the drafted inner flange Y to keep exact inner radius r1
                // Top inner arc center is innerCornerRadius *below* the drafted inner top flange line
                Point m2 = PX(xInnerWeb, yInnerTop_final - innerCornerRadius);
                // Bottom inner arc center is innerCornerRadius *above* the drafted inner bottom flange line
                Point m3 = PX(xInnerWeb, yInnerBot_final + innerCornerRadius);

                DumpPoint("m1(outer top)", m1);
                DumpPoint("m4(outer bot)", m4);
                DumpPoint("m2(inner top)", m2);
                DumpPoint("m3(inner bot)", m3);

                // === Build straight edges ===
                AddLine(curves, "top outside", p1, p2);
                AddLine(curves, "outer TR short vert", p2, p3);
                AddLine(curves, "TOP inner flange (drafted on UPN)", p4, p5);
                AddLine(curves, "web inner vertical", p6, p7);
                AddLine(curves, "BOTTOM inner flange (drafted on UPN)", p8, p9);
                AddLine(curves, "outer BR short vert", p10, p11);
                AddLine(curves, "bottom outside", p11, p12);
                AddLine(curves, "left outside", p12, p1);

                // Arc orientation (Z points out of the sketch plane)
                Direction dirZArc = frame.DirZ;

                // OUTER arcs (top-right & bottom-right)
                if (outerCornerRadius > 0)
                {
                    AddArc(curves, "outer top-right", m1, p3, p4, -dirZArc, outerCornerRadius);
                    AddArc(curves, "outer bot-right", m4, p9, p10, -dirZArc, outerCornerRadius);
                }
                else
                {
                    Logger.Log("[UProfile] outerCornerRadius == 0 → no outer arcs.");
                }

                // INNER arcs (top-left & bottom-left along the web)
                if (innerCornerRadius > 0)
                {
                    AddArc(curves, "inner top-left", m2, p5, p6, dirZArc, innerCornerRadius);
                    AddArc(curves, "inner bot-left", m3, p7, p8, dirZArc, innerCornerRadius);
                }
                else
                {
                    Logger.Log("[UProfile] innerCornerRadius == 0 → no inner arcs.");
                }

                Logger.Log($"[UProfile::GetProfileCurves] Done. Curves count = {curves.Count}");
            }
            catch (Exception ex)
            {
                Logger.Log($"[UProfile::GetProfileCurves][ERR] Exception: {ex}");
            }

            return curves;
        }

        // ---------- helpers ----------

        private static string F(double v) => v.ToString("0.##########", System.Globalization.CultureInfo.InvariantCulture);

        private static void DumpFrame(string label, Frame f)
        {
            try
            {
                Logger.Log($"[{label}] origin=({F(f.Origin.X)},{F(f.Origin.Y)},{F(f.Origin.Z)}) " +
                           $"X=({F(f.DirX.X)},{F(f.DirX.Y)},{F(f.DirX.Z)}) " +
                           $"Y=({F(f.DirY.X)},{F(f.DirY.Y)},{F(f.DirY.Z)}) " +
                           $"Z=({F(f.DirZ.X)},{F(f.DirZ.Y)},{F(f.DirZ.Z)})");
            }
            catch (Exception ex)
            {
                Logger.Log($"[DumpFrame][ERR] {ex}");
            }
        }

        private static void DumpPoint(string name, Point p)
        {
            try
            {
                Logger.Log($"[pt:{name}] ({p.X:0.##########},{p.Y:0.##########},{p.Z:0.##########})");
            }
            catch (Exception ex)
            {
                Logger.Log($"[DumpPoint:{name}][ERR] {ex}");
            }
        }

        private static void AddLine(List<ITrimmedCurve> curves, string tag, Point a, Point b)
        {
            try
            {
                var seg = CurveSegment.Create(a, b);
                curves.Add(seg);
                Logger.Log($"[line:{tag}] A=({F(a.X)},{F(a.Y)},{F(a.Z)}) → B=({F(b.X)},{F(b.Y)},{F(b.Z)}) " +
                           $"len={F((b - a).Magnitude)}");
            }
            catch (Exception ex)
            {
                Logger.Log($"[AddLine:{tag}][ERR] {ex}");
            }
        }

        private static void AddArc(List<ITrimmedCurve> curves, string tag, Point center, Point start, Point end, Direction normal, double r)
        {
            try
            {
                var arc = CurveSegment.CreateArc(center, start, end, normal);
                curves.Add(arc);
                Logger.Log($"[arc:{tag}] C=({F(center.X)},{F(center.Y)},{F(center.Z)}), " +
                           $"S=({F(start.X)},{F(start.Y)},{F(start.Z)}), " +
                           $"E=({F(end.X)},{F(end.Y)},{F(end.Z)}), " +
                           $"r≈{F(r)}, n=({F(normal.X)},{F(normal.Y)},{F(normal.Z)})");
            }
            catch (Exception ex)
            {
                Logger.Log($"[AddArc:{tag}][ERR] {ex}");
            }
        }
    }
}


//using SpaceClaim.Api.V242.Geometry;
//using System.Collections.Generic;

//namespace AESCConstruct25.FrameGenerator.Modules.Profiles
//{
//    public class UProfile : ProfileBase
//    {
//        private readonly double height;
//        private readonly double width;
//        private readonly double webThickness;
//        private readonly double flangeThickness;
//        private readonly double innerCornerRadius;
//        private readonly double outerCornerRadius;
//        private readonly double offsetX;
//        private readonly double offsetY;
//        private readonly bool isUPN;

//        public UProfile(double height, double width, double webThickness, double flangeThickness, double innerCornerRadius, double outerCornerRadius, double offsetX, double offsetY, bool isUPN)
//        {
//            this.height = height;
//            this.width = width;
//            this.webThickness = webThickness;
//            this.flangeThickness = flangeThickness;
//            this.innerCornerRadius = innerCornerRadius;
//            this.outerCornerRadius = outerCornerRadius;
//            this.offsetX = offsetX;
//            this.offsetY = offsetY;
//            this.isUPN = isUPN;

//            // Logger.Log($"AESCConstruct25: Generating U Profile {width}x{height}, Web: {webThickness}, Flange: {flangeThickness}, InnerRadius: {innerCornerRadius}, OuterRadius: {outerCornerRadius}\n");
//        }

//        public override ICollection<ITrimmedCurve> GetProfileCurves(Plane profilePlane)
//        {
//            List<ITrimmedCurve> curves = new List<ITrimmedCurve>();
//            Frame frame = profilePlane.Frame;
//            Vector offsetVector = offsetX * frame.DirX + offsetY * frame.DirY;
//            Point center = frame.Origin;// + offsetVector;

//            // Define 12 key points for U Profile (open side to the right)
//            Point p1 = center + (-width / 2) * frame.DirX + (height / 2) * frame.DirY;
//            Point p2 = center + (width / 2) * frame.DirX + (height / 2) * frame.DirY;

//            Point p3 = center + (width / 2) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;
//            Point p4 = center + (width / 2 - outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;

//            Point p5 = center + (-width / 2 + webThickness + innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;
//            Point p6 = center + (-width / 2 + webThickness) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;

//            Point p7 = center + (-width / 2 + webThickness) * frame.DirX + (-height / 2 + flangeThickness + innerCornerRadius) * frame.DirY;
//            Point p8 = center + (-width / 2 + webThickness + innerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness) * frame.DirY;

//            Point p9 = center + (width / 2 - outerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness) * frame.DirY;
//            Point p10 = center + (width / 2) * frame.DirX + (-height / 2 + flangeThickness - outerCornerRadius) * frame.DirY;

//            Point p11 = center + (width / 2) * frame.DirX + (-height / 2) * frame.DirY;
//            Point p12 = center + (-width / 2) * frame.DirX + (-height / 2) * frame.DirY;

//            // Midpoints for arcs
//            Point m1 = center + (width / 2 - outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;
//            Point m2 = center + (-width / 2 + webThickness + innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;
//            Point m3 = center + (-width / 2 + webThickness + innerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness + innerCornerRadius) * frame.DirY;
//            Point m4 = center + (width / 2 - outerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness - outerCornerRadius) * frame.DirY;

//            // Create straight edges
//            curves.Add(CurveSegment.Create(p1, p2));
//            curves.Add(CurveSegment.Create(p2, p3));
//            curves.Add(CurveSegment.Create(p4, p5));
//            curves.Add(CurveSegment.Create(p6, p7));
//            curves.Add(CurveSegment.Create(p8, p9));
//            curves.Add(CurveSegment.Create(p10, p11));
//            curves.Add(CurveSegment.Create(p11, p12));
//            curves.Add(CurveSegment.Create(p12, p1));

//            // Ensure arcs are oriented correctly
//            Direction dirZArc = frame.DirZ;

//            // OUTER arcs (top-right & bottom-right)
//            if (outerCornerRadius > 0)
//            {
//                curves.Add(CurveSegment.CreateArc(m1, p3, p4, -dirZArc)); // top-right
//                curves.Add(CurveSegment.CreateArc(m4, p9, p10, -dirZArc)); // bottom-right
//            }

//            // INNER arcs (top-left & bottom-left along the web)
//            if (innerCornerRadius > 0)
//            {
//                curves.Add(CurveSegment.CreateArc(m2, p5, p6, dirZArc)); // top-left (inner)
//                curves.Add(CurveSegment.CreateArc(m3, p7, p8, dirZArc)); // bottom-left (inner)
//            }

//            return curves;
//        }
//    }
//}
