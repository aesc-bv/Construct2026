using SpaceClaim.Api.V242.Geometry;
using System.Collections.Generic;

namespace AESCConstruct25.FrameGenerator.Modules.Profiles
{
    public class TProfile : ProfileBase
    {
        private readonly double height;
        private readonly double width;
        private readonly double webThickness;
        private readonly double flangeThickness;
        private readonly double innerCornerRadius;
        private readonly double outerCornerRadius;
        private readonly double flangeEndCornerRadius;
        private readonly double offsetX;
        private readonly double offsetY;

        public TProfile(double height, double width, double webThickness, double flangeThickness, double innerCornerRadius, double outerCornerRadius, double flangeEndCornerRadius, double offsetX, double offsetY)
        {
            this.height = height;
            this.width = width;
            this.webThickness = webThickness;
            this.flangeThickness = flangeThickness;
            this.innerCornerRadius = innerCornerRadius;
            this.outerCornerRadius = outerCornerRadius;
            this.flangeEndCornerRadius = flangeEndCornerRadius;
            this.offsetX = offsetX;
            this.offsetY = offsetY;

            // Logger.Log($"AESCConstruct25: Generating T Profile {width}x{height}, Web: {webThickness}, Flange: {flangeThickness}, InnerRadius: {innerCornerRadius}, OuterRadius: {outerCornerRadius}, FlangeEndRadius: {flangeEndCornerRadius}\n");
        }

        public override ICollection<ITrimmedCurve> GetProfileCurves(Plane profilePlane)
        {
            List<ITrimmedCurve> curves = new List<ITrimmedCurve>();
            Frame frame = profilePlane.Frame;
            Vector offsetVector = offsetX * frame.DirX + offsetY * frame.DirY;
            Point center = frame.Origin;// + offsetVector;
                                        // Logger.Log($"AESCConstruct25: OffsetT: {offsetX},  {offsetY}\n");

            // Define 14 key points
            Point p1 = center + (-width / 2) * frame.DirX + (height / 2) * frame.DirY;
            Point p2 = center + (width / 2) * frame.DirX + (height / 2) * frame.DirY;
            Point p3 = center + (width / 2) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;

            Point p4 = center + (width / 2 - outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;
            Point p5 = center + (webThickness / 2 + innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;
            Point p6 = center + (webThickness / 2) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;

            Point p7 = center + (webThickness / 2) * frame.DirX + (-height / 2 + flangeEndCornerRadius) * frame.DirY;
            Point p8 = center + (webThickness / 2 - flangeEndCornerRadius) * frame.DirX + (-height / 2) * frame.DirY;
            Point p9 = center + (-webThickness / 2 + flangeEndCornerRadius) * frame.DirX + (-height / 2) * frame.DirY;

            Point p10 = center + (-webThickness / 2) * frame.DirX + (-height / 2 + flangeEndCornerRadius) * frame.DirY;
            Point p11 = center + (-webThickness / 2) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;
            Point p12 = center + (-webThickness / 2 - innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;

            Point p13 = center + (-width / 2 + outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;
            Point p14 = center + (-width / 2) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;

            //// Midpoints for arcs
            Point m1 = center + (width / 2 - outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;
            Point m2 = center + (webThickness / 2 + innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;

            Point m3 = center + (webThickness / 2 - flangeEndCornerRadius) * frame.DirX + (-height / 2 + flangeEndCornerRadius) * frame.DirY;
            Point m4 = center + (-webThickness / 2 + flangeEndCornerRadius) * frame.DirX + (-height / 2 + flangeEndCornerRadius) * frame.DirY;

            Point m5 = center + (-webThickness / 2 - innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;
            Point m6 = center + (-width / 2 + outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;

            // Create straight edges
            curves.Add(CurveSegment.Create(p1, p2));
            curves.Add(CurveSegment.Create(p2, p3));
            curves.Add(CurveSegment.Create(p4, p5));
            curves.Add(CurveSegment.Create(p6, p7));
            curves.Add(CurveSegment.Create(p8, p9));
            curves.Add(CurveSegment.Create(p10, p11));
            curves.Add(CurveSegment.Create(p12, p13));
            curves.Add(CurveSegment.Create(p14, p1));

            // Ensure arcs are oriented correctly
            Direction dirZArc = frame.DirZ;

            // Top-right arc
            curves.Add(CurveSegment.CreateArc(m1, p3, p4, -dirZArc));

            // Inner right arc
            curves.Add(CurveSegment.CreateArc(m2, p5, p6, dirZArc));

            // Bottom-right arc
            curves.Add(CurveSegment.CreateArc(m3, p7, p8, -dirZArc));

            // Bottom-left arc
            curves.Add(CurveSegment.CreateArc(m4, p9, p10, -dirZArc));

            // Inner left arc
            curves.Add(CurveSegment.CreateArc(m5, p11, p12, dirZArc));

            // Top-left arc
            curves.Add(CurveSegment.CreateArc(m6, p13, p14, -dirZArc));
            //}

            return curves;
        }
    }
}
