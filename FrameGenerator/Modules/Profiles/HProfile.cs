using SpaceClaim.Api.V242.Geometry;
using System.Collections.Generic;

namespace AESCConstruct25.FrameGenerator.Modules.Profiles
{
    public class HProfile : ProfileBase
    {
        private readonly double height;
        private readonly double width;
        private readonly double webThickness;
        private readonly double flangeThickness;
        private readonly double innerCornerRadius;
        private readonly double outerCornerRadius;
        private readonly double offsetX;
        private readonly double offsetY;

        public HProfile(double height, double width, double webThickness, double flangeThickness, double innerCornerRadius, double outerCornerRadius, double offsetX, double offsetY)
        {
            this.height = height;
            this.width = width;
            this.webThickness = webThickness;
            this.flangeThickness = flangeThickness;
            this.innerCornerRadius = innerCornerRadius;
            this.outerCornerRadius = outerCornerRadius;
            this.offsetX = offsetX;
            this.offsetY = offsetY;

            // Logger.Log($"AESCConstruct25: Generating H Profile {width}x{height}, Web: {webThickness}, Flange: {flangeThickness}, InnerRadius: {innerCornerRadius}, OuterRadius: {outerCornerRadius}\n");
        }

        public override ICollection<ITrimmedCurve> GetProfileCurves(Plane profilePlane)
        {
            List<ITrimmedCurve> curves = new List<ITrimmedCurve>();
            Frame frame = profilePlane.Frame;
            Vector offsetVector = offsetX * frame.DirX + offsetY * frame.DirY;
            Point center = frame.Origin;// + offsetVector;

            // Define 20 key points for the H Profile
            Point p1 = center + (-width / 2) * frame.DirX + (height / 2) * frame.DirY;
            Point p2 = center + (width / 2) * frame.DirX + (height / 2) * frame.DirY;

            Point p3 = center + (width / 2) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;
            Point p4 = center + (width / 2 - outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;

            Point p5 = center + (webThickness / 2 + innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;
            Point p6 = center + (webThickness / 2) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;

            Point p7 = center + (webThickness / 2) * frame.DirX + (-height / 2 + flangeThickness + innerCornerRadius) * frame.DirY;
            Point p8 = center + (webThickness / 2 + innerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness) * frame.DirY;

            Point p9 = center + (width / 2 - outerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness) * frame.DirY;
            Point p10 = center + (width / 2) * frame.DirX + (-height / 2 + flangeThickness - outerCornerRadius) * frame.DirY;

            Point p11 = center + (width / 2) * frame.DirX + (-height / 2) * frame.DirY;
            Point p12 = center + (-width / 2) * frame.DirX + (-height / 2) * frame.DirY;

            Point p13 = center + (-width / 2) * frame.DirX + (-height / 2 + flangeThickness - outerCornerRadius) * frame.DirY;
            Point p14 = center + (-width / 2 + outerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness) * frame.DirY;

            Point p15 = center + (-webThickness / 2 - innerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness) * frame.DirY;
            Point p16 = center + (-webThickness / 2) * frame.DirX + (-height / 2 + flangeThickness + innerCornerRadius) * frame.DirY;

            Point p17 = center + (-webThickness / 2) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;
            Point p18 = center + (-webThickness / 2 - innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;

            Point p19 = center + (-width / 2 + outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;
            Point p20 = center + (-width / 2) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;

            // Midpoints for arcs
            Point m1 = center + (width / 2 - outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;
            Point m2 = center + (webThickness / 2 + innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;
            Point m3 = center + (webThickness / 2 + innerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness + innerCornerRadius) * frame.DirY;
            Point m4 = center + (width / 2 - outerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness - outerCornerRadius) * frame.DirY;

            Point m5 = center + (-width / 2 + outerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness - outerCornerRadius) * frame.DirY;
            Point m6 = center + (-webThickness / 2 - innerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness + innerCornerRadius) * frame.DirY;
            Point m7 = center + (-webThickness / 2 - innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;
            Point m8 = center + (-width / 2 + outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;

            if (innerCornerRadius > 0 || outerCornerRadius > 0)
            {
                // Ensure arcs are oriented correctly
                Direction dirZArc = frame.DirZ;

                // Create straight edges
                curves.Add(CurveSegment.Create(p1, p2));
                curves.Add(CurveSegment.Create(p2, p3));
                curves.Add(CurveSegment.CreateArc(m1, p3, p4, -dirZArc));
                curves.Add(CurveSegment.Create(p4, p5));
                curves.Add(CurveSegment.CreateArc(m2, p5, p6, dirZArc));
                curves.Add(CurveSegment.Create(p6, p7));
                curves.Add(CurveSegment.CreateArc(m3, p7, p8, dirZArc));
                curves.Add(CurveSegment.Create(p8, p9));
                curves.Add(CurveSegment.CreateArc(m4, p9, p10, -dirZArc));
                curves.Add(CurveSegment.Create(p10, p11));
                curves.Add(CurveSegment.Create(p11, p12));
                curves.Add(CurveSegment.Create(p12, p13));
                curves.Add(CurveSegment.CreateArc(m5, p13, p14, -dirZArc));
                curves.Add(CurveSegment.Create(p14, p15));
                curves.Add(CurveSegment.CreateArc(m6, p15, p16, dirZArc));
                curves.Add(CurveSegment.Create(p16, p17));
                curves.Add(CurveSegment.CreateArc(m7, p17, p18, dirZArc));
                curves.Add(CurveSegment.Create(p18, p19));
                curves.Add(CurveSegment.CreateArc(m8, p19, p20, -dirZArc));
                curves.Add(CurveSegment.Create(p20, p1));
            }
            //if (innerCornerRadius > 0 || outerCornerRadius > 0)
            //{
            //    Direction dirZArc = frame.DirZ;

            //    curves.Add(CurveSegment.CreateArc(m1, p2, p3, -dirZArc));

            //    // Bottom-right arc
            //    curves.Add(CurveSegment.CreateArc(m2, p4, p5, dirZArc));

            //    // Inner corner arc
            //    curves.Add(CurveSegment.CreateArc(m3, p6, p7, -dirZArc));
            //}

            return curves;
        }
    }
}
