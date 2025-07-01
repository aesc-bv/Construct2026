using System;
using System.Collections.Generic;
using System.IO;
using SpaceClaim.Api.V242.Geometry;

namespace AESCConstruct25.FrameGenerator.Modules.Profiles
{
    public class UProfile : ProfileBase
    {
        private readonly double height;
        private readonly double width;
        private readonly double webThickness;
        private readonly double flangeThickness;
        private readonly double innerCornerRadius;
        private readonly double outerCornerRadius;
        private readonly double offsetX;
        private readonly double offsetY;
        private static string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AESCConstruct25_Log.txt");

        public UProfile(double height, double width, double webThickness, double flangeThickness, double innerCornerRadius, double outerCornerRadius, double offsetX, double offsetY)
        {
            this.height = height;
            this.width = width;
            this.webThickness = webThickness;
            this.flangeThickness = flangeThickness;
            this.innerCornerRadius = innerCornerRadius;
            this.outerCornerRadius = outerCornerRadius;
            this.offsetX = offsetX;
            this.offsetY = offsetY;

            File.AppendAllText(logPath, $"AESCConstruct25: Generating U Profile {width}x{height}, Web: {webThickness}, Flange: {flangeThickness}, InnerRadius: {innerCornerRadius}, OuterRadius: {outerCornerRadius}\n");
        }

        public override ICollection<ITrimmedCurve> GetProfileCurves(Plane profilePlane)
        {
            List<ITrimmedCurve> curves = new List<ITrimmedCurve>();
            Frame frame = profilePlane.Frame;
            Vector offsetVector = offsetX * frame.DirX + offsetY * frame.DirY;
            Point center = frame.Origin;// + offsetVector;

            // Define 12 key points for U Profile (open side to the right)
            Point p1 = center + (-width / 2) * frame.DirX + (height / 2) * frame.DirY;
            Point p2 = center + (width / 2) * frame.DirX + (height / 2) * frame.DirY;

            Point p3 = center + (width / 2) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;
            Point p4 = center + (width / 2 - outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;

            Point p5 = center + (-width / 2 + webThickness + innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness) * frame.DirY;
            Point p6 = center + (-width / 2 + webThickness) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;

            Point p7 = center + (-width / 2 + webThickness) * frame.DirX + (-height / 2 + flangeThickness + innerCornerRadius) * frame.DirY;
            Point p8 = center + (-width / 2 + webThickness + innerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness) * frame.DirY;

            Point p9 = center + (width / 2 - outerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness) * frame.DirY;
            Point p10 = center + (width / 2) * frame.DirX + (-height / 2 + flangeThickness - outerCornerRadius) * frame.DirY;

            Point p11 = center + (width / 2) * frame.DirX + (-height / 2) * frame.DirY;
            Point p12 = center + (-width / 2) * frame.DirX + (-height / 2) * frame.DirY;

            // Midpoints for arcs
            Point m1 = center + (width / 2 - outerCornerRadius) * frame.DirX + (height / 2 - flangeThickness + outerCornerRadius) * frame.DirY;
            Point m2 = center + (-width / 2 + webThickness + innerCornerRadius) * frame.DirX + (height / 2 - flangeThickness - innerCornerRadius) * frame.DirY;
            Point m3 = center + (-width / 2 + webThickness + innerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness + innerCornerRadius) * frame.DirY;
            Point m4 = center + (width / 2 - outerCornerRadius) * frame.DirX + (-height / 2 + flangeThickness - outerCornerRadius) * frame.DirY;

            // Create straight edges
            curves.Add(CurveSegment.Create(p1, p2));
            curves.Add(CurveSegment.Create(p2, p3));
            curves.Add(CurveSegment.Create(p4, p5));
            curves.Add(CurveSegment.Create(p6, p7));
            curves.Add(CurveSegment.Create(p8, p9));
            curves.Add(CurveSegment.Create(p10, p11));
            curves.Add(CurveSegment.Create(p11, p12));
            curves.Add(CurveSegment.Create(p12, p1));

            // Ensure arcs are oriented correctly
            if (innerCornerRadius > 0 || outerCornerRadius > 0)
            {
                Direction dirZArc = frame.DirZ;
                Vector normal = frame.DirZ.UnitVector;

                // Top-right arc
                curves.Add(CurveSegment.CreateArc(m1, p3, p4, -dirZArc));

                // Bottom-right arc
                curves.Add(CurveSegment.CreateArc(m2, p5, p6, dirZArc));

                // Bottom-left arc
                curves.Add(CurveSegment.CreateArc(m3, p7, p8, dirZArc));

                // Top-left arc
                curves.Add(CurveSegment.CreateArc(m4, p9, p10, -dirZArc));
            }

            return curves;
        }
    }
}
