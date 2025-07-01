using System;
using System.Collections.Generic;
using System.IO;
using SpaceClaim.Api.V251.Geometry;

namespace AESCConstruct25.FrameGenerator.Modules.Profiles
{
    public class LProfile : ProfileBase
    {
        private readonly double height;
        private readonly double width;
        private readonly double thickness;
        private readonly double innerCornerRadius;
        private readonly double outerCornerRadius;
        private readonly double offsetX;
        private readonly double offsetY;
        private static string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AESCConstruct25_Log.txt");

        public LProfile(double height, double width, double thickness, double innerCornerRadius, double outerCornerRadius, double offsetX, double offsetY)
        {
            this.height = height;
            this.width = width;
            this.thickness = thickness;
            this.innerCornerRadius = innerCornerRadius;
            this.outerCornerRadius = outerCornerRadius;
            this.offsetX = offsetX;
            this.offsetY = offsetY;

            File.AppendAllText(logPath, $"AESCConstruct25: Generating L Profile {width}x{height}x{thickness}, InnerRadius: {innerCornerRadius}, OuterRadius: {outerCornerRadius}\n");
        }

        public override ICollection<ITrimmedCurve> GetProfileCurves(Plane profilePlane)
        {
            List<ITrimmedCurve> curves = new List<ITrimmedCurve>();
            Frame frame = profilePlane.Frame;
            Vector offsetVector = offsetX * frame.DirX + offsetY * frame.DirY;
            Point center = frame.Origin;// + offsetVector;

            File.AppendAllText(logPath, $"AESCConstruct25: OffsetL: {offsetX},  {offsetY}\n");
            // Define 9 key points
            Point p1 = center + (-width / 2) * frame.DirX + (height / 2) * frame.DirY;
            Point p2 = center + (-width / 2 + thickness - outerCornerRadius) * frame.DirX + (height / 2) * frame.DirY;

            Point p3 = center + (-width / 2 + thickness) * frame.DirX + (height / 2 - outerCornerRadius) * frame.DirY;
            Point p4 = center + (-width / 2 + thickness) * frame.DirX + (-height / 2 + thickness + innerCornerRadius) * frame.DirY;

            Point p5 = center + (-width / 2 + thickness + innerCornerRadius) * frame.DirX + (-height / 2 + thickness) * frame.DirY;
            Point p6 = center + (width / 2 - outerCornerRadius) * frame.DirX + (-height / 2 + thickness) * frame.DirY;

            Point p7 = center + (width / 2) * frame.DirX + (-height / 2 + thickness - outerCornerRadius) * frame.DirY;
            Point p8 = center + (width / 2) * frame.DirX + (-height / 2) * frame.DirY;

            Point p9 = center + (-width / 2) * frame.DirX + (-height / 2) * frame.DirY;

            // Midpoints for arcs
            Point m1 = center + (-width / 2 + thickness - outerCornerRadius) * frame.DirX + (height / 2 - outerCornerRadius) * frame.DirY;
            Point m2 = center + (-width / 2 + thickness + innerCornerRadius) * frame.DirX + (-height / 2 + thickness + innerCornerRadius) * frame.DirY;
            Point m3 = center + (width / 2 - outerCornerRadius) * frame.DirX + (-height / 2 + thickness - outerCornerRadius) * frame.DirY;

            // Create straight edges
            curves.Add(CurveSegment.Create(p1, p2));
            curves.Add(CurveSegment.Create(p3, p4));
            curves.Add(CurveSegment.Create(p5, p6));
            curves.Add(CurveSegment.Create(p7, p8));
            curves.Add(CurveSegment.Create(p8, p9));
            curves.Add(CurveSegment.Create(p9, p1));

            // Ensure arcs are oriented correctly
            if (innerCornerRadius > 0 || outerCornerRadius > 0)
            {
                Direction dirZArc = frame.DirZ;

                curves.Add(CurveSegment.CreateArc(m1, p2, p3, -dirZArc));

                // Bottom-right arc
                curves.Add(CurveSegment.CreateArc(m2, p4, p5, dirZArc));

                // Inner corner arc
                curves.Add(CurveSegment.CreateArc(m3, p6, p7, -dirZArc));
            }

            return curves;
        }
    }
}
