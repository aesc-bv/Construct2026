using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242.Geometry;
using System.Collections.Generic;

namespace AESCConstruct25.FrameGenerator.Modules.Profiles
{
    public class RectangularProfile : ProfileBase
    {
        //public double Height { get; }
        //public double Width { get; }
        //public double Thickness { get; }
        //public double OuterRadius { get; }
        //public double InnerRadius { get; }
        //public bool IsHollow { get; }

        private readonly double Height;
        private readonly double Width;
        private readonly double Thickness;
        private readonly double OuterRadius;
        private readonly double InnerRadius;
        private readonly bool IsHollow;
        //private readonly double offsetX;
        //private readonly double offsetY;

        public RectangularProfile(double height, double width, double thickness, double outerRadius, double innerRadius, bool isHollow/*, double offsetX, double offsetY*/)
        {
            this.Height = height;
            this.Width = width;
            this.Thickness = thickness;
            this.OuterRadius = outerRadius;
            this.InnerRadius = innerRadius;
            this.IsHollow = isHollow;
            //this.offsetX = offsetX;
            //this.offsetY = offsetY;
        }

        public override ICollection<ITrimmedCurve> GetProfileCurves(Plane profilePlane)
        {
            return CreateRectangleCurves(profilePlane, Width, Height, OuterRadius/*, offsetX, offsetY*/);
        }

        public override ICollection<ITrimmedCurve> GetInnerProfile(Plane profilePlane)
        {
            return CreateRectangleCurves(profilePlane, Width - 2 * Thickness, Height - 2 * Thickness, InnerRadius/*, offsetX, offsetY*/);
        }


        private static List<ITrimmedCurve> CreateRectangleCurves(Plane profilePlane, double width, double height, double radius/*, double offsetX, double offsetY*/)
        {
            List<ITrimmedCurve> curves = new List<ITrimmedCurve>();
            Frame frame = profilePlane.Frame;
            //Vector offsetVector = offsetX * frame.DirX + offsetY * frame.DirY;
            Point center = frame.Origin;// + offsetVector;
            Logger.Log($"rect radius = ${radius}");
            // Define key points
            Point p1 = center + (-width / 2 + radius) * frame.DirX + (height / 2) * frame.DirY;
            Point p2 = center + (width / 2 - radius) * frame.DirX + (height / 2) * frame.DirY;
            Point p3 = center + (width / 2) * frame.DirX + (height / 2 - radius) * frame.DirY;
            Point p4 = center + (width / 2) * frame.DirX + (-height / 2 + radius) * frame.DirY;
            Point p5 = center + (width / 2 - radius) * frame.DirX + (-height / 2) * frame.DirY;
            Point p6 = center + (-width / 2 + radius) * frame.DirX + (-height / 2) * frame.DirY;
            Point p7 = center + (-width / 2) * frame.DirX + (-height / 2 + radius) * frame.DirY;
            Point p8 = center + (-width / 2) * frame.DirX + (height / 2 - radius) * frame.DirY;

            // Midpoints for arcs
            Point m1 = center + (-width / 2 + radius) * frame.DirX + (height / 2 - radius) * frame.DirY;
            Point m2 = center + (width / 2 - radius) * frame.DirX + (height / 2 - radius) * frame.DirY;
            Point m3 = center + (width / 2 - radius) * frame.DirX + (-height / 2 + radius) * frame.DirY;
            Point m4 = center + (-width / 2 + radius) * frame.DirX + (-height / 2 + radius) * frame.DirY;

            Direction dirZArc = frame.DirZ; // Default arc direction

            // Create straight edges
            curves.Add(CurveSegment.Create(p1, p2));
            if (radius > 0)
                curves.Add(CurveSegment.CreateArc(m2, p2, p3, -dirZArc));
            curves.Add(CurveSegment.Create(p3, p4));
            if (radius > 0)
                curves.Add(CurveSegment.CreateArc(m3, p4, p5, -dirZArc));
            curves.Add(CurveSegment.Create(p5, p6));
            if (radius > 0)
                curves.Add(CurveSegment.CreateArc(m4, p6, p7, -dirZArc));
            curves.Add(CurveSegment.Create(p7, p8));
            if (radius > 0)
                curves.Add(CurveSegment.CreateArc(m1, p8, p1, -dirZArc));

            return curves;
        }
    }
}
