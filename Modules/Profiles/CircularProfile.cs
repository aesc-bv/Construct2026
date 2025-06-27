using System;
using System.Collections.Generic;
using System.IO;
using SpaceClaim.Api.V251.Geometry;

namespace AESCConstruct25.Modules.Profiles
{
    public class CircularProfile : ProfileBase
    {
        private readonly double diameter;
        private readonly double thickness;
        private readonly bool isHollow;
        private readonly double offsetX;
        private readonly double offsetY;
        private static string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AESCConstruct25_Log.txt");

        public CircularProfile(double diameter, double thickness, bool isHollow, double offsetX, double offsetY)
        {
            this.diameter = diameter;
            this.thickness = thickness;
            this.isHollow = isHollow;
            this.offsetX = offsetX;
            this.offsetY = offsetY;

            File.AppendAllText(logPath, $"AESCConstruct25: Generating Circular Profile with diameter {diameter}m and thickness {thickness}m, Hollow: {isHollow}\n");
        }

        public override ICollection<ITrimmedCurve> GetProfileCurves(Plane profilePlane)
        {
            return CreateCircularCurves(profilePlane, diameter, offsetX, offsetY);
        }

        public override ICollection<ITrimmedCurve> GetInnerProfile(Plane profilePlane)
        {
            double innerDiameter = diameter - 2 * thickness;
            if (innerDiameter <= 0)
            {
                File.AppendAllText(logPath, "AESCConstruct25: ERROR - Inner diameter too small, skipping hollow profile.\n");
                return new List<ITrimmedCurve>(); // No inner profile
            }

            return CreateCircularCurves(profilePlane, innerDiameter, offsetX, offsetY);
        }

        private static List<ITrimmedCurve> CreateCircularCurves(Plane profilePlane, double circleDiameter, double offsetX, double offsetY)
        {
            List<ITrimmedCurve> curves = new List<ITrimmedCurve>();
            Frame frame = profilePlane.Frame;
            Vector offsetVector = offsetX * frame.DirX + offsetY * frame.DirY;
            Point center = frame.Origin;// + offsetVector;

            Circle circle = Circle.Create(Frame.Create(center, frame.DirZ), circleDiameter / 2);
            curves.Add(CurveSegment.Create(circle));

            return curves;
        }
    }
}
