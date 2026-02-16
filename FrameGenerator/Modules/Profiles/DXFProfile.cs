using SpaceClaim.Api.V242.Geometry;
using System.Collections.Generic;

namespace AESCConstruct2026.FrameGenerator.Modules.Profiles
{
    public class DXFProfile : ProfileBase
    {
        public override string FilePath { get; }
        private readonly List<ITrimmedCurve> originalCurves;
        private readonly double offsetX;
        private readonly double offsetY;

        public DXFProfile(string filePath, List<ITrimmedCurve> curves, double offsetX, double offsetY)
        {
            this.FilePath = filePath;
            this.offsetX = offsetX;
            this.offsetY = offsetY;
            this.originalCurves = curves;
        }

        public override ICollection<ITrimmedCurve> GetProfileCurves(Plane profilePlane)
        {
            List<ITrimmedCurve> alignedCurves = new List<ITrimmedCurve>();

            foreach (ITrimmedCurve curve in originalCurves)
            {
                if (curve is CurveSegment segment)
                {
                    Point newStart = profilePlane.Frame.Origin
                                      + (segment.StartPoint.X/* + offsetX*/) * profilePlane.Frame.DirX
                                      + (segment.StartPoint.Y/* + offsetY*/) * profilePlane.Frame.DirY;

                    Point newEnd = profilePlane.Frame.Origin
                                    + (segment.EndPoint.X/* + offsetX*/) * profilePlane.Frame.DirX
                                    + (segment.EndPoint.Y/* + offsetY*/) * profilePlane.Frame.DirY;

                    alignedCurves.Add(CurveSegment.Create(newStart, newEnd));
                }
            }

            return alignedCurves;
        }
    }
}
