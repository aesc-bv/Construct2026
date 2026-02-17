using AESCConstruct2026.FrameGenerator.Utilities;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;

namespace AESCConstruct2026.FrameGenerator.Modules.Profiles
{
    public class CSVProfile : ProfileBase
    {
        private readonly List<ITrimmedCurve> originalCurves;
        private readonly double offsetX;
        private readonly double offsetY;

        public CSVProfile(List<ITrimmedCurve> curves, double offsetX, double offsetY)
        {
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
                                      + segment.StartPoint.X * profilePlane.Frame.DirX
                                      + segment.StartPoint.Y * profilePlane.Frame.DirY;

                    Point newEnd = profilePlane.Frame.Origin
                                    + segment.EndPoint.X * profilePlane.Frame.DirX
                                    + segment.EndPoint.Y * profilePlane.Frame.DirY;

                    if (segment.Geometry is Circle circle)
                    {
                        Point newCenter = profilePlane.Frame.Origin
                                          + circle.Axis.Origin.X * profilePlane.Frame.DirX
                                          + circle.Axis.Origin.Y * profilePlane.Frame.DirY;

                        if ((newStart - newEnd).Magnitude < 1e-6)
                        {
                            // Full circle
                            var newCircle = Circle.Create(
                                Frame.Create(newCenter, profilePlane.Frame.DirZ),
                                circle.Radius);
                            alignedCurves.Add(CurveSegment.Create(newCircle));
                        }
                        else
                        {
                            // Arc — map the original normal to the profile plane:
                            // original curves live in XY so the arc normal is +Z or -Z
                            Direction newNormal = circle.Axis.Direction.Z > 0
                                ? profilePlane.Frame.DirZ
                                : -profilePlane.Frame.DirZ;

                            try
                            {
                                alignedCurves.Add(CurveSegment.CreateArc(
                                    newCenter, newStart, newEnd, newNormal));
                            }
                            catch (Exception ex)
                            {
                                Logger.Log("CSVProfile arc creation failed, retrying with flipped normal: " + ex.ToString());
                                alignedCurves.Add(CurveSegment.CreateArc(
                                    newCenter, newStart, newEnd, -newNormal));
                            }
                        }
                    }
                    else
                    {
                        alignedCurves.Add(CurveSegment.Create(newStart, newEnd));
                    }
                }
            }

            return alignedCurves;
        }
    }
}
