using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;
using System.Globalization;
using Application = SpaceClaim.Api.V242.Application;

namespace AESCConstruct25.FrameGenerator.Modules.Profiles
{
    public abstract class ProfileBase
    {
        public virtual string FilePath { get; } = null; // Allow file path for DXF profiles

        public abstract ICollection<ITrimmedCurve> GetProfileCurves(Plane profilePlane);
        public virtual ICollection<ITrimmedCurve> GetInnerProfile(Plane profilePlane)
        {
            return new List<ITrimmedCurve>(); // Default: No inner profile unless overridden
        }
        public static ProfileBase CreateProfile(
            string profileType,
            string[] sizeValues,
            bool isHollow,
            double offsetX,
            double offsetY,
            string dxfFilePath = "",
            List<ITrimmedCurve> dxfContours = null)
        {
            try
            {
                double[] convertedSizes = new double[sizeValues.Length];
                for (int i = 0; i < sizeValues.Length; i++)
                {
                    var normalized = sizeValues[i].Replace(',', '.');
                    convertedSizes[i] = double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)
                        ? parsed / 1000.0
                        : 0.0;
                }
                if (profileType == "DXF" && dxfContours != null)
                    return new DXFProfile(dxfFilePath, dxfContours, offsetX, offsetY);
                if (profileType == "CSV" && dxfContours != null)
                    return new CSVProfile(dxfContours, offsetX, offsetY);
                if (profileType == "Circular" && sizeValues.Length == 2)
                    return new CircularProfile(convertedSizes[0], convertedSizes[1], isHollow, offsetX, offsetY);

                if (profileType == "Rectangular" && sizeValues.Length == 5)
                    return new RectangularProfile(convertedSizes[0], convertedSizes[1], convertedSizes[2], convertedSizes[3], convertedSizes[4], isHollow);

                if (profileType == "H" && sizeValues.Length == 6)
                    return new HProfile(convertedSizes[0], convertedSizes[1], convertedSizes[2], convertedSizes[3], convertedSizes[4], convertedSizes[5], offsetX, offsetY);

                if (profileType == "L" && sizeValues.Length == 5)
                    return new LProfile(convertedSizes[0], convertedSizes[1], convertedSizes[2], convertedSizes[3], convertedSizes[4], offsetX, offsetY);

                if (profileType == "T" && sizeValues.Length == 7)
                    return new TProfile(convertedSizes[0], convertedSizes[1], convertedSizes[2], convertedSizes[3], convertedSizes[4], convertedSizes[5], convertedSizes[6], offsetX, offsetY);

                if (profileType == "U" && sizeValues.Length == 6)
                    return new UProfile(convertedSizes[0], convertedSizes[1], convertedSizes[2], convertedSizes[3], convertedSizes[4], convertedSizes[5], offsetX, offsetY);
            }
            catch (Exception ex)
            {
                var msg = $"Invalid profile size format:\n  {string.Join(", ", sizeValues)}\n\n{ex.GetType().Name}: {ex.Message}";
                Application.ReportStatus(msg, StatusMessageType.Error, null);
            }

            return null;
        }
    }
}
