using System;
using System.Globalization;
using AESCConstruct2026.Localization;

namespace AESCConstruct2026.Connector2
{
    /// <summary>
    /// Pre-flight geometry validation for Connector2 inputs. Returns a localized
    /// human-readable error message for the first violated constraint, or null
    /// when all constraints pass. Side-effect free so it can be reused by the
    /// hard Create-click gate and the live UI hints.
    /// </summary>
    internal static class ConnectorValidator
    {
        public static string Validate(
            double width1, double width2, double height,
            double tolerance, double endRelief,
            double radius, bool hasRounding,
            bool hasCornerCutout, double cornerCutoutRadius,
            bool radiusInCutOut, double radiusInCutOutRadius)
        {
            if (width1 <= 0) return L.T("Connector_Err_Width1Positive");
            if (width2 <= 0) return L.T("Connector_Err_Width2Positive");
            if (height <= 0) return L.T("Connector_Err_HeightPositive");

            if (tolerance < 0) return L.T("Connector_Err_TolerancePositive");
            if (endRelief < 0) return L.T("Connector_Err_EndReliefPositive");

            if (endRelief >= height)
                return L.F("Connector_Err_EndReliefLessThanHeight", Fmt(endRelief), Fmt(height));

            if (radius > 0)
            {
                if (hasRounding)
                {
                    double maxByWidth = width2 * 0.5;
                    if (radius > maxByWidth)
                        return L.F("Connector_Err_FilletRadiusExceedsHalfWidth2", Fmt(radius), Fmt(maxByWidth));
                    if (radius > height)
                        return L.F("Connector_Err_FilletRadiusExceedsHeight", Fmt(radius), Fmt(height));
                }
                else
                {
                    double maxByWidth = width2 / Math.Sqrt(2.0);
                    if (radius > maxByWidth)
                        return L.F("Connector_Err_ChamferExceedsWidth2", Fmt(radius), Fmt(maxByWidth));
                    double maxByHeight = height * Math.Sqrt(2.0);
                    if (radius > maxByHeight)
                        return L.F("Connector_Err_ChamferExceedsHeight", Fmt(radius), Fmt(maxByHeight));
                }
            }

            if (hasCornerCutout && cornerCutoutRadius > 0)
            {
                double maxByWidth = width1 * 0.5;
                if (cornerCutoutRadius >= maxByWidth)
                    return L.F("Connector_Err_CornerCutoutExceedsHalfWidth1", Fmt(cornerCutoutRadius), Fmt(maxByWidth));
                if (cornerCutoutRadius >= height)
                    return L.F("Connector_Err_CornerCutoutExceedsHeight", Fmt(cornerCutoutRadius), Fmt(height));
            }

            if (radiusInCutOut && radiusInCutOutRadius > 0)
            {
                double maxByWidth = width2 * 0.5;
                if (radiusInCutOutRadius >= maxByWidth)
                    return L.F("Connector_Err_RadiusInCutOutExceedsHalfWidth2", Fmt(radiusInCutOutRadius), Fmt(maxByWidth));
                if (radiusInCutOutRadius >= height)
                    return L.F("Connector_Err_RadiusInCutOutExceedsHeight", Fmt(radiusInCutOutRadius), Fmt(height));
            }

            return null;
        }

        private static string Fmt(double v) => v.ToString("0.##", CultureInfo.InvariantCulture);
    }
}
