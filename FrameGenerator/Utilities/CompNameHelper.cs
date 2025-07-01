using System;
using System.Collections.Generic;
using System.Globalization;
using AESCConstruct25.Properties;     // for Settings.Default
using SpaceClaim.Api.V242;           // for Component
using AESCConstruct25.FrameGenerator.Modules;       // for CustomPartProperty

namespace AESCConstruct25.FrameGenerator.Utilities
{
    public static class CompNameHelper
    {
        /// <summary>
        /// Renames the component according to the user’s template and type-name overrides,
        /// and writes the Construct_Length property.
        /// </summary>
        /// <param name="component">The newly created Component.</param>
        /// <param name="profileType">Raw type key (e.g. "Rectangular").</param>
        /// <param name="profileData">
        /// Must contain "w" and optionally "h" (both in mm) as strings.
        /// </param>
        /// <param name="length">Length in meters (will be multiplied by 1000 for mm).</param>
        public static void SetNameAndLength(
            Component component,
            string profileType,
            Dictionary<string, string> profileData,
            double length)
        {
            // ─── 1) Build override map from Settings.Default.TypeString ────────────
            var overrideMap = new Dictionary<string, string>();
            var rawTypeString = Settings.Default.TypeString ?? "";
            foreach (var entry in rawTypeString
                         .Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var parts = entry.Split(new[] { '@' }, 2);
                if (parts.Length == 2)
                    overrideMap[parts[0]] = parts[1];
            }
            // choose the display name, or fallback to the raw type
            var displayName = overrideMap.TryGetValue(profileType, out var d) && !string.IsNullOrWhiteSpace(d)
                              ? d
                              : profileType;

            // ─── 2) Parse w/h from profileData (they’re already in mm) ───────────
            double wMm = 0.0, hMm = 0.0;
            if (profileData.TryGetValue("w", out var wRaw) &&
                double.TryParse(wRaw.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var w))
                wMm = w;
            if (profileData.TryGetValue("h", out var hRaw) &&
                double.TryParse(hRaw.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var h))
                hMm = h;

            // ─── 3) Grab the naming template ─────────────────────────────────────
            //    e.g. "[Name]_[p1]x[p2]_[length]"
            var template = Settings.Default.NameString?.Trim()
                           ?? "[Name]_[p1]x[p2]_[length]";

            // ─── 4) Build token values ──────────────────────────────────────────
            var p1 = ((int)wMm).ToString(CultureInfo.InvariantCulture);
            var p2 = ((int)hMm).ToString(CultureInfo.InvariantCulture);
            var lengthStr = ((int)(length * 1000)).ToString(CultureInfo.InvariantCulture);

            // ─── 5) Replace tokens in the template ───────────────────────────────
            var finalName = template
                .Replace("[Name]", displayName)
                .Replace("[p1]", p1)
                .Replace("[p2]", p2)
                .Replace("[length]", lengthStr);

            // ─── 6) Apply the name to both Component and its Part template ─────
            //component.Name = finalName;
            component.Template.Name = finalName;

            // ─── 7) Stamp the Construct_Length property on the Part ─────────────
            var part = component.Template;
            if (part.CustomProperties.ContainsKey("Construct_Length"))
                part.CustomProperties["Construct_Length"].Value = length;
            else
                CustomPartProperty.Create(part, "Construct_Length", length);
        }
    }
}
