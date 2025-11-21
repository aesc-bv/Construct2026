using AESCConstruct25.Properties;     // Settings.Default
using SpaceClaim.Api.V242;           // Component
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace AESCConstruct25.FrameGenerator.Utilities
{
    public static class CompNameHelper
    {
        /// <summary>
        /// Renames the component according to the user’s template and type-name overrides,
        /// and writes the Construct_Length property (meters).
        /// 
        /// Supported tokens in Settings.Default.NameString (case-insensitive):
        ///   [name] [w] [h] [t] [s] [length]
        /// Backwards-compat:
        ///   [p1] -> [w], [p2] -> [h]
        ///
        /// profileData should contain numbers in **mm** as strings (as you already do):
        ///   "w", "h", "t", "s" (some shapes may omit some keys; we fallback to 0/empty).
        /// </summary>
        public static void SetNameAndLength(
            Component component,
            string profileType,
            Dictionary<string, string> profileData,
            double lengthMeters)
        {
            // --- 1) Build type→override map from Settings.TypeString ("Type@Name|Type@Name|...")
            var overrideMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var rawTypeString = Settings.Default.TypeString ?? "";
            foreach (var entry in rawTypeString.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var parts = entry.Split(new[] { '@' }, 2);
                if (parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[1]))
                    overrideMap[parts[0].Trim()] = parts[1].Trim();
            }

            // --- 2) Resolve base display name
            string overrideName = null;
            if (!overrideMap.TryGetValue(profileType, out overrideName))
            {
                var fuzzy = overrideMap.Keys.FirstOrDefault(k => k.StartsWith(profileType, StringComparison.OrdinalIgnoreCase));
                if (fuzzy != null) overrideName = overrideMap[fuzzy];
            }

            var part = component.Template;

            // If not overridden, use the Part custom "Name" (you already set that in CreateComponent), else literal profileType
            string customNameProp = null;
            if (part.CustomProperties.TryGetValue("Name", out var nameProp) &&
                !string.IsNullOrWhiteSpace(nameProp.Value?.ToString()))
            {
                customNameProp = nameProp.Value.ToString().Trim();
            }

            var baseName = overrideName ?? customNameProp ?? profileType;

            // --- 3) Gather numeric params from profileData (mm as strings)
            double wMm = GetMm(profileData, "w");
            double hMm = GetMm(profileData, "h");
            double tMm = GetMm(profileData, "t");
            double sMm = GetMm(profileData, "s");

            // Circular often only has D→we stored it as "w" already. If h missing, mirror w.
            if (hMm <= 0 && wMm > 0) hMm = wMm;

            // --- 4) Determine template (default if empty)
            var rawTemplate = (Settings.Default.NameString ?? "").Trim();
            if (string.IsNullOrWhiteSpace(rawTemplate))
                rawTemplate = "[name]_[w]x[h]_[length]";

            // --- 5) Prepare token map with formatted strings
            var tokens = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = baseName ?? string.Empty,
                ["w"] = FormatMm(wMm),
                ["h"] = FormatMm(hMm),
                ["t"] = FormatMm(tMm),
                ["s"] = FormatMm(sMm),
                ["length"] = FormatLengthMm(lengthMeters)
            };

            // --- 6) Apply template (case-insensitive token replacement)
            string finalName = ApplyNameTemplate(rawTemplate, tokens);

            // --- 7) Write to Part and stamp Construct_Length (meters)
            part.Name = finalName;

            if (part.CustomProperties.ContainsKey("Construct_Length"))
                part.CustomProperties["Construct_Length"].Value = lengthMeters;
            else
                CustomPartProperty.Create(part, "Construct_Length", lengthMeters);
        }

        // ---------- helpers ----------

        private static double GetMm(Dictionary<string, string> pd, string key)
        {
            if (pd == null) return 0.0;
            if (!pd.TryGetValue(key, out var raw)) return 0.0;

            raw = (raw ?? "").Replace(',', '.');
            if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var v))
                return v;

            return 0.0;
        }

        private static string ApplyNameTemplate(string template, IDictionary<string, string> values)
        {
            // Matches tokens like [name], [w], [h], [t], [s], [length] in any casing
            return Regex.Replace(template, @"\[(name|w|h|t|s|length)\]", m =>
            {
                var key = m.Groups[1].Value;
                return values.TryGetValue(key, out var val) ? (val ?? "") : "";
            }, RegexOptions.IgnoreCase);
        }

        private static string FormatMm(double valueMm)
        {
            // Kill sub-micron noise, then snap to integer mm if ~equal
            double v = Math.Round(valueMm, 6);
            double nearest = Math.Round(v, MidpointRounding.AwayFromZero);
            if (Math.Abs(v - nearest) < 1e-6) v = nearest;

            return v.ToString(v % 1 == 0 ? "0" : "0.###", CultureInfo.InvariantCulture);
        }

        private static string FormatLengthMm(double lengthMeters)
        {
            return FormatMm(lengthMeters * 1000.0);
        }
    }
}
