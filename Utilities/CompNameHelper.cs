using AESCConstruct2026.Properties;     // Settings.Default
using SpaceClaim.Api.V242;           // Component
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace AESCConstruct2026.FrameGenerator.Utilities
{
    public static class CompNameHelper
    {
        private static readonly Regex NameTokenRegex = new Regex(@"\[(name|w|h|t|s|d|length)\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Renames the component according to the user's template and type-name overrides,
        /// and writes the Construct_Length property (meters).
        ///
        /// Supported tokens (case-insensitive):
        ///   [name] [w] [h] [t] [s] [d] [length]
        ///
        /// Each profile type can have its own naming template (stored in TypeString as
        /// Type@Name@Template). If the per-profile template is empty, falls back to the
        /// global Settings.Default.NameString.
        ///
        /// profileData should contain numbers in **mm** as strings:
        ///   "w", "h", "t", "s", "D" (some shapes may omit some keys; we fallback to 0/empty).
        /// </summary>
        public static void SetNameAndLength(
            Component component,
            string profileType,
            Dictionary<string, string> profileData,
            double lengthMeters)
        {
            // --- 1) Build type→(Name, Template) map from TypeString ("Type@Name@Template|...")
            var nameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var templateMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var rawTypeString = Settings.Default.TypeString ?? "";
            foreach (var entry in rawTypeString.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var parts = entry.Split('@');
                var type = parts[0].Trim();
                if (parts.Length >= 2 && !string.IsNullOrWhiteSpace(parts[1]))
                    nameMap[type] = parts[1].Trim();
                if (parts.Length >= 3 && !string.IsNullOrWhiteSpace(parts[2]))
                    templateMap[type] = parts[2].Trim();
            }

            // --- 2) Resolve base display name
            string overrideName = null;
            if (!nameMap.TryGetValue(profileType, out overrideName))
            {
                var fuzzy = nameMap.Keys.FirstOrDefault(k => k.StartsWith(profileType, StringComparison.OrdinalIgnoreCase));
                if (fuzzy != null) overrideName = nameMap[fuzzy];
            }

            var part = component.Template;

            // If not overridden, use the Part custom "Name", else literal profileType
            string customNameProp = null;
            if (part.CustomProperties.TryGetValue("Name", out var nameProp) &&
                !string.IsNullOrWhiteSpace(nameProp.Value?.ToString()))
            {
                customNameProp = nameProp.Value.ToString().Trim();
            }

            var baseName = overrideName ?? customNameProp ?? profileType;

            // --- 3) Gather numeric params from profileData (mm as strings)
            int decimals = Math.Max(0, Settings.Default.NameDecimals);

            double dMm = GetMm(profileData, "D");
            double wMm = GetMm(profileData, "w");
            if (wMm <= 0) wMm = dMm;   // Circular: D → w fallback
            double hMm = GetMm(profileData, "h");
            double tMm = GetMm(profileData, "t");
            double sMm = GetMm(profileData, "s");

            // Circular often only has D. If h missing, mirror w.
            if (hMm <= 0 && wMm > 0) hMm = wMm;

            // --- 4) Determine template: per-profile → global → hardcoded default
            string rawTemplate = null;
            if (!templateMap.TryGetValue(profileType, out rawTemplate) || string.IsNullOrWhiteSpace(rawTemplate))
            {
                var fuzzy = templateMap.Keys.FirstOrDefault(k => k.StartsWith(profileType, StringComparison.OrdinalIgnoreCase));
                if (fuzzy != null) rawTemplate = templateMap[fuzzy];
            }
            if (string.IsNullOrWhiteSpace(rawTemplate))
                rawTemplate = (Settings.Default.NameString ?? "").Trim();
            if (string.IsNullOrWhiteSpace(rawTemplate))
                rawTemplate = "[name]_[w]x[h]_[length]";

            // --- 5) Prepare token map with formatted strings
            var tokens = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = baseName ?? string.Empty,
                ["w"] = FormatMm(wMm, decimals),
                ["h"] = FormatMm(hMm, decimals),
                ["t"] = FormatMm(tMm, decimals),
                ["s"] = FormatMm(sMm, decimals),
                ["d"] = FormatMm(dMm, decimals),
                ["length"] = FormatLengthMm(lengthMeters, decimals)
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
            return NameTokenRegex.Replace(template, m =>
            {
                var key = m.Groups[1].Value;
                return values.TryGetValue(key, out var val) ? (val ?? "") : "";
            });
        }

        private static string FormatMm(double valueMm, int decimals)
        {
            // Kill sub-micron noise, then snap to integer mm if ~equal
            double v = Math.Round(valueMm, 6);
            double nearest = Math.Round(v, MidpointRounding.AwayFromZero);
            if (Math.Abs(v - nearest) < 1e-6) v = nearest;

            if (decimals <= 0 || v % 1 == 0)
                return v.ToString("0", CultureInfo.InvariantCulture);

            // "0.##" for decimals=2, "0.###" for decimals=3, etc.
            var fmt = "0." + new string('#', decimals);
            return v.ToString(fmt, CultureInfo.InvariantCulture);
        }

        private static string FormatLengthMm(double lengthMeters, int decimals)
        {
            return FormatMm(lengthMeters * 1000.0, decimals);
        }
    }
}
