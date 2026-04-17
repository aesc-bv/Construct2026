using AESCConstruct2026.FrameGenerator.Modules;
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
        private static readonly Regex NameTokenRegex = new Regex(@"\[([A-Za-z][A-Za-z0-9_]*)\]", RegexOptions.Compiled);

        /// <summary>
        /// Renames the component according to the user's template and type-name overrides,
        /// and writes the Construct_Length property (meters).
        ///
        /// Supported tokens (all case-insensitive):
        ///   [name]                 — base display name (from TypeString override, Part "Name" custom property, or profileType)
        ///   [length]               — tube length in mm, formatted with NameDecimals
        ///   [p1] [p2] [p3] ...     — positional dimension tokens in the order returned by ProfileModule.GetArgs
        ///   [w] [h] [t] [s] [d]    — direct profileData keys (case-insensitive)
        ///   any other CSV column   — e.g. [D], [a], [b], [r1], [r2], [r3] — resolved directly from profileData
        ///
        /// Fallbacks so the name never shows a literal `[x]` or a zero where a real value exists:
        ///   - Missing `w`/`h`/`d` → falls back to profileData["D"] (Circular diameter).
        ///   - Missing single-letter dimension token → first non-zero positional value.
        ///   - Unknown token → empty string (drops the literal `[...]`).
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

            // --- 3) Determine template: per-profile → global → hardcoded default
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

            // --- 4) Apply template (generic token resolution)
            int decimals = Math.Max(0, Settings.Default.NameDecimals);
            string finalName = ApplyTemplate(rawTemplate, profileType, profileData, baseName, lengthMeters, decimals);

            // --- 5) Write to Part and stamp Construct_Length (meters)
            part.Name = finalName;

            if (part.CustomProperties.ContainsKey("Construct_Length"))
                part.CustomProperties["Construct_Length"].Value = lengthMeters;
            else
                CustomPartProperty.Create(part, "Construct_Length", lengthMeters);
        }

        // ---------- template engine ----------

        private static string ApplyTemplate(
            string template,
            string profileType,
            Dictionary<string, string> profileData,
            string baseName,
            double lengthMeters,
            int decimals)
        {
            var ciData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (profileData != null)
            {
                foreach (var kv in profileData)
                {
                    if (!ciData.ContainsKey(kv.Key))
                        ciData[kv.Key] = kv.Value;
                }
            }

            string[] positional;
            try { positional = ProfileModule.GetArgs(profileType, profileData) ?? Array.Empty<string>(); }
            catch { positional = Array.Empty<string>(); }

            return NameTokenRegex.Replace(template, m =>
            {
                var key = m.Groups[1].Value;
                return ResolveToken(key, ciData, positional, baseName, lengthMeters, decimals);
            });
        }

        private static string ResolveToken(
            string key,
            Dictionary<string, string> ciData,
            string[] positional,
            string baseName,
            double lengthMeters,
            int decimals)
        {
            if (string.Equals(key, "name", StringComparison.OrdinalIgnoreCase))
                return baseName ?? "";
            if (string.Equals(key, "length", StringComparison.OrdinalIgnoreCase))
                return FormatLengthMm(lengthMeters, decimals);

            // Positional [p1]..[pN]
            if (key.Length >= 2 && (key[0] == 'p' || key[0] == 'P')
                && int.TryParse(key.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)
                && n >= 1 && n <= positional.Length)
            {
                var v = FormatIfNumeric(positional[n - 1], decimals);
                if (!IsEffectivelyZero(v)) return v;
                return v; // return "0" as-is if that's genuinely what the profile reports
            }

            // Direct profileData lookup (case-insensitive)
            if (ciData.TryGetValue(key, out var raw))
            {
                var formatted = FormatIfNumeric(raw, decimals);
                if (!IsEffectivelyZero(formatted))
                    return formatted;
                // Key exists but value is zero — fall through to dimension fallbacks below
            }

            // Fallback: `[w]`, `[h]`, `[d]` on profiles that only carry `D` (e.g. Circular)
            bool isWHD = string.Equals(key, "w", StringComparison.OrdinalIgnoreCase)
                      || string.Equals(key, "h", StringComparison.OrdinalIgnoreCase)
                      || string.Equals(key, "d", StringComparison.OrdinalIgnoreCase);
            if (isWHD && ciData.TryGetValue("D", out var dv))
            {
                var f = FormatIfNumeric(dv, decimals);
                if (!IsEffectivelyZero(f)) return f;
            }

            // Last-resort for any dimension-like token: first non-zero positional value.
            if (IsDimensionLetter(key) && positional.Length > 0)
            {
                foreach (var p in positional)
                {
                    var f = FormatIfNumeric(p, decimals);
                    if (!IsEffectivelyZero(f)) return f;
                }
            }

            // Return what direct lookup gave us (even if zero) before giving up.
            if (ciData.TryGetValue(key, out var rawAgain))
                return FormatIfNumeric(rawAgain, decimals);

            return "";
        }

        // ---------- helpers ----------

        private static bool IsDimensionLetter(string key)
        {
            if (string.IsNullOrEmpty(key) || key.Length != 1) return false;
            char c = char.ToLowerInvariant(key[0]);
            return c == 'w' || c == 'h' || c == 'd' || c == 'a' || c == 'b' || c == 's' || c == 't';
        }

        private static bool IsEffectivelyZero(string formatted)
        {
            if (string.IsNullOrEmpty(formatted)) return true;
            if (formatted == "0") return true;
            return false;
        }

        private static string FormatIfNumeric(string raw, int decimals)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            var normalized = raw.Replace(',', '.');
            if (double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out var v))
                return FormatMm(v, decimals);
            return raw;
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
