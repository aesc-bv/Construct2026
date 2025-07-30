using AESCConstruct25.Properties;     // for Settings.Default
using SpaceClaim.Api.V242;           // for Component
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

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
            // ─── LOG ENTRY ──────────────────────────────────────────────────────
            // Logger.Log($"[SetNameAndLength] ENTER: profileType='{profileType}', length={length}");
            // Logger.Log($"[SetNameAndLength] Settings.TypeString = '{Settings.Default.TypeString}'");
            // Logger.Log($"[SetNameAndLength] Settings.NameString = '{Settings.Default.NameString}'");

            // ─── 1) Build override map ───────────────────────────────────────────
            var overrideMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var rawTypeString = Settings.Default.TypeString ?? "";
            foreach (var entry in rawTypeString
                         .Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var parts = entry.Split(new[] { '@' }, 2);
                if (parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[1]))
                {
                    var key = parts[0].Trim();
                    overrideMap[key] = parts[1].Trim();
                    // Logger.Log($"[SetNameAndLength] overrideMap['{key}'] = '{parts[1].Trim()}'");
                }
            }

            // ─── 2) Lookup overrideName ──────────────────────────────────────────
            string overrideName = null;
            if (!overrideMap.TryGetValue(profileType, out overrideName))
            {
                // fallback: any key that starts with the profileType
                var fuzzyKey = overrideMap.Keys
                    .FirstOrDefault(k => k.StartsWith(profileType, StringComparison.OrdinalIgnoreCase));
                if (fuzzyKey != null)
                {
                    overrideName = overrideMap[fuzzyKey];
                    // Logger.Log($"[SetNameAndLength] Fuzzy-matched overrideMap['{fuzzyKey}'] = '{overrideName}'");
                }
            }
            //if (overrideName != null)
            // Logger.Log($"[SetNameAndLength] Using overrideName = '{overrideName}'");
            //else
            // Logger.Log($"[SetNameAndLength] No override found for '{profileType}'");

            // ─── 3) If still no override, pull custom 'Name' property ────────────
            var part = component.Template;
            string customName = null;
            if (overrideName == null
                && part.CustomProperties.TryGetValue("Name", out var nameProp)
                && !string.IsNullOrWhiteSpace(nameProp.Value?.ToString()))
            {
                customName = nameProp.Value.ToString().Trim();
                // Logger.Log($"[SetNameAndLength] Using custom Part.Name = '{customName}'");
            }
            else if (overrideName == null)
            {
                // Logger.Log($"[SetNameAndLength] No custom Part.Name property found");
            }

            // ─── 4) Base name = override → custom → literal profileType ─────────
            var baseName = overrideName ?? customName ?? profileType;
            // Logger.Log($"[SetNameAndLength] baseName = '{baseName}'");

            // ─── 5) Parse w/h (mm) ───────────────────────────────────────────────
            double wMm = 0, hMm = 0;
            if (profileData.TryGetValue("w", out var wRaw)
                && double.TryParse(wRaw.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var w))
                wMm = w;
            if (profileData.TryGetValue("h", out var hRaw)
                && double.TryParse(hRaw.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var h))
                hMm = h;
            // Logger.Log($"[SetNameAndLength] parsed wMm={wMm}, hMm={hMm}");

            // ─── 6) Template (fallback if empty) ─────────────────────────────────
            var template = Settings.Default.NameString?.Trim();
            if (string.IsNullOrWhiteSpace(template))
            {
                template = "[Name]_[p1]x[p2]_[length]";
                // Logger.Log($"[SetNameAndLength] empty NameString → using default template '{template}'");
            }
            else
            {
                // Logger.Log($"[SetNameAndLength] using template '{template}'");
            }

            // ─── 7) Token values ─────────────────────────────────────────────────
            var p1 = ((int)wMm).ToString(CultureInfo.InvariantCulture);
            var p2 = ((int)hMm).ToString(CultureInfo.InvariantCulture);
            var lengthStr = ((int)(length * 1000)).ToString(CultureInfo.InvariantCulture);
            // Logger.Log($"[SetNameAndLength] tokens: p1={p1}, p2={p2}, length={lengthStr}");

            // ─── 8) Replace tokens ────────────────────────────────────────────────
            var finalName = template
                .Replace("[Name]", baseName)
                .Replace("[p1]", p1)
                .Replace("[p2]", p2)
                .Replace("[length]", lengthStr);
            // Logger.Log($"[SetNameAndLength] finalName = '{finalName}'");

            // ─── 9) Apply to Part template ────────────────────────────────────────
            part.Name = finalName;
            // Logger.Log($"[SetNameAndLength] applied part.Name = '{part.Name}'");

            // ───10) Stamp Construct_Length ────────────────────────────────────────
            if (part.CustomProperties.ContainsKey("Construct_Length"))
            {
                part.CustomProperties["Construct_Length"].Value = length;
                // Logger.Log($"[SetNameAndLength] updated existing Construct_Length = {length}");
            }
            else
            {
                CustomPartProperty.Create(part, "Construct_Length", length);
                // Logger.Log($"[SetNameAndLength] created new Construct_Length = {length}");
            }

            // Logger.Log($"[SetNameAndLength] EXIT");
        }

    }
}
