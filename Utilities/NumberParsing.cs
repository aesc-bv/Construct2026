using System.Globalization;

namespace AESCConstruct2026.FrameGenerator.Utilities
{
    // Locale-safe numeric parsing and formatting helpers.
    //
    // CSV data files and any cross-locale persistence must always use InvariantCulture
    // (decimal mark '.', thousands mark absent or ','). UI text boxes must tolerate
    // both '.' and ',' as a decimal mark because Dutch / German / French users type the
    // latter. Mixing the two — in particular calling TryParse with CurrentCulture before
    // InvariantCulture — causes "12.5" to parse as 125.0 on nl-NL (the '.' is interpreted
    // as a thousands separator). All call sites in this codebase must route through this
    // helper.
    public static class NumberParsing
    {
        private const NumberStyles DefaultStyles = NumberStyles.Float | NumberStyles.AllowThousands;

        // Strict parse: InvariantCulture only. Use for CSV / file I/O and any value that
        // we ourselves wrote in canonical form.
        public static bool TryParseInvariant(string s, out double value)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                value = 0;
                return false;
            }
            return double.TryParse(s, DefaultStyles, CultureInfo.InvariantCulture, out value);
        }

        // Tolerant parse for user-typed text. CAD dimensions never carry thousands
        // separators (nobody types "1,234.5 mm"), so we normalise a lone ',' to '.'
        // up-front and then parse with NumberStyles.Float — no AllowThousands.
        // Two reasons for this order:
        //   1. Parsing "12,3" with InvariantCulture + AllowThousands SUCCEEDS and
        //      returns 123 (',' treated as a grouping separator), which silently
        //      multiplies every Dutch-typed dimension by ~10×. Doing the replace
        //      first sidesteps that trap.
        //   2. Never calling TryParse with CurrentCulture avoids the inverse trap
        //      where "12.5" on nl-NL parses to 125.0 (',' is decimal, '.' becomes
        //      a thousands separator).
        public static bool TryParseUserInput(string s, out double value)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                value = 0;
                return false;
            }

            // A lone ',' is always a decimal mark in this codebase. If both ',' and
            // '.' are present we assume the user typed canonical form ("1,234.5" is
            // grouped) and leave it alone for the parser to reject — that case is
            // not a supported input shape.
            string normalised = s;
            if (s.IndexOf('.') < 0 && s.IndexOf(',') >= 0)
                normalised = s.Replace(',', '.');

            return double.TryParse(normalised, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        // Integer variant of TryParseUserInput. Accepts "5" and "5.0"/"5,0" but rejects
        // "5.3". Range-checks against Int32 to avoid silent overflow. Named distinctly
        // from the double overload so that existing `out var` callers stay unambiguous.
        public static bool TryParseUserInputInt(string s, out int value)
        {
            if (TryParseUserInput(s, out double d)
                && d >= int.MinValue && d <= int.MaxValue
                && d == System.Math.Truncate(d))
            {
                value = (int)d;
                return true;
            }
            value = 0;
            return false;
        }

        // Format with InvariantCulture. Use for CSV writes, clipboard output, and any
        // string that may be read on a different locale than the one that wrote it.
        public static string FormatInvariant(double value, string format = null)
        {
            return format == null
                ? value.ToString(CultureInfo.InvariantCulture)
                : value.ToString(format, CultureInfo.InvariantCulture);
        }
    }
}
