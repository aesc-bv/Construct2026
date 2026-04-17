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

        // Tolerant parse for user-typed text: tries the canonical form first (so that
        // "12.5" always wins regardless of the current OS locale), then normalises a
        // lone ',' to '.' and retries. Does NOT fall through to CurrentCulture — doing
        // so would reintroduce the "12.5 -> 125" trap on nl-NL.
        public static bool TryParseUserInput(string s, out double value)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                value = 0;
                return false;
            }

            if (double.TryParse(s, DefaultStyles, CultureInfo.InvariantCulture, out value))
                return true;

            bool hasComma = s.IndexOf(',') >= 0;
            bool hasDot = s.IndexOf('.') >= 0;
            if (hasComma && !hasDot)
            {
                string normalised = s.Replace(',', '.');
                if (double.TryParse(normalised, DefaultStyles, CultureInfo.InvariantCulture, out value))
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
