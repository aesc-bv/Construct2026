using System.Globalization;

namespace AESCConstruct2026.Fastener.Module
{

    /// <summary>Represents nut dimension data parsed from a CSV row.</summary>
    public class Nut
    {
        public string Type { get; set; }
        public string Name { get; set; }
        public string Size { get; set; }
        public double D { get; set; }
        public double S { get; set; }
        public double E { get; set; }
        public double H { get; set; }
        public double K { get; set; }

        public static Nut FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(';');
            for (int i = 0; i < values.Length; i++)
                values[i] = values[i].Trim();

            Nut profile = new Nut
            {
                Type = values.Length > 0 ? values[0] : string.Empty,
                Name = values.Length > 1 ? values[1] : string.Empty,
                Size = values.Length > 2 ? values[2] : string.Empty,
                D = values.Length > 3 && double.TryParse(values[3], NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal) ? dVal : 0,
                S = values.Length > 4 && double.TryParse(values[4], NumberStyles.Any, CultureInfo.InvariantCulture, out double sVal) ? sVal : 0,
                E = values.Length > 5 && double.TryParse(values[5], NumberStyles.Any, CultureInfo.InvariantCulture, out double eVal) ? eVal : 0,
                H = values.Length > 6 && double.TryParse(values[6], NumberStyles.Any, CultureInfo.InvariantCulture, out double hVal) ? hVal : 0
            };

            return profile;
        }

    }
}
