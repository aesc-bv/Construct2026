using System.Globalization;

namespace AESCConstruct2026.Fastener.Module
{
    /// <summary>Represents bolt dimension data parsed from a CSV row.</summary>
    public class Bolt
    {
        public string Type { get; set; }
        public string Name { get; set; }
        public string Size { get; set; }
        public double D { get; set; }
        public double C { get; set; }
        public double K { get; set; }
        public double S { get; set; }
        public double L { get; set; }
        public double T { get; set; }

        public static Bolt FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(';');
            // Trim all fields
            for (int i = 0; i < values.Length; i++)
                values[i] = values[i].Trim();

            Bolt profile = new Bolt
            {
                Type = values.Length > 0 ? values[0] : string.Empty,
                Name = values.Length > 1 ? values[1] : string.Empty,
                Size = values.Length > 2 ? values[2] : string.Empty,
                D = values.Length > 3 && double.TryParse(values[3], NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal) ? dVal : 0,
                C = values.Length > 4 && double.TryParse(values[4], NumberStyles.Any, CultureInfo.InvariantCulture, out double cVal) ? cVal : 0,
                K = values.Length > 5 && double.TryParse(values[5], NumberStyles.Any, CultureInfo.InvariantCulture, out double kVal) ? kVal : 0,
                S = values.Length > 6 && double.TryParse(values[6], NumberStyles.Any, CultureInfo.InvariantCulture, out double sVal) ? sVal : 0,
                L = values.Length > 7 && double.TryParse(values[7], NumberStyles.Any, CultureInfo.InvariantCulture, out double lVal) ? lVal : 0,
                T = values.Length > 8 && double.TryParse(values[8], NumberStyles.Any, CultureInfo.InvariantCulture, out double tVal) ? tVal : 0
            };

            return profile;
        }
    }
}
