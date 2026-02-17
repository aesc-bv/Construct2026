using System.Globalization;

namespace AESCConstruct2026.Fastener.Module
{

    /// <summary>Represents washer dimension data parsed from a CSV row.</summary>
    public class Washer
    {
        public string Type { get; set; }
        public string Name { get; set; }
        public string Size { get; set; }
        public double D1 { get; set; }
        public double D2 { get; set; }
        public double S { get; set; }

        public static Washer FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(';');
            for (int i = 0; i < values.Length; i++)
                values[i] = values[i].Trim();

            Washer profile = new Washer
            {
                Type = values.Length > 0 ? values[0] : string.Empty,
                Name = values.Length > 1 ? values[1] : string.Empty,
                Size = values.Length > 2 ? values[2] : string.Empty,
                D1 = values.Length > 3 && double.TryParse(values[3], NumberStyles.Any, CultureInfo.InvariantCulture, out double d1Val) ? d1Val : 0,
                D2 = values.Length > 4 && double.TryParse(values[4], NumberStyles.Any, CultureInfo.InvariantCulture, out double d2Val) ? d2Val : 0,
                S = values.Length > 5 && double.TryParse(values[5], NumberStyles.Any, CultureInfo.InvariantCulture, out double sVal) ? sVal : 0
            };

            return profile;
        }
    }
}
