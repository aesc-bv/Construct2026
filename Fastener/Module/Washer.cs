using System.Globalization;

namespace AESCConstruct2026.Fastener.Module
{

    public class Washer
    {
        public string type { get; set; }
        public string Name { get; set; }
        public string size { get; set; }
        public double d1 { get; set; }
        public double d2 { get; set; }
        public double s { get; set; }

        public static Washer FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(';');
            for (int i = 0; i < values.Length; i++)
                values[i] = values[i].Trim();

            Washer profile = new Washer
            {
                type = values.Length > 0 ? values[0] : string.Empty,
                Name = values.Length > 1 ? values[1] : string.Empty,
                size = values.Length > 2 ? values[2] : string.Empty,
                d1 = values.Length > 3 && double.TryParse(values[3], NumberStyles.Any, CultureInfo.InvariantCulture, out double d1Val) ? d1Val : 0,
                d2 = values.Length > 4 && double.TryParse(values[4], NumberStyles.Any, CultureInfo.InvariantCulture, out double d2Val) ? d2Val : 0,
                s = values.Length > 5 && double.TryParse(values[5], NumberStyles.Any, CultureInfo.InvariantCulture, out double sVal) ? sVal : 0
            };

            return profile;
        }
    }
}
