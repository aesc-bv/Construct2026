using System.Globalization;

namespace AESCConstruct25.Fastener.Module
{
    public class Bolt
    {
        public string type { get; set; }
        public string Name { get; set; }
        public string size { get; set; }
        public double d { get; set; }
        public double c { get; set; }
        public double k { get; set; }
        public double s { get; set; }
        public double l { get; set; }
        public double t { get; set; }

        public static Bolt FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(';');
            // Trim all fields
            for (int i = 0; i < values.Length; i++)
                values[i] = values[i].Trim();

            Bolt profile = new Bolt
            {
                type = values.Length > 0 ? values[0] : string.Empty,
                Name = values.Length > 1 ? values[1] : string.Empty,
                size = values.Length > 2 ? values[2] : string.Empty,
                d = values.Length > 3 && double.TryParse(values[3], NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal) ? dVal : 0,
                c = values.Length > 4 && double.TryParse(values[4], NumberStyles.Any, CultureInfo.InvariantCulture, out double cVal) ? cVal : 0,
                k = values.Length > 5 && double.TryParse(values[5], NumberStyles.Any, CultureInfo.InvariantCulture, out double kVal) ? kVal : 0,
                s = values.Length > 6 && double.TryParse(values[6], NumberStyles.Any, CultureInfo.InvariantCulture, out double sVal) ? sVal : 0,
                l = values.Length > 7 && double.TryParse(values[7], NumberStyles.Any, CultureInfo.InvariantCulture, out double lVal) ? lVal : 0,
                t = values.Length > 8 && double.TryParse(values[8], NumberStyles.Any, CultureInfo.InvariantCulture, out double tVal) ? tVal : 0
            };

            return profile;
        }
    }
}
