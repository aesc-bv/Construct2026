using System.Globalization;

namespace AESCConstruct25.Fastener.Module
{

    public class Nut
    {
        public string type { get; set; }
        public string Name { get; set; }
        public string size { get; set; }
        public double d { get; set; }
        public double s { get; set; }
        public double e { get; set; }
        public double h { get; set; }
        public double k { get; set; }

        public static Nut FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(';');
            for (int i = 0; i < values.Length; i++)
                values[i] = values[i].Trim();

            Nut profile = new Nut
            {
                type = values.Length > 0 ? values[0] : string.Empty,
                Name = values.Length > 1 ? values[1] : string.Empty,
                size = values.Length > 2 ? values[2] : string.Empty,
                d = values.Length > 3 && double.TryParse(values[3], NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal) ? dVal : 0,
                s = values.Length > 4 && double.TryParse(values[4], NumberStyles.Any, CultureInfo.InvariantCulture, out double sVal) ? sVal : 0,
                e = values.Length > 5 && double.TryParse(values[5], NumberStyles.Any, CultureInfo.InvariantCulture, out double eVal) ? eVal : 0,
                h = values.Length > 6 && double.TryParse(values[6], NumberStyles.Any, CultureInfo.InvariantCulture, out double hVal) ? hVal : 0
            };

            return profile;
        }

    }
}
