using AESCConstruct2026.FrameGenerator.Utilities;

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

            double Col(int i) => i < values.Length && NumberParsing.TryParseUserInput(values[i], out double v) ? v : 0;

            return new Washer
            {
                Type = values.Length > 0 ? values[0] : string.Empty,
                Name = values.Length > 1 ? values[1] : string.Empty,
                Size = values.Length > 2 ? values[2] : string.Empty,
                D1 = Col(3),
                D2 = Col(4),
                S = Col(5)
            };
        }
    }
}
