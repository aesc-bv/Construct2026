using AESCConstruct2026.FrameGenerator.Utilities;

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

            double Col(int i) => i < values.Length && NumberParsing.TryParseUserInput(values[i], out double v) ? v : 0;

            return new Nut
            {
                Type = values.Length > 0 ? values[0] : string.Empty,
                Name = values.Length > 1 ? values[1] : string.Empty,
                Size = values.Length > 2 ? values[2] : string.Empty,
                D = Col(3),
                S = Col(4),
                E = Col(5),
                H = Col(6)
            };
        }

    }
}
