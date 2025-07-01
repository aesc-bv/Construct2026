using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AESCConstruct25.Fastener.Module
{
    public class Bolt
    {
        public string type { get; set; }
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
            Bolt profile = new Bolt();

            profile.type = !string.IsNullOrEmpty(values[0]) ? Convert.ToString(values[0]) : string.Empty;
            profile.size = !string.IsNullOrEmpty(values[1]) ? Convert.ToString(values[1]) : string.Empty;

            profile.d = !string.IsNullOrEmpty(values[2]) ? Convert.ToDouble(values[2]) : 0;
            profile.c = !string.IsNullOrEmpty(values[3]) ? Convert.ToDouble(values[3]) : 0;
            profile.k = !string.IsNullOrEmpty(values[4]) ? Convert.ToDouble(values[4]) : 0;
            profile.s = !string.IsNullOrEmpty(values[5]) ? Convert.ToDouble(values[5]) : 0;
            profile.l = !string.IsNullOrEmpty(values[6]) ? Convert.ToDouble(values[6]) : 0;
            profile.t = !string.IsNullOrEmpty(values[7]) ? Convert.ToDouble(values[7]) : 0;

            return profile;
        }

    }
}
