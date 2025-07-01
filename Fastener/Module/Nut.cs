using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AESCConstruct25.Fastener.Module
{

    public class Nut
    {
        public string type { get; set; }
        public string size { get; set; }
        public double d { get; set; }
        public double s { get; set; }
        public double e { get; set; }
        public double h { get; set; }
        public double k { get; set; }

        public static Nut FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(';');
            Nut profile = new Nut();

            profile.type = !string.IsNullOrEmpty(values[0]) ? Convert.ToString(values[0]) : string.Empty;
            profile.size = !string.IsNullOrEmpty(values[1]) ? Convert.ToString(values[1]) : string.Empty;

            profile.d = !string.IsNullOrEmpty(values[2]) ? Convert.ToDouble(values[2]) : 0;
            profile.s = !string.IsNullOrEmpty(values[3]) ? Convert.ToDouble(values[3]) : 0;
            profile.e = !string.IsNullOrEmpty(values[4]) ? Convert.ToDouble(values[4]) : 0;
            profile.h = !string.IsNullOrEmpty(values[5]) ? Convert.ToDouble(values[5]) : 0;
            profile.k = !string.IsNullOrEmpty(values[6]) ? Convert.ToDouble(values[6]) : 0;

            return profile;
        }
    }
}
