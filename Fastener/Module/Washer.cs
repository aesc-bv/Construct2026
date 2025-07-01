using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AESCConstruct25.Fastener.Module
{

    public class Washer
    {
        public string type { get; set; }
        public string size { get; set; }
        public double d1 { get; set; }
        public double d2 { get; set; }
        public double s { get; set; }

        public static Washer FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(';');
            Washer profile = new Washer();
            profile.type = Convert.ToString(values[0]);
            profile.size = Convert.ToString(values[1]);
            profile.d1 = Convert.ToDouble(values[2]);
            profile.d2 = Convert.ToDouble(values[3]);
            profile.s = Convert.ToDouble(values[4]);
            return profile;
        }
    }
}
