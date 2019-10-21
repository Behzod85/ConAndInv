using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConAndInv
{
    public static class Constants
    {
        public const string CONTRACT_FILE_NAME = "contracts.bin";
        public const string GOOD_FILE_NAME = "goods.bin";
        public const string PARTNER_FILE_NAME = "partners.bin";
        public const string SETTINGS_FILE_NAME = "settings.bin";
        public static double StringToDouble(string s)
        {
            var s2 = s.Replace(".", ",");
            if (double.TryParse(s2, out double d))
                return d;
            else
                return 0.0;
        }

        public static int StringToInt(string s)
        {
            if (int.TryParse(s, out int i))
                return i;
            else
                return 0;
        }

    }
}
