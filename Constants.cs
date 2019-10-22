using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConAndInv
{
    public static class Constants
    {
        public const string CONTRACT_FILE_NAME = "D:\\Data\\contracts.bin";
        public const string GOOD_FILE_NAME = "D:\\Data\\goods.bin";
        public const string PARTNER_FILE_NAME = "D:\\Data\\partners.bin";
        public const string SETTINGS_FILE_NAME = "D:\\Data\\settings.bin";
        public static decimal StringToDouble(string s)
        {
            var s2 = s.Replace(".", ",");
            if (decimal.TryParse(s2, out decimal d))
                return d;
            else
                return 0.0m;
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
