using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConAndInv
{
    [Serializable]
    public class Settings
    {
        private static Settings instance;
        public string PathToContractsFolder { get; set; }
        public string PathToInvoicesFolder { get; set; }
        public string PathToInvoicesExcellFolder { get; set; }
        public int CurrentContractNumber { get; set; }
        public int currentVad { get; set; }
        private Settings()
        {

        }
        public static Settings getInstance()
        {
            if (instance == null)
                instance = new Settings();
            return instance;
        }
    }
}
