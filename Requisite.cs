using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConAndInv
{
    [Serializable]
    class Requisite
    {
        public string ActivityType { get; set; }
        public string PartnerName { get; set; }
        public string Address { get; set; }
        public string OptionalAddress { get; set; }
        public string CurrentAccount { get; set; }
        public string Bank { get; set; }
        public string PostAndNameOfRespondent { get; set; }
        public string BasedOn { get; set; }

        public string BankIdentificationCode { get; set; }
        public string TaxpayerIdentificationNumber { get; set; }
        public string Oced { get; set; }
        public string VadNumber { get; set; }
        

        public void insertSpace(int position)
        {
            if (position < 3)
            {
                return;
            } 
            Bank = Bank.Insert(position, " ");
            insertSpace(position - 3);
        }
        public string fileNameForContract()
        {
            return "";
        }
    }
}
