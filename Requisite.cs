using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConAndInv
{
    [Serializable]
   public class Requisite
    {
        public string Descrtipion { get; set; }
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
        public string Attorney { get; set; }
        public string Telephones { get; set; }

        public void insertSpace(int position)
        {
            if (position < 3)
            {
                return;
            }
            CurrentAccount = CurrentAccount.Insert(position, " ");
            insertSpace(position - 3);
        }
        public string fileNameForContract()
        {
            return "";
        }
        public override string ToString()
        {
            var s = Address;
            s += "\r";
            if (OptionalAddress != null) s += $"Производственный адрес: {OptionalAddress} \r";
            if (Telephones != null) s += $"Телефон: {Telephones} \r";
            s += "р/сч: " + CurrentAccount;
            s += "\r" + Bank + ". " + BankIdentificationCode;
            s += "\r ИНН: " + TaxpayerIdentificationNumber;
            if (Oced == null) s += "ОКЭД: -";
            else s += "ОКЭД: " + Oced;
            s += "\r Рег. код пл-а НДС: " + VadNumber;
            return s;
        }
    }
}
