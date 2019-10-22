using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConAndInv
{
    [Serializable]
    public class Goods
    {
        public string Description { get; set; }
        public string Name { get; set; }
        public string Format { get; set; }
        public string Volume { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public int VadInPercent { get; set; }

        public decimal sum ()
        {

            return Quantity * Price;
        }
        public decimal sumWithVad ()
        {
            return sum() * (VadInPercent / 100.0m + 1);
        }
        public decimal vadToPay()
        {
            return sumWithVad() - sum();
        }
    }
}
