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
        public double Price { get; set; }
        public int VadInPercent { get; set; }

        public double sum ()
        {

            return Quantity * Price;
        }
        public double sumWithVad ()
        {
            return sum() * (VadInPercent / 100.0 + 1);
        }
        public double vadToPay()
        {
            return sumWithVad() - sum();
        }
    }
}
