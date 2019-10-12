using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConAndInv
{
    enum Material
    {
        executivePaper = 0,
        executiveMaterial = 1,
        customerPaper = 2,
        customerMaterial = 3
    }
    [Serializable]
    class CIDocument
    {
        private static string[] materials = new string[] { "бумага исполнителя", "материал исполнителя", "бумага заказчика", "материал заказчика" };
        private Material material;
        public string Number { get; set; }
        public DateTime Date { get; set; }
        //public Material Material {set; }

        public Material Material
        {
            set { material = value; }
        }

        public Goods[] Goods { get; set; }
        public Requisite Requisite { get; set; }

        public string getMaterialString ()
        {
            return materials[(int)material];
        }
        public static string[] getArrayOfMaterials()
        {
            return materials;
        }
    }
    
}
