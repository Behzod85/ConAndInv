using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConAndInv
{
    public enum Material
    {
        executivePaper = 0,
        executiveMaterial = 1,
        customerPaper = 2,
        customerMaterial = 3
    }
    [Serializable]
    public class CIDocument
    {
        public string FileName { get; set; }
        private static string[] materials = new string[] { "бумага исполнителя", "материал исполнителя", "бумага заказчика", "материал заказчика" };
        private Material material;
        public int Number { get; set; }
        public DateTime Date { get; set; }
        //public Material Material {set; }
        public List<Goods> Goods { get; set; } = new List<Goods>();
        public Material Material
        {
            set { material = value; }
        }

        
        public Requisite Requisite { get; set; }
        public int getMaterialInt()
        {
            return (int)material;
        }
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
