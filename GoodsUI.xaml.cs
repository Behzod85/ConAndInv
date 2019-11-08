using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ConAndInv
{
    /// <summary>
    /// Interaction logic for GoodsUI.xaml
    /// </summary>
    public partial class GoodsUI : UserControl
    {
        private List<Goods> MyGoods = new List<Goods>();
        public ObservableCollection<string> MyProperty2 { get; set; } = new ObservableCollection<string>();
        public List<TextBlock> MyProperty { get; set; }
        public GoodsUI()
        {
            InitializeComponent();
            setUpGoodsComponents();
        }

        private void contractGoodDescriptionCB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var i = contractGoodDescriptionCB.SelectedIndex;
            if (i < MyGoods.Count && i > -1)
            {
                contractGoodNameTB.Text = MyGoods[i].Name;
                contractGoodFormatTB.Text = MyGoods[i].Format;
                contractGoodVolumeTB.Text = MyGoods[i].Volume;
                contractGoodPriceTB.Text = MyGoods[i].Price.ToString();
                contractGoodVadTB.Text = MyGoods[i].VadInPercent.ToString();
            }
        }
        public void setUpGoodsComponents()
        {
            MyGoods = loadBinFile();
            MyGoods.Sort((i1, i2) => i1.Description.CompareTo(i2.Description));
            var goodsDescription = new string[MyGoods.Count];

            for (var i = 0; i < MyGoods.Count; i++)
            {

                goodsDescription[i] = MyGoods[i].Description;

            }
            contractGoodDescriptionCB.ItemsSource = goodsDescription;
        }

        private List<Goods> loadBinFile()
        {
            string fileName = Constants.GOOD_FILE_NAME; ;

            var binFormatter = new BinaryFormatter();
            try
            {
                using (var file = new FileStream(fileName, FileMode.Open))
                {
                    if (file != null)
                        return (List<Goods>)binFormatter.Deserialize(file);
                    else return new List<Goods>();
                }
            }
            catch (Exception)
            {

                return new List<Goods>();
            }


        }

        private void contractGoodQuantityTB_TextChanged(object sender, TextChangedEventArgs e)
        {
            contractGoodSumTBl.Text = sum().ToString();
            contractGoodSumVadTBl.Text = sumWithVad().ToString();
        }
        private decimal sum()
        {

            return Constants.StringToInt(contractGoodQuantityTB.Text) * Constants.StringToDouble(contractGoodPriceTB.Text);
        }
        private decimal sumWithVad()
        {
            return sum() * (Constants.StringToInt(contractGoodVadTB.Text) / 100.0m + 1);
        }
        private decimal vadToPay()
        {
            return sumWithVad() - sum();
        }
    }
}
