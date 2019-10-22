using System;
using System.Collections.Generic;
using System.Linq;
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
    /// Interaction logic for InGoods.xaml
    /// </summary>
    public partial class InGoods : UserControl
    {
        public InGoods()
        {
            InitializeComponent();
        }

        public decimal sum()
        {

            return Constants.StringToInt(invGoodQuantityTB.Text) * Constants.StringToDouble(invGoodPriceTB.Text);
        }
        public decimal sumWithVad()
        {
            return sum() * (Constants.StringToInt(invGoodVadTB.Text) / 100.0m + 1);
        }

        private void invGoodQuantityTB_TextChanged(object sender, TextChangedEventArgs e)
        {
            invGoodSumTBl.Text = sum().ToString();
            invGoodSumVadTBl.Text = sumWithVad().ToString();
        }
    }
}
