using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using frm = System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using word = Microsoft.Office.Interop.Word;
using excel = Microsoft.Office.Interop.Excel;

namespace ConAndInv
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<GoodsUI> MyGoodUIs { get; set; } = new List<GoodsUI>();
        public List<InGoods> MyInvGoods { get; set; } = new List<InGoods>();
        public List<CIDocument> MyContracts { get; set; } = new List<CIDocument>();
        public List<Requisite> MyRequisites { get; set; } = new List<Requisite>();
        public List<Goods> MyGoods { get; set; } = new List<Goods>();
        public Settings setting = Settings.getInstance();

        public MainWindow()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("ru-RU");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("ru-RU");
            InitializeComponent();
            this.contractDateDP.SelectedDate = DateTime.Now;
            invContractDateDP.SelectedDate = DateTime.Now;
            invDateDP.SelectedDate = DateTime.Now;
            invAttorneyDateDP.SelectedDate = DateTime.Now;

            // Initializing Settings
            setting = loadBinFile<Settings>(setting);
            setUpSettings();

            // Initializing Requisites
            MyRequisites = loadBinFile<List<Requisite>>(MyRequisites);
            setUpPartnerComponents();

            // Initializing Requisites
            MyGoods = loadBinFile<List<Goods>>(MyGoods);
            setUpGoodsComponents();

            // Initializing Mycontracts
            MyContracts = loadBinFile<List<CIDocument>>(MyContracts);
            if (MyRequisites != null) partnerNameCB.SelectedIndex = 0;
            setUpContractsComponents();
            materialCB.SelectedIndex = 0;
            


        }



        #region Serialization save, load


        private T loadBinFile<T>(T myType)
        {
            string fileName = "unknown.bin";
            if (typeof(T) == typeof(List<CIDocument>))
            {
                fileName = Constants.CONTRACT_FILE_NAME;
            }
            else if (typeof(T) == typeof(List<Goods>))
            {
                fileName = Constants.GOOD_FILE_NAME;
            }
            else if (typeof(T) == typeof(List<Requisite>))
            {
                fileName = Constants.PARTNER_FILE_NAME;
            }
            else if (typeof(T) == typeof(Settings))
            {
                fileName = Constants.SETTINGS_FILE_NAME;
            }
            var binFormatter = new BinaryFormatter();
            try
            {
                using (var file = new FileStream(fileName, FileMode.Open))
                {
                    if (file != null)
                        return (T)binFormatter.Deserialize(file);
                    else return myType;
                }
            }
            catch (Exception)
            {

                return myType;
            }


        }

        private void saveBinFile<T>(T myType)
        {
            string fileName = "unknown.bin";
            if (typeof(T) == typeof(List<CIDocument>))
            {
                fileName = Constants.CONTRACT_FILE_NAME;
            } else if (typeof(T) == typeof(List<Goods>))
            {
                fileName = Constants.GOOD_FILE_NAME;
            }
            else if (typeof(T) == typeof(List<Requisite>))
            {
                fileName = Constants.PARTNER_FILE_NAME;
            }
            else if (typeof(T) == typeof(Settings))
            {
                fileName = Constants.SETTINGS_FILE_NAME;
            }

            var binFormatter = new BinaryFormatter();

            using (var file = new FileStream(fileName, FileMode.OpenOrCreate))
            {
                binFormatter.Serialize(file, myType);
            }
        }
        #endregion

        #region Partners add, remove, edit, setup
        private void pAddBtn_Click(object sender, RoutedEventArgs e)
        {
            var newReq = new Requisite()
            {
                Descrtipion = pPartnerDescriptionTB.Text,
                PartnerName = pPartnerNameTB.Text,
                Address = pPartnerAddressTB.Text,
                OptionalAddress = pPartnerOptionalAddressTB.Text,
                CurrentAccount = pPartnerCurrentAccountTB.Text,
                Bank = pPartnerBankTB.Text,
                PostAndNameOfRespondent = pPartnerPostAndNameOfRespondentTB.Text,
                BasedOn = pPartnerBasedOnTB.Text,
                BankIdentificationCode = pPartnerBankIdentificationTB.Text,
                TaxpayerIdentificationNumber = pPartnerTaxpayerIdentificationNumberTB.Text,
                Oced = pPartnerOcedTB.Text,
                VadNumber = pPartnerVadNumberTB.Text,
                Attorney = pPartnerAttroneyTB.Text,
                Telephones = pPartnerTelephonesTB.Text
            };
            if (MyRequisites.Find(i => i.Descrtipion == newReq.Descrtipion) == null)
            {
                MyRequisites.Add(newReq);
                setUpPartnerComponents();
                saveBinFile(MyRequisites);
            }
            else
            {
                System.Windows.MessageBox.Show($"Запись с названием {newReq.Descrtipion} уже существует. Измените короткое имя партнёра!", "Повторение!!!"); ;
            }

        }

        private void pEditBtn_Click(object sender, RoutedEventArgs e)
        {
            var i = pPartnerDescriptionCB.SelectedIndex;
            if (i < MyRequisites.Count && i > -1)
            {
                MyRequisites[i].Descrtipion = pPartnerDescriptionTB.Text;
                MyRequisites[i].PartnerName = pPartnerNameTB.Text;
                MyRequisites[i].Address = pPartnerAddressTB.Text;
                MyRequisites[i].OptionalAddress = pPartnerOptionalAddressTB.Text;
                MyRequisites[i].CurrentAccount = pPartnerCurrentAccountTB.Text;
                MyRequisites[i].Bank = pPartnerBankTB.Text;
                MyRequisites[i].PostAndNameOfRespondent = pPartnerPostAndNameOfRespondentTB.Text;
                MyRequisites[i].BasedOn = pPartnerBasedOnTB.Text;
                MyRequisites[i].BankIdentificationCode = pPartnerBankIdentificationTB.Text;
                MyRequisites[i].TaxpayerIdentificationNumber = pPartnerTaxpayerIdentificationNumberTB.Text;
                MyRequisites[i].Oced = pPartnerOcedTB.Text;
                MyRequisites[i].VadNumber = pPartnerVadNumberTB.Text;
                MyRequisites[i].Attorney = pPartnerAttroneyTB.Text;
                MyRequisites[i].Telephones = pPartnerTelephonesTB.Text;
                setUpPartnerComponents();
                saveBinFile(MyRequisites);
            }
        }

        private void pRemoveBtn_Click(object sender, RoutedEventArgs e)
        {
            var i = pPartnerDescriptionCB.SelectedIndex;
            if (i < MyRequisites.Count && i > -1)
            {
                MyRequisites.RemoveAt(i);
                setUpPartnerComponents();
                saveBinFile(MyRequisites);
            }
        }
        private void setUpPartnerComponents()
        {
            if (MyRequisites == null) return;
            MyRequisites.Sort((i1, i2) => i1.Descrtipion.CompareTo(i2.Descrtipion));
            var partnersDescription = new string[MyRequisites.Count];

            for (var i = 0; i < MyRequisites.Count(); i++)
            {
                partnersDescription[i] = MyRequisites[i].Descrtipion;
            }
            this.pPartnerDescriptionCB.ItemsSource = partnersDescription;

            //as an exception contract Partners changed here
            partnerNameCB.ItemsSource = partnersDescription;
        }

        private void pPartnerDescriptionCB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var i = pPartnerDescriptionCB.SelectedIndex;
            if (i < MyRequisites.Count && i > -1)
            {
                pPartnerDescriptionTB.Text = MyRequisites[i].Descrtipion;
                pPartnerNameTB.Text = MyRequisites[i].PartnerName;
                pPartnerAddressTB.Text = MyRequisites[i].Address;
                pPartnerOptionalAddressTB.Text = MyRequisites[i].OptionalAddress;
                pPartnerCurrentAccountTB.Text = MyRequisites[i].CurrentAccount;
                pPartnerBankTB.Text = MyRequisites[i].Bank;
                pPartnerPostAndNameOfRespondentTB.Text = MyRequisites[i].PostAndNameOfRespondent;
                pPartnerBasedOnTB.Text = MyRequisites[i].BasedOn;
                pPartnerBankIdentificationTB.Text = MyRequisites[i].BankIdentificationCode;
                pPartnerTaxpayerIdentificationNumberTB.Text = MyRequisites[i].TaxpayerIdentificationNumber;
                pPartnerOcedTB.Text = MyRequisites[i].Oced;
                pPartnerVadNumberTB.Text = MyRequisites[i].VadNumber;
                pPartnerAttroneyTB.Text = MyRequisites[i].Attorney;
                pPartnerTelephonesTB.Text = MyRequisites[i].Telephones;
                setUpPartnerComponents();
            }
        }

        #endregion

        #region Settings save button, setup, dialog boxes
        private void sePathToContractsFolderBtn_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new frm.FolderBrowserDialog())
            {
                frm.DialogResult result = fbd.ShowDialog();

                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    var senderBtn = (Button)sender;
                    if (senderBtn.Name == "sePathToContractsFolderBtn")
                    {
                        sePathToContractsFolderTB.Text = fbd.SelectedPath;
                    }
                    if (senderBtn.Name == "sePathToInvoicesFolderBtn")
                    {
                        sePathToInvoicesFolderTB.Text = fbd.SelectedPath;
                    }
                    if (senderBtn.Name == "sePathToInvoicesExcellFolderBtn")
                    {
                        sePathToInvoicesExcellFolderTB.Text = fbd.SelectedPath;
                    }
                }
            }
        }

        private void seSaveBtn_Click(object sender, RoutedEventArgs e)
        {
            setting.PathToContractsFolder = sePathToContractsFolderTB.Text;
            setting.PathToInvoicesFolder = sePathToInvoicesFolderTB.Text;
            setting.PathToInvoicesExcellFolder = sePathToInvoicesExcellFolderTB.Text;
            if (int.TryParse(seCurrentContractNumberTB.Text, out int i))
            {
                setting.CurrentContractNumber = i;
            }
            if (int.TryParse(seCurrentVadTB.Text, out int v))
            {
                setting.currentVad = v;
            }
            saveBinFile<Settings>(setting);
        }
        private void setUpSettings()
        {
            sePathToContractsFolderTB.Text = setting.PathToContractsFolder;
            sePathToInvoicesFolderTB.Text = setting.PathToInvoicesFolder;
            sePathToInvoicesExcellFolderTB.Text = setting.PathToInvoicesExcellFolder;
            seCurrentContractNumberTB.Text = setting.CurrentContractNumber.ToString();
            seCurrentVadTB.Text = setting.currentVad.ToString();
        }
        #endregion


        #region MyContracts priceVad, setUP, setUPfileName, PartnerNameSelectionChanged, addButton, MinusButton, SwapButton


        private void calculateVadBtn_Click(object sender, RoutedEventArgs e)
        {
            calculateVadTB.Text = (Constants.StringToDouble(calculateVadTB.Text) / ((setting.currentVad + 100.0m) / 100)).ToString();
        }
        private void setUpContractsComponents()
        {
            if (MyContracts == null) return;
            // edit Contract option
            MyContracts.Sort((i1, i2) => i1.Number.CompareTo(i2.Number));
            var contractNumbers = new int[MyContracts.Count];

            for (var i = 0; i < MyContracts.Count; i++)
            {
                contractNumbers[i] = MyContracts[i].Number;
            }
            this.editContractCB.ItemsSource = contractNumbers;

            // contract number option
            contractNumberTB.Text = setting.CurrentContractNumber.ToString();
            setUpFileName();

            // contract materialCB
            materialCB.ItemsSource = CIDocument.getArrayOfMaterials();

            // set Up Swap Combo Boxes
            setUpSwapCBs();

        }
        private void setUpFileName()
        {
            if (MyRequisites == null) return;
            if (partnerNameCB == null) return;
            if (contractDateDP == null) return;
            int pNameCB = partnerNameCB.SelectedIndex;
            if (pNameCB < 0) return;
            fileNameTB.Text = $"№{contractNumberTB.Text}_{MyRequisites[partnerNameCB.SelectedIndex].Descrtipion}_{contractDateDP.SelectedDate.Value.Year.ToString()}";
        }

        private void partnerNameCB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int i = partnerNameCB.SelectedIndex;
            if (i < 0) return;
            setUpFileName();
            aboutPartnerTBl.Text = MyRequisites[i].ToString();
        }

        private void contractNumberTB_TextChanged(object sender, TextChangedEventArgs e)
        {
            setUpFileName();
        }

        private void contractListAddBtn_Click(object sender, RoutedEventArgs e)
        {
            MyGoodUIs.Add(setNewGoodsUI());
            contractDataContainerLV.Items.Add(MyGoodUIs.Last());
            setUpSwapCBs();
        }
        private void contractSwitchBtn_Click(object sender, RoutedEventArgs e)
        {
            int s1 = contractSwitch1CB.SelectedIndex;
            int s2 = contractSwitch2CB.SelectedIndex;
            if (s1 < 0 || s2 < 0) return;
            if (s1 == s2) return;
            MyGoodUIs[s1].contractNumberTB.Text = contractSwitch2CB.SelectedItem.ToString();
            MyGoodUIs[s2].contractNumberTB.Text = contractSwitch1CB.SelectedItem.ToString();
            MyGoodUIs.Sort((v1, v2) => v1.contractNumberTB.Text.CompareTo(v2.contractNumberTB.Text));
            contractDataContainerLV.Items.Clear();
            foreach (var item in MyGoodUIs)
            {
                contractDataContainerLV.Items.Add(item);
            }
        }

        private void contractListMinusBtn_Click(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(contractListMinusTB.Text, out int i))
            {
                if (i <= MyGoodUIs.Count && i > 0)
                {
                    MyGoodUIs.RemoveAt(i - 1);
                    contractDataContainerLV.Items.Clear();
                    var j = 1;
                    foreach (var item in MyGoodUIs)
                    {
                        item.contractNumberTB.Text = j.ToString();
                        j++;
                        contractDataContainerLV.Items.Add(item);
                    }
                    setUpSwapCBs();
                }
            }
        }

        private GoodsUI setNewGoodsUI()
        {
            var gUI = new GoodsUI();
            gUI.contractGoodVadTB.Text = setting.currentVad.ToString();
            gUI.contractNumberTB.Text = (MyGoodUIs.Count + 1).ToString();

            return gUI;
        }

        private void setUpSwapCBs()
        {
            if (MyGoodUIs == null) return;
            var i = MyGoodUIs.Count;
            if (i < 1) return;
            if (contractSwitch1CB == null || contractSwitch2CB == null) return;
            var itemSource = new int[i];
            for (var j = 0; j < i; j++)
            {
                itemSource[j] = j + 1;
            }
            contractSwitch1CB.ItemsSource = itemSource;
            contractSwitch2CB.ItemsSource = itemSource;
        }

        private void contractOKBtn_Click(object sender, RoutedEventArgs e)
        {
            if (MyContracts == null) return;

            var contract = new CIDocument()
            {
                Number = Constants.StringToInt(contractNumberTB.Text),
                Date = contractDateDP.SelectedDate.GetValueOrDefault(),
                Material = (Material)materialCB.SelectedIndex,
                Requisite = MyRequisites[partnerNameCB.SelectedIndex],
                Goods = documentGoods(),
                FileName = fileNameTB.Text

            };
            var i = MyContracts.FindIndex(j => j.Number == contract.Number);
            if (i == -1)
            {
                MyContracts.Add(contract);

            } else
            {
                MyContracts[i] = contract;
            }
            saveBinFile<List<CIDocument>>(MyContracts);

            // set next number to settings
            if (setting.CurrentContractNumber <= contract.Number)
            {
                setting.CurrentContractNumber = contract.Number + 1;
                saveBinFile<Settings>(setting);
                setUpSettings();
            }
            setUpContractsComponents();
            publishContractToWord(contract);
        }
        private List<Goods> documentGoods()
        {
            var g = new List<Goods>();
            foreach (GoodsUI item in MyGoodUIs)
            {
                var goods = new Goods()
                {
                    Name = item.contractGoodNameTB.Text,
                    Format = item.contractGoodFormatTB.Text,
                    Volume = item.contractGoodVolumeTB.Text,
                    Quantity = Constants.StringToInt(item.contractGoodQuantityTB.Text),
                    Price = Constants.StringToDouble(item.contractGoodPriceTB.Text),
                    VadInPercent = Constants.StringToInt(item.contractGoodVadTB.Text)
                };
                g.Add(goods);
            }
            return g;

        }

        private void editContractBtn_Click(object sender, RoutedEventArgs e)
        {
            var i = editContractCB.SelectedIndex;
            if (i >= MyContracts.Count || i < 0) return;
            contractNumberTB.Text = MyContracts[i].Number.ToString();
            contractDateDP.SelectedDate = MyContracts[i].Date;
            partnerNameCB.SelectedIndex = MyRequisites.FindIndex(j => j.Descrtipion == MyContracts[i].Requisite.Descrtipion);
            materialCB.SelectedIndex = MyContracts[i].getMaterialInt();
            MyGoodUIs.Clear();
            contractDataContainerLV.Items.Clear();
            foreach (var item in MyContracts[i].Goods)
            {
                var nGUI = setNewGoodsUI();
                nGUI.contractGoodNameTB.Text = item.Name;
                nGUI.contractGoodFormatTB.Text = item.Format;
                nGUI.contractGoodVolumeTB.Text = item.Volume;
                nGUI.contractGoodQuantityTB.Text = item.Quantity.ToString();
                nGUI.contractGoodPriceTB.Text = item.Price.ToString();
                nGUI.contractGoodVadTB.Text = item.VadInPercent.ToString();
                nGUI.contractGoodSumTBl.Text = item.sum().ToString();
                nGUI.contractGoodSumVadTBl.Text = item.sumWithVad().ToString();
                MyGoodUIs.Add(nGUI);
                contractDataContainerLV.Items.Add(MyGoodUIs.Last());
            }
            setUpSwapCBs();
        }
        #endregion

        #region Goods Slection, addBtn, editBtn, RemoveBtn


        private void gDescriptionCB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var i = gDescriptionCB.SelectedIndex;
            if (i < 0 || i > MyGoods.Count) return;
            gDescriptionTB.Text = MyGoods[i].Description;
            gNameTB.Text = MyGoods[i].Name;
            gFormatTB.Text = MyGoods[i].Format;
            gVolumeTB.Text = MyGoods[i].Volume;
            gPriceTB.Text = MyGoods[i].Price.ToString();
            gVadInPercentTB.Text = MyGoods[i].VadInPercent.ToString();
        }

        private void gAddBtn_Click(object sender, RoutedEventArgs e)
        {
            var good = new Goods()
            {
                Description = gDescriptionTB.Text,
                Name = gNameTB.Text,
                Format = gFormatTB.Text,
                Volume = gVolumeTB.Text,
                Price = Constants.StringToDouble(gPriceTB.Text)
            };


            if (int.TryParse(gVadInPercentTB.Text, out int i))
            {
                good.VadInPercent = i;
            }

            if (MyGoods.Find(j => j.Description == good.Description) == null)
            {
                MyGoods.Add(good);
                setUpGoodsComponents();
                saveBinFile(MyGoods);
                updateMyGoodsUIcbComponent();
            }
            else
            {
                System.Windows.MessageBox.Show($"Запись с названием {good.Description} уже существует. Измените описание товара!", "Повторение!!!"); ;
            }


        }

        private void gEditBtn_Click(object sender, RoutedEventArgs e)
        {
            var i = gDescriptionCB.SelectedIndex;
            if (i < MyGoods.Count && i > -1)
            {
                MyGoods[i].Description = gDescriptionTB.Text;
                MyGoods[i].Name = gNameTB.Text;
                MyGoods[i].Format = gFormatTB.Text;
                MyGoods[i].Volume = gVolumeTB.Text;
                MyGoods[i].Price = Constants.StringToDouble(gPriceTB.Text);

                if (int.TryParse(gVadInPercentTB.Text, out int j))
                {
                    MyGoods[i].VadInPercent = j;
                }

                setUpGoodsComponents();
                saveBinFile(MyGoods);
                updateMyGoodsUIcbComponent();
            }
        }

        private void gRemoveBtn_Click(object sender, RoutedEventArgs e)
        {
            var i = gDescriptionCB.SelectedIndex;
            if (i < MyGoods.Count && i > -1)
            {
                MyGoods.RemoveAt(i);
                setUpGoodsComponents();
                saveBinFile(MyGoods);
                updateMyGoodsUIcbComponent();
            }
        }

        private void setUpGoodsComponents()
        {
            if (MyGoods == null) return;
            MyGoods.Sort((i1, i2) => i1.Description.CompareTo(i2.Description));
            var goodsDescription = new string[MyGoods.Count];

            for (var i = 0; i < MyGoods.Count; i++)
            {
                goodsDescription[i] = MyGoods[i].Description;
            }
            gDescriptionCB.ItemsSource = goodsDescription;
        }

        private void updateMyGoodsUIcbComponent()
        {
            foreach (var item in MyGoodUIs)
            {
                item.setUpGoodsComponents();
            }
        }

        #endregion

        private void publishContractToWord(CIDocument contract)
        {
            var wordApp = new word.Application();
            wordApp.Visible = true;
            var path = "D:\\Data\\dogovorNDS.dotm";

            var wDoc = wordApp.Documents.Add(path);
            wDoc.Variables["dognomer"].Value = contract.Number.ToString();
            wDoc.Variables["data1"].Value = contract.Date.ToString("«dd» MMMM yyyy");
            wDoc.Variables["nazvanieFirmi"].Value = contract.Requisite.PartnerName;
            wDoc.Variables["licoFirmi"].Value = contract.Requisite.PostAndNameOfRespondent;
            wDoc.Variables["osnovanieFirmi"].Value = contract.Requisite.BasedOn;

            var wTable1 = wDoc.Tables[1];
            int row = 2;
            var allSum = 0.0m;
            var allSumVat = 0.0m;
            foreach (Goods item in contract.Goods)
            {
                wordTable(wTable1, row, item);
                allSum += item.sum();
                allSumVat += item.sumWithVad();
                if (contract.Goods.Count > (row - 1))
                {
                    wTable1.Rows.Add(wTable1.Rows.Last);
                }
                row++;
            }
            wDoc.Variables["itogZ"].Value = doubleToString(allSum);
            decimal vat = allSumVat - allSum;
            if (vat != 0)
            {
                wDoc.Variables["onlyNDS"].Value = doubleToString(vat);
            } else
            {
                wDoc.Variables["onlyNDS"].Value = "Без НДС";
            }
            
            wDoc.Variables["nds"].Value = doubleToString(allSumVat);


            wDoc.Variables["summaPropisyu"].Value = Summa.Пропись(allSumVat, Currenc.Сум, Case.Первая).Replace(" 00 тийин", "");

            
            if (allSumVat != allSum)
            {
                wDoc.Variables["summaPropisyu"].Value += " в т.ч. НДС " + doubleSumAndTiyin(allSumVat - allSum);
            }
            else
            {
                wDoc.Variables["summaPropisyu"].Value += ".";
            }

            wDoc.Variables["data2"].Value = contract.Date.ToString("dd.MM.yyyy");
            wDoc.Variables["material"].Value = CIDocument.getArrayOfMaterials()[contract.getMaterialInt()];

            wDoc.Variables["adresFirmi"].Value = contract.Requisite.ToString();



            wDoc.Fields.Update();

            wDoc.Fields.UpdateSource();
            wDoc.Fields.ToggleShowCodes();
            wDoc.Fields.Unlink();

            wDoc.SaveAs2(setting.PathToContractsFolder + "\\" + contract.FileName + ".docx");


        }
        private bool isDoubleFloating(decimal d)
        {
            decimal myInt = Math.Floor(d);
            decimal myFloat = d - myInt;
            if (myFloat > 0)
            {
                return true;
            } else
            {
                return false;
            }
        }
        private string doubleToString (decimal d)
        {
            
            decimal myInt = Math.Floor(d);
            decimal myFloat = d - myInt;

            string s = myInt.ToString("### ### ### ### ### ###").TrimStart();

            if (myFloat > 0 )
            {
                s += myFloat.ToString().Replace("0,", ",");
            } 



            return s;
        }
        private string doubleToString2(decimal d)
        {

            decimal myInt = Math.Floor(d);
            decimal myFloat = d - myInt;

            string s = myInt.ToString();

            if (myFloat > 0)
            {
                s += myFloat.ToString().Replace("0,", ",");
            }



            return s;
        }
        private string doubleSumAndTiyin (decimal d) {
            var s = "";
            s = d.ToString("N", CultureInfo.CreateSpecificCulture("fr-CA"));
            var k = "";
            int zap = s.IndexOf(',');
            if (zap >= 0)
            {
                k = s.Substring(zap);
                k = k.Replace(",", "");
                s = s.Substring(0, zap);
            }
            if (k.Replace("0","").Length > 0)
            {
                return s + " сум " + k + " тийин.";
            } else
            {
                return s + " сум.";
            }

            
        }
        private void wordTable(word.Table t, int row, Goods g)
        {
            t.Cell(row, 1).Range.Text = (row-1).ToString();
            t.Cell(row, 2).Range.Text = g.Name;
            t.Cell(row, 3).Range.Text = g.Format;
            t.Cell(row, 4).Range.Text = g.Volume;
            t.Cell(row, 5).Range.Text = g.Quantity.ToString("### ### ### ### ### ###").TrimStart();
            t.Cell(row, 6).Range.Text = doubleToString(g.Price);
            t.Cell(row, 7).Range.Text = doubleToString(g.sum());
            decimal vat = g.sumWithVad() - g.sum();
            if (vat == 0)
            {
                t.Cell(row, 8).Range.Text = "Без НДС";
            } else
            {
                t.Cell(row, 8).Range.Text = doubleToString(vat);
            }
            
            t.Cell(row, 9).Range.Text = doubleToString(g.sumWithVad());

            t.Cell(row, 1).Range.Bold = 0;
            t.Cell(row, 2).Range.Bold = 0;
            t.Cell(row, 3).Range.Bold = 0;
            t.Cell(row, 4).Range.Bold = 0;
            t.Cell(row, 5).Range.Bold = 0;
            t.Cell(row, 6).Range.Bold = 0;
            t.Cell(row, 7).Range.Bold = 0;
            t.Cell(row, 8).Range.Bold = 0;
            t.Cell(row, 9).Range.Bold = 0;
        }
        private void publishInvoiceToWord(CIDocument contract)
        {

        }
        private void publishInvoiceToExcel(CIDocument contract)
        {

        }

        #region Invoice setups

        
        private void invContractNumber_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var i = invContractNumber.SelectedIndex;
            if (i >= MyContracts.Count || i < 0) return;

            invPartner.SelectedIndex = MyRequisites.FindIndex(j => j.Descrtipion == MyContracts[i].Requisite.Descrtipion);
            invContractDateDP.SelectedDate = MyContracts[i].Date;
            invAttorneyPersonTB.Text = MyRequisites[invPartner.SelectedIndex].Attorney;

            MyInvGoods.Clear();
            invDataContainerLV.Items.Clear();
            foreach (var item in MyContracts[i].Goods)
            {
                var nGUI = setNewInvGoods();
                nGUI.invGoodNameTB.Text = item.Name;
                nGUI.invGoodQuantityTB.Text = item.Quantity.ToString();
                nGUI.invGoodPriceTB.Text = item.Price.ToString();
                nGUI.invGoodVadTB.Text = item.VadInPercent.ToString();
                nGUI.invGoodSumTBl.Text = item.sum().ToString();
                nGUI.invGoodSumVadTBl.Text = item.sumWithVad().ToString();
                MyInvGoods.Add(nGUI);
                invDataContainerLV.Items.Add(MyInvGoods.Last());
            }
            setUpInvSwapCBs();
        }

        private void invListAddBtn_Click(object sender, RoutedEventArgs e)
        {
            MyInvGoods.Add(setNewInvGoods());
            invDataContainerLV.Items.Add(MyInvGoods.Last());
            setUpInvSwapCBs();
        }

        private void setUpInvSwapCBs()
        {
            if (MyInvGoods == null) return;
            var i = MyInvGoods.Count;
            if (i < 1) return;
            if (invSwap1CB == null || invSwap2CB == null) return;
            var itemSource = new int[i];
            for (var j = 0; j < i; j++)
            {
                itemSource[j] = j + 1;
            }
            invSwap1CB.ItemsSource = itemSource;
            invSwap2CB.ItemsSource = itemSource;
        }

        private InGoods setNewInvGoods()
        {
            var gUI = new InGoods();
            gUI.invGoodVadTB.Text = setting.currentVad.ToString();
            gUI.invGoodNumberTB.Text = (MyInvGoods.Count + 1).ToString();

            return gUI;
        }

        private void invListMinusBtn_Click(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(invListMinusTB.Text, out int i))
            {
                if (i <= MyInvGoods.Count && i > 0)
                {
                    MyInvGoods.RemoveAt(i - 1);
                    invDataContainerLV.Items.Clear();
                    var j = 1;
                    foreach (var item in MyInvGoods)
                    {
                        item.invGoodNumberTB.Text = j.ToString();
                        j++;
                        invDataContainerLV.Items.Add(item);
                    }
                    setUpInvSwapCBs();
                }
            }
        }

        private void invSwapBtn_Click(object sender, RoutedEventArgs e)
        {
            int s1 = invSwap1CB.SelectedIndex;
            int s2 = invSwap2CB.SelectedIndex;
            if (s1 < 0 || s2 < 0) return;
            if (s1 == s2) return;
            MyInvGoods[s1].invGoodNumberTB.Text = invSwap2CB.SelectedItem.ToString();
            MyInvGoods[s2].invGoodNumberTB.Text = invSwap1CB.SelectedItem.ToString();
            MyInvGoods.Sort((v1, v2) => v1.invGoodNumberTB.Text.CompareTo(v2.invGoodNumberTB.Text));
            invDataContainerLV.Items.Clear();
            foreach (var item in MyInvGoods)
            {
                invDataContainerLV.Items.Add(item);
            }
        }

        #endregion

        private void invExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            var eApp = new excel.Application();
            var path = "D:\\Data\\factura.xltm";
            var eDoc = eApp.Workbooks.Add(path);
            int row = 10;
            excel.Worksheet wSheet = (excel.Worksheet)eApp.ActiveSheet;
            foreach (InGoods item in MyInvGoods)
            {

                wSheet.Cells[row, 1] = item.invGoodNumberTB.Text;
                wSheet.Cells[row, 2] = item.invGoodNameTB.Text;
                wSheet.Cells[row, 3] = "1";
                wSheet.Cells[row, 4] = item.invGoodQuantityTB.Text;
                wSheet.Cells[row, 5] = item.invGoodPriceTB.Text;
                wSheet.Cells[row, 6] = doubleToString2(item.sum());
                wSheet.Cells[row, 7] = item.invGoodVadTB.Text;
                wSheet.Cells[row, 8] = doubleToString2(item.sumWithVad() - item.sum());
                wSheet.Cells[row, 9] = doubleToString2(item.sumWithVad());
                wSheet.Cells[row, 10] = "0";
                wSheet.Cells[row, 11] = "0";
                wSheet.Cells[row, 12] = doubleToString2(item.sumWithVad());
                row++;

            }
            eDoc.SaveAs(setting.PathToInvoicesExcellFolder + "\\D№" + invContractNumber.Text + "_" + invContractDateDP.SelectedDate.Value.ToString("yyyy") + ".xlsx");
            eDoc.Close();
            eApp.Quit();
            MessageBox.Show("Создан Excel файл");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            List<CIDocument> contracts = new List<CIDocument>();
            List<FindContract> findContracts = new List<FindContract>();
			if (findContractTB.Text == "")
			{
                contracts = MyContracts;

            } else
            {
                if (findContractPNRB.IsChecked == true)
                {
                    contracts = MyContracts.FindAll(cI => cI.Requisite.PartnerName == findContractTB.Text);

                }
                else if (findContractCIRB.IsChecked == true)
                {
                    contracts = MyContracts.FindAll(cI => cI.Goods.Find(g => g.Name.Contains(findContractTB.Text)) != null);
                }
                else if (findContractDRB.IsChecked == true)
                {
                    contracts = MyContracts.FindAll(cI => cI.Requisite.Descrtipion == findContractTB.Text);
                }
            }
            if (contracts.Count > 0)
            {

                foreach (var item in contracts)
                {
                    FindContract findContract = new FindContract();
                    Decimal d = 0;
                    Decimal dv = 0;
                    foreach (var goods in item.Goods)
                    {
                        d += goods.sum();
                        dv += goods.sumWithVad();
                    }
                    findContract.MyNumber = item.Number;
                    findContract.findContractTB.Text = item.Requisite.PartnerName + ". Договор №" + item.Number.ToString() + " от " + item.Date.ToString("dd.MM.yyyy") + " общая сумма договора: " + d.ToString() + " общая сумма договора с НДС: " + dv.ToString();
                    findContract.fcContracts.ItemsSource = item.Goods;
                    findContracts.Add(findContract);
                }
            }
            findContractLB.ItemsSource = findContracts;
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            FindContract con = new FindContract();
            if (findContractLB.SelectedItem != null)
            con = (FindContract)findContractLB.SelectedItem;
            if (con != null)
            {
                MyContracts.Remove(MyContracts.Find(item => item.Number == con.MyNumber));
            }

            saveBinFile<List<CIDocument>>(MyContracts);
            setUpContractsComponents();
            Button_Click(null, null);
        }
    }
}
