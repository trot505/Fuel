using Microsoft.Win32;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;

namespace Fuel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xamld
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        ObservableCollection<Company> company = new ObservableCollection<Company>();
        private const string DIR_SEPARATOR = @"\";
        private string companyPatch = @"company.json";
        private string cellPatch = @"optionxls.json";
        cellExcel cell = new cellExcel();
        List<Out> outArr = new List<Out>();

        public MainWindow()
        {
            InitializeComponent();            
            if (File.Exists(companyPatch))
            {
                string fileC = File.ReadAllText(companyPatch, UTF8Encoding.UTF8);
                company = JsonConvert.DeserializeObject<ObservableCollection<Company>>(fileC);
                CompanyGrid.ItemsSource = company;
                CompanyGrid.UnselectAllCells();
            } else {
                System.Windows.MessageBox.Show("Файл со списком организайи не существует!", "ВНИМАНИЕ", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            
        }
        
        private void button_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog OPF = new System.Windows.Forms.OpenFileDialog();
            OPF.Filter = "Файлы excel|*.xls;*.xlsx;*.xlsm";
            OPF.Title = "Выберите файл excel";
            if (OPF.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName.Text = OPF.FileName;
                //var exelP = new Microsoft.Office.Interop.Excel.Application();
                //var WorkBookExcel = exelP.Workbooks.Open(OPF.FileName);
                //var WorkSheetExcel = (Microsoft.Office.Interop.Excel.Worksheet)WorkBookExcel.Sheets[1];
                //object[,] arr = (object[,])WorkSheetExcel.UsedRange.Value;
                //dataEx.ItemsSource = arr;
            }

        }




        private void add_Click(object sender, RoutedEventArgs e)
        {
            addcompany tr = new addcompany();
            if (tr.ShowDialog() == true)
            {
                company.Add(new Company(tr.name.Text, tr.nameBash.Text, tr.nameLuk.Text));
                tr.Close();
            }
            searchText.Text = "";

        }

        private void deleteCompany()
        {
            if (CompanyGrid.SelectedIndex >= 0)
            {
                Company c = CompanyGrid.SelectedItem as Company;
                company.Remove(c);
            }
        }

        private void updateCompany(string n, string nb, string nl)
        {
            if (CompanyGrid.SelectedIndex >= 0)
            {
                Company newC = new Company(n, nb, nl);
                Company c = CompanyGrid.SelectedItem as Company;
                company.Remove(c);
                company.Add(newC);
            }
        }

        private void searchText_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            var s = company.Where<Company>(r => (r.Name + r.NameBash + r.NameLuk).ToLower().Contains(searchText.Text.ToLower()));
            CompanyGrid.ItemsSource = (s.SequenceEqual(company)) ? company:s;
            
        }

        private void updateC_Click(object sender, RoutedEventArgs e)
        {
            if (CompanyGrid.SelectedIndex >= 0)
            {
                addcompany tr = new addcompany();
                Company c = CompanyGrid.SelectedItem as Company;
                tr.name.Text = c.Name;
                tr.nameBash.Text = c.NameBash;
                tr.nameLuk.Text = c.NameLuk;
                if (tr.ShowDialog() == true)
                {
                    updateCompany(tr.name.Text, tr.nameBash.Text, tr.nameLuk.Text);
                }
                tr.Close();
            }
            else {
                System.Windows.MessageBox.Show("Необходимо выбрать фирму из списка.", "ВНИМЕНИЕ !!!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            CompanyGrid.UnselectAllCells();
        }

        private void delC_Click(object sender, RoutedEventArgs e)
        {
            if (CompanyGrid.SelectedIndex >= 0)
            {
                var res = System.Windows.MessageBox.Show("Вы собираетесьб удалить запись: \n"
                + (CompanyGrid.SelectedItem as Company).Name, "УДАЛЕНИЕ", MessageBoxButton.YesNo);
                if (res == MessageBoxResult.Yes)
                {
                    deleteCompany();
                    System.Windows.MessageBox.Show("Удаление записи прошло успешно", "СООБЩЕНИЕ", MessageBoxButton.OK);
                }
            }
            else {
                System.Windows.MessageBox.Show("Необходимо выбрать фирму из списка.", "ВНИМЕНИЕ !!!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            CompanyGrid.UnselectAllCells();
            searchText.Text = "";
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            
            if (!File.Exists(companyPatch))
            {
                File.Create(companyPatch);
            }
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(companyPatch, false))
            {
                file.WriteLine(JsonConvert.SerializeObject(company));
            }

        }

        private void folderBtn_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            if (folder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folderPatch.Text = folder.SelectedPath;
            }
        }

        private void clearSearch_Click(object sender, RoutedEventArgs e)
        {
            searchText.Text = "";
            CompanyGrid.ItemsSource = company;
        }

        private void CellText()
        {
            card.Text = cell.CellCard.ToString();
            holder.Text = cell.CellCompany.ToString();
            azs.Text = cell.CellAzs.ToString();
            adres.Text = cell.CellAdressAzs.ToString();
            date.Text = cell.CellDateFill.ToString();
            fuelT.Text = cell.CellFuelT.ToString();
            countF.Text = cell.CellCountF.ToString();
            operation.Text = cell.CellOperation.ToString();
            first.Text = cell.FirstRow.ToString();
            last.Text = cell.LastRow.ToString();
            folderPatch.Text = cell.FolderPatch.ToString();
            folderMonth.SelectedIndex = cell.FolderMonth;
        }

        private void radioBash_Checked(object sender, RoutedEventArgs e)
        {
           
            Newtonsoft.Json.Linq.JObject arr = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(cellPatch, Encoding.UTF8));
            cell = JsonConvert.DeserializeObject<cellExcel>(arr["Bash"].ToString());

            CellText();
            
        }

        private void radioLuk_Checked(object sender, RoutedEventArgs e)
        {
            
            Newtonsoft.Json.Linq.JObject arr = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(cellPatch, Encoding.UTF8));
            cell = JsonConvert.DeserializeObject<cellExcel>(arr["Luk"].ToString());
            adres.Text = cell.CellAdressAzs.ToString();

            CellText();
        }



        private void parseBtn_Click(object sender, RoutedEventArgs e)
        {

            if (File.Exists(fileName.Text))
            {
                using (ExcelPackage pack = new ExcelPackage(new FileInfo(fileName.Text)))
                {
                    ExcelWorksheet page = pack.Workbook.Worksheets[cell.ListExl];

                    for (int i = cell.LastRow; i >= cell.FirstRow; i--)
                    {
                        var c = page.Cells[i, cell.CellCard].Value.ToString(); var s = page.Cells[i, cell.CellAzs].Value.ToString();
                        var a = page.Cells[i, cell.CellAdressAzs].Value.ToString(); var d = page.Cells[i, cell.CellDateFill].Value.ToString();
                        var o = page.Cells[i, cell.CellOperation].Value.ToString(); var t = page.Cells[i, cell.CellFuelT].Value.ToString();
                        var co = page.Cells[i, cell.CellCountF].Value.ToString(); var n = page.Cells[i, cell.CellCompany].Value.ToString();

                        outArr.Add(new Out(c, s, a, d, o, t, co, n));
                    }
                    pack.Dispose();
                }
            }

            string provider = (radioBash.IsChecked.Value) ? "Башнефть" : (radioLuk.IsChecked.Value) ? "Лукойл" : "";
            string outDir = folderPatch.Text + DIR_SEPARATOR + folderMonth.Text + DIR_SEPARATOR + provider;
            if (!Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }
            var oneC = outArr.GroupBy(f => f.NameCompany).Distinct();
            foreach (var row in oneC)
            {
                var s = (radioBash.IsChecked.Value) ? company.Where(c => c.NameBash.ToLower() == row.Key.ToLower()).Select(k => k.Name) 
                    : (radioLuk.IsChecked.Value) ? company.Where(c => c.NameLuk.ToLower() == row.Key.ToLower()).Select(k => k.Name) 
                    : company.Where(c => c.Name.ToLower() == row.Key.ToLower()).Select(k => k.Name);
               
                if (s.Count() == 1)
                {
                    string nameFile = s.ElementAt(0).ToString();
                } else if (s.Count() > 1)
                {
                   var res = System.Windows.MessageBox.Show("В списке оргинизаций имеются \nдублированные записи.\nПерейти к редактированию", "", MessageBoxButton.OK,MessageBoxImage.Information);
                    if(res == MessageBoxResult.OK)
                    {
                        CompanyTab.IsSelected = true;
                        searchText.Text = row.Key;
                        break;
                    }
                } else {
                    var res = System.Windows.MessageBox.Show("В списке оргинизаций отсутствует взаимосвязь \n с \""+ row.Key +"\".\nПерейти к добавлению организации", "", MessageBoxButton.OK, MessageBoxImage.Information);
                    if(res == MessageBoxResult.OK)
                    {
                        addcompany addC = new addcompany();
                        if (radioBash.IsChecked.Value)
                        {
                            addC.nameBash.Text = row.Key;
                        }
                        if (radioLuk.IsChecked.Value)
                        {
                            addC.nameLuk.Text = row.Key;
                        }
                        if (addC.ShowDialog() == true)
                        {
                            company.Add(new Company(addC.name.Text, addC.nameBash.Text, addC.nameLuk.Text));
                            addC.Close();
                        }
                    }
                    break;
                }
                
                ExcelPackage pack = new ExcelPackage(new FileInfo(outDir + DIR_SEPARATOR + s.ElementAt(0) + ".xlsx"));
                var str = outArr.Where(r => r.NameCompany == row.Key);
                //dg.ItemsSource = str;
                ExcelWorksheet page = pack.Workbook.Worksheets.Add("Отчет по картам");
                page.Column(1).Style.Font.UnderLine = true;
                int i = 1;
                foreach (var r in str)
                {
                    page.Cells[i, 1].Value = r.Card;
                    page.Cells[i, 2].Value = r.DateFill;
                    page.Cells[i, 3].Value = r.Azs;
                    page.Cells[i, 4].Value = r.AdressAzs;
                    page.Cells[i, 5].Value = r.TypeFuel;
                    page.Cells[i, 6].Value = r.Operation;
                    page.Cells[i, 7].Value = r.CountFuel;
                    

                    i++;
                }
                pack.Save();
                pack.Dispose();
            }
               
            


        }

        
    }
}