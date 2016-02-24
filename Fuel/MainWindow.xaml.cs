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

        public MainWindow()
        {
            InitializeComponent();
            var patch = @"company.json";
            if (File.Exists(patch))
            {
                string fileC = File.ReadAllText(patch, UTF8Encoding.UTF8);
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
            OPF.Filter = "Файлы excel|*.xls;*.xlsx";
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
            var patch = @"company.json";
            if (!File.Exists(patch))
            {
                File.Create(patch);
            }
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(patch, false))
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

        //[JsonObject(MemberSerialization.OptIn)]
        class cellExcel
        {

           // [JsonProperty("CellCard")]
            public int CellCard { get; set; }
            //[JsonProperty("CellServicePoint")]
            public int CellServicePoint { get; set; }
           // [JsonProperty("CellCompany")]
            public int CellCompany { get; set; }
           // [JsonProperty("CellAdressAzs")]
            public int CellAdressAzs { get; set; }
           // [JsonProperty("CellDateFill")]
            public int CellDateFill { get; set; }
           // [JsonProperty("CellOperation")]
            public int CellOperation { get; set; }
           // [JsonProperty("CellTypeFuel")]
            public int CellTypeFuel { get; set; }
           // [JsonProperty("CellCountFuel")]
            public int CellCountFuel { get; set; }
        }

        private void radioBash_Checked(object sender, RoutedEventArgs e)
        {
            var patch = @"optionxls.json";            
            cellExcel cell= new cellExcel();
            Newtonsoft.Json.Linq.JObject arr = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(patch, Encoding.UTF8));
            cell = JsonConvert.DeserializeObject<cellExcel>(arr["Bash"].ToString());           
            adres.Text = cell.CellAdressAzs.ToString();
        }

        private void radioLuk_Checked(object sender, RoutedEventArgs e)
        {
            var patch = @"optionxls.json";
            cellExcel cell = new cellExcel();
            Newtonsoft.Json.Linq.JObject arr = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(patch, Encoding.UTF8));
            cell = JsonConvert.DeserializeObject<cellExcel>(arr["Luk"].ToString());
            adres.Text = cell.CellAdressAzs.ToString();
        }
    }
}