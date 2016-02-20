using Microsoft.Win32;
using System.Windows;
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

namespace Fuel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        ObservableCollection<Company> company = new ObservableCollection<Company>();
       //Company[] str;
        public MainWindow()
        {
            InitializeComponent();
            string fileC = File.ReadAllText(@"company.json", UTF8Encoding.UTF8);
            Company[] str = JsonConvert.DeserializeObject<Company[]>(fileC);

            foreach (Company s in str)
            {
                company.Add(s);
            }
            //str = JsonConvert.DeserializeObject<Company[]>(fileC);

            //Company s = Array.Find(str, r => r.NameBash == "Регион 1");

            //textN.Text = s.Name;

            CompanyGrid.ItemsSource = company;
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog OPF = new OpenFileDialog();
            OPF.Filter = "Файлы excel|*.xls;*.xlsx";
            OPF.Title = "Выберите файл excel";
            if (OPF.ShowDialog() == true)
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
                Company newC = new Company(n,nb,nl);
                Company c = CompanyGrid.SelectedItem as Company;
                company.Remove(c);
                company.Add(newC);
            }
        }

        private void CompanyGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            upDel f = new upDel();
            if (f.ShowDialog() == true)
            {
                addcompany tr = new addcompany();
                Company c = CompanyGrid.SelectedItem as Company;
                tr.name.Text = c.Name;
                tr.nameBash.Text = c.NameBash;
                tr.nameLuk.Text = c.NameLuk;
                if(tr.ShowDialog() == true)
                {
                    updateCompany(tr.name.Text, tr.nameBash.Text, tr.nameLuk.Text);
                }
                tr.Close();

            } else {
                deleteCompany();                    
            }
            CompanyGrid.SelectedIndex = -1;
            f.Close();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"company.json", false))
            {
                file.WriteLine(JsonConvert.SerializeObject(company));
            }
           
        }
    }
}
