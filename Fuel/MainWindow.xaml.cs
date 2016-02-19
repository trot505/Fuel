using Microsoft.Win32;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using System.Threading.Tasks;
using Newtonsoft.Json;


namespace Fuel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        Company[] str;
        public MainWindow()
        {
            InitializeComponent();
            string fileC = File.ReadAllText(@"company.json", UTF8Encoding.UTF8);

            Company[] str = JsonConvert.DeserializeObject<Company[]>(fileC);

            Company s = Array.Find(str, r => r.NameBash == "Регион 1");

            textN.Text = s.Name;
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog OPF = new OpenFileDialog();
            OPF.Filter = "Файлы excel|*.xls;*.xlsx";
            OPF.Title = "Выберите файл excel";
            if (OPF.ShowDialog() != null)
            {
                fileName.Text = OPF.FileName;
            }

        }

        
        private void Jp_Click(object sender, RoutedEventArgs e)
        {
           
         
        }


       
    }
}
