using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Fuel
{
    /// <summary>
    /// Логика взаимодействия для upDel.xaml
    /// </summary>
    public partial class exC : Window
    {
        
        public exC()
        {
            InitializeComponent();
            
        }

        private void save_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void addF_MouseDown(object sender, MouseButtonEventArgs e)
        {
            OpenFileDialog OPF = new OpenFileDialog();
            OPF.Filter = "Файлы excel|*.xls;*.xlsx;*.xlsm";
            OPF.Title = "Выберите файл excel";
            if (OPF.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileC.Text = OPF.FileName;
            }

        }        
    }
}
