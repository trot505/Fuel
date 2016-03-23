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
using System.Windows.Shapes;

namespace Fuel
{
    /// <summary>
    /// Логика взаимодействия для WinSelect.xaml
    /// </summary>
    public partial class WinSelect : Window
    {
        public WinSelect()
        {
            InitializeComponent();
        }
        
        private void selecC_Click(object sender, RoutedEventArgs e)
        {
            if(selectGrid.SelectedIndex >= 0)
            {
                DialogResult = true;    
            }
            else
            {
                System.Windows.MessageBox.Show("Необходимо выбрать организацию из списка.", "ВНИМЕНИЕ !!!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void cancelC_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
