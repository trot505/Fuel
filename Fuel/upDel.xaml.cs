﻿using System;
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
    /// Логика взаимодействия для upDel.xaml
    /// </summary>
    public partial class upDel : Window
    {
        
        public upDel()
        {
            InitializeComponent();
            
        }

        private void updateC_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void delC_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
