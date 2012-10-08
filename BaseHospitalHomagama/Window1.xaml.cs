using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace BaseHospitalHomagama
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        
        public Window1()
        {
            InitializeComponent();
        }

        private void Grid_MouseDown(object sender, MouseButtonEventArgs e)
        {
            MainWindow.canceled = true;
            this.Close();
        }

        private void grid1_MouseDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.printReqD = (bool)checkBox1.IsChecked;
            MainWindow.printTestedD = (bool)checkBox2.IsChecked;
            MainWindow.canceled = false;
            this.Close();
        }

        private void cancel_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.canceled = true;
            this.Close();
        }
    }
}
