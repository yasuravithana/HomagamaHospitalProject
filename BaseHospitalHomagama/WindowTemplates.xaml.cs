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
    /// Interaction logic for WindowTemplates.xaml
    /// </summary>
    public partial class WindowTemplates : Window
    {
        public WindowTemplates()
        {
            InitializeComponent();
        }
        
        public void setList(String[] templ)
        {
            for (int i = 0; i < 10; i++)
            {
                listBox1.Items[i] = templ[i];
            }
        }

        private void listBox1_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (this.listBox1.SelectedIndex != -1)
            {
                MainWindow.template=listBox1.SelectedItem.ToString();
                this.Close();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MainWindow.timer2.Start();  
        }



    }
}
