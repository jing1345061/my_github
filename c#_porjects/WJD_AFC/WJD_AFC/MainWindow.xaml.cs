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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WJD_AFC
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            control_.Content = new danyi();
            button_1.Background = Brushes.LimeGreen;
            button_2.Background = Brushes.WhiteSmoke;
            button_3.Background = Brushes.WhiteSmoke;
            button_4.Background = Brushes.WhiteSmoke;
            
        }
        Control danyi_main;
        Control ruyuan_main;
        private void button_1_Click(object sender, RoutedEventArgs e)
        {

            if (danyi_main != null)
            { control_.Content = danyi_main; }
            else
            {
                danyi_main = new danyi();
                control_.Content = danyi_main;
            }
            
            


            button_1.Background = Brushes.LimeGreen;
            button_2.Background = Brushes.WhiteSmoke;
            button_3.Background = Brushes.WhiteSmoke;
            button_4.Background = Brushes.WhiteSmoke;
        }

        private void button_2_Click(object sender, RoutedEventArgs e)
        {
            if (ruyuan_main != null)
            { control_.Content = ruyuan_main; }
            else
            {
                ruyuan_main = new ruyuan();
                control_.Content = ruyuan_main;
            }
            button_1.Background = Brushes.WhiteSmoke;
            button_2.Background = Brushes.LimeGreen;
            button_3.Background = Brushes.WhiteSmoke;
            button_4.Background = Brushes.WhiteSmoke;
        }


        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            danyi.colse_browser();
        }
    }
}
