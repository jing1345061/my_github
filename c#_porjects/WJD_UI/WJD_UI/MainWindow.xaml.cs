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
using System.Windows.Navigation;
using System.Windows.Shapes;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using Newtonsoft;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;
using System.Data;
using Microsoft.Win32;
using System.Collections;


namespace WJD_UI
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            control_.Content = new danyi();
        }


        private void button_danyi_Click(object sender, RoutedEventArgs e)
        {
            control_.Content = new danyi();
        }

        private void button_ruyuan_Click(object sender, RoutedEventArgs e)
        {
            control_.Content = new ruyuan();
            danyi.colse_browser();
        }

        private void button_canchu_Click(object sender, RoutedEventArgs e)
        {
            control_.Content = new cancu();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            danyi.colse_browser();
        }
    }
}
