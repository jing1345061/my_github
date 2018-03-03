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

namespace WJD_UI
{
    /// <summary>
    /// ruyuan.xaml 的交互逻辑
    /// </summary>
    public partial class ruyuan : UserControl
    {
        public ruyuan()
        {
            InitializeComponent();
        }

        private void openfile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
         
            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                this.textBox.Text = openFileDialog.FileName;
            }

        }

        private void button_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
