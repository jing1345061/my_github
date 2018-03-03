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

namespace WJD_UI
{
    /// <summary>
    /// T.xaml 的交互逻辑
    /// </summary>
    public partial class T : Window
    {
        public T()
        {
            InitializeComponent();
        }

        private string _temp = "";
        private string yanzhenma = "";

        public string Temp
        {
            get
            {
                return _temp;
            }
            set
            {
                _temp = value;


            }
        }
        public string _yanzhenma
        {
            set
            {

                yanzhenma = System.Environment.CurrentDirectory + @"\" + value;
                BitmapImage image = new BitmapImage(new Uri(yanzhenma, UriKind.Absolute));//打开图片
                yzm.Source = image;//将控件和图片绑定，logo为Image控件名称
            }



        }




        private void button1_Click(object sender, RoutedEventArgs e)
        {
            if (textBox1.Text != null)
            {
                this.Temp = textBox1.Text.ToUpper();
            }
            else
            {
                MessageBox.Show("验证码不正确");

            }
            Close();
        }
    }
}
