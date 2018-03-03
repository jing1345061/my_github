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
using System.IO;


namespace WJD_UI
{
    /// <summary>
    /// danyi.xaml 的交互逻辑
    /// </summary>
    public partial class danyi : UserControl
    {
        public danyi()
        {
            InitializeComponent();
            den_lu_zhuan_tai = false;
        }

      static  IWebDriver browser;
        ChromeOptions options = new ChromeOptions();
        bool den_lu_zhuan_tai = false;

        private void set_chrome()
        {
             
            var driverService = ChromeDriverService.CreateDefaultService();
           options.AddArgument("headless");
            options.AddArgument("disable-gpu");
            options.AddArgument("silent");
            options.AddArgument("--disable-plugins"); // disable flash
            options.LeaveBrowserRunning = true;
            driverService.HideCommandPromptWindow = true;
            browser = new ChromeDriver(driverService, options);
           
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            if (den_lu_zhuan_tai == true)
            {

                step_one();
            }
            else
            {
                string fp = System.IO.Directory.GetCurrentDirectory() + "\\cookies.json";
                JObject json1 = (JObject)JsonConvert.DeserializeObject(File.ReadAllText(fp));
                JArray array = (JArray)json1["danyi_cookies"];
                set_chrome();
                browser.Navigate().GoToUrl("http://www.singlewindow.gd.cn");
                browser.Manage().Cookies.DeleteAllCookies();
                Cookie ck;
                for (int i = 0; i < array.Count; i++)
                {
                    ck = new Cookie(array[i]["name"].ToString(), array[i]["value"].ToString(), "");
                    browser.Manage().Cookies.AddCookie(ck);

                }
                browser.Navigate().GoToUrl("http://www.singlewindow.gd.cn/index/swProxy/deskserver/sw/deskIndex?menu_id=dec001");
                try
                {

                    textBox.Text += browser.FindElements(By.XPath("//span[@style='color: #fff; font-family: Microsoft YaHei; font-size: 14px;']"))[1].Text + "\n";
                    den_lu_zhuan_tai = true;
                    step_one();
                }
                catch
                {

                    denlu_danyi();
                    step_one();

                }
            }

        }
       

        private void denlu_danyi()
        {
            //set_chrome();
             browser.Navigate().GoToUrl("http://app.singlewindow.cn/cas/login?_local_login_flag=1&service=http://app.singlewindow.cn/cas/jump.jsp%3FtoUrl%3DaHR0cDovL2FwcC5zaW5nbGV3aW5kb3cuY24vY2FzL29hdXRoMi4wL2F1dGhvcml6ZT9jbGllbnRfaWQ9Z2QwMDAwMDAwMEt0NGRTeDAwMiZyZXNwb25zZV90eXBlPWNvZGUmcmVkaXJlY3RfdXJpPWh0dHAlM0ElMkYlMkZ3d3cuc2luZ2xld2luZG93LmdkLmNuJTJGT0F1dGhMb2dpbkNvbnRyb2xsZXI=&localServerUrl=http://www.singlewindow.gd.cn&colorA1=fff&colorA2=255,204,51,%200.2&localRegistryUrl=aHR0cHM6Ly9hcHAuc2luZ2xld2luZG93LmNuL3VzZXJzZXJ2ZXIvdXNlci91c2VyRXRwc1JlZ2lzdGVyL2Nob3NlUmd0V2F5");
            Thread.Sleep(1000);
            textBox.Text += browser.Title + "\n";
            IWebElement denlu_name = browser.FindElement(By.XPath("//input[@id='swy']"));
            denlu_name.SendKeys("wjd13650116965");
            IWebElement denlu_password = browser.FindElement(By.XPath("//input[@id='swm2']"));
            denlu_password.SendKeys("wjd13650116965");
            IWebElement denlu_yan_zhen_ma = browser.FindElement(By.XPath("//input[@id='verifyCode']"));
            var takesScreenshot = (ITakesScreenshot)browser; //截图
            var screenshot = takesScreenshot.GetScreenshot();
            screenshot.SaveAsFile("screenshot.png", ScreenshotImageFormat.Png);
            caijianpic("screenshot.png", 190, 170, 107, 44);
            denlu_yan_zhen_ma.SendKeys(yanzhenma("screenshot.png"));
            IWebElement denlu_button = browser.FindElement(By.XPath("//input[@id='sub3']"));
            denlu_button.Click();
            Thread.Sleep(1000);
            browser.Navigate().GoToUrl("http://www.singlewindow.gd.cn/index/swProxy/deskserver/sw/deskIndex?menu_id=dec001");
            den_lu_zhuan_tai = true;
        }


        private void step_one()
        {
            string entryid = "";
           string datebegin = "";
            string dateend = "";
            if (begin.SelectedDate != null & end.SelectedDate != null)
            {
               // entryid = cx_haiguandanhao.Text;
                if (begin.SelectedDate != null)
                {
                    datebegin = Convert.ToDateTime(begin.Text).ToString("yyyy-MM-dd");
                }
                if (end.SelectedDate != null)
                {
                    dateend = Convert.ToDateTime(end.Text).ToString("yyyy-MM-dd");
                }
                //MessageBox.Show(datebegin+"     "+dateend);

            }
            else
            {
                MessageBox.Show("亲！你什么都不填，你想让我怎么查呢？");
                return;
            }

            string step_one_url = "{\"cusCiqNo\":\"\",\"entryId\":\"" + entryid + "\",\"billNo\":\"\",\"entDeclNo\":\"\",\"copNo\":\"\",\"cusDecStatus\":\"\",\"cusDecStatusName\":\"\",\"ciqDecStatus\":\"\",\"dclTrnRelFlag\":\"0\",\"dclTrnRelFlagName\":\"一般报关单\",\"rcvgdTradeCode\":\"\",\"agentCode\":\"\",\"ownerCode\":\"\",\"contrNo\":\"\",\"supvModeCdde\":\"\",\"tradeModeCode\":\"\",\"customMaster\":\"\",\"orgCode\":\"\",\"splitBillLadNo\":\"\",\"updateTime\":\"" + datebegin + "\",\"updateTimeEnd\":\"" + dateend + "\"}";
            step_one_url = System.Net.WebUtility.UrlEncode(step_one_url);
            step_one_url = System.Net.WebUtility.UrlEncode(step_one_url);
            step_one_url = System.Net.WebUtility.UrlEncode(step_one_url);

            Thread.Sleep(1000);
            browser.Navigate().GoToUrl("http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/common/cusQuery?limit=1500&offset=0&sortName=updateTime&sortOrder=desc&decStatusInfo="+step_one_url);

            string cx_guanjianguanlianhao = browser.FindElement(By.XPath("//body")).Text;
            //textBox.Text = cx_guanjianguanlianhao;           

            JObject json1 = (JObject)JsonConvert.DeserializeObject(cx_guanjianguanlianhao);
            JArray array = (JArray)json1["rows"];
            //textBox.Text += array.Count.ToString()+"\n";


            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("关检关联号", typeof(string));
            dt.Columns.Add("报关单号", typeof(string));
            dt.Columns.Add("商品项数", typeof(string));
            dt.Columns.Add("报关状态", typeof(string));
            dt.Columns.Add("收发单位编码", typeof(string));
            dt.Columns.Add("收发货人", typeof(string));
            dt.Columns.Add("毛重", typeof(string));
            dt.Columns.Add("件数", typeof(string));
            dt.Columns.Add("监管方式", typeof(string));        
            dt.Columns.Add("进出口标志", typeof(string));
            dt.Columns.Add("进出口日期", typeof(string));
            dt.Columns.Add("申报单位代码", typeof(string));
            dt.Columns.Add("申报单位名称", typeof(string));
            dt.Columns.Add("生产销售/消费使用单位编码", typeof(string));
            dt.Columns.Add("生产销售 / 消费使用单位名称", typeof(string));
            DataRow row = dt.NewRow();
            for (int i = 0; i < array.Count; i++)
            {
                row = dt.NewRow();
                row["关检关联号"]=array[i]["cusCiqNo"].ToString();
                row["报关单号"] = array[i]["entryId"].ToString();
                row["商品项数"] = array[i]["goodsNum"].ToString();
                row["报关状态"] = array[i]["cusDecStatusName"].ToString();
                if (array[i]["cusIEFlagName"].ToString() == "进口")
                {
                    row["收发单位编码"] = array[i]["rcvgdTradeCode"].ToString();
                    row["收发货人"] = array[i]["consigneeCname"].ToString();
                }
                else
                {
                    row["收发单位编码"] = array[i]["cnsnTradeCode"].ToString();
                    row["收发货人"] = array[i]["consignorCname"].ToString();
                }

                row["毛重"] = array[i]["grossWt"].ToString();
                row["件数"] = array[i]["packNo"].ToString();
                row["监管方式"] = array[i]["supvModeCdde"].ToString();
                row["进出口标志"] = array[i]["cusIEFlagName"].ToString();
                row["进出口日期"] = array[i]["iEDate"].ToString();
                row["申报单位代码"] = array[i]["agentCode"].ToString();
                row["申报单位名称"] = array[i]["agentName"].ToString();
                row["生产销售/消费使用单位编码"] = array[i]["ownerCode"].ToString();
                row["生产销售 / 消费使用单位名称"] = array[i]["ownerName"].ToString();

                dt.Rows.Add(row);
            }

            dataGrid1.ItemsSource = dt.DefaultView;

            ICookieJar danyi_cookie = browser.Manage().Cookies;

            string cookies_str = "{\"danyi_cookies\":[";
            for (int i = 0; i < danyi_cookie.AllCookies.Count; i++)
            {
                if (i != danyi_cookie.AllCookies.Count - 1)
                {
                    cookies_str += "{\"name\":" + "\"" + danyi_cookie.AllCookies[i].Name.ToString() + "\",";
                    cookies_str += "\"value\":" + "\"" + danyi_cookie.AllCookies[i].Value.ToString() + "\"},";
                }
                else
                {

                    cookies_str += "{\"name\":" + "\"" + danyi_cookie.AllCookies[i].Name.ToString() + "\",";
                    cookies_str += "\"value\":" + "\"" + danyi_cookie.AllCookies[i].Value.ToString() + "\"}";

                }
            }
            cookies_str += "]}";

            JObject cookies_json = (JObject)JsonConvert.DeserializeObject(cookies_str);

            string fp = System.IO.Directory.GetCurrentDirectory()+ "\\cookies.json";

            File.WriteAllText(fp, JsonConvert.SerializeObject(cookies_json));

            //textBox.Text += cookies_str;


             // browser.Quit();   */
        }

        public static void caijianpic(String picPath, int x, int y, int width, int height)   //处理截图
        {
            //图片路径
            String oldPath = picPath;
            //新图片路径
            String newPath = System.IO.Path.GetExtension(oldPath);
            //计算新的文件名，在旧文件名后加_new
            newPath = oldPath.Substring(0, oldPath.Length - newPath.Length) + newPath;
            //定义截取矩形
            System.Drawing.Rectangle cropArea = new System.Drawing.Rectangle(x, y, width, height);
            //要截取的区域大小
            //加载图片
            System.Drawing.Image img = System.Drawing.Image.FromStream(new System.IO.MemoryStream(System.IO.File.ReadAllBytes(oldPath)));
            //判断超出的位置否
            if ((img.Width < x + width) || img.Height < y + height)
            {
                MessageBox.Show("裁剪尺寸超出原有尺寸！");
                img.Dispose();
                return;
            }
            //定义Bitmap对象
            System.Drawing.Bitmap bmpImage = new System.Drawing.Bitmap(img);
            //进行裁剪
            System.Drawing.Bitmap bmpCrop = bmpImage.Clone(cropArea, bmpImage.PixelFormat);
            //保存成新文件

            bmpCrop.Save(newPath);
            //释放对象
            img.Dispose(); bmpCrop.Dispose();
        }

        public string yanzhenma(string path)
        {
            T NewForm = new T();
            NewForm._yanzhenma = "screenshot.png";
            NewForm.ShowDialog();
            string yzm = NewForm.Temp;
            if (yzm != null)
            {
               // MessageBox.Show(yzm);

            }

            else
            {
                MessageBox.Show("验证码不正确");


            }

            return yzm;
        }   //处理验证码

        private void button_jijinjichu_Click(object sender, RoutedEventArgs e) //点击即进即出按钮
        {

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbooks wbks = app.Workbooks;
            _Workbook _wbk = wbks.Add(System.IO.Directory.GetCurrentDirectory() + "\\Templates\\jijinjichu_Template.xls");
            Sheets shs = _wbk.Sheets;
            _Worksheet sheet1 = (_Worksheet)shs.get_Item(1);
            _Worksheet sheet2 = (_Worksheet)shs.get_Item(2);
            _Worksheet sheet4 = (_Worksheet)shs.get_Item(4);
            _Worksheet sheet5 = (_Worksheet)shs.get_Item(5);
           


            string[] guanjianhao = new string[2];


            if (dataGrid1.SelectedItems.Count == 2)
            {

                DataRowView row_0 = (DataRowView)dataGrid1.SelectedItems[0];
                guanjianhao[0] = row_0[0].ToString();
               textBox.Text+= guanjianhao[0].Substring(0, 1);
                DataRowView row_1 = (DataRowView)dataGrid1.SelectedItems[1];
                guanjianhao[1] = row_1[1].ToString();
                double qty1 = 0.0;
                if (row_0[9].ToString() == "出口")
                {
                    qty1 = 0.0;
                    browser.Navigate().GoToUrl("http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/cus/queryBCusDecList?limit=500&offset=0&cusCiqNo=" + guanjianhao[0]);
                    //textBox.Text += browser.FindElement(By.TagName("body")).Text + "\n";
                    JObject json_0 = (JObject)JsonConvert.DeserializeObject(browser.FindElement(By.TagName("body")).Text);
                    JArray body_0 = (JArray)json_0["rows"];
                    for (int i = 0; i < body_0.Count; i++)
                    {
                        int rows = i + 2;
                       sheet2.Cells[rows,1]  =body_0[i]["gNo"].ToString();
                        sheet2.Cells[rows, 4] = body_0[i]["gNo"].ToString();
                        sheet2.Cells[rows, 5] = body_0[i]["codeTs"].ToString();
                        sheet2.Cells[rows, 6] = body_0[i]["gName"].ToString();
                        sheet2.Cells[rows, 7] = body_0[i]["gModel"].ToString();                 
                        sheet2.Cells[rows, 8] = "11";
                        sheet2.Cells[rows, 9] = body_0[i]["gQty"].ToString();
                        sheet2.Cells[rows, 10] = body_0[i]["gUnit"].ToString();
                       sheet2.Cells[rows, 11] = body_0[i]["qty1"].ToString();
                        sheet2.Cells[rows, 12] = body_0[i]["unit1"].ToString();
                        sheet2.Cells[rows, 13] = body_0[i]["declPrice"].ToString();
                        sheet2.Cells[rows, 14] = body_0[i]["declTotal"].ToString();
                        sheet2.Cells[rows, 15] = body_0[i]["cusOriginCountry"].ToString();
                        sheet2.Cells[rows, 16] = body_0[i]["tradeCurr"].ToString();
                        sheet2.Cells[rows, 17] = body_0[i]["dutyMode"].ToString();
                        //sheet2.Cells[rows, 17] ="";
                       qty1 +=double.Parse(body_0[i]["qty1"].ToString());
                    }
                    sheet1.Cells[3, 2] = row_0[4].ToString();
                    sheet1.Cells[3, 4] = row_0[5].ToString();
                    sheet1.Cells[3, 8] = row_0[10].ToString();
                    sheet1.Cells[4, 2] = row_0[11].ToString();
                    sheet1.Cells[4, 4] = row_0[12].ToString();
                    sheet1.Cells[4, 8] = "5222";
                    sheet1.Cells[5, 2] = "142";                  
                    sheet1.Cells[5, 4] = row_0[7].ToString();
                    sheet1.Cells[5, 6] = row_0[6].ToString();
                    sheet1.Cells[5, 8] = qty1.ToString();
                    sheet1.Cells[6, 6] = row_0[1].ToString();
                    sheet1.Cells[7,6] = row_0[8].ToString();
                    sheet1.Cells[4,6] = "YY5222176K001001";
                    sheet1.Cells[6,2] = "2";
                    sheet1.Cells[6,4] = "3";
                    sheet1.Cells[7,8] = "进区";


                    browser.Navigate().GoToUrl("http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/cus/queryBCusDecList?limit=500&offset=0&cusCiqNo=" + guanjianhao[1]);
                    //textBox.Text += browser.FindElement(By.TagName("body")).Text + "\n";
                    JObject json_1 = (JObject)JsonConvert.DeserializeObject(browser.FindElement(By.TagName("body")).Text);
                    JArray body_1 = (JArray)json_0["rows"];
                    qty1 = 0.0;
                    for (int i = 0; i < body_0.Count; i++)
                    {
                        int rows = i + 2;
                        sheet5.Cells[rows, 1] = body_1[i]["gNo"].ToString();
                        sheet5.Cells[rows, 4] = body_1[i]["gNo"].ToString();
                        sheet5.Cells[rows, 5] = body_1[i]["codeTs"].ToString();
                        sheet5.Cells[rows, 6] = body_1[i]["gName"].ToString();
                        sheet5.Cells[rows, 7] = body_1[i]["gModel"].ToString();
                        sheet5.Cells[rows, 8] = "11";
                        sheet5.Cells[rows, 9] = body_1[i]["gQty"].ToString();
                        sheet5.Cells[rows, 10] = body_1[i]["gUnit"].ToString();
                        sheet5.Cells[rows, 11] = body_1[i]["qty1"].ToString();
                        sheet5.Cells[rows, 12] = body_1[i]["unit1"].ToString();
                        sheet5.Cells[rows, 13] = body_1[i]["declPrice"].ToString();
                        sheet5.Cells[rows, 14] = body_1[i]["declTotal"].ToString();
                        sheet5.Cells[rows, 15] = body_1[i]["cusOriginCountry"].ToString();
                        sheet5.Cells[rows, 16] = body_1[i]["tradeCurr"].ToString();
                        sheet5.Cells[rows, 17] = body_1[i]["dutyMode"].ToString();
                        qty1 += double.Parse(body_0[i]["qty1"].ToString());
                    }
                    sheet4.Cells[3, 2] = row_1[4].ToString();
                    sheet4.Cells[3, 4] = row_1[5].ToString();
                    sheet4.Cells[3, 8] = row_1[10].ToString();
                    sheet4.Cells[4, 2] = row_1[11].ToString();
                    sheet4.Cells[4, 4] = row_1[12].ToString();
                    sheet4.Cells[4, 8] = "5222";
                    sheet4.Cells[5, 2] = "142";
                    sheet4.Cells[5, 4] = row_1[7].ToString();
                    sheet4.Cells[5, 6] = row_1[6].ToString();
                    sheet4.Cells[5, 8] = qty1.ToString();
                    sheet4.Cells[6, 6] = row_1[1].ToString();
                    sheet4.Cells[7, 6] = row_1[8].ToString();
                    sheet4.Cells[4, 6] = "YY5222176K001001";
                    sheet4.Cells[6, 2] = "2";
                    sheet4.Cells[6, 4] = "3";
                    sheet4.Cells[7, 8] = "出区";
                }
                else
                {
                    qty1 = 0.0;
                    browser.Navigate().GoToUrl("http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/cus/queryBCusDecList?limit=500&offset=0&cusCiqNo=" + guanjianhao[0]);
                    //textBox.Text += browser.FindElement(By.TagName("body")).Text + "\n";
                    JObject json_0 = (JObject)JsonConvert.DeserializeObject(browser.FindElement(By.TagName("body")).Text);
                    JArray body_0 = (JArray)json_0["rows"];
                    for (int i = 0; i < body_0.Count; i++)
                    {
                        int rows = i + 2;
                        sheet5.Cells[rows, 1] = body_0[i]["gNo"].ToString();
                        sheet5.Cells[rows, 4] = body_0[i]["gNo"].ToString();
                        sheet5.Cells[rows, 5] = body_0[i]["codeTs"].ToString();
                        sheet5.Cells[rows, 6] = body_0[i]["gName"].ToString();
                        sheet5.Cells[rows, 7] = body_0[i]["gModel"].ToString();
                        sheet5.Cells[rows, 8] = "11";
                        sheet5.Cells[rows, 9] = body_0[i]["gQty"].ToString();
                        sheet5.Cells[rows, 10] = body_0[i]["gUnit"].ToString();
                        sheet5.Cells[rows, 11] = body_0[i]["qty1"].ToString();
                        sheet5.Cells[rows, 12] = body_0[i]["unit1"].ToString();
                        sheet5.Cells[rows, 13] = body_0[i]["declPrice"].ToString();
                        sheet5.Cells[rows, 14] = body_0[i]["declTotal"].ToString();
                        sheet5.Cells[rows, 15] = body_0[i]["cusOriginCountry"].ToString();
                        sheet5.Cells[rows, 16] = body_0[i]["tradeCurr"].ToString();
                        sheet5.Cells[rows, 17] = body_0[i]["dutyMode"].ToString();
                        qty1 += double.Parse(body_0[i]["qty1"].ToString());
                    }
                    sheet4.Cells[3, 2] = row_0[4].ToString();
                    sheet4.Cells[3, 4] = row_0[5].ToString();
                    sheet4.Cells[3, 8] = row_0[10].ToString();
                    sheet4.Cells[4, 2] = row_0[11].ToString();
                    sheet4.Cells[4, 4] = row_0[12].ToString();
                    sheet4.Cells[4, 8] = "5222";
                    sheet4.Cells[5, 2] = "142";
                    sheet4.Cells[5, 4] = row_0[7].ToString();
                    sheet4.Cells[5, 6] = row_0[6].ToString();
                    sheet4.Cells[5, 8] = qty1.ToString();
                    sheet4.Cells[6, 6] = row_0[1].ToString();
                    sheet4.Cells[7, 6] = row_0[8].ToString();
                    sheet4.Cells[4, 6] = "YY5222176K001001";
                    sheet4.Cells[6, 2] = "2";
                    sheet4.Cells[6, 4] = "3";
                    sheet4.Cells[7, 8] = "出区";



                    qty1 = 0.0;
                    browser.Navigate().GoToUrl("http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/cus/queryBCusDecList?limit=500&offset=0&cusCiqNo=" + guanjianhao[1]);
                    //textBox.Text += browser.FindElement(By.TagName("body")).Text + "\n";
                    JObject json_1 = (JObject)JsonConvert.DeserializeObject(browser.FindElement(By.TagName("body")).Text);
                    JArray body_1 = (JArray)json_0["rows"];
                    for (int i = 0; i < body_0.Count; i++)
                    {
                        int rows = i + 2;
                        sheet2.Cells[rows, 1] = body_1[i]["gNo"].ToString();
                        sheet2.Cells[rows, 4] = body_1[i]["gNo"].ToString();
                        sheet2.Cells[rows, 5] = body_1[i]["codeTs"].ToString();
                        sheet2.Cells[rows, 6] = body_1[i]["gName"].ToString();
                        sheet2.Cells[rows, 7] = body_1[i]["gModel"].ToString();
                        sheet2.Cells[rows, 8] = "11";
                        sheet2.Cells[rows, 9] = body_1[i]["gQty"].ToString();
                        sheet2.Cells[rows, 10] = body_1[i]["gUnit"].ToString();
                        sheet2.Cells[rows, 11] = body_1[i]["qty1"].ToString();
                        sheet2.Cells[rows, 12] = body_1[i]["unit1"].ToString();
                        sheet2.Cells[rows, 13] = body_1[i]["declPrice"].ToString();
                        sheet2.Cells[rows, 14] = body_1[i]["declTotal"].ToString();
                        sheet2.Cells[rows, 15] = body_1[i]["cusOriginCountry"].ToString();
                        sheet2.Cells[rows, 16] = body_1[i]["tradeCurr"].ToString();
                        sheet2.Cells[rows, 17] = body_1[i]["dutyMode"].ToString();
                        qty1 += double.Parse(body_0[i]["qty1"].ToString());
                    }
                    sheet1.Cells[3, 2] = row_1[4].ToString();
                    sheet1.Cells[3, 4] = row_1[5].ToString();
                    sheet1.Cells[3, 8] = row_1[10].ToString();
                    sheet1.Cells[4, 2] = row_1[11].ToString();
                    sheet1.Cells[4, 4] = row_1[12].ToString();
                    sheet1.Cells[4, 8] = "5222";
                    sheet1.Cells[5, 2] = "142";
                    sheet1.Cells[5, 4] = row_1[7].ToString();
                    sheet1.Cells[5, 6] = row_1[6].ToString();
                    sheet1.Cells[5, 8] = qty1.ToString();
                    sheet1.Cells[6, 6] = row_1[1].ToString();
                    sheet1.Cells[7, 6] = row_1[8].ToString();
                    sheet1.Cells[4, 6] = "YY5222176K001001";
                    sheet1.Cells[6, 2] = "2";
                    sheet1.Cells[6, 4] = "3";
                    sheet1.Cells[7, 8] = "进区";
                }
                //MessageBox.Show(sheet2.UsedRange.Rows.Count.ToString()+"  "+sheet5.UsedRange.Rows.Count.ToString());

                string usedrows = "X" + sheet2.UsedRange.Rows.Count.ToString();
                Range excelRange = sheet2.get_Range("A1",usedrows);//Borders.LineStyle 单元格边框线
                excelRange.Borders.LineStyle = 1; //单元格边框线类型(线型,虚线型)
                //excelRange.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                usedrows = "X" + sheet5.UsedRange.Rows.Count.ToString();
                excelRange = sheet5.get_Range("A1", usedrows);  
                excelRange.Borders.LineStyle = 1;
              //  excelRange.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                _wbk.SaveAs("d:\\automatic_for_customs\\"+row_0[1]+"--"+row_1[1]+".xls", XlFileFormat.xlExcel8);
                _wbk.Save();
                _wbk.Close();
                app.Quit();

                /*     

                     browser.Navigate().GoToUrl("http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/cus/queryBCusDecList?limit=500&offset=0&cusCiqNo=" + guanjianhao[0]);
                     //  textBox.Text += browser.FindElement(By.TagName("body")).Text + "\n";
                     JObject json_0 = (JObject)JsonConvert.DeserializeObject(browser.FindElement(By.TagName("body")).Text);
                     JArray body_0 = (JArray)json_0["rows"];

                     textBox.Text += body_0[0].ToString();   */

            }
            else
            {
                MessageBox.Show("即进即出导入生成必须选择两项且为进口和出口");

            }
           
        }
        private void to_excel(System.Data.DataTable dt, int a)
        {


            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelWB = excelApp.Workbooks.Add(System.Type.Missing);    //创建工作簿（WorkBook：即Excel文件主体本身）  
            Worksheet excelWS = (Worksheet)excelWB.Worksheets[a];   //创建工作表（即Excel里的子表sheet） 1表示在子表sheet1里进行数据导出  

            excelWS.Cells.NumberFormat = "@";     //  如果数据中存在数字类型 可以让它变文本格式显示  
            //将数据导入到工作表的单元格  
            for (int ii = 0; ii < dt.Rows.Count; ii++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    excelWS.Cells[ii + 1, j + 1] = dt.Rows[ii][j].ToString();   //Excel单元格第一个从索引1开始  
                }
            }
            excelApp.AlertBeforeOverwriting = true;
            try
            {
                SaveFileDialog sfd = new SaveFileDialog();
                //设置这个对话框的起始保存路径  
                sfd.InitialDirectory = @"D:\";
                //设置保存的文件的类型，注意过滤器的语法  
                sfd.Filter = "Excel 工作簿|*.xlsx";
                //调用ShowDialog()方法显示该对话框，该方法的返回值代表用户是否点击了确定按钮  
                if (sfd.ShowDialog() == true)
                {
                    excelWB.SaveAs(sfd.FileName);  //将其进行保存到指定的路径
                    MessageBox.Show("保存成功");
                }
                else
                {
                    MessageBox.Show("取消保存");
                }
                //excelWB.SaveAs("D:\\" + "" + ".xlsx");  //将其进行保存到指定的路径 


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            excelWB.Close();
            excelApp.Quit();  //KillAllExcel(excelApp); 释放可能还没释放的进程
            
        }
 

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            browser.Quit();
            den_lu_zhuan_tai = false;
        }
        public static void colse_browser()
        {
            if (browser != null)
            {
                browser.Quit();
            }


        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            
            string fp = System.IO.Directory.GetCurrentDirectory() + "\\cookies.json";
            JObject json1 = (JObject)JsonConvert.DeserializeObject(File.ReadAllText(fp));
            JArray array = (JArray)json1["danyi_cookies"];
            string ck = "";

            for (int i = 0; i < array.Count; i++)
            {
                if (i == array.Count-1)
                {
                    ck += array[i]["name"].ToString() +"="+ array[i]["value"].ToString();
                }
                else
                {
                    ck += array[i]["name"].ToString() +"="+ array[i]["value"].ToString() + ";";
                }               
            }
            string requestUrl = "http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/cus/queryDCusDecHead";
            var post_json = new { cusCiqNo = "I20180000018309867", cusIEFlag="I"};
            string data = JsonConvert.SerializeObject(post_json);
            System.Net. HttpWebRequest request = System.Net. WebRequest.Create(requestUrl) as System.Net. HttpWebRequest;
            request.Method = "POST";
            request.Accept = "application/json, text/javascript, */*; q=0.01";
            request.ContentType = "application/json; charset=UTF-8";
            request.Headers.Add("X-Requested-With", "XMLHttpRequest");
            request.Referer = "http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/cusImport?operationType=I&cusCiqNo=I20180000018309867&ngBasePath=http://www.singlewindow.gd.cn:80/swProxy/decserver/";
            request.Headers.Add("Accept-Language", "zh-cn");
            request.Headers.Add("Cookie", ck);
            request.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:27.0) Gecko/20100101 Firefox/27.0";     
            byte[] data_bytes = Encoding.UTF8.GetBytes(data);
            //byte[] data_bytes = Encoding.GetEncoding("gb2312").GetBytes(post);
            request.ContentLength = data_bytes.Length;

            request.ContentLength = data.Length;
           byte[]  data_byte = System.Text.Encoding.UTF8.GetBytes(data);
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(data_byte, 0, data.Length);
            requestStream.Close();

            //获取当前Http请求的响应实例
            string htmlStr = "";
            System.Net. HttpWebResponse response = request.GetResponse() as System.Net.HttpWebResponse;
            Stream responseStream = response.GetResponseStream();
            using (StreamReader reader = new StreamReader(responseStream, Encoding.GetEncoding("UTF-8")))
            {
                htmlStr = reader.ReadToEnd();
            }
            responseStream.Close();

          textBox.Text+= htmlStr+"\n";

        }

 

    }
   
}
