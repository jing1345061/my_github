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
namespace WJD_AFC
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

        static IWebDriver browser;
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

                   // textBox.Text += browser.FindElements(By.XPath("//span[@style='color: #fff; font-family: Microsoft YaHei; font-size: 14px;']"))[1].Text + "\n";
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
            //textBox.Text += browser.Title + "\n";
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

            string fp = System.IO.Directory.GetCurrentDirectory() + "\\cookies.json";

            File.WriteAllText(fp, JsonConvert.SerializeObject(cookies_json));
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


            string fp = System.IO.Directory.GetCurrentDirectory() + "\\cookies.json";
            JObject json1 = (JObject)JsonConvert.DeserializeObject(File.ReadAllText(fp));
            JArray ck_array = (JArray)json1["danyi_cookies"];
            string ck = "";
            for (int i = 0; i < ck_array.Count; i++)
            {
                if (i == ck_array.Count - 1)
                {
                    ck += ck_array[i]["name"].ToString() + "=" + ck_array[i]["value"].ToString();
                }
                else
                {
                    ck += ck_array[i]["name"].ToString() + "=" + ck_array[i]["value"].ToString() + ";";
                }
            }


            string get_url = "http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/common/cusQuery?limit=1500&offset=0&sortName=updateTime&sortOrder=desc&decStatusInfo=" + step_one_url;
            System.Net.HttpWebRequest request_get = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(get_url);
            request_get.Method = "GET";
            request_get.ContentType = "text/html;charset=UTF-8";
            request_get.Headers.Add("Cookie", ck);
            request_get.UserAgent = null;
            request_get.Timeout = 100000;
            System.Net.HttpWebResponse response_get = (System.Net.HttpWebResponse)request_get.GetResponse();
            Stream myResponseStream = response_get.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            string cx_guanjianguanlianhao = retString;
       

            JObject cx_guanjianguanlianhao_to_json= (JObject)JsonConvert.DeserializeObject(cx_guanjianguanlianhao);
            JArray array = (JArray)cx_guanjianguanlianhao_to_json["rows"];
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
            DataRow row = dt.NewRow();
            for (int i = 0; i < array.Count; i++)
            {
                row = dt.NewRow();
                row["关检关联号"] = array[i]["cusCiqNo"].ToString();
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

                dt.Rows.Add(row);
            }

            dataGrid1.ItemsSource = dt.DefaultView;

           

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
            string fp = System.IO.Directory.GetCurrentDirectory() + "\\cookies.json";
            JObject json1 = (JObject)JsonConvert.DeserializeObject(File.ReadAllText(fp));
            JArray array = (JArray)json1["danyi_cookies"];
            string ck = "";
            for (int i = 0; i < array.Count; i++)
            {
                if (i == array.Count - 1)
                {
                    ck += array[i]["name"].ToString() + "=" + array[i]["value"].ToString();
                }
                else
                {
                    ck += array[i]["name"].ToString() + "=" + array[i]["value"].ToString() + ";";
                }
            }
            if (dataGrid1.SelectedItems.Count == 2)
            {
                string[] guanjianhao = new string[2];
                DataRowView row_a = (DataRowView)dataGrid1.SelectedItems[0];
                DataRowView row_b= (DataRowView)dataGrid1.SelectedItems[1];

                if (row_a[0].ToString().Substring(0, 1) != row_b[0].ToString().Substring(0, 1))
                {
                    if (row_a[0].ToString().Substring(0, 1) == "I")
                    {
                        guanjianhao[0] = row_a[0].ToString();
                        
                        guanjianhao[1] = row_b[0].ToString();
                    }
                    else
                    {
                        guanjianhao[0] = row_b[0].ToString();
                        guanjianhao[1] = row_a[0].ToString();               
                    }

                }
                else
                {
                    MessageBox.Show("必须为一进一出报关单");
                    return;
                }
                //textBox.Text += guanjianhao[0].Substring(0, 1) + "\n";//取字符串的第一个字符


                string[] header_str = new string[2];
                string[] details_str = new string[2];

                

                for (int i =0;i<header_str.Length;i++)
                {
                    var post_json = new { cusCiqNo = guanjianhao[i], cusIEFlag = guanjianhao[i].Substring(0, 1) };
                    string data = JsonConvert.SerializeObject(post_json);
                    string requestUrl = "http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/cus/queryDCusDecHead";
                    System.Net.HttpWebRequest request_post = System.Net.WebRequest.Create(requestUrl) as System.Net.HttpWebRequest;
                    request_post.Method = "POST";
                    request_post.Accept = "application/json, text/javascript, */*; q=0.01";
                    request_post.ContentType = "application/json; charset=UTF-8";
                    request_post.Headers.Add("X-Requested-With", "XMLHttpRequest");
                    request_post.Headers.Add("Accept-Language", "zh-cn");
                    request_post.Headers.Add("Cookie", ck);
                    request_post.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:27.0) Gecko/20100101 Firefox/27.0";
                    request_post.ContentLength = data.Length;
                    byte[] data_byte = System.Text.Encoding.UTF8.GetBytes(data);
                    Stream requestStream = request_post.GetRequestStream();
                    requestStream.Write(data_byte, 0, data.Length);
                    requestStream.Close();
                   
                    System.Net.HttpWebResponse response_post = request_post.GetResponse() as System.Net.HttpWebResponse;  //获取当前Http请求的响应实例
                    Stream responseStream = response_post.GetResponseStream();
                    using (StreamReader reader = new StreamReader(responseStream, Encoding.GetEncoding("UTF-8")))
                    {
                        header_str[i] = reader.ReadToEnd();
                    }
                    responseStream.Close();
                    Thread.Sleep(1000);
                }

                string get_url = "";
                
                for (int i = 0; i < details_str.Length; i++)
                {

                    get_url = "http://www.singlewindow.gd.cn/swProxy/decserver/sw/dec/cus/queryBCusDecList?limit=500&offset=0&cusCiqNo=" + guanjianhao[i];
                    System.Net.HttpWebRequest request_get = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(get_url);
                    request_get.Method = "GET";
                    request_get.ContentType = "text/html;charset=UTF-8";
                    request_get.Headers.Add("Cookie", ck);
                    request_get.UserAgent = null;
                    request_get.Timeout = 10000;
                    System.Net.HttpWebResponse response_get = (System.Net.HttpWebResponse)request_get.GetResponse();
                    Stream myResponseStream = response_get.GetResponseStream();
                    StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
                    details_str[i]= myStreamReader.ReadToEnd();
                    myStreamReader.Close();
                    myResponseStream.Close();

                }


                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Workbooks wbks = app.Workbooks;
                _Workbook _wbk = wbks.Add(System.IO.Directory.GetCurrentDirectory() + "\\Templates\\jijinjichu_Template.xls");
                Sheets shs = _wbk.Sheets;
                _Worksheet sheet1 = (_Worksheet)shs.get_Item(1);
                _Worksheet sheet2 = (_Worksheet)shs.get_Item(2);
                _Worksheet sheet4 = (_Worksheet)shs.get_Item(4);
                _Worksheet sheet5 = (_Worksheet)shs.get_Item(5);
                sheet1.Cells.NumberFormat = "@"; //以文本显示
                sheet2.Cells.NumberFormat = "@"; //以文本显示
                sheet4.Cells.NumberFormat = "@"; //以文本显示
                sheet5.Cells.NumberFormat = "@"; //以文本显示


                header_str [0]= "{\"header\":[" + header_str[0] + "]}";
                JObject header_to_json = (JObject)JsonConvert.DeserializeObject(header_str[0]);
                JArray header_json = (JArray)header_to_json["header"];
                JObject details_to_json = (JObject)JsonConvert.DeserializeObject(details_str[0]);
                JArray details_json = (JArray)details_to_json["rows"];

                sheet1.Cells[3, 2] = header_json[0]["data"]["preDecHead"]["rcvgdTradeCode"].ToString();
                sheet1.Cells[3, 4] = header_json[0]["data"]["preDecHead"]["consigneeCname"].ToString();
                if (row_a[0].ToString().Substring(0, 1) == "I")
                {
                    sheet1.Cells[3, 8] = row_a[10].ToString();
                }
                else
                {
                    sheet1.Cells[3, 8] = row_b[10].ToString();
                }
               
                sheet1.Cells[4, 2] = header_json[0]["data"]["preDecHead"]["ownerCode"].ToString();
                sheet1.Cells[4, 4] = header_json[0]["data"]["preDecHead"]["ownerName"].ToString();
                sheet1.Cells[4, 6] = "YY5222176K001001";
                sheet1.Cells[4, 8] = "5222";
                sheet1.Cells[5, 2] = "142";
                sheet1.Cells[5, 4] = header_json[0]["data"]["preDecHead"]["packNo"].ToString();
                sheet1.Cells[5, 6] = header_json[0]["data"]["preDecHead"]["grossWt"].ToString();
                sheet1.Cells[5, 8] = header_json[0]["data"]["preDecHead"]["netWt"].ToString();
                sheet1.Cells[6, 2] = header_json[0]["data"]["preDecHead"]["wrapType"].ToString();
                sheet1.Cells[6, 4] = header_json[0]["data"]["preDecHead"]["transMode"].ToString();
                sheet1.Cells[6, 6] = header_json[0]["data"]["preDecHead"]["entryId"].ToString();
                sheet1.Cells[6, 8] = "3";
                sheet1.Cells[7, 2] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                sheet1.Cells[7, 6] = header_json[0]["data"]["preDecHead"]["supvModeCdde"].ToString();
                sheet1.Cells[7, 8] = "进区";
                sheet1.Cells[9 ,2] = header_json[0]["data"]["preDecHead"]["cusTrafModeName"].ToString();
                sheet1.Cells[9, 4] = header_json[0]["data"]["preDecHead"]["trafName"].ToString();
                sheet1.Cells[9, 6] = header_json[0]["data"]["preDecHead"]["cusVoyageNo"].ToString();
                sheet1.Cells[9, 8] = header_json[0]["data"]["preDecHead"]["billNo"].ToString();
                sheet1.Cells[10, 2] = header_json[0]["data"]["preDecHead"]["cutMode"].ToString();
                sheet1.Cells[10, 4] = header_json[0]["data"]["preDecHead"]["inRatio"].ToString();
                sheet1.Cells[10, 6] = header_json[0]["data"]["preDecHead"]["paymentMarkName"].ToString();
                sheet1.Cells[10, 8] = header_json[0]["data"]["preDecHead"]["licenseNo"].ToString();
                sheet1.Cells[11, 2] = header_json[0]["data"]["preDecHead"]["distinatePortName"].ToString();
                sheet1.Cells[11, 4] = header_json[0]["data"]["preDecHead"]["districtCodeName"].ToString();
                //sheet1.Cells[11, 6] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                //sheet1.Cells[11, 8] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                //sheet1.Cells[12, 2] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                //sheet1.Cells[12, 4] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
               // sheet1.Cells[12, 6] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                //sheet1.Cells[12, 8] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                sheet1.Cells[13, 2] = header_json[0]["data"]["preDecHead"]["entryTypeName"].ToString();
                sheet1.Cells[13, 4] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                sheet1.Cells[13, 8] = header_json[0]["data"]["preDecHead"]["contrNo"].ToString();
                for (int i = 0; i < details_json.Count; i++)
                {
                    int rows = i + 2;
                    sheet2.Cells[rows, 1] = details_json[i]["gNo"].ToString();
                    sheet2.Cells[rows, 4] = details_json[i]["gNo"].ToString();
                    sheet2.Cells[rows, 5] = details_json[i]["codeTs"].ToString();
                    sheet2.Cells[rows, 6] = details_json[i]["gName"].ToString();
                    sheet2.Cells[rows, 7] = details_json[i]["gModel"].ToString();
                    sheet2.Cells[rows, 8] = "11";
                    sheet2.Cells[rows, 9] = details_json[i]["gQty"].ToString();
                    sheet2.Cells[rows, 10] = details_json[i]["gUnit"].ToString();
                    sheet2.Cells[rows, 11] = details_json[i]["qty1"].ToString();
                    sheet2.Cells[rows, 12] = details_json[i]["unit1"].ToString();
                    sheet2.Cells[rows, 13] = details_json[i]["declPrice"].ToString();
                    sheet2.Cells[rows, 14] = details_json[i]["declTotal"].ToString();
                    sheet2.Cells[rows, 15] = details_json[i]["cusOriginCountry"].ToString();
                    sheet2.Cells[rows, 16] = details_json[i]["tradeCurr"].ToString();
                    sheet2.Cells[rows, 17] = details_json[i]["dutyMode"].ToString();
                    //sheet2.Cells[rows, 21] = details_json[i]["cusTradeCountry"].ToString();
                    //sheet2.Cells[rows, 17] ="";

                }


                header_str[1] = "{\"header\":[" + header_str[1] + "]}";
                 header_to_json = (JObject)JsonConvert.DeserializeObject(header_str[1]);
                header_json = (JArray)header_to_json["header"];
                 details_to_json = (JObject)JsonConvert.DeserializeObject(details_str[1]);
                details_json = (JArray)details_to_json["rows"];
               // textBox.Text += header_json[0]["data"]["preDecHead"].ToString();
               sheet4.Cells[3, 2] = header_json[0]["data"]["preDecHead"]["cnsnTradeCode"].ToString();
               sheet4.Cells[3, 4] = header_json[0]["data"]["preDecHead"]["consignorCname"].ToString();
                if (row_a[0].ToString().Substring(0, 1) == "I")
                {
                    sheet4.Cells[3, 8] = row_b[10].ToString();
                }
                else
                {
                    sheet4.Cells[3, 8] = row_a[10].ToString();
                }
                sheet4.Cells[4, 2] = header_json[0]["data"]["preDecHead"]["ownerCode"].ToString();
               sheet4.Cells[4, 4] = header_json[0]["data"]["preDecHead"]["ownerName"].ToString();
               sheet4.Cells[4, 6] = "YY5222176K001001";
               sheet4.Cells[4, 8] = "5222";
               sheet4.Cells[5, 2] = "142";
               sheet4.Cells[5, 4] = header_json[0]["data"]["preDecHead"]["packNo"].ToString();
               sheet4.Cells[5, 6] = header_json[0]["data"]["preDecHead"]["grossWt"].ToString();
               sheet4.Cells[5, 8] = header_json[0]["data"]["preDecHead"]["netWt"].ToString();
               sheet4.Cells[6, 2] = header_json[0]["data"]["preDecHead"]["wrapType"].ToString();
               sheet4.Cells[6, 4] = header_json[0]["data"]["preDecHead"]["transMode"].ToString();
               sheet4.Cells[6, 6] = header_json[0]["data"]["preDecHead"]["entryId"].ToString();
               sheet4.Cells[6, 8] = "3";
               sheet4.Cells[7, 2] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
               sheet4.Cells[7, 6] = header_json[0]["data"]["preDecHead"]["supvModeCdde"].ToString();
               sheet4.Cells[7, 8] = "进区";
               sheet4.Cells[9, 2] = header_json[0]["data"]["preDecHead"]["cusTrafModeName"].ToString();
               sheet4.Cells[9, 4] = header_json[0]["data"]["preDecHead"]["trafName"].ToString();
               sheet4.Cells[9, 6] = header_json[0]["data"]["preDecHead"]["cusVoyageNo"].ToString();
               sheet4.Cells[9, 8] = header_json[0]["data"]["preDecHead"]["billNo"].ToString();
               sheet4.Cells[10, 2] = header_json[0]["data"]["preDecHead"]["cutMode"].ToString();
               sheet4.Cells[10, 4] = header_json[0]["data"]["preDecHead"]["inRatio"].ToString();
               sheet4.Cells[10, 6] = header_json[0]["data"]["preDecHead"]["paymentMarkName"].ToString();
               sheet4.Cells[10, 8] = header_json[0]["data"]["preDecHead"]["licenseNo"].ToString();
               sheet4.Cells[11, 2] = header_json[0]["data"]["preDecHead"]["distinatePortName"].ToString();
               sheet4.Cells[11, 4] = header_json[0]["data"]["preDecHead"]["districtCodeName"].ToString();
                //sheet1.Cells[11, 6] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                //sheet1.Cells[11, 8] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                //sheet1.Cells[12, 2] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                //sheet1.Cells[12, 4] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                //sheet4.Cells[12, 6] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
                //sheet1.Cells[12, 8] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
               sheet4.Cells[13, 2] = header_json[0]["data"]["preDecHead"]["entryTypeName"].ToString();
               sheet4.Cells[13, 4] = header_json[0]["data"]["preDecHead"]["noteS"].ToString();
               sheet4.Cells[13, 8] = header_json[0]["data"]["preDecHead"]["contrNo"].ToString();
                for (int i = 0; i < details_json.Count; i++)
                {
                    int rows = i + 2;
                   sheet5.Cells[rows, 1] = details_json[i]["gNo"].ToString();
                   sheet5.Cells[rows, 4] = details_json[i]["gNo"].ToString();
                   sheet5.Cells[rows, 5] = details_json[i]["codeTs"].ToString();
                   sheet5.Cells[rows, 6] = details_json[i]["gName"].ToString();
                   sheet5.Cells[rows, 7] = details_json[i]["gModel"].ToString();
                   sheet5.Cells[rows, 8] = "11";
                   sheet5.Cells[rows, 9] = details_json[i]["gQty"].ToString();
                   sheet5.Cells[rows, 10] = details_json[i]["gUnit"].ToString();
                   sheet5.Cells[rows, 11] = details_json[i]["qty1"].ToString();
                   sheet5.Cells[rows, 12] = details_json[i]["unit1"].ToString();
                   sheet5.Cells[rows, 13] = details_json[i]["declPrice"].ToString();
                   sheet5.Cells[rows, 14] = details_json[i]["declTotal"].ToString();
                   sheet5.Cells[rows, 15] = details_json[i]["cusOriginCountry"].ToString();
                   sheet5.Cells[rows, 16] = details_json[i]["tradeCurr"].ToString();
                   sheet5.Cells[rows, 17] = details_json[i]["dutyMode"].ToString();
                    //sheet5.Cells[rows, 21] = details_json[i]["cusTradeCountry"].ToString();
                    //sheet2.Cells[rows, 17] ="";

                }





                string usedrows = "X" + sheet2.UsedRange.Rows.Count.ToString();
                Range excelRange = sheet2.get_Range("A1", usedrows);//Borders.LineStyle 单元格边框线
                excelRange.Borders.LineStyle = 1; //单元格边框线类型(线型,实线型)
                //excelRange.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                usedrows = "X" + sheet5.UsedRange.Rows.Count.ToString();
                excelRange = sheet5.get_Range("A1", usedrows);
                excelRange.Borders.LineStyle = 1;
                //  excelRange.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                _wbk.SaveAs("d:\\WJD_AFC\\" + row_a[1] + "--" + row_b[1] + ".xls", XlFileFormat.xlExcel8);
                _wbk.Save();
                _wbk.Close();
                app.Quit();





            }
            else
            {
                MessageBox.Show("入园系统即进即出必须选择进口和出口两项报关单");

            }
            
          
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


        private void button_Click_1(object sender, RoutedEventArgs e)
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

                    string welcome = browser.FindElements(By.XPath("//span[@style='color: #fff; font-family: Microsoft YaHei; font-size: 14px;']"))[1].Text + "\n";
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

        
    }
}
