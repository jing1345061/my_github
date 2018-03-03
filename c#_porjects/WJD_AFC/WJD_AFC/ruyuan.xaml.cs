using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
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
using CsharpHttpHelper;
using CsharpHttpHelper.Enum;

namespace WJD_AFC
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

        string file_path;
        private void button_openfiles_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();

            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                this.textBox.Text = openFileDialog.FileName;
                file_path = openFileDialog.FileName;
            }
        }



        static IWebDriver browser;
        ChromeOptions options = new ChromeOptions();
       
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
        public void denlu_ruyuan()
        {
            set_chrome();

            browser.Navigate().GoToUrl("http://www.dgqxbswlzx.com");
            Thread.Sleep(1000);
            //textBox.Text = browser.Title;
            IWebElement dianji = browser.FindElement(By.XPath("//h1"));
            dianji.Click();
            Thread.Sleep(300);
            IWebElement username = browser.FindElement(By.XPath("//input[@id='username']"));
            IWebElement password = browser.FindElement(By.XPath("//input[@id='password']"));
            IWebElement submit = browser.FindElement(By.XPath("//input[@name='commit']"));
            username.SendKeys("4469W6K001001");
            password.SendKeys("wjd13650116965wjd");
            submit.Click();
            browser.Navigate().GoToUrl("http://www.dgqxbswlzx.com/html/homePage.html?username=4469W6K001001");
            Thread.Sleep(1000);
            ICookieJar danyi_cookie = browser.Manage().Cookies;
            string cookies_str = "{\"ruyuan_cookies\":[";
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

            string fp = System.IO.Directory.GetCurrentDirectory() + "\\ruyuan_cookies.json";

            File.WriteAllText(fp, JsonConvert.SerializeObject(cookies_json));

            
        }
        private bool denluzhuantai()
        {
            string fp = System.IO.Directory.GetCurrentDirectory() + "\\ruyuan_cookies.json";
            JObject json1 = (JObject)JsonConvert.DeserializeObject(File.ReadAllText(fp));
            JArray ck_array = (JArray)json1["ruyuan_cookies"];
            string ck = "";
            ck = ck_array[0]["name"].ToString() + "=" + ck_array[0]["value"].ToString();

            string get_url = "http://www.dgqxbswlzx.com/getAccountInfo.action";
            System.Net.HttpWebRequest request_get = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(get_url);
            request_get.Method = "GET";
            request_get.Accept = "application/json, text/javascript, */*; q=0.01";
            //request_get.ContentType = "application/json;charset=UTF-8";
            request_get.Headers.Add("ContentType", "application/json;charset=UTF-8");
            request_get.Headers.Add("Cookie", ck);
            request_get.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36";
            request_get.Timeout = 10000;
            System.Net.HttpWebResponse response_get = (System.Net.HttpWebResponse)request_get.GetResponse();
            Stream myResponseStream = response_get.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            string username = "4469W6K001001";
            if (retString.Contains(username))
            {
                return true;
            }
            else
            {
                return false;

            }
        }
        private void button_daoru_Click(object sender, RoutedEventArgs e)
        {
            
            
            if (denluzhuantai()==true)
            {
               MessageBox.Show( HttpUploadFile("http://www.dgqxbswlzx.com/html/dayTravel/loadMergEntExcel.action",file_path));
            }
            else
            {
                denlu_ruyuan();
                MessageBox.Show(HttpUploadFile("http://www.dgqxbswlzx.com/html/dayTravel/loadMergEntExcel.action", file_path));
            }





        }

        public static string HttpUploadFile(string url, string path)
        {
            string fp = System.IO.Directory.GetCurrentDirectory() + "\\ruyuan_cookies.json";
            JObject json1 = (JObject)JsonConvert.DeserializeObject(File.ReadAllText(fp));
            JArray ck_array = (JArray)json1["ruyuan_cookies"];
            string ck = "";
            ck = ck_array[0]["name"].ToString() + "=" + ck_array[0]["value"].ToString();

            // 设置参数
            System.Net.HttpWebRequest request = System.Net.WebRequest.Create(url) as System.Net.HttpWebRequest;
           // System.Net.CookieContainer cookieContainer = new System.Net.CookieContainer();
           // request.CookieContainer = cookieContainer;
           // request.AllowAutoRedirect = true;
            request.Method = "POST";
            request.Headers.Add("Cookie", ck);
            string boundary = "----WebKitFormBoundarytl5ic61YjIGTI4d9"; // 分隔线
            request.ContentType = "multipart/form-data;charset=utf-8;boundary=" + boundary;
            byte[] itemBoundaryBytes = Encoding.UTF8.GetBytes("\r\n--" + boundary + "\r\n");
            byte[] endBoundaryBytes = Encoding.UTF8.GetBytes("\r\n--" + boundary + "--\r\n");
            int pos = path.LastIndexOf("\\");
            string fileName = path.Substring(pos + 1);

            //请求头部信息 
            StringBuilder sbHeader = new StringBuilder(string.Format("Content-Disposition:form-data;name=\"uploadFile\";filename=\"" +fileName+"\"\r\nContent-Type:application/vnd.ms-excel\r\n\r\n"));
            byte[] postHeaderBytes = Encoding.UTF8.GetBytes(sbHeader.ToString());

            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            byte[] bArr = new byte[fs.Length];
            fs.Read(bArr, 0, bArr.Length);
            fs.Close();

            Stream postStream = request.GetRequestStream();
            postStream.Write(itemBoundaryBytes, 0, itemBoundaryBytes.Length);
            postStream.Write(postHeaderBytes, 0, postHeaderBytes.Length);
            postStream.Write(bArr, 0, bArr.Length);
            postStream.Write(endBoundaryBytes, 0, endBoundaryBytes.Length);
            postStream.Close();

            //发送请求并获取相应回应数据
            System.Net.HttpWebResponse response = request.GetResponse() as System.Net.HttpWebResponse;
            //直到request.GetResponse()程序才开始向目标网页发送Post请求
            Stream instream = response.GetResponseStream();
            StreamReader sr = new StreamReader(instream, Encoding.UTF8);
            //返回结果网页（html）代码
            string content = sr.ReadToEnd();
            return content+fileName;
        }
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            denlu_ruyuan();
        }
    }
}
