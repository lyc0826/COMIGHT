using Microsoft.Web.WebView2.Core;
using OfficeOpenXml;
using System.Data;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static COMIGHT.Methods;

namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for BrowserWindow.xaml
    /// </summary>
    public partial class BrowserWindow : Window
    {
        public BrowserWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  //定义EPPlus库许可证类型为非商用！！！

            //Dispatcher.Invoke(() => EnsureCoreWebView2());
            //async void EnsureCoreWebView2()
            //{
            //    await webView2.EnsureCoreWebView2Async(null); //在UI线程上调用，显式初始化WebView控件
            //}

            Dispatcher.Invoke(async () =>
                {
                    await webView2.EnsureCoreWebView2Async(null); //在UI线程上调用，显式初始化WebView控件
                });

            //浏览器控件初始化完成后，触发WebView_CoreWebView2InitializationCompleted过程
            webView2.CoreWebView2InitializationCompleted += WebView_CoreWebView2InitializationCompleted!;

        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            webView2.GoBack(); //后退一个网页
        }

        private async void BtnClean_Click(object sender, RoutedEventArgs e)
        {
            await CleanCache();
        }

        private void BtnForward_Click(object sender, RoutedEventArgs e)
        {
            webView2.GoForward(); //前进一个网页
        }

        private void BtnReload_Click(object sender, RoutedEventArgs e)
        {
            webView2.Reload(); //重新载入网址
        }

        private void CmbUrl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? url = cmbURL.SelectedItem as string; //将组合框被选项的文字赋值给网址变量
            WebView_OpenNewUrl(url); //打开网址
        }

        private void TxtbxURL_KeyDown(object sender, KeyEventArgs e)
        {
            string url = txtbxURL.Text;
            if (e.Key == Key.Enter) //如果按下的是回车键，则打开网址
            {
                WebView_OpenNewUrl(url);
            }
        }

        private void TxtbxURL_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            txtbxURL.SelectAll();
        }

        private async Task CleanCache()
        {
            MessageBoxResult result = MessageBox.Show("确定要清理所有浏览数据和cookie吗?", "确认", MessageBoxButton.YesNo); //弹出对话框，询问是否确定清理
            if (result == MessageBoxResult.Yes) //如果对话框返回Yes（选择了“是”）
            {
                //清理的数据类型赋值为所有网站和所有Cookies
                CoreWebView2BrowsingDataKinds dataKinds = CoreWebView2BrowsingDataKinds.AllSite | CoreWebView2BrowsingDataKinds.Cookies;

                //清理时间范围赋值为从1970年1月1日0点0分0秒开始到当前
                DateTime startTime = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                DateTime endTime = DateTime.UtcNow;
                await webView2.CoreWebView2.Profile.ClearBrowsingDataAsync(dataKinds, startTime, endTime); //清理数据
            }
        }

        private void WebView_CoreWebView2InitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs e) //WebView初始化完成后，执行此过程
        {

            DataTable? dataTable = ReadExcelWorksheetIntoDataTableAsString(dataBaseFilePath, "Websites"); //读取数据库Excel工作簿的“网址”工作表，赋值给DataTable变量

            if (dataTable != null) //如果DataTable变量不为null
            {
                foreach (DataRow dataRow in dataTable.Rows) //遍历所有数据行
                {
                    cmbURL.Items.Add(Convert.ToString(dataRow["Website"]));  //将当前数据行的"Website"数据列的数据添加到网址组合框
                }
            }

            cmbURL.SelectedIndex = 0; //网址列表组合框选择0号（第1）项
            string startupURL = Convert.ToString(dataTable!.Rows[0]["Website"])!; //将DataTable第0行的"Website"数据赋值给起始URL变量
            WebView_OpenNewUrl(startupURL); //打开起始URL

            //添加webView事件响应过程
            webView2.NavigationCompleted += WebView_NavigationCompleted; //打开网站完成后，触发WebView_NavigationCompleted过程
            webView2.CoreWebView2.NewWindowRequested += WebView_NewWindowRequested; //出现新建浏览窗口请求时，触发CoreWebView2_NewWindowRequested过程
        }

        private async void WebView_NavigationCompleted(object? sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e) //网站加载完成后，执行此过程
        {
            txtbxURL.Text = webView2.Source.ToString(); //将正打开的网址赋值给网址文本框

            /*添加JS代码：
              向网页添加鼠标悬停事件响应，如果鼠标悬停处的标签为div标记，在该处加上绿色阴影；
              向网页添加鼠标移出事件响应，如果鼠标移出处的标签为div标记，将该处的阴影取消；
              向网页添加鼠标双击事件响应，如果鼠标双击处的标签为div标记，复制该处文字，在该处加上红色阴影，0.5秒后复原；
            */

            await webView2.ExecuteScriptAsync(@"

                document.body.addEventListener('mouseover', 
                    function(event) 
                    {
                        if (event.target.tagName.toLowerCase() === 'div') 
                        {
                            event.target.style.boxShadow = 'inset 0 0 0 3px rgba(80, 255, 80, 0.8)'; 
                        }  
                    });

                document.body.addEventListener('mouseout', 
                    function(event) 
                    {
                        if (event.target.tagName.toLowerCase() === 'div') 
                        {
                            event.target.style.boxShadow = '' 
                        }
                    });

                document.body.addEventListener('dblclick', 
                    function(event) 
                    {
                        if (event.target.tagName.toLowerCase() === 'div') 
                        {
                            var originalBoxShadow = event.target.style.boxShadow;
                            var textToCopy = event.target.innerText;

                            var textArea = document.createElement('textarea');
                            textArea.value = textToCopy;
                            document.body.appendChild(textArea);
                            textArea.select();
                            var successful = document.execCommand('copy');
                            document.body.removeChild(textArea);

                            if(successful)
                            {
                                event.target.style.boxShadow = 'inset 0 0 0 3px rgba(255, 80, 80, 0.8)' 
                                setTimeout(
                                    function() 
                                    {
                                        event.target.style.boxShadow = originalBoxShadow 
                                    }
                                    , 500);    
                            } 
                            else 
                            {
                                alert('复制失败！');
                            }
                        }
                    });
                ");

        }

        private void WebView_NewWindowRequested(object? sender, CoreWebView2NewWindowRequestedEventArgs e)
        {
            string url = e.Uri.ToString(); //将新建浏览窗口请求事件中的网址转换成字符串
            if (!url.Contains("oauth")) //如果网址中不含输入用户名密码的验证标志
            {
                WebView_OpenNewUrl(url); //打开网址
                e.Handled = true; //将Handled属性设为true，表明打开新网址事件已处理，禁止弹窗 
            }
        }

        public void WebView_OpenNewUrl(string? url)
        {
            if (url != null && url.Length > 0)  //如果网址变量不为null且字数大于0
            {
                //正则表达式匹配模式设为：开头标记，“http”，“s”至多一个，“://”，如果输入网址匹配失败，则在输入网址前加上"http://"
                if (!Regex.IsMatch(url, @"^http[s]?://"))
                {
                    url = "http://" + url;
                }
                webView2.CoreWebView2.Navigate(url); //打开网址，WebView.Source = new Uri(url) 
                txtbxURL.Text = url; //将正打开的网址赋值给网址文本框 

            }
        }

        private void BrowserWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Top = 50.0;
            this.Left = 100.0;
        }


    }
}
