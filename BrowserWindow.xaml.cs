using Microsoft.Web.WebView2.Core;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static COMIGHT.PublicVariables;
using static COMIGHT.MainWindow;


namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for BrowserWindow.xaml
    /// </summary>
    public partial class BrowserWindow : Window
    {

        internal class WebsiteData // 定义网址数据类
        {
            public List<string> Urls { get; set; } = new List<string>(); // 定义网址列表属性（Json反序列化时，属性访问权限必须为public才能正常访问）
        }

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

        private void BrowserWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Top = 20.0;
            this.Left = 50.0;
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

        private void CmbbxUrl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? url = cmbbxUrl.SelectedItem?.ToString() ?? ""; //将组合框被选项的文字赋值给网址变量
            WebView_OpenNewUrl(url); //打开网址
        }

        private void TxtbxUrl_KeyDown(object sender, KeyEventArgs e)
        {
            string url = txtbxUrl.Text;
            if (e.Key == Key.Enter) //如果按下的是回车键，则打开网址
            {
                WebView_OpenNewUrl(url);
            }
        }

        private void TxtbxURL_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            txtbxUrl.SelectAll();
        }

        private async Task CleanCache()
        {
            MessageBoxResult result = MessageBox.Show("Are you sure to clear all the browsing history and cookies?", "Inquiry", MessageBoxButton.YesNo); //弹出对话框，询问是否确定清理
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
            try
            {
                // 检查网址Json文件是否存在
                if (!File.Exists(websitesJsonFilePath)) // 如果网址Json文件不存在，则抛出异常
                {
                    throw new Exception("JSON file not found.");
                }

                string jsonContent = File.ReadAllText(websitesJsonFilePath); // 读取网址Json文件内容

                // 反序列化网址Json内容，赋值给网址数据对象变量
                WebsiteData? websiteData = JsonConvert.DeserializeObject<WebsiteData>(jsonContent);

                if (websiteData?.Urls != null && websiteData.Urls.Count > 0) // 如果网址列表属性不为null也不为空
                {
                    // 将网址添加到网址组合框中
                    foreach (string url in websiteData.Urls) // 遍历网址列表
                    {
                        cmbbxUrl.Items.Add(url); // 将当前网址添加到网址组合框中
                    }
                }

                // 获取起始网址：如果用户使用记录中的最近打开网址不为null或全空白字符，则得到该网址；否则，如果网址组合框的选项数大于0且0号选项字符串不为null或全空白字符，则获取该0号项目字符串
                string? startupUrl = !string.IsNullOrWhiteSpace(latestRecords.LatestUrl) ? latestRecords.LatestUrl : ( cmbbxUrl.Items.Count > 0 && !string.IsNullOrWhiteSpace(cmbbxUrl.Items[0].ToString()) ) ? cmbbxUrl.Items[0].ToString() : string.Empty;

                WebView_OpenNewUrl(startupUrl); //打开起始网址

                //添加webView事件响应过程
                webView2.NavigationCompleted += WebView_NavigationCompleted; //打开网站完成后，触发WebView_NavigationCompleted过程
                webView2.CoreWebView2.NewWindowRequested += WebView_NewWindowRequested; //出现新建浏览窗口请求时，触发CoreWebView2_NewWindowRequested过程
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private async void WebView_NavigationCompleted(object? sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e) //网站加载完成后，执行此过程
        {
            txtbxUrl.Text = webView2.Source.ToString(); //将正打开的网址赋值给网址文本框

            /*添加JS代码：
              向网页添加鼠标悬停事件响应，如果鼠标悬停处的标签为div标记，在该处加上绿色阴影；
              向网页添加鼠标移出事件响应，如果鼠标移出处的标签为div标记，将该处的阴影取消；
              向网页添加鼠标双击事件响应，如果鼠标双击处的标签为div标记，复制该处文字，在该处加上红色阴影，0.5秒后复原；
            */

            string jsScript = @"

                document.body.addEventListener('mouseover', 
                    function(event) 
                    {
                        if (event.target.tagName.toLowerCase() === 'div') 
                        {
                            event.target.style.boxShadow = 'inset 0 0 0 3px rgba(80, 255, 80, 0.8)'; // 内阴影，水平偏移0，垂直偏移0，模糊半径0，颜色rgba(80, 255, 80, 0.8)
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

                            textToCopy = textToCopy.replace(/^\s*[\r\n]/gm, '')  // 移除空白段落
                                                   .replace(/^\s+|\s+$/gm, '');   // 移除每行首尾空白 （全局+多行模式）

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
                                alert('Copying Failed!');
                            }
                        }
                    });
                ";

            await webView2.ExecuteScriptAsync(jsScript);
            
            //await webView2.ExecuteScriptAsync(@"

            //    document.body.addEventListener('mouseover', 
            //        function(event) 
            //        {
            //            if (event.target.tagName.toLowerCase() === 'div') 
            //            {
            //                event.target.style.boxShadow = 'inset 0 0 0 3px rgba(80, 255, 80, 0.8)'; // 内阴影，水平偏移0，垂直偏移0，模糊半径0，颜色rgba(80, 255, 80, 0.8)
            //            }  
            //        });

            //    document.body.addEventListener('mouseout', 
            //        function(event) 
            //        {
            //            if (event.target.tagName.toLowerCase() === 'div') 
            //            {
            //                event.target.style.boxShadow = '' 
            //            }
            //        });

            //    document.body.addEventListener('dblclick', 
            //        function(event) 
            //        {
            //            if (event.target.tagName.toLowerCase() === 'div') 
            //            {
            //                var originalBoxShadow = event.target.style.boxShadow;
            //                var textToCopy = event.target.innerText;

            //                textToCopy = textToCopy.replace(/^\s*[\r\n]/gm, '')  // 移除空白段落
            //                                       .replace(/^\s+|\s+$/gm, '');   // 移除每行首尾空白 （全局+多行模式）

            //                var textArea = document.createElement('textarea');
            //                textArea.value = textToCopy;
            //                document.body.appendChild(textArea);
            //                textArea.select();
            //                var successful = document.execCommand('copy');
            //                document.body.removeChild(textArea);

            //                if(successful)
            //                {
            //                    event.target.style.boxShadow = 'inset 0 0 0 3px rgba(255, 80, 80, 0.8)' 
            //                    setTimeout(
            //                        function() 
            //                        {
            //                            event.target.style.boxShadow = originalBoxShadow 
            //                        }
            //                        , 500);    
            //                } 
            //                else 
            //                {
            //                    alert('Copying Failed!');
            //                }
            //            }
            //        });
            //    ");

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
                txtbxUrl.Text = url; //将正打开的网址赋值给网址文本框
                latestRecords.LatestUrl = url; // 将正打开的网址保存到用户使用记录中
                recordsManager.SaveSettings(latestRecords); // 保存用户使用记录

            }
        }

    }
}
