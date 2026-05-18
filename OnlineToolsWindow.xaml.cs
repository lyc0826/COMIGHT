using Microsoft.Web.WebView2.Core;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static COMIGHT.Methods;
using static COMIGHT.Settings;



namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for OnlineToolsWindow.xaml
    /// </summary>
     
    public partial class OnlineToolsWindow : Window
    {
        // 单个网址标签数据类（标签名 + 对应网址）
        private class WebsiteTag
        {
            public string Label { get; set; } = string.Empty;
            public string Url { get; set; } = string.Empty;
        }

        // 网址数据类（标签集合），需有公共无参构造函数以满足 SettingsManager<T> 的 new() 约束
        private class WebsiteData
        {
            public List<WebsiteTag> Tags { get; set; } = new List<WebsiteTag>();
        }

        public OnlineToolsWindow()
        {
            InitializeComponent();

            Dispatcher.Invoke(async () =>
            {
                await webView2.EnsureCoreWebView2Async(null); // 在UI线程上调用，显式初始化WebView控件
            });

            // 浏览器控件初始化完成后，触发 WebView_CoreWebView2InitializationCompleted 过程
            webView2.CoreWebView2InitializationCompleted += WebView_CoreWebView2InitializationCompleted!;
        }

        private void OnlineToolsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Top = 20.0;
            this.Left = 50.0;
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            webView2.GoBack(); // 后退一个网页
        }

        private async void BtnClean_Click(object sender, RoutedEventArgs e)
        {
            await CleanCache();
        }

        private void BtnForward_Click(object sender, RoutedEventArgs e)
        {
            webView2.GoForward(); // 前进一个网页
        }

        private void BtnReload_Click(object sender, RoutedEventArgs e)
        {
            webView2.Reload(); // 重新载入网址
        }

        private async Task CleanCache()
        {
            bool result = ShowMessage("Are you sure to clear all the browsing history and cookies?"); // 弹出对话框，询问是否确定清理
            if (result == true) // 如果对话框返回Yes（选择了"是"）
            {
                // 清理的数据类型赋值为所有网站和所有Cookies
                CoreWebView2BrowsingDataKinds dataKinds = CoreWebView2BrowsingDataKinds.AllSite | CoreWebView2BrowsingDataKinds.Cookies;

                // 清理时间范围赋值为从1970年1月1日0点0分0秒开始到当前
                DateTime startTime = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                DateTime endTime = DateTime.UtcNow;
                await webView2.CoreWebView2.Profile.ClearBrowsingDataAsync(dataKinds, startTime, endTime); // 清理数据
            }

            ShowSuccessMessage();
        }

        // WebView初始化完成后，执行此过程
        private void WebView_CoreWebView2InitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            try
            {

                // 使用 SettingsManager 加载网址Json文件（文件不存在时会返回默认空对象）
                SettingsManager<WebsiteData> websitesManager = new SettingsManager<WebsiteData>(websitesJsonFilePath);
                WebsiteData websiteData = websitesManager.GetSettings();

                // 在 WrapPanel 中生成标签云
                BuildTagCloud(websiteData);


                // 添加webView事件响应过程
                webView2.NavigationCompleted += WebView_NavigationCompleted;       // 打开网站完成后触发
                webView2.CoreWebView2.NewWindowRequested += WebView_NewWindowRequested; // 出现新建浏览窗口请求时触发
            }
            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        // 在 WrapPanel 中生成标签云：每个标签为一个按钮，字号与底色随机变化以呈现标签云效果
        private void BuildTagCloud(WebsiteData websiteData)
        {
            wrpnlTags.Children.Clear(); // 先清空已有标签

            if (websiteData?.Tags == null || websiteData.Tags.Count == 0)
            {
                return;
            }


            foreach (WebsiteTag tag in websiteData.Tags)
            {
                // 跳过标签名或网址为空白的项
                if (string.IsNullOrWhiteSpace(tag.Label) || string.IsNullOrWhiteSpace(tag.Url))
                {
                    continue;
                }

                // 创建标签按钮：Content 为显示文字，Tag 属性存放对应网址，便于点击时取用
                Button tagButton = new Button
                {
                    Content = tag.Label,
                    Tag = tag.Url, // 存放网址，便于点击时取用
                    Margin = new Thickness(3),
                    Padding = new Thickness(3, 3, 3, 3),
                    Width = double.NaN,
                    FontSize = 13,
                    BorderThickness = new Thickness(1),
                    Cursor = Cursors.Hand,
                   
                };

                tagButton.Click += TagButton_Click; // 绑定点击事件
                wrpnlTags.Children.Add(tagButton);
            }
        }

        // 标签按钮被点击时打开对应网址
        private void TagButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string url)
            {
                WebView_OpenNewUrl(url);
            }
        }

        // 网站加载完成后，执行此过程
        private async void WebView_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            /* 添加JS代码：
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
                            event.target.style.boxShadow = ' 
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
                                event.target.style.boxShadow = 'inset 0 0 3px rgba(255, 80, 80, 0.8)' 
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
        }

        private void WebView_NewWindowRequested(object? sender, CoreWebView2NewWindowRequestedEventArgs e)
        {
            string url = e.Uri.ToString(); // 将新建浏览窗口请求事件中的网址转换成字符串
            if (!url.Contains("oauth")) // 如果网址中不含输入用户名密码的验证标志
            {
                WebView_OpenNewUrl(url); // 打开网址
                e.Handled = true;        // 将Handled属性设为true，表明打开新网址事件已处理，禁止弹窗 
            }
        }

        public void WebView_OpenNewUrl(string? url)
        {
            if (!string.IsNullOrWhiteSpace(url)) // 如果网址变量不为null也不为空白
            {
                // 正则表达式匹配模式设为：开头标记，"http"，"s"至多一个，"://"，
                // 如果输入网址匹配失败，则在输入网址前加上 "http://"
                if (!Regex.IsMatch(url, @"^http[s]?://"))
                {
                    url = "http://" + url;
                }
                webView2.CoreWebView2.Navigate(url); // 打开网址
                // latestRecords.LatestUrl = url;     // 将正打开的网址赋值给用户使用记录
            }
        }
    }
}
