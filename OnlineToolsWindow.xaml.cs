using Microsoft.Web.WebView2.Core;
using System.Windows.Media;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static COMIGHT.Methods;
using static COMIGHT.AppDataManager;


namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for OnlineToolsWindow.xaml
    /// </summary>
     
    public partial class OnlineToolsWindow : Window
    {

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
                ShowSuccessMessage();
            }  
        }

        // WebView初始化完成后，执行此过程
        private void WebView_CoreWebView2InitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            try
            {
                // 禁用WebView的状态栏
                webView2.CoreWebView2.Settings.IsStatusBarEnabled = false;
                // 在 WrapPanel 中生成网址标签，数据来源为网址Json文件
                BuildWebsiteTags(websiteData);

                // 添加webView事件响应过程
                webView2.NavigationCompleted += WebView_NavigationCompleted;       // 打开网站完成后触发
                webView2.CoreWebView2.NewWindowRequested += WebView_NewWindowRequested; // 出现新建浏览窗口请求时触发
                WebView_OpenNewUrl(userRecords.OpenedWebsite); // 打开用户使用记录中记录的网址
            }
            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        // 生成网址标签：每个标签为一个按钮
        private void BuildWebsiteTags(WebsiteData websiteData)
        {
            wrpnlTags.Children.Clear(); // 清空WrapPanel中已有的标签

            // 如果网址标签列表为空，则不执行后续操作
            if (websiteData?.WebsiteTags == null || websiteData.WebsiteTags.Count == 0)
            {
                return;
            }

            foreach (WebsiteTag websiteTag in websiteData.WebsiteTags)
            {
                // 跳过标签名或网址为空白的项
                if (string.IsNullOrWhiteSpace(websiteTag.Label) || string.IsNullOrWhiteSpace(websiteTag.Url))
                {
                    continue;
                }

                // 创建标签按钮：Content 为显示文字，Tag 属性存放对应网址，便于点击时取用
                Button tagButton = new Button
                {
                    Content = websiteTag.Label,
                    Tag = websiteTag.Url, // 存放网址，便于点击时取用
                    Margin = new Thickness(3),
                    Padding = new Thickness(3, 3, 3, 3),
                    Width = double.NaN, // 自动宽度
                    FontSize = 13,
                    BorderThickness = new Thickness(1),
                    Cursor = Cursors.Hand,
                    Background = Brushes.LightBlue,
                    Foreground = Brushes.Black,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
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
        //private async void WebView_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        //{
        //    /* 添加JS代码：
        //      向网页添加鼠标悬停事件响应，如果鼠标悬停处的标签为div标记，在该处加上绿色阴影；
        //      向网页添加鼠标移出事件响应，如果鼠标移出处的标签为div标记，将该处的阴影取消；
        //      向网页添加鼠标双击事件响应，如果鼠标双击处的标签为div标记，复制该处文字，在该处加上红色阴影，0.5秒后复原；
        //    */
        //    string jsScript = @"

        //        document.body.addEventListener('mouseover', 
        //            function(event) 
        //            {
        //                if (event.target.tagName.toLowerCase() === 'div') 
        //                {
        //                    event.target.style.boxShadow = 'inset 0 0 0 2px rgba(80, 255, 80, 0.8)'; // 内阴影，水平偏移0，垂直偏移0，模糊半径0，宽度3px，颜色：红、绿、蓝、透明度(80, 255, 80, 0.7) 
        //                }  
        //            });

        //        document.body.addEventListener('mouseout', 
        //            function(event) 
        //            {
        //                if (event.target.tagName.toLowerCase() === 'div') 
        //                {
        //                    event.target.style.boxShadow = '';
        //                }
        //            });

        //        document.body.addEventListener('dblclick', 
        //            function(event) 
        //            {
        //                if (event.target.tagName.toLowerCase() === 'div') 
        //                {
        //                    var target = event.target;
        //                    var originalBoxShadow = target.style.boxShadow;
        //                    var textToCopy = target.innerText
        //                        .replace(/^\s*[\r\n]/gm, '')
        //                        .replace(/^\s+|\s+$/gm, '');
        //                    navigator.clipboard.writeText(textToCopy).then(
        //                        function() {
        //                            target.style.boxShadow = 'inset 0 0 0 2px rgba(255, 80, 80, 0.6)';
        //                            setTimeout(function() {
        //                                target.style.boxShadow = originalBoxShadow;
        //                            }, 500);
        //                        },
        //                        function(err) {
        //                            alert('Copying Failed: ' + err);
        //                        }
        //                    );
        //                }
        //            }); 

        //    ";

        //    await webView2.ExecuteScriptAsync(jsScript);
        //}

        private async void WebView_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            /* 功能清单：
               1. div鼠标悬停绿阴影、移出取消、双击复制文字+红色闪烁阴影
               2. 屏蔽外部链接、邮件(mailto)、电话(tel)链接
               3. 终极屏蔽：普通广告 + 弹出式广告/模态弹窗/遮罩弹窗（无闪烁、无残留）
            */
            string jsScript = @"

                // ==================== 精准屏蔽广告（不影响正常弹出菜单） ====================
                (function() {
                    // 1. 【核心】纯广告专属选择器（仅匹配广告，不碰正常UI）
                    const adBlockSelectors = [
                        '[class*=ad-],[class*=-ad],[class*=ads],[class*=advert],[class*=adzone],[class*=banner]',
                        '[id*=ad-],[id*=-ad],[id*=ads],[id*=advert],[id*=adzone],[id*=banner]',
                        'iframe[src*=ad],iframe[src*=ads],iframe[src*=advertisement]',
                        '[class*=gg],[class*=guanggao]'
                    ].join(',');

                    // 2. 【核心】正常UI白名单（不处理：弹出菜单/下拉框/模态框/提示框等）
                    const safeUISelectors = [
                        // 所有正常弹出式菜单、下拉、模态框、提示层（彻底避免误伤）
                        '[class*=dropdown],[class*=menu],[class*=popup],[class*=modal],[class*=tooltip],[class*=popover]',
                        '[class*=nav],[class*=select],[class*=option],[class*=list],[class*=hover]'
                    ].join(',');

                    // 3. 安全隐藏广告（仅隐藏，不删除DOM，不破坏正常UI）
                    function safeHideAd(el) {
                        // 白名单1：排除网页根节点，避免崩溃
                        const rootTags = ['HTML', 'BODY', 'MAIN', 'SECTION', 'ARTICLE', 'NAV', 'HEADER', 'FOOTER', 'DIV'];
                        if (!el || !el.style || rootTags.includes(el.tagName)) return;

                        // 白名单2：【核心】如果是正常弹出菜单/UI，直接跳过，绝不处理
                        if (el.matches(safeUISelectors) || el.closest(safeUISelectors)) return;

                        //// 仅温和隐藏广告，不清空HTML、不删除节点（避免破坏网页结构）
                        //el.style.display = 'none';
                        //el.style.visibility = 'hidden';
                        //el.style.pointerEvents = 'none';
                        //el.style.opacity = '0';
                        
                        // 隐藏、清空、删除DOM
                        el.style.display = 'none';
                        el.style.visibility = 'hidden';
                        el.style.pointerEvents = 'none';
                        el.style.opacity = '0';
                        el.innerHTML = ''; 
                        el.remove(); 
                    }

                    // 4. 批量屏蔽：仅匹配「广告元素」且「排除正常UI」
                    function blockAllAds() {
                        document.querySelectorAll(`${adBlockSelectors}:not(${safeUISelectors})`).forEach(item => safeHideAd(item));
                    }

                    // 5. 轻量化监听（仅处理新增的广告节点，不干扰正常UI）
                    const adObserver = new MutationObserver(mutations => {
                        for (let mutation of mutations) {
                            for (let node of mutation.addedNodes) {
                                if (node.nodeType === 1) { // 仅处理元素节点
                                    // 先跳过正常UI，再匹配广告
                                    if (node.matches(safeUISelectors)) continue;
                                    const adNode = node.closest(adBlockSelectors);
                                    if (adNode) safeHideAd(adNode);
                                }
                            }
                        }
                    });

                    // 启动监听
                    adObserver.observe(document.body, {
                        childList: true,
                        subtree: true
                    });

                    // 初始执行屏蔽
                    blockAllAds();

                    // 6. 精准拦截恶意广告弹窗（不影响正常window.open）
                    const originalOpen = window.open;
                    window.open = function(...args) {
                        const url = args[0] || '';
                        // 仅拦截：无地址 / 纯广告链接，放行所有正常弹窗
                        const isAdPopup = !url || url.includes('advertisement') || url.includes('guanggao') || url.includes('gg');
                        return isAdPopup ? null : originalOpen.apply(this, args);
                    };

                })();


                // ==================== 【优化版】Header 净化 + Footer 全隐藏 ====================
                (function() {
                    // 1. 【核心白名单】100%保护所有导航菜单，绝不误隐藏
                    const safeMenuSelectors = [
                        // 导航核心
                        'nav, [class*=nav],[id*=nav],[class*=navigation],[id*=navigation]',
                        // 菜单
                        '[class*=menu],[id*=menu],[class*=menus],[id*=menus]',
                        // 下拉/按钮/链接
                        '[class*=dropdown],[class*=btn],[class*=button],[class*=link],[class*=item]',
                        '[class*=nav-item],[class*=nav-link],[class*=menu-item]',
                        // 列表容器
                        'ul, li, a'
                    ].join(',');

                    // 2. 【核心Logo匹配】强制隐藏所有Logo图片/文字/容器
                    const logoSelectors = [
                        '[class*=logo],[id*=logo]',
                        '[class*=brand],[id*=brand]',
                        '[class*=site-logo],[class*=brand-logo],[class*=logotype]',
                        'img[class*=logo],img[id*=logo],img[alt*=logo],img[alt*=品牌]'
                    ].join(',');

                    // 处理 Header：只隐藏Logo，保留所有菜单
                    function cleanHeader() {
                        const headers = document.querySelectorAll('header');
                        headers.forEach(header => {
                            // 第一步：强制隐藏所有Logo（图片/容器/文字）
                            document.querySelectorAll(logoSelectors).forEach(logo => {
                                if (logo.closest('header')) {
                                    logo.style.display = 'none';
                                    logo.style.visibility = 'hidden';
                                    logo.style.width = '0';
                                    logo.style.height = '0';
                                }
                            });

                            // 第二步：保护所有导航菜单，绝不隐藏
                            header.querySelectorAll('*').forEach(child => {
                                // 如果是菜单相关元素，直接跳过，不做任何隐藏
                                if (child.matches(safeMenuSelectors) || child.closest(safeMenuSelectors)) {
                                    return;
                                }
                            });
                        });
                    }

                    // 处理 Footer：全部隐藏
                    function hideFooter() {
                        document.querySelectorAll('footer').forEach(footer => {
                            footer.style.display = 'none';
                            footer.style.visibility = 'hidden';
                            footer.style.height = '0';
                            footer.style.margin = '0';
                            footer.style.padding = '0';
                        });
                    }

                    // 初始执行
                    cleanHeader();
                    hideFooter();

                    // 监听动态加载
                    const layoutObserver = new MutationObserver(() => {
                        cleanHeader();
                        hideFooter();
                    });

                    layoutObserver.observe(document.body, { childList: true, subtree: true });
                })();

                
                //// ==================== 屏蔽邮件/电话/外部网站超链接 ====================
                //document.body.addEventListener('click', function(event) {
                //    const targetLink = event.target.closest('a');
                //    if (!targetLink) return;

                //    const linkHref = targetLink.href || '';
                //    const currentSiteOrigin = window.location.origin;
                //    let needBlock = false;

                //    if (linkHref.startsWith('mailto:')) needBlock = true;
                //    else if (linkHref.startsWith('tel:')) needBlock = true;
                //    else if (targetLink.origin !== currentSiteOrigin) needBlock = true;

                //    if (needBlock) {
                //        event.preventDefault();
                //    }
                //});

                // ==================== 隐藏指定链接（邮件/电话/代码仓库/社交媒体） ====================
                // 保留元素占位 = 不破坏网页布局；透明+禁用交互 = 完全不可见不可点
                (function() {

                    const currentHost = window.location.hostname;

                    // 匹配规则：协议头 + 域名关键词（全覆盖主流平台）
                    const linkSelectors = [
                        // 1. 邮件/电话/短信 链接
                        'a[href^=""mailto:""]',
                        'a[href^=""tel:""]',
                        'a[href^=""sms:""]',

                        //// 2. 代码仓库
                        //'a[href*=""github.com""]',
                        //'a[href*=""gitlab.com""]',
                        //'a[href*=""gitee.com""]',
                        //'a[href*=""bitbucket.org""]',
                        //'a[href*=""gitcode.net""]',
                        //'a[href*=""coding.net""]',
                        //'a[href*=""codeberg.org""]',

                        //// 3. 社交媒体
                        //'a[href*=""weixin.qq.com""]',
                        //'a[href*=""weibo.com""]',
                        //'a[href*=""qq.com""]',
                        //'a[href*=""twitter.com""]',
                        //'a[href*=""x.com""]',
                        //'a[href*=""facebook.com""]',
                        //'a[href*=""instagram.com""]',
                        //'a[href*=""youtube.com""]',
                        //'a[href*=""tiktok.com""]',
                        //'a[href*=""linkedin.com""]',
                        //'a[href*=""instagram.com""]',
                        //'a[href*=""discord.com""]',
                        //'a[href*=""t.me""]',
                        //'a[href*=""whatsapp.com""]',
                        //'a[href*=""bsky.app""]',
                        //'a[href*=""douyin.com""]',
                        //'a[href*=""xiaohongshu.com""]',
                        //'a[href*=""bilibili.com""]',

                    ].join(',');

                    // 安全隐藏：保留布局占位，仅隐藏视觉+禁用交互
                    function hideTargetLink(el) {
                        if (!el || !el.style) return;
                        // 核心样式：visibility保留占位，opacity透明，禁止点击
                        el.style.display = 'none';
                        el.style.visibility = 'hidden';
                        el.style.opacity = '0';
                        el.style.pointerEvents = 'none';
                        el.style.userSelect = 'none';
                    }

                    // 用 endsWith 判断本站/子域名还是外部链接
                    function hideExternalLinks() {
                        document.querySelectorAll('a[href]').forEach(link => {
                            try {
                                // 跳过：无效链接、锚点、相对路径（本站内部链接）
                                if (!link.href || link.href.startsWith('#') || link.href.startsWith('/')) return;
                        
                                const linkHost = link.hostname;
                                // 放行规则：
                                // 1. 完全相同主机名 → 本站
                                // 2. 以 .当前主机名 结尾 → 子域名（如 www.abc.com → abc.com）
                                const isSelfSite = linkHost === currentHost || linkHost.endsWith(`.${currentHost}`);
                        
                                // 非本站 → 隐藏
                                if (!isSelfSite) {
                                    hideTargetLink(link);
                                }
                            } catch (e) {}
                        });
                    }
                    
                    // 初始隐藏所有匹配的链接
                    function initHideLinks() {
                        document.querySelectorAll(linkSelectors).forEach(hideTargetLink);
                        hideExternalLinks();
                    }

                    // 监听动态加载的链接（网页异步加载的元素也能处理）
                    const linkObserver = new MutationObserver(mutations => {
                        for (let mutation of mutations) {
                            for (let node of mutation.addedNodes) {
                                if (node.nodeType === 1) {
                                    const target = node.closest(linkSelectors);
                                    if (target) hideTargetLink(target);
                                    hideExternalLinks();  
                                }
                            }
                        }
                    });

                    // 启动监听
                    linkObserver.observe(document.body, { childList: true, subtree: true });
                    initHideLinks();

                })();


                // ==================== div标签鼠标交互 ====================

                document.body.addEventListener('mouseover', function(event) {
                    if (event.target.tagName.toLowerCase() === 'div') {
                        event.target.style.boxShadow = 'inset 0 0 0 2px rgba(80, 255, 80, 0.8)';
                    }
                });

                document.body.addEventListener('mouseout', function(event) {
                    if (event.target.tagName.toLowerCase() === 'div') {
                        event.target.style.boxShadow = '';
                    }
                });

                document.body.addEventListener('dblclick', function(event) {
                    if (event.target.tagName.toLowerCase() === 'div') {
                        var target = event.target;
                        var originalBoxShadow = target.style.boxShadow;
                        var textToCopy = target.innerText
                            .replace(/^\s*[\r\n]/gm, '')
                            .replace(/^\s+|\s+$/gm, '');
                        navigator.clipboard.writeText(textToCopy).then(
                            function() {
                                target.style.boxShadow = 'inset 0 0 0 2px rgba(255, 80, 80, 0.6)';
                                setTimeout(function() {
                                    target.style.boxShadow = originalBoxShadow;
                                }, 500);
                            },
                        );
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
                userRecords.OpenedWebsite = url; // 将打开的网址存入用户使用记录中
            }
        }
    }
}
