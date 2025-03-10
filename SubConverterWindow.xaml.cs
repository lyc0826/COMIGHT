﻿using System.Windows;
using System.Windows.Input;
using static COMIGHT.MainWindow;
using static COMIGHT.Methods;

namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for SubConverter.xaml
    /// </summary>
    public partial class SubConverterWindow : Window
    {

        Dictionary<string, string> dicConversionTypes = new Dictionary<string, string>() //定义转换类型字典，键名为类型，键值为代号
            {
                { "Clash", "clash" },
                { "ClashR", "clashr" },
                { "Loon", "loon" },
                { "SS", "ss" },
                { "SSR", "ssr" },
                { "Surfboard", "surfboard" },
                { "Surge 2", "surge&ver=2" },
                { "Surge 3", "surge&ver=3" },
                { "Surge 4", "surge&ver=4" },
                { "Trojan", "trojan" },
                { "V2Ray", "v2ray" },
                { "Mixed", "mixed" },
                { "Auto", "auto" }
            };

        private ExternalAppManager externalAppManager;

        public SubConverterWindow()
        {
            InitializeComponent();

            List<string> lstConversionTypesKeys = dicConversionTypes.Keys.ToList(); //将转换类型字典的键名转换为List

            cmbbxConversionType.ItemsSource = lstConversionTypesKeys; // 将转换类型字典的键名列表赋值给转换类型组合框
            cmbbxConversionType.SelectedIndex = 0;

            txtbxConverterBackEndURL.Text = latestRecords.LatestSubConverterBackEndUrl; // 将用户使用记录中的订阅转换器后端URL赋值给订阅转换器后端URL文本框
            txtbxOriginalSubUrls.Text = latestRecords.LatestOriginalSubUrls; // 将用户使用记录中的订阅URL赋值给源订阅URL文本框
            txtbxExternalConfigUrl.Text = latestRecords.LatestExternalConfigUrl; // 将用户使用记录中的外部配置URL赋值给外部配置URL文本框

            string appPath = appSettings.SubConverterPath; // 获取订阅转换器程序路径
            externalAppManager = new ExternalAppManager(appPath); // 创建外部应用程序管理器对象，并赋值给外部应用程序管理器对象变量（打开订阅转换器）
            externalAppManager.StartMonitoring(); // 启动应用程序并监控运行状态

            lblStatus.DataContext = externalAppManager; // 将状态标签控件的数据环境设为外部应用程序管理器对象
        }

        //protected override void OnClosed(EventArgs e) // 重写 OnClosed 方法，该方法在窗口关闭时调用
        //{
        //    base.OnClosed(e); // 调用基类的 OnClosed 方法
        //    externalAppManager.StopMonitoring(); // 调用 _appMonitor 的 StopMonitoring 方法，停止监控任务
        //}

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string subConverterBackEndUrl = txtbxConverterBackEndURL.Text.Trim(); // 获取订阅转换器后端URL
                string originalSubUrls = txtbxOriginalSubUrls.Text.Trim(); // 获取源订阅Url
                string externalConfigUrl = txtbxExternalConfigUrl.Text.Trim(); // 获取外部配置Url

                if (cmbbxConversionType.SelectedItem == null || string.IsNullOrWhiteSpace(subConverterBackEndUrl) || string.IsNullOrWhiteSpace(originalSubUrls)) // 如果订阅转换器后端Url、源订阅Url或转换类型组合框已选项有一个为null，则抛出异常
                {
                    throw new Exception("Invalid url or conversion type.");
                }

                latestRecords.LatestSubConverterBackEndUrl = subConverterBackEndUrl; // 将用户输入的订阅转换器后端URL赋值给用户使用记录
                latestRecords.LatestOriginalSubUrls = originalSubUrls; // 将用户输入的订阅URL赋值给用户使用记录
                latestRecords.LatestExternalConfigUrl = externalConfigUrl; // 将用户输入的外部配置URL赋值给用户使用记录

                string targetType = dicConversionTypes[cmbbxConversionType.SelectedItem.ToString()!]; // 从转换类型字典中获取对应的转换类型代码
                string encodedOriginalSubUrls = Uri.EscapeDataString(originalSubUrls); // 获取经Url编码后的源订阅Url
                string encodedExternalConfigUrl = Uri.EscapeDataString(externalConfigUrl); // 获取经Url编码后的外部配置Url

                string convertedSubUrl = $"{subConverterBackEndUrl}sub?target={targetType}&url={encodedOriginalSubUrls}"; // 拼接生成转换后的订阅链接
                if (!string.IsNullOrWhiteSpace(encodedExternalConfigUrl)) // 如果经Url编码后的外部配置Url不为null或全空白字符，则将该段Url拼接到订阅链接最后
                {
                    convertedSubUrl += $"&config={encodedExternalConfigUrl}";
                }

                txtbxConvertedSubUrl.Text = convertedSubUrl; // 将转换后的链接赋值给转换后链接文本框
                txtbxConvertedSubUrl.SelectAll(); //全选转换后链接文本框文字
                txtbxConvertedSubUrl.Focus(); //转换后链接文本框获取焦点

            }
            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void TxtbxConvertedSubUrl_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Clipboard.SetText(txtbxConvertedSubUrl.Text); // 将转换后链接文本框的文字复制到剪贴板
            ShowMessage("Converted subscription copied to the clipboard.");
        }

    }
}
