using System.Windows;
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


        public SubConverterWindow()
        {
            InitializeComponent();

            List<string> lstConversionTypesKeys = dicConversionTypes.Keys.ToList(); //将转换类型字典的键名转换为List

            cmbbxConversionType.ItemsSource = lstConversionTypesKeys; // 将转换类型字典的键名列表赋值给转换类型组合框
            cmbbxConversionType.SelectedIndex = 0;

            txtbxConverterBackEndUrl.Text = latestRecords.LatestSubConverterBackEndUrl; // 将用户使用记录中的订阅转换器后端URL赋值给订阅转换器后端URL文本框
            txtbxOriginalSubUrls.Text = latestRecords.LatestOriginalSubUrls; // 将用户使用记录中的订阅URL赋值给源订阅URL文本框
            txtbxExternalConfigUrl.Text = latestRecords.LatestExternalConfigUrl; // 将用户使用记录中的外部配置URL赋值给外部配置URL文本框

        }


        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string subConverterBackEndUrl = txtbxConverterBackEndUrl.Text.Trim(); // 获取订阅转换器后端URL
                string originalSubUrls = txtbxOriginalSubUrls.Text.Trim(); // 获取源订阅Url
                string externalConfigUrl = txtbxExternalConfigUrl.Text.Trim(); // 获取外部配置Url

                if (cmbbxConversionType.SelectedItem == null || string.IsNullOrWhiteSpace(subConverterBackEndUrl) || string.IsNullOrWhiteSpace(originalSubUrls)) // 如果订阅转换器后端Url、源订阅Url或转换类型组合框已选项有一个为null，则抛出异常
                {
                    throw new Exception("Invalid conversion type or url(s).");
                }

                latestRecords.LatestSubConverterBackEndUrl = subConverterBackEndUrl; // 将用户输入的订阅转换器后端URL赋值给用户使用记录
                latestRecords.LatestOriginalSubUrls = originalSubUrls; // 将用户输入的订阅URL赋值给用户使用记录
                latestRecords.LatestExternalConfigUrl = externalConfigUrl; // 将用户输入的外部配置URL赋值给用户使用记录

                string targetType = dicConversionTypes[cmbbxConversionType.SelectedItem.ToString()!]; // 从转换类型字典中获取对应的转换类型代码
                string encodedOriginalSubUrls = Uri.EscapeDataString(originalSubUrls); // 获取经Url编码后的源订阅Url
                string encodedExternalConfigUrl = Uri.EscapeDataString(externalConfigUrl); // 获取经Url编码后的外部配置Url

                string convertedSubUrl = $"{subConverterBackEndUrl}sub?target={targetType}&url={encodedOriginalSubUrls}"; // 拼接生成转换后的订阅链接
                if (!string.IsNullOrWhiteSpace(encodedExternalConfigUrl)) // 如果经Url编码后的外部配置Url不为null或全空白字符，则将该段Url拼接到转换后的订阅链接最后
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
            ShowMessage("Converted subscription url copied.");
        }

    }
}
