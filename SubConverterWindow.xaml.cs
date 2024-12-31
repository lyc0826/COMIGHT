using System.Windows;
using System.Windows.Input;

namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for SubConverter.xaml
    /// </summary>
    public partial class SubConverterWindow : Window
    {

        Dictionary<string, string> dicConversionTypes = new Dictionary<string, string>() //定义转换类型字典，键名为类型，键值为代码
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
        }

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string originalUrl = txtbxOriginalSubscription.Text.Trim(); // 获取源Url
                if (string.IsNullOrWhiteSpace(originalUrl) || cmbbxConversionType.SelectedItem == null) // 如果源Url为null或转换类型组合框已选项为null，则抛出异常
                {
                    throw new Exception("Invalid URL or conversion type.");
                }

                string encodedUrl = Uri.EscapeDataString(originalUrl); // 编码源Url
                string targetType = dicConversionTypes[cmbbxConversionType.SelectedItem.ToString()!]; // 从转换类型字典中获取对应的转换类型代码
                string convertedUrl = $"http://127.0.0.1:25500/sub?target={targetType}&url={encodedUrl}"; // 拼接生成转换后的链接

                txtbxConvertedSubscription.Text = convertedUrl; // 将转换后的链接赋值给转换后链接文本框
                txtbxConvertedSubscription.SelectAll(); //全选转换后链接文本框文字
                txtbxConvertedSubscription.Focus(); //转换后链接文本框获取焦点

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void TxtbxConvertedSubscription_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Clipboard.SetText(txtbxConvertedSubscription.Text); // 复制链接到剪贴板
            MessageBox.Show("Converted subscription copied to the clipboard.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
