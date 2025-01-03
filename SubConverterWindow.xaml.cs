﻿using System.Windows;
using System.Windows.Input;
using static COMIGHT.MainWindow;

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
            
            txtbxConverterBackEndURL.Text = latestRecords.LatestSubConverterBackEndUrl; // 将用户使用记录中的订阅转换器后端URL赋值给订阅转换器后端URL文本框
            txtbxOriginalSubUrls.Text = latestRecords.LatestOriginalSubUrls; // 将用户使用记录中的订阅URL赋值给源订阅URL文本框
        }

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string subConverterBackEndUrl = txtbxConverterBackEndURL.Text.Trim(); // 获取订阅转换器后端URL
                string originalSubUrls = txtbxOriginalSubUrls.Text.Trim(); // 获取源订阅Url
                if (string.IsNullOrWhiteSpace(subConverterBackEndUrl) || string.IsNullOrWhiteSpace(originalSubUrls) || cmbbxConversionType.SelectedItem == null) // 如果订阅转换器后端Url、源订阅Url或转换类型组合框已选项有一个为null，则抛出异常
                {
                    throw new Exception("Invalid Url or conversion type.");
                }

                latestRecords.LatestSubConverterBackEndUrl = subConverterBackEndUrl; // 将用户输入的订阅转换器后端URL赋值给用户使用记录
                latestRecords.LatestOriginalSubUrls = originalSubUrls; // 将用户输入的订阅URL赋值给用户使用记录
                recordsManager.SaveSettings(latestRecords); // 保存用户使用记录

                string encodedOriginalSubUrls = Uri.EscapeDataString(originalSubUrls); // 获取经Url编码后的源订阅Url
                string targetType = dicConversionTypes[cmbbxConversionType.SelectedItem.ToString()!]; // 从转换类型字典中获取对应的转换类型代码

                string convertedSubUrl = $"{subConverterBackEndUrl}sub?target={targetType}&url={encodedOriginalSubUrls}"; // 拼接生成转换后的订阅链接

                txtbxConvertedSubUrl.Text = convertedSubUrl; // 将转换后的链接赋值给转换后链接文本框
                txtbxConvertedSubUrl.SelectAll(); //全选转换后链接文本框文字
                txtbxConvertedSubUrl.Focus(); //转换后链接文本框获取焦点
               
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

        private void TxtbxConvertedSubUrl_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Clipboard.SetText(txtbxConvertedSubUrl.Text); // 复制链接到剪贴板
            MessageBox.Show("Converted subscription copied to the clipboard.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
