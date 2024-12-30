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
using System.Windows.Shapes;

namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for SubConverter.xaml
    /// </summary>
    public partial class SubConverterWindow : Window
    {
        public SubConverterWindow()
        {
            InitializeComponent();
            cmbbxConversionType.ItemsSource = new List<string> { "Clash", "ClashR", "Loon", "SS", "SSR", "Surfboard", "Surge&ver=2", "Surge&ver=3", "Surge&ver=4", "Trojan", "V2Ray", "Mixed", "Auto" };
            cmbbxConversionType.SelectedIndex = 0;
        }

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string originalUrl = txtbxOriginalSubscription.Text.Trim();
                if (string.IsNullOrEmpty(originalUrl) || cmbbxConversionType.SelectedItem == null)
                {
                    throw new Exception("Invalid URL or conversion type.");
                }

                string encodedUrl = Uri.EscapeDataString(originalUrl); // 编码源Url
                string targetType = cmbbxConversionType.SelectedItem.ToString()!.ToLower();
                string convertedUrl = $"http://127.0.0.1:25500/sub?target={targetType}&url={encodedUrl}"; // 拼接转换后的链接

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
            MessageBox.Show("Converted subscription copied to the clipboard.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
