using System;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using QRCoder; // 引用 QRCoder 命名空间

namespace COMIGHT
{
    public partial class QRCodeWindow : Window
    {
        public QRCodeWindow()
        {
            InitializeComponent();

            // 窗体加载时，如果文本框有默认值，也可以触发一次生成
            // GenerateQRCode(InputTextBox.Text); 
        }

        // 文本框内容改变时触发
        private void InputTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            GenerateQRCode(InputTextBox.Text);
        }

        // 点击 OK 按钮触发
        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            // 这里可以根据需求修改，比如返回 DialogResult 或直接关闭
            //this.DialogResult = true;
            this.Close();
        }

        /// <summary>
        /// 核心方法：生成二维码并显示
        /// </summary>
        /// <param name="content">要编码的文本内容</param>
        private void GenerateQRCode(string content)
        {
            // 如果内容为空，清空图片并返回
            if (string.IsNullOrEmpty(content))
            {
                QRImage.Source = null;
                return;
            }

            try
            {
                // 1. 创建二维码生成器实例
                QRCodeGenerator qrGenerator = new QRCodeGenerator();

                // 2. 创建二维码数据
                // ECCLevel 是纠错等级 (L, M, Q, H)，Q 级约为 25% 的纠错能力
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.Q);

                // 3. 使用 PngByteQRCode 渲染器 (生成字节数组，适合 WPF 使用，无需依赖 System.Drawing)
                PngByteQRCode qrCode = new PngByteQRCode(qrCodeData);

                // GetGraphic(20) 中的 20 是指每个模块(像素点)的大小
                byte[] qrCodeBytes = qrCode.GetGraphic(20);

                // 4. 将字节数组转换为 WPF 可识别的 ImageSource
                QRImage.Source = ByteToImage(qrCodeBytes);
            }
            catch (Exception ex)
            {
                // 简单的错误处理，防止输入特殊字符导致崩溃
                System.Diagnostics.Debug.WriteLine($"生成二维码错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 辅助方法：将字节数组转换为 BitmapImage
        /// </summary>
        private BitmapImage ByteToImage(byte[] blob)
        {
            if (blob == null || blob.Length == 0) return null;

            BitmapImage bitmap = new BitmapImage();
            using (MemoryStream ms = new MemoryStream(blob))
            {
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad; // 必须设置，否则流关闭后图片会消失
                bitmap.StreamSource = ms;
                bitmap.EndInit();
                bitmap.Freeze(); // 冻结对象，使其可以跨线程访问（虽然这里是在UI线程）
            }
            return bitmap;
        }
    }
}
