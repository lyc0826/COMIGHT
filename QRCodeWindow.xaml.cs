using iText.Signatures.Validation.Lotl;
using Microsoft.Win32;
using QRCoder; 
using System;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using static COMIGHT.MainWindow;
using static COMIGHT.Methods; 

namespace COMIGHT
{
    public partial class QRCodeWindow : Window
    {
        private byte[]? qrCodeBytes = null;

        public QRCodeWindow()
        {
            InitializeComponent();

            // 窗体加载时，如果文本框有默认值，也可以触发一次生成
            // GenerateQRCode(InputTextBox.Text); 
        }

        // 文本框内容改变时触发
        private void InputTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            GenerateQRCode(InputTextBox.Text); // 调用核心方法，生成二维码并显示
        }

        // 点击 OK 按钮触发
        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// 核心方法：生成二维码并显示
        /// <param name="content">要编码的文本内容</param>
        private void GenerateQRCode(string content)
        {
            try
            {
                // 如果内容为空，清空图片并返回
                if (string.IsNullOrEmpty(content))
                {
                    qrCodeBytes = null;
                    QRImage.Source = null;
                    return;
                }
            
                // 创建二维码生成器实例
                QRCodeGenerator qrGenerator = new QRCodeGenerator();

                // 创建二维码数据
                // ECCLevel 是纠错等级 (L, M, Q, H)，Q 级约为 25% 的纠错能力
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.Q);

                // 使用 PngByteQRCode 渲染器 (生成字节数组，适合 WPF 使用，无需依赖 System.Drawing)
                PngByteQRCode qrCode = new PngByteQRCode(qrCodeData);

                // GetGraphic(20) 中的 20 是指每个模块(像素点)的大小
                qrCodeBytes = qrCode.GetGraphic(20);

                // 将字节数组转换为 WPF 可识别的 ImageSource
                QRImage.Source = ByteToImage(qrCodeBytes);
            }
            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        /// <summary>
        /// 辅助方法：将字节数组转换为 BitmapImage
        /// </summary>
        private BitmapImage? ByteToImage(byte[] blob)
        {
            if (blob == null || blob.Length == 0) // 如果图块字节数组为空，则返回 null
            {
                return null;
            }

            BitmapImage bitmap = new BitmapImage(); // 创建 BitmapImage 对象
            using (MemoryStream ms = new MemoryStream(blob)) // 创建内存流
            {
                bitmap.BeginInit(); // 开始初始化
                bitmap.CacheOption = BitmapCacheOption.OnLoad; // 缓存选项, 表示图片加载完成后缓存图片（必须设置，否则流关闭后图片会消失）
                bitmap.StreamSource = ms; // 设置流源
                bitmap.EndInit(); // 结束初始化
                bitmap.Freeze(); // 冻结对象，使其可以跨线程访问（虽然这里是在UI线程）
            }
            return bitmap;
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            SavePic();
        }

        private void SavePic()
        {
            try
            {
                string targetFolderPath = appSettings.SavingFolderPath; //获取目标文件夹路径
                string targetPNGFile = Path.Combine(targetFolderPath!, $"{CleanFileAndFolderName(InputTextBox.Text)}.png"); //获取目标图片文件路径全名

                // 将图片保存
                // 直接保存字节数组（最简单）

                if (qrCodeBytes == null)
                {
                    throw new Exception("No QR Code to Save.");
                }

                File.WriteAllBytes(targetPNGFile, qrCodeBytes);

                ShowSuccessMessage();

                //SaveFileDialog saveDialog = new SaveFileDialog
                //{
                //    Filter = "PNG 图片|*.png",
                //    FileName = "QRCode.png"
                //};
                //if (saveDialog.ShowDialog() == true)
                //{
                //    File.WriteAllBytes(saveDialog.FileName, _currentQRCodeBytes);
                //    MessageBox.Show("保存成功！");
                //}

            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

    }
}
