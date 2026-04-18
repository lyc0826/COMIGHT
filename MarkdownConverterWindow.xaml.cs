using DocSharp.Markdown;
using System.Data;
using System.IO;
using System.Windows;
using static COMIGHT.Methods;
using static COMIGHT.Settings;
using Window = System.Windows.Window;


namespace COMIGHT
{
    /// <summary>
    /// MarkdownConverterWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MarkdownConverterWindow : Window
    {

        public MarkdownConverterWindow(string defaultText = "")
        {
            InitializeComponent();
            txtbxMarkdown.Text = defaultText;

        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            ConvertMarkdownIntoWordDocument();

            if (this.IsModal())
            {
                this.DialogResult = true; //对话框返回值设为true
            }
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsModal())
            {
                this.DialogResult = false; //对话框返回值设为false
            }
            this.Close();
        }

        private void ConvertMarkdownIntoWordDocument()
        {
            try
            {

                if (string.IsNullOrWhiteSpace(txtbxMarkdown.Text)) //如果Markdown文本框为空白，则结束本过程
                {
                    throw new Exception("No text found.");
                }

                string mdText = txtbxMarkdown.Text; //获取Markdown文本框的的文本，赋值给Markdown文本变量
                mdText = appSettings.KeepEmojisInMarkdown ? mdText : mdText.RemoveEmojis(); //获取Markdown文本变量：如果程序设置允许Office文件中存在Emoji字符，则得到Markdown文本变量原值；否则，得到删除Markdown文本中Emoji后的值

                //将导出文本框的文字按换行符拆分为数组（删除每个元素前后空白字符，并删除空白元素），转换成列表
                List<string> lstParagraphs = mdText
                    .Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();

                if (lstParagraphs.Count == 0) //如果段落列表元素数为0，则抛出异常
                {
                    throw new Exception("No text found.");
                }

                string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径
                // 获取目标文件主名：将段落列表所有元素的Markdown标记和不能作为文件名的字符删除后，将不为null或空的字符串的元素的第一个，作为目标文件主名
                string targetFileMainName = lstParagraphs.ConvertAll(e => CleanPathName(e.RemoveMarkdownMarks())).Where(e => !string.IsNullOrWhiteSpace(e)).First();

                //将目标Markdown文档转换为目标Word文档
                string targetWordFilePath = Path.Combine(targetFolderPath, $"{targetFileMainName}.docx"); //获取目标Word文档文件路径

                MarkdownSource markdown = MarkdownSource.FromMarkdownString(mdText); // 创建Markdown源对象
                MarkdownConverter converter = new MarkdownConverter() //  创建Markdown转换器对象
                {
                    //ImagesBaseUri = Path.GetDirectoryName(targetMDFilePath)  // 设置图片的路径
                };
                converter.ToDocx(markdown, targetWordFilePath, append: false); // 将Markdown文档转换成Word文档

                // 提取目标Word文档中的表格并转存为目标Excel文档
                string targetExcelFilePath = Path.Combine(targetFolderPath, $"{CleanPathName($"Tbl_{targetFileMainName}")}.xlsx"); //获取目标Excel文件路径

                bool tableExtracted = ExtractTablesFromWordToExcel(targetWordFilePath, targetExcelFilePath); // 提取目标Word文档中的表格并转存为目标Excel文档，如果成功则将true赋值给“表格已提取”变量

                // 获取结果消息
                string resultMessage = $"File saved as '{targetWordFilePath}'{(tableExtracted ? $" and '{targetExcelFilePath}'" : "")}.";

                ShowSuccessMessage(resultMessage);
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }


    }
}