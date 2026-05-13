using DocSharp.Markdown;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.Data;
using System.IO;
using System.Windows;
using static COMIGHT.Methods;
using static COMIGHT.Settings;
using Window = System.Windows.Window;



namespace COMIGHT
{
    /// <summary>
    /// MarkupTextConverterWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MarkupTextConverterWindow : Window
    {

        public MarkupTextConverterWindow(string defaultText = "")
        {
            InitializeComponent();
            txtbxMarkup.Text = defaultText;
            
        }

        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            ConvertMarkupIntoWordDocument();
        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsModal())
            {
                this.DialogResult = false; //对话框返回值设为false
            }
            this.Close();
        }

        private void ConvertMarkupIntoWordDocument()
        {
            try
            {

                if (string.IsNullOrWhiteSpace(txtbxMarkup.Text)) //如果Markdown文本框为空白，则结束本过程
                {
                    throw new Exception("No text found.");
                }

                string markupText = txtbxMarkup.Text; //获取Markdown文本框的的文本，赋值给Markdown文本变量
                markupText = appSettings.KeepEmojisInMarkupText ? markupText : markupText.RemoveEmojis(); //获取Markdown文本变量：如果程序设置允许Office文件中存在Emoji字符，则得到Markdown文本变量原值；否则，得到删除Markdown文本中Emoji后的值

                //将导出文本框的文字按换行符拆分为数组（删除每个元素前后空白字符，并删除空白元素），转换成列表
                List<string> lstParagraphs = markupText.RemoveMarkdownMarks().RemoveHtmlTags()
                    .Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();

                if (lstParagraphs.Count == 0) //如果段落列表元素数为0，则抛出异常
                {
                    throw new Exception("No text found.");
                }

                string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径
                // 获取目标文件主名：将段落列表所有元素的Markdown标记和不能作为文件名的字符删除后，将不为null或空的字符串的元素的第一个，作为目标文件主名
                string targetFileMainName = lstParagraphs.ConvertAll(e => CleanPathAndFileName(e)).Where(e => !string.IsNullOrWhiteSpace(e)).First();

                //将目标Markdown文档转换为目标Word文档
                string targetWordFilePath = Path.Combine(targetFolderPath, $"{targetFileMainName}.docx"); //获取目标Word文档文件路径

                if (!Enum.TryParse((string?)cmbbxMarkupType.SelectedItem, out EnumMarkupType enumMarkupType)) //获取Markdown类型枚举变量
                {
                    throw new Exception("No valid Markup Type selected.");
                }

                switch (enumMarkupType)
                {
                    case EnumMarkupType.Markdown:

                        MarkdownSource markdown = MarkdownSource.FromMarkdownString(markupText); // 创建Markdown源对象
                        MarkdownConverter markdownConverter = new MarkdownConverter() //  创建Markdown转换器对象
                        {
                            //ImagesBaseUri = Path.GetDirectoryName(targetMDFilePath)  // 设置图片的路径
                        };
                        markdownConverter.ToDocx(markdown, targetWordFilePath, append: false); // 将Markdown文档转换成Word文档
                        
                        break;
                    
                    case EnumMarkupType.HTML:
                        //无需转换，因为已经是HTML了
                        using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(targetWordFilePath, WordprocessingDocumentType.Document))
                        {
                            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart(); //创建主文档部分
                            mainPart.Document = new Document(new Body()); //创建 Word 文档基本结构（Document中包括Body）
                            HtmlConverter htmlConverter = new HtmlConverter(mainPart); //创建 HTML 转换器

                            // 把 HTML 中的段落、标题、表格、列表、图片等尽量转换为 Word 可识别的 OpenXML 元素。
                            var openXmlElements = htmlConverter.Parse(markupText);
                            //将转换后的 OpenXML 元素加入 Word 文档 Body
                            Body? body = mainPart.Document.Body;
                            if (body != null)
                            {
                                foreach (var element in openXmlElements)
                                {
                                    body.Append(element);
                                }
                            }

                            mainPart.Document.Save(); // 保存 Word 文档
                        }

                        break;

                    default:
                        throw new Exception("No valid Markup Type selected.");
                }

                // 提取目标Word文档中的表格并转存为目标Excel文档
                string targetExcelFilePath = Path.Combine(targetFolderPath, $"{CleanPathAndFileName($"Tbl_{targetFileMainName}")}.xlsx"); //获取目标Excel文件路径

                bool tableExtracted = ExtractTablesFromWordToExcel(targetWordFilePath, targetExcelFilePath); // 提取目标Word文档中的表格并转存为目标Excel文档，如果成功则将true赋值给“表格已提取”变量

                // 获取结果消息
                string resultMessage = $"File saved as:\n'{targetWordFilePath}'{(tableExtracted ? $"\n'{targetExcelFilePath}'" : "")}";

                ShowSuccessMessage(resultMessage);
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        private void MarkupTextConverterWindow_Loaded(object sender, RoutedEventArgs e)
        {
            
            // 设置下拉列表框的项源为枚举类型EnumMarkupType的所有值
            List<string> lstMarkupTypes = Enum.GetValues(typeof(EnumMarkupType))
                                 .Cast<EnumMarkupType>()
                                 .Select(type => type.ToString())
                                 .ToList();
            var listItemsSource = (ListItemsSource)this.Resources["ListItemsSource"]; // 将窗体资源中的ListItemsSource对象赋值给listItemsSource对象
            listItemsSource.MarkupTypeList =lstMarkupTypes; // 将标记文本类型列表赋值给listItemsSource对象中的标记文本类型属性

            this.DataContext = userRecords; // 将应用设置窗口的数据环境设为用户使用记录对象

        }
    }
}