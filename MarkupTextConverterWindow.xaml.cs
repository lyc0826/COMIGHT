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

        enum EnumMarkupType { Markdown, HTML };

        public MarkupTextConverterWindow(string defaultText = "")
        {
            InitializeComponent();
            txtbxMarkup.Text = defaultText;
            // 设置下拉列表框的项源为枚举类型EnumMarkupType的所有值，并将该项源转换为字符串列表
            //cmbbxMarkupType.ItemsSource = new List<string>{ EnumMarkupType.Markdown.ToString(), EnumMarkupType.HTML.ToString() };
            cmbbxMarkupType.ItemsSource = Enum.GetValues(typeof(EnumMarkupType))
                                 .Cast<EnumMarkupType>()
                                 .Select(type => type.ToString())
                                 .ToList();
            cmbbxMarkupType.SelectedIndex = 0; // 设置下拉列表框的选中项为第一个
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

        private void ConvertMarkdownIntoWordDocument()
        {
            try
            {

                if (string.IsNullOrWhiteSpace(txtbxMarkup.Text)) //如果Markdown文本框为空白，则结束本过程
                {
                    throw new Exception("No text found.");
                }

                string mdText = txtbxMarkup.Text; //获取Markdown文本框的的文本，赋值给Markdown文本变量
                mdText = appSettings.KeepEmojisInMarkup ? mdText : mdText.RemoveEmojis(); //获取Markdown文本变量：如果程序设置允许Office文件中存在Emoji字符，则得到Markdown文本变量原值；否则，得到删除Markdown文本中Emoji后的值

                //将导出文本框的文字按换行符拆分为数组（删除每个元素前后空白字符，并删除空白元素），转换成列表
                List<string> lstParagraphs = mdText
                    .Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();

                if (lstParagraphs.Count == 0) //如果段落列表元素数为0，则抛出异常
                {
                    throw new Exception("No text found.");
                }

                string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径
                // 获取目标文件主名：将段落列表所有元素的Markdown标记和不能作为文件名的字符删除后，将不为null或空的字符串的元素的第一个，作为目标文件主名
                string targetFileMainName = lstParagraphs.ConvertAll(e => CleanPathAndFileName(e.RemoveMarkdownMarks())).Where(e => !string.IsNullOrWhiteSpace(e)).First();

                //将目标Markdown文档转换为目标Word文档
                string targetWordFilePath = Path.Combine(targetFolderPath, $"{targetFileMainName}.docx"); //获取目标Word文档文件路径

                MarkdownSource markdown = MarkdownSource.FromMarkdownString(mdText); // 创建Markdown源对象
                MarkdownConverter converter = new MarkdownConverter() //  创建Markdown转换器对象
                {
                    //ImagesBaseUri = Path.GetDirectoryName(targetMDFilePath)  // 设置图片的路径
                };
                converter.ToDocx(markdown, targetWordFilePath, append: false); // 将Markdown文档转换成Word文档

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


        private void ConvertHTMLIntoWordDocument()
        {
            try
            {

                if (string.IsNullOrWhiteSpace(txtbxMarkup.Text)) //如果Markdown文本框为空白，则结束本过程
                {
                    throw new Exception("No text found.");
                }

                string htmlText = txtbxMarkup.Text; //获取Markdown文本框的的文本，赋值给Markdown文本变量
                htmlText = appSettings.KeepEmojisInMarkup ? htmlText : htmlText.RemoveEmojis(); //获取Markdown文本变量：如果程序设置允许Office文件中存在Emoji字符，则得到Markdown文本变量原值；否则，得到删除Markdown文本中Emoji后的值

                //将导出文本框的文字按换行符拆分为数组（删除每个元素前后空白字符，并删除空白元素），转换成列表
                List<string> lstParagraphs = htmlText.RemoveHtmlTags()
                    .Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();

                if (lstParagraphs.Count == 0) //如果段落列表元素数为0，则抛出异常
                {
                    throw new Exception("No text found.");
                }

                string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径
                // 获取目标文件主名：将段落列表所有元素的Markdown标记和不能作为文件名的字符删除后，将不为null或空的字符串的元素的第一个，作为目标文件主名
                string targetFileMainName = lstParagraphs.ConvertAll(e => CleanPathAndFileName(e.RemoveMarkdownMarks())).Where(e => !string.IsNullOrWhiteSpace(e)).First();

                //将目标Markdown文档转换为目标Word文档
                string targetWordFilePath = Path.Combine(targetFolderPath, $"{targetFileMainName}.docx"); //获取目标Word文档文件路径


                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(targetWordFilePath, WordprocessingDocumentType.Document))
                {
                    /*
                     * 9.1 创建主文档部分
                     */
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    /*
                     * 9.2 创建 Word 文档基本结构
                     * 
                     * 一个最基本的 Word 文档结构是：
                     * Document
                     *   Body
                     */
                    mainPart.Document = new Document(new Body());
                    /*
                     * 9.3 创建 HTML 转换器
                     */
                    HtmlConverter converter = new HtmlConverter(mainPart);
                    /*
                     * 9.4 如果 HTML 中包含相对路径图片，可以设置 BaseImageUrl
                     * 
                     * 例如 HTML 中有：
                     * <img src="images/a.png" />
                     * 
                     * 那么可以指定图片相对路径的基准目录。
                     * 如果你的图片和程序、HTML 文件或某个固定目录有关，可以自行修改这里。
                     * 
                     * 注意：
                     * BaseImageUrl 需要是 Uri 类型。
                     */
                    //htmlConverter.BaseImageUrl = new Uri(targetFolderPath + Path.DirectorySeparatorChar);
                    /*
                     * 9.5 将 HTML 字符串转换为 OpenXml 元素集合
                     * 
                     * Parse 方法会把 HTML 中的段落、标题、表格、列表、图片等
                     * 尽量转换为 Word 可识别的 OpenXML 元素。
                     */
                    var openXmlElements = converter.Parse(htmlText);
                    /*
                     * 9.6 将转换后的 OpenXML 元素加入 Word 文档 Body
                     */
                    Body? body = mainPart.Document.Body;
                    if (body != null)
                    {
                        foreach (var element in openXmlElements)
                        {
                            body.Append(element);
                        }
                    }
                    /*
                     * 9.7 保存 Word 文档
                     */
                    mainPart.Document.Save();
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

        private void ConvertMarkupIntoWordDocument()
        {
            try
            {

                if (string.IsNullOrWhiteSpace(txtbxMarkup.Text)) //如果Markdown文本框为空白，则结束本过程
                {
                    throw new Exception("No text found.");
                }

                string markupText = txtbxMarkup.Text; //获取Markdown文本框的的文本，赋值给Markdown文本变量
                markupText = appSettings.KeepEmojisInMarkup ? markupText : markupText.RemoveEmojis(); //获取Markdown文本变量：如果程序设置允许Office文件中存在Emoji字符，则得到Markdown文本变量原值；否则，得到删除Markdown文本中Emoji后的值

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


    }
}