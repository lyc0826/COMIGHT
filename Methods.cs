using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.Style;
using System.Data;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Interop;
using static COMIGHT.Settings;
using Application = System.Windows.Application;
using DataTable = System.Data.DataTable;
using MSWord = Microsoft.Office.Interop.Word;
using MSWordDocument = Microsoft.Office.Interop.Word.Document;
using MSWordParagraph = Microsoft.Office.Interop.Word.Paragraph;
using MSWordSection = Microsoft.Office.Interop.Word.Section;
using MSWordTable = Microsoft.Office.Interop.Word.Table;
using Task = System.Threading.Tasks.Task;
using Window = System.Windows.Window;
using static COMIGHT.Constants;


namespace COMIGHT
{
    public static partial class Methods
    {
        public static async Task BatchFormatWordDocumentsHelperAsync(List<string> filePaths)
        {
            Task task = Task.Run(() => Process()); // 创建一个异步任务，执行过程为process()
            void Process()
            {
                MSWord.Application msWordApp = new MSWord.Application(); //打开Word应用程序并赋值给word应用程序变量
                msWordApp.ScreenUpdating = false; //关闭屏幕更新
                msWordApp.DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone; //关闭警告
                msWordApp.Visible = false; //“程序窗口可见”设为否
                MSWordDocument? msWordDocument = null; //定义Word文档变量

                try
                {

                    // 定义页边距、行距、字体、字号等的值

                    double topMargin = msWordApp.CentimetersToPoints((float)3.7); // 顶端页边距
                    double bottomMargin = msWordApp.CentimetersToPoints((float)3.5); // 底端页边距
                    double leftMargin = msWordApp.CentimetersToPoints((float)2.8); // 左页边距
                    double rightMargin = msWordApp.CentimetersToPoints((float)2.6); // 右页边距
                    float lineSpace = (float)appSettings.CnLineSpace; // 行间距

                    string titleFontName = appSettings.CnTitleFontName; // 大标题字体
                    string bodyFontName = appSettings.CnBodyFontName; // 正文字体

                    string cnHeading0FontName = appSettings.CnHeading0FontName; // 中文0级小标题
                    string cnHeading1FontName = appSettings.CnHeading1FontName; // 中文1级小标题
                    string cnHeading2FontName = appSettings.CnHeading2FontName;  // 中文2级小标题
                    string cnHeading3_4FontName = appSettings.CnHeading3_4FontName;  // 通用小标题
                    string cnItemNumFontName = cnHeading1FontName; // 中文条款项小标题字体

                    string tableTitleFontName = cnHeading1FontName; // 表格标题字体
                    string tableBodyFontName = bodyFontName; // 表格正文字体

                    string footerFontName = "Times New Roman"; // 页脚字体

                    float titleFontSize = (float)appSettings.CnTitleFontSize; // 大标题字号
                    float bodyFontSize = (float)appSettings.CnBodyFontSize; // 正文字号

                    float cnHeading0FontSize = (float)appSettings.CnHeading0FontSize; // 中文0级小标题
                    float cnHeading1FontSize = (float)appSettings.CnHeading1FontSize; // 中文1级小标题
                    float cnHeading2FontSize = (float)appSettings.CnHeading2FontSize; // 中文2级小标题
                    float cnHeading3_4FontSize = (float)appSettings.CnHeading3_4FontSize; // 中文3-4级小标题
                    float cnItemNumFontSize = cnHeading1FontSize; // 中文条款项小标题字号

                    float tableTitleFontSize = cnHeading1FontSize; // 表格标题字号
                    float tableBodyFontSize = bodyFontSize - 2; // 表格正文字号

                    float footerFontSize = 14; // 页脚字号为四号


                    // 定义正则表达式

                    // 定义大标题正则表达式变量，匹配模式为：从开头开始，不含2个及以上连续的换行符回车符（允许不连续的换行符回车符）、不含“附件/录”、Appendix注释、非“。”分页符的字符1-80个，换行符回车符，后方出现：换行符回车符
                    Regex regExTitle = new Regex(@"(?<=^|\n|\r)(?:(?![\n\r]{2,})(?!(?:附[ |\t]*[件录]|appendix)[^。\f\n\r]{0,3}[\n\r])[^。\f]){1,80}[\n\r](?=[\n\r])", RegexOptions.Multiline | RegexOptions.IgnoreCase);

                    // 定义中文发往单位正则表达式变量，匹配模式为：从开头开始，换行符回车符（一个空行），不含“附件/录”注释、不含小标题编号、不含“如下：”、非“。：:；;”分页符换行符回车符的字符1个及以上，“：:”，换行符回车符
                    Regex regExCnAddressee = new Regex(@"(?<=^|\n|\r)[\n\r](?:(?!附[ |\t]*[件录][^。\f\n\r]{0,3}[\n\r])(?![（\(]?[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*[、\.，,）\)])(?!如下[：:])[^。：:；;\f\n\r]){1,}[：:][\n\r]", RegexOptions.Multiline);

                    // 定义中文0级小标题正则表达式变量，匹配模式为：从开头开始，“第”，空格制表符任意多个，阿拉伯数字中文数字1个及以上，空格制表符任意多个，“部分、篇、章、节”，非“。；;”分页符换行符回车符的字符0-40个，换行符回车符
                    Regex regExCnHeading0 = new Regex(@"(?<=^|\n|\r)第[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*(?:部分|篇|章|节)[^。；;\f\n\r]{0,40}[\n\r]", RegexOptions.Multiline);

                    // 定义中文1、2级小标题正则表达式变量，匹配模式为：从开头开始，“（(”至多一个（捕获组），中文数字1个及以上，空格制表符任意多个，“、.，,）)”，非“。；;”分页符换行符回车符的字符1-40个，“。”换行符回车符
                    Regex regExCnHeading1_2 = new Regex(@"(?<=^|\n|\r)(（|\()?[ |\t]*[一二三四五六七八九十〇零]+[ |\t]*[、\.，,）\)][^。；;\f\n\r]{1,40}[。\n\r]", RegexOptions.Multiline);

                    // 定义中文3、4级小标题正则表达式变量，匹配模式为：从开头开始，“（(”至多一个（捕获组），空格制表符任意多个，阿拉伯数字1个及以上，空格制表符任意多个，“、.，,）)”，非“。；;”分页符换行符回车符的字符1-40个，“。”换行符回车符
                    Regex regExCnHeading3_4 = new Regex(@"(?<=^|\n|\r)(（|\()?[ |\t]*\d+[ |\t]*[、\.，,）\)][^。；;\f\n\r]{1,40}[。\n\r]", RegexOptions.Multiline);

                    // 定义中文“X是”编号正则表达式变量，匹配模式为：前方出现换行符回车符“。：:；;，,”，空格制表符任意多个，中文数字1个及以上，空格制表符任意多个，“是”；后方出现非分页符换行符回车符的字符1个及以上
                    Regex regExCnShiNum = new Regex(@"(?<=[\n\r。：:；;，,][ |\t]*)[一二三四五六七八九十〇零]+[ |\t]*是(?=[^\f\n\r]{1,})", RegexOptions.Multiline);

                    // 定义中文“条款项”编号正则表达式变量，匹配模式为：从开头开始，“第”，空格制表符任意多个，阿拉伯数字或中文数字1个及以上，空格制表符任意多个，“条款项”，“：:”空格制表符
                    Regex regExCnItemNum = new Regex(@"(?<=^|\n|\r)第[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*[条款项][：:| |\t]", RegexOptions.Multiline); // 将正则匹配模式设为条款项编号

                    // 定义清单数字编号列表，包含1、2、3、4级编号匹配模式
                    List<string> listNums = new List<string>() { @"[一二三四五六七八九十〇零]+[ |\t]*[、\.，,]", @"[（\(][ |\t]*[一二三四五六七八九十〇零]+[ |\t]*[）\)]", @"\d+[ |\t]*[、\.，,）\)]", @"[（\(][ |\t]*\d+[ |\t]*[、\.，,）\)]" };

                    // 定义中文附件注释正则表达式变量，匹配模式为：从开头开始，“附”，空格制表符任意多个，“件录”，非“。”分页符换行符回车符的字符0-3个，换行符回车符
                    Regex regExCnAppendix = new Regex(@"(?<=^|\n|\r)附[ |\t]*[件录][^。\f\n\r]{0,3}[\n\r]", RegexOptions.Multiline);

                    // 定义括号注释正则表达式变量，匹配模式为：从开头开始，“（(”，非“（）()。”分页符换行符回车符的字符1-40个，“）)”，换行符回车符
                    Regex regExBracket = new Regex(@"(?<=^|\n|\r)[（\(][^（）\(\)。\f\n\r]{1,40}[）\)][\n\r]", RegexOptions.Multiline);

                    // 定义中文落款字符串变量，匹配模式为：签名至少1个，最后为日期
                    Regex regExSignOff = new Regex(@"(?<=^|\n|\r)[\n\r](?:[\u4e00-\u9fa5][^。：:；;，,\f\n\r]{1,}[\n\r])+[12]\d{3}[ |\t]*年[月日期\d：:\.\-/| |\t]{0,10}[\n\r]", RegexOptions.Multiline);


                    // 批量处理Word文档

                    foreach (string filePath in filePaths) //遍历文件路径全名列表所有元素
                    {
                        msWordDocument = msWordApp.Documents.Open(filePath); // 打开word文档并赋值给Word文档变量

                        // 判断是否为空文档
                        if (msWordDocument.Content.Text.Trim().Length <= 1) // 如果将Word换行符全部删除后，剩下的字符数小于等于1，则结束本过程
                        {
                            return;
                        }

                        // 接受并停止修订
                        msWordDocument.AcceptAllRevisions();
                        msWordDocument.TrackRevisions = false;
                        msWordDocument.ShowRevisions = false;

                        string documentText = msWordDocument.Content.Text; // 全文文字变量赋值

                        // 清除原文格式，替换空格、换行符等

                        // 设置查找模式
                        MSWord.Selection selection = msWordApp.Selection; //将选区赋值给选区变量
                        MSWord.Find find = msWordApp.Selection.Find; //将选区查找赋值给查找变量

                        find.ClearFormatting(); // 清除格式
                        find.Wrap = WdFindWrap.wdFindStop; // 到文档结尾后停止查找
                        find.Forward = true; // 正向查找
                        find.MatchByte = false; // 区分全角半角为False
                        find.MatchWildcards = false; // 使用通配符为False

                        // 全文空格替换为半角空格，换行符替换为回车符
                        selection.WholeStory();

                        find.Text = " "; // 查找空格
                        find.Replacement.Text = " "; // 将空格替换为半角空格
                        find.Execute(Replace: WdReplace.wdReplaceAll);

                        //find.Text = "\t"; // 查找制表符
                        //find.Replacement.Text = "    "; // 将制表符替换为4个空格
                        //find.Execute(Replace: WdReplace.wdReplaceAll);

                        find.Text = "\v"; // 查找换行符（垂直制表符），^l"
                        find.Replacement.Text = "\r"; // 将换行符（垂直制表符）替换为回车符
                        find.Execute(Replace: WdReplace.wdReplaceAll);

                        // 清除段首、段尾多余空格和制表符，段落自动编号转文本
                        for (int i = msWordDocument.Paragraphs.Count; i >= 1; i--) // 从末尾往开头遍历所有段落
                        {
                            MSWordParagraph paragraph = msWordDocument.Paragraphs[i];

                            //正则表达式匹配模式设为：前方出现开头标记、换行符回车符，空格或制表符；如果段落文字被匹配成功，则继续循环
                            while (Regex.IsMatch(paragraph.Range.Text, @"(?<=^|\n|\r)[ |\t]"))
                            {
                                paragraph.Range.Characters[1].Delete(); // 删除开头第一个字符
                            }

                            //正则表达式匹配模式设为：空格或制表符，后方出现换行符回车符、结尾标记；如果段落文字被匹配成功，则继续循环
                            while (Regex.IsMatch(paragraph.Range.Text, @"[ |\t](?=\n|\r|$)"))
                            {
                                paragraph.Range.Select();
                                selection.EndKey(WdUnits.wdLine); // 光标移动到段落结尾换行符之前
                                selection.TypeBackspace(); // 删除前一个字符
                            }

                            // 如果当前段落不在表格内，且含有自动编号
                            if (!paragraph.Range.Information[WdInformation.wdWithInTable] && !string.IsNullOrEmpty(paragraph.Range.ListFormat.ListString))
                            {
                                // 自动编号正则表达式匹配模式设为：中文数字、阿拉伯数字；如果自动编号被匹配成功
                                if (Regex.IsMatch(paragraph.Range.ListFormat.ListString, @"[一二三四五六七八九十〇零\d]"))
                                {
                                    paragraph.Range.InsertBefore(paragraph.Range.ListFormat.ListString + " "); // 在段落文字前添加自动编号和一个空格
                                }
                            }
                        }

                        // 清除文首和文末的空白段
                        while (msWordDocument.Paragraphs[1].Range.Text == "\r") // 如果第1段文字为回车符，则继续循环
                        {
                            msWordDocument.Paragraphs[1].Range.Delete(); // 删除第1段
                        }

                        while (msWordDocument.Paragraphs[msWordDocument.Paragraphs.Count].Range.Text == "\r"
                            && msWordDocument.Paragraphs[msWordDocument.Paragraphs.Count - 1].Range.Text == "\r") // 如果最后一段和倒数第二段文字均为回车符，则继续循环
                        {
                            msWordDocument.Paragraphs[msWordDocument.Paragraphs.Count].Range.Delete(); // 删除最后一段
                        }


                        // 对齐缩进
                        selection.WholeStory();
                        selection.ClearFormatting(); // 清除全部格式、样式
                        MSWord.ParagraphFormat paragraphFormat = msWordApp.Selection.ParagraphFormat; //将选区段落格式赋值给段落格式变量
                        paragraphFormat.Reset(); // 段落格式清除
                        paragraphFormat.CharacterUnitFirstLineIndent = 2; // 设置首行缩进：如果为中文文档，则缩进2个字符；否则为0个字符
                        paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0pt
                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify; // 对齐方式设为两端对齐
                        //paragraphFormat.IndentFirstLineCharWidth((short)(isCnDocument ? 3 : 0)); // 设置首行缩进：如果为中文文档，则缩进3个字符；否则为0个字符


                        // 全文格式初始化
                        selection.WholeStory(); // 选择word所有文档

                        MSWord.PageSetup pageSetup = selection.PageSetup; // 将选区页面设置赋值给页面设置变量
                        pageSetup.PageWidth = msWordApp.CentimetersToPoints((float)21); // 页面宽度设为21cm
                        pageSetup.PageHeight = msWordApp.CentimetersToPoints((float)29.7); // 页面高度设为29.7cm
                        pageSetup.TopMargin = (float)topMargin; // 顶端边距设为预设值
                        pageSetup.BottomMargin = (float)bottomMargin; // 底端边距设为预设值
                        pageSetup.LeftMargin = (float)leftMargin; // 左边距设为预设值
                        pageSetup.RightMargin = (float)rightMargin; // 右边距设为预设值

                        selection.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight; // 突出显示文本取消

                        MSWord.Paragraphs paragraphs = selection.Paragraphs; // 将选区段落赋值给段落变量

                        paragraphs.AutoAdjustRightIndent = 0; // 不自动调整右缩进
                        paragraphs.DisableLineHeightGrid = -1; //取消“如果定义了网格，则对齐到网格”
                        paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly; // 行距设置为固定值
                                                                                       // '.LineSpacingRule = wdLineSpace1pt5 '行距固定1.5
                        paragraphs.LineSpacing = lineSpace; // 行距设为预设值
                        paragraphs.SpaceBefore = msWordApp.CentimetersToPoints(0); // 段落前间距设为0
                        paragraphs.SpaceAfter = msWordApp.CentimetersToPoints(0); // 段落后间距设为0

                        // 基础正文字体设置
                        MSWord.Font font = msWordApp.Selection.Font; //将选区字体赋值给字体变量
                        font.Name = bodyFontName; // 正文字体设为预设值
                        font.Size = bodyFontSize; // 正文字号设为预设值
                        font.ColorIndex = WdColorIndex.wdBlack; // 颜色设为黑色
                        font.Bold = 0; // “是否粗体”设为0
                        font.Kerning = 0; // “为字体调整字符间距”值设为0
                        font.DisableCharacterSpaceGrid = true;  //取消“如果定义了文档网格,则对齐到网格”，忽略字体的每行字符数

                        documentText = msWordDocument.Content.Text; // 全文文字变量重赋值（前期对文档进行过处理，内容可能已经改变）


                        // 文档大标题设置
                        selection.HomeKey(WdUnits.wdStory);

                        int referencePageNum = 0; //参考页码赋值为0
                        MatchCollection matchesTitles = regExTitle.Matches(documentText); // 获取全文文字经过大标题正则表达式匹配后的结果

                        foreach (Match matchTitle in matchesTitles) // 遍历所有匹配到的大标题文字
                        {
                            selection.HomeKey(WdUnits.wdStory);
                            find.Text = matchTitle.Value; // 查找大标题
                            find.Execute();
                            int pageNum = selection.Information[WdInformation.wdActiveEndPageNumber]; // 当前页码变量赋值
                            if (!selection.Information[WdInformation.wdWithInTable] && pageNum != referencePageNum) //如果当前大标题不在表格内，且与之前已确定的大标题不在同一页（一页最多一个大标题）
                            {
                                bool formatTitle = false; // “设置大标题格式”变量赋值为False
                                if (pageNum == 1) // 如果大标题候选文字在第一页
                                {
                                    formatTitle = true; // “设置大标题格式”变量赋值为True
                                }
                                else // 否则
                                {
                                    selection.MoveStart(WdUnits.wdLine, -5); // 将搜索到大标题候选文字选区向上扩展5行
                                    if (selection.Text.Contains("\f")) // 如果选区内含有分页符，则候选文字判断为大标题，“设置大标题格式”变量赋值为True
                                    {
                                        formatTitle = true;
                                    }
                                    selection.MoveStart(WdUnits.wdLine, 5); // 选区起点复原
                                }

                                if (formatTitle) // 如果要设置大标题格式
                                {
                                    paragraphFormat.CharacterUnitFirstLineIndent = 0;
                                    paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0
                                    paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 居中对齐
                                    font.Name = titleFontName; // 设置字体为预设值
                                    font.Size = titleFontSize; // 设置字号为预设值
                                    font.Bold = (int)WdConstants.wdToggle; // 字体加粗
                                    selection.EndKey(WdUnits.wdLine); // 光标移到选区的最后一个字（换行符之前）

                                    // 中文发往单位设置

                                    selection.MoveDown(WdUnits.wdLine, 1, WdMovementType.wdMove); // 光标下移到下方一行
                                    selection.Expand(WdUnits.wdLine); // 全选一行
                                    selection.MoveEnd(WdUnits.wdLine, 5); // 选区向下扩大5行

                                    MatchCollection matchesCnAddressees = regExCnAddressee.Matches(selection.Text); // 获取选区文字经过中文发往单位正则表达式匹配的结果
                                    foreach (Match matchCnAddressee in matchesCnAddressees) // 遍历所有匹配到的中文发往单位文字结果
                                    {
                                        find.Text = matchCnAddressee.Value; // 查找发往单位
                                        find.Execute(); // 执行查找

                                        if (!selection.Information[WdInformation.wdWithInTable]) // 如果找到的文字不在表格内
                                        {
                                            paragraphFormat.CharacterUnitFirstLineIndent = 0;
                                            paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 段落首行缩进为0
                                        }
                                        selection.Collapse(WdCollapseDirection.wdCollapseEnd); // 将选区折叠到末尾
                                    }


                                    referencePageNum = selection.Information[WdInformation.wdActiveEndPageNumber]; // 获取大标题所在页码并赋值给参考页码变量，为以后提供参考
                                }
                            }
                        }


                        // 中文0级（部分、篇、章、节）小标题设置
                        selection.HomeKey(WdUnits.wdStory);

                        MatchCollection matchesCnHeading0s = regExCnHeading0.Matches(documentText); // 获取全文文字经过中文0级小标题正则表达式匹配的结果

                        foreach (Match matchCnHeading0 in matchesCnHeading0s)
                        {
                            find.Text = matchCnHeading0.Value;
                            find.Execute();
                            if (paragraphs[1].Range.Sentences.Count == 1) // 如果中文小标题所在段落只有一句
                            {
                                paragraphs[1].OutlineLevel = WdOutlineLevel.wdOutlineLevel1; // 将当前中文小标题的大纲级别设为1级
                            }
                            paragraphFormat.CharacterUnitFirstLineIndent = 0;
                            paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进为0
                            paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 居中对齐
                            font.Name = cnHeading0FontName;
                            font.Size = cnHeading0FontSize;
                            font.Bold = 1;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 中文1、2级小标题设置
                        selection.HomeKey(WdUnits.wdStory);

                        MatchCollection matchesCnHeading1_2s = regExCnHeading1_2.Matches(documentText); // 获取全文文字经过中文1、2级小标题正则表达式匹配的结果

                        foreach (Match matchCnHeading1_2 in matchesCnHeading1_2s)
                        {
                            find.Text = matchCnHeading1_2.Value;
                            find.Execute();
                            if (paragraphs[1].Range.Sentences.Count == 1)
                            {
                                if (!matchCnHeading1_2.Groups[1].Success) //如果正则表达式匹配捕获组失败（编号开头不含“（” ）
                                {
                                    paragraphs[1].OutlineLevel = WdOutlineLevel.wdOutlineLevel1; // 将当前中文小标题的大纲级别设为1级
                                    font.Name = cnHeading1FontName;
                                    font.Size = cnHeading1FontSize;
                                }
                                else // 否则
                                {
                                    paragraphs[1].OutlineLevel = WdOutlineLevel.wdOutlineLevel2; // 将当前中文小标题的大纲级别设为2级
                                    font.Name = cnHeading2FontName;
                                    font.Size = cnHeading2FontSize;
                                }
                            }

                            font.Bold = 1;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 中文3、4级小标题设置
                        selection.HomeKey(WdUnits.wdStory);

                        MatchCollection matchesCnHeading3_4s = regExCnHeading3_4.Matches(documentText); // 获取全文文字经过中文3、4级小标题正则表达式匹配的结果

                        foreach (Match matchCnHeading3_4 in matchesCnHeading3_4s)
                        {
                            find.Text = matchCnHeading3_4.Value;
                            find.Execute();

                            if (paragraphs[1].Range.Sentences.Count == 1)
                            {
                                paragraphs[1].OutlineLevel = !matchCnHeading3_4.Groups[1].Success ? WdOutlineLevel.wdOutlineLevel3 : WdOutlineLevel.wdOutlineLevel4; //设置小标题所在段落的大纲级别：如果正则表达式匹配捕获组失败（ 编号开头不含“（” ），则设为3级；否则，设为4级  
                            }

                            font.Name = cnHeading3_4FontName;
                            font.Size = cnHeading3_4FontSize;
                            font.Bold = 1;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 中文“X是”编号设置
                        selection.HomeKey(WdUnits.wdStory);

                        MatchCollection matchesCnShiNums = regExCnShiNum.Matches(documentText); // 获取全文文字经过“X是”编号正则表达式匹配的结果

                        foreach (Match matchCnShiNum in matchesCnShiNums)
                        {
                            find.Text = matchCnShiNum.Value;
                            find.Execute();
                            font.Bold = 1;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 中文“条款项”编号设置
                        selection.HomeKey(WdUnits.wdStory);

                        MatchCollection matchesCnItemNums = regExCnItemNum.Matches(documentText); // 获取全文文字经过条款项编号正则表达式匹配的结果

                        foreach (Match matchCnItemNum in matchesCnItemNums)
                        {
                            find.Text = matchCnItemNum.Value;
                            find.Execute();
                            font.Name = cnItemNumFontName;
                            font.Size = cnItemNumFontSize;
                            font.Bold = 1;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        //将前期被识别为小标题的数字编号清单恢复为正文文字格式

                        foreach (string listNum in listNums)  //遍历清单数字编号正则表达式列表
                        {
                            selection.HomeKey(WdUnits.wdStory);

                            // 定义数字编号清单块正则表达式变量，匹配模式为：（从开头开始，数字编号，非分页符换行符回车符的字符至少一个，换行符回车符），以上字符串（捕获组）2个及以上
                            Regex regExListBlock = new Regex(@"((?<=^|\n|\r)" + listNum + @"[^\f\n\r]+[\n\r]){2,}", RegexOptions.Multiline);

                            MatchCollection matchesListBlocks = regExListBlock.Matches(documentText); // 获取全文文字经过数字编号清单块正则表达式匹配的结果

                            foreach (Match matchListBlock in matchesListBlocks) // 遍历数字编号清单块正则表达式匹配结果集合
                            {
                                //如果数字编号清单块正则表达式匹配到的字符串长度/捕获组匹配数的商（即每个条目的平均字数）大于等于指定数值（中文文档100），则不视为清单条目，直接跳过当前循环并进入下一个循环
                                if (matchListBlock.Value.Length / (matchListBlock.Groups[1].Captures.Count) >= 100)
                                {
                                    continue;
                                }

                                // 文本片段正则表达式匹配模式设为：含换行符回车符的任意字符的字符1-255个；获取当前数字编号清单块字符串经匹配后的第一个结果（截取前部最多255个字符，避免超出Interop库Find方法的限制）
                                Match matchTextSection = Regex.Match(matchListBlock.Value, @"(?:.|[\n\r]){1,255}", RegexOptions.Multiline);

                                find.Text = matchTextSection.Value;
                                find.Execute();

                                selection.MoveEnd(WdUnits.wdCharacter, matchListBlock.Value.Length - matchTextSection.Value.Length); //将搜索结果选区的末尾向后扩展至数字编号清单块的末尾
                                paragraphs.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText; // 将选区所在段落的大纲级别设为正文级别

                                //将选区设为正文文字格式
                                font.Name = bodyFontName;
                                font.Size = bodyFontSize;
                                font.Bold = 0;
                                selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                            }

                        }

                        // 中文附件注释设置
                        selection.HomeKey(WdUnits.wdStory);

                        MatchCollection matchesCnAppendixes = regExCnAppendix.Matches(documentText); // 获取全文文字经过附件注释正则表达式匹配的结果

                        foreach (Match matchCnAppendix in matchesCnAppendixes)
                        {
                            find.Text = matchCnAppendix.Value;
                            find.Execute();
                            if (selection.Information[WdInformation.wdWithInTable] == false) // 如果查找结果不在表格内
                            {
                                paragraphFormat.CharacterUnitFirstLineIndent = 0;
                                paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0
                                paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; // 左对齐
                            }
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 设置表格格式

                        // 遍历所有表格
                        foreach (MSWordTable table in msWordDocument.Tables)
                        {
                            // 表格上方标题、注释设置
                            table.Cell(1, 1).Select(); // 选择第1行第1列的单元格
                            selection.MoveUp(WdUnits.wdLine, 1, WdMovementType.wdMove); // 光标上移到表格上方一行
                            selection.Expand(WdUnits.wdLine); // 全选表格上方一行
                            selection.MoveStart(WdUnits.wdLine, -5); // 选区向上扩大5行

                            // 定义表格上方标题正则表达式变量
                            //Regex regExTableTitle = new Regex(tableTitleRegEx, RegexOptions.Multiline | RegexOptions.IgnoreCase);

                            MatchCollection matchesTableTitles = regExTableTitle.Matches(selection.Text); // 获取选区文字经过表格上方标题正则表达式匹配的结果

                            if (matchesTableTitles.Count > 0) // 如果匹配到的结果集合元素数大于0
                            {
                                find.Text = matchesTableTitles[0].Value;
                                find.Execute();
                                paragraphFormat.CharacterUnitFirstLineIndent = 0;
                                paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0);
                                paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                font.Name = tableTitleFontName;
                                font.Size = tableTitleFontSize;
                                font.Bold = 1;
                                selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                            }

                            // 表格设置
                            table.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic; // 前景色设为自动
                            table.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic; // 背景色设为自动
                            table.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight; // 高亮色设为无高亮
                            table.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight; // 高亮色设为无高亮

                            // 单元格边距
                            table.TopPadding = msWordApp.PixelsToPoints(0, true); // 上边距为0
                            table.BottomPadding = msWordApp.PixelsToPoints(0, true); // 下边距为0
                            table.LeftPadding = msWordApp.PixelsToPoints(0, true); // 左边距为0
                            table.RightPadding = msWordApp.PixelsToPoints(0, true); // 右边距为0
                            table.Spacing = msWordApp.PixelsToPoints(0, true); // 单元格间距为0
                            table.AllowPageBreaks = true; // 允许表格断页
                            table.AllowAutoFit = true; // 允许自动重调尺寸

                            // 设置边框：内外单线条，0.5磅粗
                            //table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle; // 内部线条样式为单线条
                            //table.Borders.InsideLineWidth = WdLineWidth.wdLineWidth050pt; // 内部线条粗细为0.5磅
                            //table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle; // 外部线条样式为单线条
                            //table.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth050pt; // 外部线条粗细为0.5磅

                            // 设置行格式
                            table.Rows.WrapAroundText = 0; // 取消文字环绕
                            table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter; // 表水平居中
                            table.Rows.AllowBreakAcrossPages = -1; // 允许行断页
                            table.Rows.HeightRule = WdRowHeightRule.wdRowHeightAuto; // 行高设为自动
                            table.Rows.LeftIndent = msWordApp.CentimetersToPoints(0); // 左面缩进量为0

                            // 设置字体、段落格式
                            table.Range.Font.Name = tableBodyFontName; // 字体为仿宋
                            table.Range.Font.Color = WdColor.wdColorAutomatic; // 字体颜色设为自动
                            table.Range.Font.Size = tableBodyFontSize; // 字号为四号
                            table.Range.Font.Kerning = 0; // “为字体调整字符间距”值设为0
                            table.Range.Font.DisableCharacterSpaceGrid = true;

                            table.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                            table.Range.ParagraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0);
                            table.Range.ParagraphFormat.AutoAdjustRightIndent = 0; // 自动调整右缩进为false
                            //table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 单元格内容水平居中
                            table.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle; // 单倍行距

                            table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter; // 单元格内容垂直居中

                            // 自动调整表格
                            table.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthAuto; // 列宽度设为自动
                            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent); // 根据内容调整表格
                            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow); // 根据窗口调整表格
                        }

                        // 括号注释设置
                        selection.HomeKey(WdUnits.wdStory);

                        MatchCollection matchesBrakets = regExBracket.Matches(documentText); // 获取全文文字经过括号注释正则表达式匹配的结果

                        foreach (Match matchBraket in matchesBrakets)
                        {
                            find.Text = matchBraket.Value;
                            find.Execute();
                            if (selection.Information[WdInformation.wdWithInTable] == false) // 如果查找结果不在表格内
                            {
                                paragraphFormat.CharacterUnitFirstLineIndent = 0;
                                paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0
                                paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 居中对齐
                            }
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }


                        // 落款设置
                        selection.HomeKey(WdUnits.wdStory);

                        MatchCollection matchesSignOffs = regExSignOff.Matches(documentText); // 获取全文文字经过签名和日期落款正则表达式匹配的结果

                        foreach (Match matchSignOff in matchesSignOffs)
                        {
                            find.Text = matchSignOff.Value;
                            find.Execute();
                            if (selection.Information[WdInformation.wdWithInTable] == false) // 如果查找结果不在表格内
                            {
                                foreach (MSWordParagraph paragraph in selection.Paragraphs) // 遍历所有落款中的段落
                                {
                                    float rightIndentation = Math.Max(0, 10 - paragraph.Range.Text.Length / 2); // 计算右缩进量，如果小于0，则限定为0
                                    paragraph.Format.CharacterUnitRightIndent = rightIndentation; // 右缩进设为之前计算值
                                    paragraph.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight; // 右对齐
                                }
                            }
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }


                        // 页脚页码设置
                        foreach (MSWordSection section in msWordDocument.Sections) // 遍历所有节
                        {
                            section.PageSetup.DifferentFirstPageHeaderFooter = 0;     // “首页页眉页脚不同”设为否
                            section.PageSetup.OddAndEvenPagesHeaderFooter = 0;        // “奇偶页页眉页脚不同”设为否

                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete(); // 删除页脚中的内容
                            // 设置页码
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleNumberInDash;  // 页码左右带横线； wdPageNumberStyleArabicFullWidth 阿拉伯数字全宽
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;  // 不续前节
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = 1;  // 从1开始编号
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.Add(WdPageNumberAlignment.wdAlignPageNumberOutside, FirstPage: true); // 页码奇数页靠右，偶数页靠左； wdAlignPageNumberInside  奇左偶右 wdAlignPageNumberCenter 页码居中
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name = footerFontName;
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Size = footerFontSize;

                            section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete(); // 删除页眉中的内容
                            section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone; // 段落下边框线设为无
                        }

                        msWordDocument.Save(); // 保存Word文档
                        msWordDocument.Close(); // 关闭Word文档

                    }

                }

                catch (Exception)
                {
                    throw; // 抛出异常，包含异步过程中发生的异常的所有信息，以便让调用者处理
                }

                finally
                {
                    msWordApp.ScreenUpdating = true;
                    if (msWordDocument != null) Marshal.ReleaseComObject(msWordDocument); // 释放Word文档对象
                    KillOfficeApps(new object[] { msWordApp }); // 关闭Word应用程序进程
                }

            }

            await task;

        }

        public static async Task BatchRepairWordDocumentsHelperAsync(List<string> filePaths)
        {
            Task task = Task.Run(() => Process());
            void Process()
            {
                MSWord.Application msWordApp = new MSWord.Application(); //打开Word应用程序并赋值给word应用程序变量
                try
                {
                    foreach (string filePath in filePaths) //遍历文件路径全名列表所有元素
                    {
                        string targetFilePath = Path.Combine(Path.GetDirectoryName(filePath)!, $"{Path.GetFileNameWithoutExtension(filePath)}.docx"); //获取目标Word文件路径全名
                        MSWordDocument msWordDocument = msWordApp!.Documents.Open(filePath); //打开Word文档，赋值给Word文档变量
                        MSWordDocument targetMSWordDocument = msWordApp.Documents.Add(); // 新建Word文档并赋值给目标Word文档变量
                        msWordDocument.Content.Copy(); // 复制当前Word文档全文
                        targetMSWordDocument.Content.PasteSpecial(WdPasteOptions.wdUseDestinationStyles); // 粘贴到目标Word文档，使用目标文档格式
                        msWordDocument.Close(); //关闭当前Word文档
                        //目标Word文件另存为docx格式，使用最新Word版本兼容模式
                        targetMSWordDocument.SaveAs2(FileName: targetFilePath, FileFormat: WdSaveFormat.wdFormatDocumentDefault, CompatibilityMode: WdCompatibilityMode.wdCurrent);
                        targetMSWordDocument.Close(); //关闭目标Word文件
                    }
                }
                catch (Exception)
                {
                    throw;
                }

                finally
                {
                    KillOfficeApps(new object[] { msWordApp });
                }
            }

            await task;

        }

        public static T Clamp<T>(this T value, T min, T max) where T : IComparable<T> //泛型参数T，T必须实现IComparable<T>接口
        {
            //赋值给函数返回值：如果输入值比最小值小，则得到最小值；如果比最大值大，则得到最大值；否则，得到输入值
            return value.CompareTo(min) < 0 ? min : value.CompareTo(max) > 0 ? max : value;
        }

        public static string CleanFileAndFolderName(string inputName, int maxLength = 254)
        {
            string cleanedName = inputName.Trim(); // 去除首尾空白字符

            // 定义文件名和文件夹名中不允许出现的字符，赋值给非法字符变量
            string invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());

            // 遍历文件名和文件夹名中的字符：如果为不允许出现的字符，则得到'_'；否则，得到原字符；将以上字符形成数组，再转换成字符串，赋值给清理后的名称变量
            cleanedName = new string(inputName.Select(c => invalidChars.Contains(c) ? '_' : c).ToArray());

            // 截取到指定长度
            //return cleanedName.Length > maxLength ? cleanedName.Substring(0, maxLength) : cleanedName;
            cleanedName = cleanedName[..Math.Min(maxLength, cleanedName.Length)]; //截取目标字数
            return cleanedName; // 将清理后的文件名和文件夹名赋值给函数返回值
        }

        public static string CleanWorksheetName(string inputName, int maxLength = 30)
        {
            string cleanedName = inputName.Trim(); //去除首尾空白字符

            // 清理工作表名中非中文、非英文、非数字或下划线的字符
            cleanedName = Regex.Replace(cleanedName, @"[^\u4e00-\u9fa5\w| ]+", "");
            cleanedName = cleanedName[..Math.Min(maxLength, cleanedName.Length)]; //截取目标字数
            return cleanedName; // 将清理后的工作表名赋值给函数返回值
        }

        public static int ConvertColumnLettersIntoIndex(string columnLetters)
        {
            //将输入列符转换为大写，从左到右逐字与字符"A"的ASCII编码取差值，并以26进制的方式累加，赋值给函数返回值
            return columnLetters.ToUpper().Aggregate(0, (tempColumnIndex, columnLetter) => tempColumnIndex * 26 + (columnLetter - 'A' + 1));
        }

        public static string ConvertArabicNumberIntoChinese(int numbers)
        {
            // 定义中文位数字字典，包含阿拉伯数字0到9对应的中文数字
            Dictionary<char, string> dicChineseDigits = new Dictionary<char, string>
                { { '0', "零" }, { '1', "一" }, { '2', "二" }, { '3', "三" }, { '4', "四" }, { '5', "五" },
                { '6', "六" }, { '7', "七" }, { '8', "八" }, { '9', "九" } };

            // 定义中文数字单位数组，包含单位（个，十，百，千，万...）
            string[] arrChineseUnits = new string[] { "", "十", "百", "千", "万", "十", "百", "千", "亿" };

            string arabicNumberStr = numbers.ToString(); // 将输入的阿拉伯数字转换为字符串
            int n = arabicNumberStr.Length; // 获取阿拉伯数字字符串的字数

            //从左到右逐字将阿拉伯位数字转换成中文位数字、获取其中文数字单位，并逐步合并，赋值给中文数字字符串变量
            string chineseNumberStr = arabicNumberStr.Select((arabicDigit, i) =>
                {
                    string chineseDigit = dicChineseDigits[arabicDigit]; //获取当前阿拉伯位数字对应的中文位数字
                    string chineseUnit = arrChineseUnits[n - i - 1]; //获取当前阿拉伯位数字对应的中文单位 （假设是个3位数，从左往右数当i到达第2位（1号）的十位数字时，3-1-1=1，在中文数字单位数组中索引号为1的元素为“十”，依此类推）
                    return chineseDigit + chineseUnit; //返回当前中文位数字和中文单位的组合
                }).Aggregate("", (tempChineseNumberStr, addedChineseNumberStr) => tempChineseNumberStr + addedChineseNumberStr);

            //正则表达式匹配模式为：从开头开始，“一”，后方出现“十”；将匹配到的字符串替换为空（删除二位数的“十”前面的“一”）
            chineseNumberStr = Regex.Replace(chineseNumberStr, @"^一(?=十)", "");
            //正则表达式匹配模式为：前方出现“零”，“十百千”；将匹配到的字符串替换为空（删除“零”后面的“十百千”）
            chineseNumberStr = Regex.Replace(chineseNumberStr, @"(?<=零)[十百千]", "");
            //正则表达式匹配模式为：前方出现任意字符，“零”一个及以上，后方出现“零十百千万亿”或结尾标记；将匹配到的字符串替换为空（删除重复的“零”和结尾的“零”）
            chineseNumberStr = Regex.Replace(chineseNumberStr, @"(?<=.)零+(?=零|十|百|千|万|亿|$)", "");

            return chineseNumberStr; // 将中文数字字符串赋值给函数返回值
        }

        public static void ExtractTablesFromWordToExcel(string wordFilePath, string targetExcelFilePath)
        {
            try
            {
                if (new FileInfo(wordFilePath).Length == 0) //如果当前文件大小为0，则直接结束本过程
                {
                    return;
                }

                // 使用 NPOI 处理 
                using (FileStream wordFileStream = File.OpenRead(wordFilePath)) //打开目标Word文档，赋值给Word文档文件流变量
                {
                    using XWPFDocument wordDocument = new XWPFDocument(wordFileStream); //创建Word文档对象，赋值给Word文档变量
                    {
                        if (wordDocument.Tables.Count > 0) // 如果目标Word文档中包含表格
                        {
                            FileInfo outputExcelFile = new FileInfo(targetExcelFilePath);
                            using (var excelPackage = new ExcelPackage()) // 创建一个Excel包对象
                            {
                                var workbook = excelPackage.Workbook; // 获取Excel工作簿对象
                                int wordTableIndex = 0;
                                for (int i = 0; i < wordDocument.BodyElements.Count; i++) // 遍历目标Word文档中的所有元素
                                {
                                    var wordElement = wordDocument.BodyElements[i]; // 获取目标Word文档中当前元素，并赋值给Word元素变量
                                    if (wordElement is XWPFTable wordTable) // 如果当前Word元素是表格类型，则将其赋值给新变量 wordTable，然后：
                                    {
                                        string tableTitle = "Sheet" + (wordTableIndex + 1); // 定义表格标题，默认为“Sheet”与当前word文档表格索引号加1
                                        // 获取表格标题 (这部分逻辑与Word文档读取相关，保持不变)
                                        if (i > 0) // 如果当前Word元素不是0号元素
                                        {
                                            List<string> lstTableTitle = new List<string>();
                                            for (int k = 1; k <= 5 && i - k >= 0; k++) // 从当前Word元素开始，向前遍历5个元素，直到0号元素为止
                                            {
                                                if (wordDocument.BodyElements[i - k] is XWPFParagraph) // 如果前方当前Word元素是Word段落
                                                {
                                                    XWPFParagraph paragraph = (XWPFParagraph)wordDocument.BodyElements[i - k]; // 获取前方当前Word元素，并赋值给段落变量

                                                    // 表格标题正则表达式模式设为：开头标记，不含“。；;”的字符1-100个，结尾标记；如果段落文字被匹配成功，将被增加到表格标题列表中
                                                    if (Regex.IsMatch(paragraph.Text, @"^[^。；;]{1,100}$", RegexOptions.Multiline))
                                                    {
                                                        lstTableTitle.Add(paragraph.Text);
                                                    }
                                                }
                                            }
                                            // 获取表格标题：如果表格标题列表不为空，则得到其中长度最短的字符串元素；否则，得到表格标题变量原值
                                            tableTitle = lstTableTitle.Count > 0 ? lstTableTitle.OrderBy(s => s.Length).First() : tableTitle;
                                        }

                                        // 创建Excel工作表，使用序号加表格标题作为工作表的名称
                                        ExcelWorksheet worksheet = workbook.Worksheets.Add(CleanWorksheetName($"{wordTableIndex + 1}_{tableTitle}"));
                                        int columnCount = wordTable.Rows.Max(r => r.GetTableCells().Count); //获取Word文档表格所有行里包含单元格数量最多的那一行的单元格数量，即Word文档表格列数，赋值给表格列数变量

                                        worksheet.Cells[1, 1, 1, columnCount].Merge = true; // 合并Excel工作表第一行单元格（EPPlus的行和列索引从1开始）
                                        worksheet.Cells[1, 1].Value = tableTitle; // 将表格标题赋值给Excel工作表1行1列的单元格
                                        int excelRowIndex = 2; // 从Excel工作表2号（第2）行开始写入表格数据
                                        foreach (XWPFTableRow wordTableRow in wordTable.Rows) // 遍历当前Word文档表格中的所有行
                                        {
                                            int excelColumnIndex = 1; // Excel列索引从1开始
                                            foreach (XWPFTableCell wordTableCell in wordTableRow.GetTableCells()) // 遍历当前Word文档表格当前行中的所有单元格
                                            {
                                                worksheet.Cells[excelRowIndex, excelColumnIndex++].Value = wordTableCell.GetText(); // 将当前Word文档表格的当前行当前单元格的文字赋值给当前行当前列的Excel单元格
                                            }
                                            excelRowIndex++; // Excel行索引累加1
                                        }

                                        FormatExcelWorksheet(worksheet, 2, 0); // 格式化表格数据区域（表头为2行）

                                        wordTableIndex++; // Word文档表格索引号累加1
                                    }
                                }

                                excelPackage.SaveAs(outputExcelFile); // 将Excel包存入文件中
                            }
                        }
                    }
                }
            }

            catch (Exception)
            {
                throw;
            }

        }

        public static void FormatExcelWorksheet(ExcelWorksheet excelWorksheet, int headerRowCount = 0, int footerRowCount = 0)
        {
            try
            {
                if (excelWorksheet.Dimension == null) //如果Excel工作表为空，则结束本过程
                {
                    return;
                }

                foreach (ExcelRangeBase cell in excelWorksheet.Cells[excelWorksheet.Dimension.Address]) //遍历所有已使用的单元格
                {
                    //如果当前单元格是合并单元格、值是字符串且不含公式，则将文字中的换行符替换为空格后，重新赋值给单元格（避免自动调整行高时文字显示不全）
                    if (cell.Merge && cell.Value is string && string.IsNullOrWhiteSpace(cell.Formula))
                    {
                        cell.Value = cell.Text.Replace('\n', ' ');
                    }
                }

                // 获取Excel工作表行数和列数
                int rowCount = excelWorksheet.Dimension.End.Row;
                int columnCount = excelWorksheet.Dimension.End.Column;

                //设置表头格式、自动筛选
                if (headerRowCount >= 1) //如果表头行数大于等于1
                {
                    ExcelRange headerRange = excelWorksheet.Cells[1, 1, headerRowCount, columnCount]; //将表头区域赋值给表头区域变量

                    // 设置表头区域字体、对齐
                    headerRange.Style.Font.Name = appSettings.WorksheetFontName; // 获取应用程序设置中的字体名称
                    headerRange.Style.Font.Size = (float)appSettings.WorksheetFontSize; // 获取应用程序设置中的字体大小
                    headerRange.Style.Font.Bold = true; //表头区域字体加粗
                    headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; //单元格内容水平居中对齐
                    headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center; //单元格内容垂直居中对齐
                    headerRange.Style.WrapText = true; //设置文字自动换行

                    if (excelWorksheet.AutoFilter.Address == null) // 如果自动筛选区域为null（未开启自动筛选），则将表头最后一行的自动筛选设为true
                    {
                        excelWorksheet.Cells[headerRowCount, 1, headerRowCount, columnCount].AutoFilter = true;
                    }

                    for (int i = 1; i <= headerRowCount; i++) //遍历表头所有行
                    {
                        ExcelRange headerRowCells = excelWorksheet.Cells[i, 1, i, columnCount]; //将当前行所有单元格赋值给表头行单元格变量

                        int mergedCellCount = headerRowCells.Count(cell => cell.Merge); // 计算当前表头行单元格中被合并的单元格数量
                        //获取“行单元格是否合并”值：如果被合并的单元格数量占当前行所有单元格的75%以上，得到true；否则得到false
                        bool isRowMerged = mergedCellCount >= headerRowCells.Count() * 0.75 ? true : false;
                        //获取边框样式：如果行单元格被合并，则得到无边框样式；否则得到细线边框样式
                        ExcelBorderStyle borderStyle = isRowMerged ? ExcelBorderStyle.None : ExcelBorderStyle.Thin;

                        //设置当前行所有单元格的边框
                        headerRowCells.Style.Border.BorderAround(borderStyle); //设置当前单元格最外侧的边框为之前获取的边框样式
                        headerRowCells.Style.Border.Top.Style = borderStyle; //设置当前单元格顶部的边框为之前获取的边框样式
                        headerRowCells.Style.Border.Left.Style = borderStyle;
                        headerRowCells.Style.Border.Right.Style = borderStyle;
                        headerRowCells.Style.Border.Bottom.Style = borderStyle;

                        excelWorksheet.Rows[i].CustomHeight = false; //设置当前行“是否手动调整行高”为false（即为自动）

                    }

                }

                // 将Excel工作表除去表头、表尾的区域赋值给记录区域变量
                ExcelRange recordRange = excelWorksheet.Cells[headerRowCount + 1, 1, rowCount - footerRowCount, columnCount];

                // 设置记录区域字体、对齐
                recordRange.Style.Font.Name = appSettings.WorksheetFontName;
                recordRange.Style.Font.Size = (float)appSettings.WorksheetFontSize;
                recordRange.Style.Font.Bold = false;
                recordRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; //单元格内容水平居中对齐
                recordRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center; //单元格内容垂直居中对齐
                recordRange.Style.WrapText = true; //设置文字自动换行

                // 设置记录区域边框、内部单元格边框为单细线
                recordRange.Style.Border.BorderAround(ExcelBorderStyle.Thin); //设置整个区域最外侧的边框
                recordRange.Style.Border.Top.Style = ExcelBorderStyle.Thin; //设置区域内部所有单元格的边框
                recordRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                recordRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                recordRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                //设置列宽
                double fullWidth = 0; //全表格宽度赋值为0

                int firstRefRowIndex = Math.Max(1, headerRowCount); //获取起始参考行的索引号：表头最末行的索引号，如果小于1，则限定为1
                //获取最末参考行的索引号：除去表尾后余下行的最后一行的索引号，如果小于起始参考行的索引号，则限定为起始参考行的索引号
                int lastRefRowIndex = Math.Max(firstRefRowIndex, rowCount - footerRowCount);

                for (int j = 1; j <= columnCount; j++) //遍历所有列
                {
                    if (!excelWorksheet.Columns[j].Hidden) //如果当前列不为隐藏列
                    {
                        ExcelRange columnCells = excelWorksheet.Cells[firstRefRowIndex, j, lastRefRowIndex, j]; //将当前列起始参考行至最末参考行的单元格赋值给列单元格集合变量
                        //从列单元格集合变量中筛选出文本不为null或全空白字符且不是合并的单元格，赋值给合格单元格集合变量
                        IEnumerable<ExcelRangeBase> qualifiedCells = columnCells.Where(cell => !string.IsNullOrWhiteSpace(cell.Text) && !cell.Merge);
                        //计算当前列所有合格单元格的字符数平均值：如果合格单元格集合不为空，则得到所有单元格字符数的平均值，否则得到0
                        double averageCharacterCount = qualifiedCells.Any() ? qualifiedCells.Average(cell => cell.Text.Length) : 0;
                        excelWorksheet.Columns[j].Style.WrapText = false; //设置当前列文字自动换行为false
                        excelWorksheet.Columns[j].AutoFit(); //设置当前列自动调整列宽（能完整显示文字的最适合列宽）
                        excelWorksheet.Columns[j].Style.WrapText = true; //设置当前列文字自动换行
                        //在当前列最合适列宽、基于单元格字符数平均值计算出的列宽中取较小值（并限制在8-40的范围），赋值给列宽变量
                        double columnWidth = Math.Min(excelWorksheet.Columns[j].Width, averageCharacterCount * 2 + 2).Clamp<double>(6, 36);
                        excelWorksheet.Columns[j].Width = columnWidth; //设置当前列的列宽

                        fullWidth += columnWidth; //将当前列列宽累加至全表格宽度
                    }
                }

                //设置记录区域行高
                for (int i = headerRowCount + 1; i <= rowCount - footerRowCount; i++) //遍历除去表头、表尾的所有行
                {
                    if (!excelWorksheet.Rows[i].Hidden)  // 如果当前行没有被隐藏，设置当前行“是否手动调整行高”为false（即为自动）
                    {
                        excelWorksheet.Rows[i].CustomHeight = false;
                    }
                }

                // 设置表尾格式
                if (footerRowCount >= 1) //如果表尾行数大于等于1
                {
                    ExcelRange footerRange = excelWorksheet.Cells[rowCount - footerRowCount + 1, 1, rowCount, columnCount]; //将表尾区域赋值给表尾区域变量

                    // 设置表尾区域字体、对齐
                    footerRange.Style.Font.Name = appSettings.WorksheetFontName; // 获取应用程序设置中的字体名称
                    footerRange.Style.Font.Size = (float)appSettings.WorksheetFontSize; // 获取应用程序设置中的字体大小

                    footerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; //单元格内容水平左对齐
                    footerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center; //单元格内容垂直居中对齐
                    footerRange.Style.WrapText = true; //设置文字自动换行

                    for (int i = rowCount; i >= rowCount - footerRowCount + 1; i--) //遍历表尾所有行
                    {
                        ExcelRange footerRowCells = excelWorksheet.Cells[i, 1, i, columnCount]; //将当前行所有单元格赋值给表尾行单元格变量

                        int mergedCellCount = footerRowCells.Count(cell => cell.Merge); // 计算当前表尾行单元格中被合并的单元格数量
                        //获取“行单元格是否合并”值：如果被合并的单元格数量占当前行所有单元格的75%以上，得到true；否则得到false
                        bool isRowMerged = mergedCellCount >= footerRowCells.Count() * 0.75 ? true : false;
                        //获取边框样式：如果行单元格被合并，则得到无边框样式；否则得到细线边框样式
                        ExcelBorderStyle borderStyle = isRowMerged ? ExcelBorderStyle.None : ExcelBorderStyle.Thin;

                        //设置当前行所有单元格的边框
                        footerRowCells.Style.Border.BorderAround(borderStyle); //设置当前单元格最外侧的边框为之前获取的边框样式
                        footerRowCells.Style.Border.Top.Style = borderStyle; //设置当前单元格顶部的边框为之前获取的边框样式
                        footerRowCells.Style.Border.Left.Style = borderStyle;
                        footerRowCells.Style.Border.Right.Style = borderStyle;
                        footerRowCells.Style.Border.Bottom.Style = borderStyle;

                        excelWorksheet.Rows[i].CustomHeight = false; //设置当前行“是否手动调整行高”为false（即为自动）

                    }

                }

                //调整纸张、方向、对齐
                ExcelPrinterSettings printerSettings = excelWorksheet.PrinterSettings; //将Excel工作表打印设置赋值给打印设置变量
                printerSettings.PaperSize = ePaperSize.A4; // 纸张设置为A4
                printerSettings.Orientation = fullWidth < 120 ? eOrientation.Portrait : eOrientation.Landscape; //设置纸张方向：如果全表格宽度小于120，为纵向；否则，为横向
                //printerSettings.PrintArea = usedRange; //设置打印区域为已使用范围
                printerSettings.HorizontalCentered = true; //表格水平居中对齐
                printerSettings.VerticalCentered = false; //表格垂直居中对齐为false

                //设置页边距
                printerSettings.LeftMargin = 1.2 / 2.54;
                printerSettings.RightMargin = 1.2 / 2.54;
                printerSettings.TopMargin = 1.2 / 2.54;
                printerSettings.BottomMargin = 1.2 / 2.54;
                printerSettings.HeaderMargin = 0.8 / 2.54;
                printerSettings.FooterMargin = 0.8 / 2.54;

                //设定打印顶端标题行：如果表头行数大于等于1，则设为第1行起到表头最后一行的区域；否则设为空（取消顶端标题行）
                printerSettings.RepeatRows = headerRowCount >= 1 ? new ExcelAddress($"$1:${headerRowCount}") : new ExcelAddress("");
                //设定打印左侧重复列为A列
                //printerSettings.RepeatColumns = new ExcelAddress($"$A:$A");

                // 设置页脚
                string footerText = "P&P / &N"; //设置页码
                excelWorksheet.HeaderFooter.OddFooter.CenteredText = footerText; // 设置奇数页页脚
                excelWorksheet.HeaderFooter.EvenFooter.CenteredText = footerText; // 设置偶数页页脚

                // 设置视图和打印版式
                ExcelWorksheetView view = excelWorksheet.View; //将Excel工作表视图设置赋值给视图设置变量
                view.UnFreezePanes(); //取消冻结窗格
                view.FreezePanes(headerRowCount + 1, 1); // 冻结表头行（参数指定第一个不要冻结的单元格）
                view.PageLayoutView = true; // 将工作表视图设置为页面布局视图
                printerSettings.FitToPage = true; // 启用适应页面的打印设置
                printerSettings.FitToWidth = 1; // 设置缩放为几页宽，1代表即所有列都将打印在一页上
                printerSettings.FitToHeight = 0; // 设置缩放为几页高，0代表打印页数不受限制，可能会跨越多页
                printerSettings.PageOrder = ePageOrder.OverThenDown; // 将打印顺序设为“先行后列”
                view.PageLayoutView = false; // 将页面布局视图设为false（即普通视图）

            }

            catch (Exception)
            {
                throw;
            }
        }

        public static string? GetKeyColumnLetter()
        {
            string latestColumnLetter = userRecords.LatestKeyColumnLetter; //读取设置中保存的主键列符
            InputDialog inputDialog = new InputDialog(question: "Input the key column letter (e.g. \"A\"）", defaultAnswer: latestColumnLetter); //弹出对话框，输入主键列符
            if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则函数返回值赋值为null
            {
                return null;
            }
            string columnLetter = inputDialog.Answer;
            userRecords.LatestKeyColumnLetter = columnLetter; // 将对话框返回的列符存入设置

            return columnLetter; //将列符赋值给函数返回值
        }

        public static List<string>? GetWorksheetOperatingRangeAddresses()
        {
            string latestOperatingRangeAddresses = userRecords.LatestOperatingRangeAddresses; //读取用户使用记录中保存的操作区域
            InputDialog inputDialog = new InputDialog(question: "Input the operating range addresses (separated by a comma, e.g. \"B2:C3,B4:C5\")", defaultAnswer: latestOperatingRangeAddresses); //弹出对话框，输入操作区域
            if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则函数返回值赋值为null
            {
                return null;
            }
            string operatingRangeAddresses = inputDialog.Answer; //获取对话框返回的操作区域
            userRecords.LatestOperatingRangeAddresses = operatingRangeAddresses; //将对话框返回的操作区域赋值给用户使用记录

            //将操作区域地址拆分为数组，转换成列表，并移除每个元素的首尾空白字符，赋值给函数返回值
            return operatingRangeAddresses.Split(',').ToList().ConvertAll(e => e.Trim());
        }

        public static (int startIndex, int endIndex) GetWorksheetRange()
        {
            string latestExcelWorksheetIndexesStr = userRecords.LatestExcelWorksheetIndexesStr; //读取用户使用记录中保存的Excel工作表索引号范围字符串
            InputDialog inputDialog = new InputDialog(question: "Input the index number or range of worksheets to be processed (a single number, e.g. \"1\", or 2 numbers separated by a hyphen, e.g. \"1-3\")", defaultAnswer: latestExcelWorksheetIndexesStr); //弹出对话框，输入工作表索引号范围

            if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则工作表索引号范围起始值均-1，赋值给函数返回值
            {
                return (-1, -1);
            }

            string excelWorksheetIndexesStr = inputDialog.Answer;
            userRecords.LatestExcelWorksheetIndexesStr = excelWorksheetIndexesStr; // 将对话框返回的Excel工作表索引号范围字符串赋值给用户使用记录
            //将Excel工作表索引号字符串拆分成数组，转换成列表，移除每个元素的首尾空白字符，转换成数值，减去1（EPPlus工作表索引号从0开始，Excel从1开始），赋值给Excel工作表索引号列表
            List<int> lstExcelWorksheetIndexesStr = excelWorksheetIndexesStr.Split('-').ToList().ConvertAll(e => Convert.ToInt32(e.Trim())).ConvertAll(e => e - 1);
            int index1 = lstExcelWorksheetIndexesStr[0]; //获取Excel工作表索引号界值1：列表的0号元素的值
            int index2 = lstExcelWorksheetIndexesStr.Count() == 1 ? index1 : lstExcelWorksheetIndexesStr[1]; //获取Excel工作表索引号界值2：如果Excel工作表索引号列表只有一个元素（界值1和2相同），则得到Excel工作表索引号界值1；否则，得到列表的1号元素的值
            return (Math.Min(index1, index2), Math.Max(index1, index2)); // 将Excel工作表索引号的2个界值中较小的和较大的值分别作为起始值和结束值赋值给函数返回值元组
        }

        public static (int headerRowCount, int footerRowCount) GetHeaderAndFooterRowCount()
        {
            string lastestHeaderFooterRowCountStr = userRecords.LastestHeaderAndFooterRowCountStr; //读取设置中保存的表头表尾行数字符串
            InputDialog inputDialog = new InputDialog(question: "Input the row count of the table header and footer (separated by a comma, e.g. \"2,0\")", defaultAnswer: lastestHeaderFooterRowCountStr); //弹出对话框，输入表头表尾行数
            if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则表头、表尾行数均为-1，赋值给函数返回值元组
            {
                return (-1, -1);
            }

            string headerFooterRowCountStr = inputDialog.Answer; //获取对话框返回的表头、表尾行数字符串
            userRecords.LastestHeaderAndFooterRowCountStr = headerFooterRowCountStr; // 将对话框返回的表头、表尾行数字符串存入设置

            //将表头、表尾字符串拆分成数组，转换成列表，移除每个元素的首尾空白字符，转换成数值，如果小于0则限定为0，并赋值给表头表尾行数列表
            List<int> lstHeaderFooterRowCount = headerFooterRowCountStr.Split(',').ToList().ConvertAll(e => Convert.ToInt32(e.Trim())).ConvertAll(e => Math.Max(0, e));
            //将表头表尾行数列表0号、1号元素，赋值给函数返回值
            return (lstHeaderFooterRowCount[0], lstHeaderFooterRowCount[1]);
        }

        public static int GetInstanceCountByHandle<T>() where T : Window //泛型参数T，T必须是Window的实例
        {
            int count = 0;
            foreach (Window window in Application.Current.Windows) //遍历所有的窗口
            {
                if (window is T && new WindowInteropHelper(window).Handle != IntPtr.Zero) //如果当前窗口是指定类型（窗口）的实例，且句柄不为0（窗口打开状态）
                {
                    count++; //计数器加1
                }
            }
            return count; //计数器值赋给函数返回值
        }

        public static void KillOfficeApps(object[] apps)
        {
            foreach (Object app in apps)
            {
                if (app != null)
                {
                    dynamic dynamicApp = app;
                    dynamicApp.Quit(); //退出应用程序
                    Marshal.ReleaseComObject(app); //释放COM对象
                    Marshal.FinalReleaseComObject(app);
                }
            }
            GC.Collect(); //垃圾回收
            GC.WaitForPendingFinalizers();
        }

        public static async Task<DataTable?> ReadExcelWorksheetIntoDataTableAsync(string filePath, object worksheetID)
        {
            Task<DataTable?> task = Task.Run(() => Process());
            DataTable? Process()
            {
                try
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath))) // 打开Excel文件，赋值给Excel包变量
                    {
                        ExcelWorksheet? excelWorksheet = null;
                        switch (worksheetID) //根据worksheetID变量类型进入相应的分支
                        {
                            case int index: //如果为整数，则赋值给索引号变量
                                excelWorksheet = excelPackage.Workbook.Worksheets[index]; //将指定索引号的Excel工作表赋值给Excel工作表变量（Excel工作表索引号从1开始，EPPlus从0开始）
                                break;
                            case string name: //如果为字符串，则赋值给名称变量
                                excelWorksheet = excelPackage.Workbook.Worksheets[name]; //将指定名称的Excel工作表赋值给Excel工作表变量
                                break;
                            default: //以上均不符合，则抛出异常

                                throw new Exception("Parameter error.");
                        }

                        TrimCellStrings(excelWorksheet!, true); //删除Excel工作表内所有单元格值的首尾空格，并全部转换为文本型
                        RemoveWorksheetEmptyRowsAndColumns(excelWorksheet!); //删除Excel工作表内所有空白行和空白列
                        if ((excelWorksheet.Dimension?.Rows ?? 0) <= 1) //如果Excel工作表已使用行数（如果工作表为空，则为0）小于等于1，则函数返回值赋值为null
                        {
                            return null;
                        }

                        foreach (ExcelRangeBase cell in excelWorksheet.Cells[excelWorksheet.Dimension!.Address]) //遍历已使用区域的所有单元格
                        {
                            //移除当前单元格文本首尾空白字符后重新赋值给当前单元格（所有单元格均转为文本型）
                            cell.Value = cell.Text.Trim();
                        }

                        DataTable dataTable = new DataTable(); // 定义DataTable变量
                                                               //读取Excel工作表并载入DataTable（第一行为表头，将所有错误值视为空值，总是允许无效值）
                        dataTable = excelWorksheet.Cells[excelWorksheet.Dimension.Address].ToDataTable(
                            o =>
                            {
                                o.FirstRowIsColumnNames = true;
                                //o.SkipNumberOfRowsEnd = footerRowCount; // 跳过表尾指定行数
                                o.ExcelErrorParsingStrategy = ExcelErrorParsingStrategy.HandleExcelErrorsAsBlankCells;
                                o.AlwaysAllowNull = true;
                            });

                        dataTable = RemoveDataTableEmptyRowsAndColumns(dataTable); // 删除DataTable内所有空白行和空白列

                        //将DataTable赋值给函数返回值：如果DataTable的数据行和列数均不为0，则得到DataTable；否则得到null
                        return (dataTable.Rows.Count * dataTable.Columns.Count > 0) ? dataTable : null;
                    }
                }

                catch (Exception) // 捕获错误
                {
                    return null; //函数返回值赋值为null
                }
            }

            return await task;
        }

        public static DataTable RemoveDataTableEmptyRowsAndColumns(DataTable dataTable)
        {
            //清除空白数据行
            for (int i = dataTable.Rows.Count - 1; i >= 0; i--) // 遍历DataTable所有数据行
            {
                // 如果当前数据行的所有数据列的值均为数据库空值，或为null或全空白字符，则删除当前数据行
                if (dataTable.Rows[i].ItemArray.All(value => value == DBNull.Value || string.IsNullOrWhiteSpace(value?.ToString())))
                {
                    dataTable.Rows[i].Delete();
                }

                dataTable.AcceptChanges();
            }

            //清除空白数据列
            for (int j = dataTable.Columns.Count - 1; j >= 0; j--) // 遍历DataTable所有数据列
            {
                //如果所有数据行的当前数据列的值均为数据库空值，或为null或全空白字符，则删除当前数据列
                if (dataTable.AsEnumerable().All(dataRow => dataRow[j] == DBNull.Value || string.IsNullOrWhiteSpace(dataRow[j].ToString())))
                {
                    dataTable.Columns.RemoveAt(j);
                }
            }
            dataTable.AcceptChanges(); //接受上述更改

            return dataTable; // 将DataTable赋值给函数返回值
        }

        public static string RemoveMarkdownMarks(this string inText)
        {
            string outText = inText;

            // 行首尾空白字符正则表达式匹配模式为：开头标记，不为非空白字符也不为换行符的字符（不为换行符的空白字符）至少一个/或前述字符至少一个，结尾标记；将匹配到的字符串替换为空
            //[^\S\n]+与(?:(?!\n)\s)+等同
            outText = Regex.Replace(outText, @"^[^\S\n]+|[^\S\n]+$", "", RegexOptions.Multiline);

            // 将行内换行符号替换为换行符
            outText = Regex.Replace(outText, @"<br>", "\n", RegexOptions.Multiline);

            // 水平分隔线符号正则表达式匹配模式为：开头标记，“*-_”至少一个，结尾标记；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^[\*\-_]+$", "", RegexOptions.Multiline);

            // 标题符号正则表达式匹配模式为：开头标记，“#”（同行标题标记）至少一个，空格至少一个/或开头标记，“=-”（上一行标题标记）至少一个，结尾标记；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^#+[ ]+|^[=\-]+$", "", RegexOptions.Multiline);

            // 无序列表符号正则表达式匹配模式为：开头标记，“*-+”，空格至少一个；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^[\*\-\+][ ]+", "", RegexOptions.Multiline);

            // 引用符号正则表达式匹配模式为：开头标记，“>”，空格任意多个；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^>[ ]*", "", RegexOptions.Multiline);

            // 移除行内代码和代码块引用符号 - 代码引用符号正则表达式匹配模式为：任意字符任意多个（尽量少匹配）（捕获组1），“`”1-3个（捕获组2），任意字符任意多个（尽量少匹配）（捕获组3），捕获组2，任意字符任意多个（尽量少匹配）；将匹配到的字符串替换为捕获组1、3、4合并后的字符串
            outText = Regex.Replace(outText, @"(.*?)(`{1,3})(.*?)\2(.*?)", "$1$3$4", RegexOptions.Singleline); // 文本可能跨多行，"."需包含换行符，故使用单行匹配模式

            // 移除行内公式和公式块引用符号
            outText = Regex.Replace(outText, @"(.*?)(\${1,2})(.*?)\2(.*?)", "$1$3$4", RegexOptions.Singleline);

            // 移除删除线符号
            outText = Regex.Replace(outText, @"(.*?)(~~)(.*?)\2(.*?)", "$1$3$4", RegexOptions.Multiline);

            // 表格表头分隔线符号正则表达式匹配模式为：开头标记，“|-:”至少一个，结尾标记；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^[\|\-:]+$", "", RegexOptions.Multiline);

            // 表格行开头和结尾符号正则表达式匹配模式为：开头标记，“|”/或“|”，结尾标记；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^\||\|$", "", RegexOptions.Multiline);

            // 表格内部多余空白字符正则表达式匹配模式为：前方出现“|”，不为换行符的空白字符至少一个/或前述字符至少一个，后方出现“|”；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"(?<=\|)[^\S\n]+|[^\S\n]+(?=\|)", "", RegexOptions.Multiline);

            // 移除斜体或粗体符号（1个代表斜体，2个代表粗体，3个代表粗斜体）
            //outText = Regex.Replace(outText, @"(^|.*?)([\*_]{1,3})(.*?)\2(.*?|$)", "$1$3$4", RegexOptions.Singleline);
            outText = Regex.Replace(outText, @"(.*?)([\*_]{1,3})(.*?)\2(.*?)", "$1$3$4", RegexOptions.Multiline);

            // 空白行正则表达式匹配模式设为：开头标记，空白字符任意多个；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^\s*", "", RegexOptions.Multiline);

            // 再次将每行首尾空白字符替换为空
            outText = Regex.Replace(outText, @"^[^\S\n]+|[^\S\n]+$", "", RegexOptions.Multiline);

            return outText; //将输出文字赋值给函数返回值
        }

        public static void RemoveWorksheetEmptyRowsAndColumns(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet.Dimension == null) //如果Excel工作表为空，结束本过程
            {
                return;
            }

            // 删除空白行
            for (int i = excelWorksheet.Dimension.End.Row; i >= 1; i--) //遍历所有行
            {
                //如果当前行第1列到最末列所有单元格均为null或全空白字符，则删除当前行
                if (excelWorksheet.Cells[i, 1, i, excelWorksheet.Dimension.End.Column].All(c => string.IsNullOrWhiteSpace(c.Text)))
                {
                    excelWorksheet.DeleteRow(i);
                }
            }

            // 删除空白列
            for (int j = excelWorksheet.Dimension.End.Column; j >= 1; j--) //遍历所有列
            {
                //如果当前列第1行到最末行所有单元格均为null或全空白字符，则删除当前列
                if (excelWorksheet.Cells[1, j, excelWorksheet.Dimension.End.Row, j].All(c => string.IsNullOrWhiteSpace(c.Text)))
                {
                    excelWorksheet.DeleteColumn(j);
                }
            }

        }

        // 定义Emoji正则表达式字符串
        static string emojiRegEx = @"\p{So}" + //"Symbol, Other" 的所有字符，涵盖了绝大多数BMP平面内的符号。
                @"|[\uD800-\uDBFF][\uDC00-\uDFFF]" + //匹配补充平面的主要Emoji范围，包括表情、交通、国旗、卡牌等。
                @"|[\u200D\uFE0F]"; //匹配用于组合Emoji的零宽连字和变体选择器。

        // 创建一个静态的Regex对象，用于匹配Emoji字符
        static Regex regExEmoji = new Regex(emojiRegEx, RegexOptions.Compiled);

        public static string RemoveEmojis(this string text)
        {

            return regExEmoji.Replace(text, string.Empty); // 正则表达式匹配模式设为所有Emoji字符；将匹配到的字符串替换为空，赋值给函数返回值
        }

        public enum EnumFileType { Excel, Word, DocumentAndTable, Convertible, Pdf, All } //定义文件类型枚举

        public static List<string>? SelectFiles(EnumFileType fileType, bool isMultiselect, string dialogTitle)
        {
            string filter = fileType switch //根据文件类型枚举，返回相应的文件类型和扩展名的过滤项
            {
                EnumFileType.Excel => "Excel Files(*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files(*.*)|*.*",
                EnumFileType.Word => "Word Files(*.docx;*.docm)|*.docx;*.docm|All Files(*.*)|*.*",
                EnumFileType.DocumentAndTable => "Document And Table Files(*.docx;*.xlsx;*.docm;*.xlsm;*.pdf)|*.docx;*.xlsx;*.docm;*.xlsm;*.pdf|All Files(*.*)|*.*",
                EnumFileType.Convertible => "Convertible Files(*.doc;*.xls;*.wps;*.et)|*.doc;*.xls;*.wps;*.et|All Files(*.*)|*.*",
                //EnumFileType.Pdf => "PDF Files(*.pdf)|*.pdf|All Files(*.*)|*.*",
                _ => "All Files(*.*)|*.*"
            };

            string initialDirectory = userRecords.LatestFolderPath; //获取保存在设置中的文件夹路径
            //重新赋值给初始文件夹路径变量：如果初始文件夹路径存在，则得到初始文件夹路径原值；否则得到C盘根目录
            initialDirectory = Directory.Exists(initialDirectory) ? initialDirectory : "C:" + Path.DirectorySeparatorChar;
            OpenFileDialog openFileDialog = new OpenFileDialog() //打开文件选择对话框
            {
                Multiselect = isMultiselect, //是否可多选
                Title = dialogTitle, //对话框标题
                Filter = filter, //文件类型和相应扩展名的过滤项
                InitialDirectory = initialDirectory //初始文件夹路径
            };

            if (openFileDialog.ShowDialog() == true) //如果对话框返回true（选择了OK）
            {
                userRecords.LatestFolderPath = Path.GetDirectoryName(openFileDialog.FileNames[0])!; // 将本次选择的文件的文件夹路径保存到设置中

                return openFileDialog.FileNames.ToList(); // 将被选中的文件数组转换成列表，赋给函数返回值
            }
            return null; //如果上一个if未执行，没有文件列表赋给函数返回值，则函数返回值赋值为null
        }

        public static string? SelectFolder(string dialogTitle)
        {
            string initialDirectory = userRecords.LatestFolderPath; // 读取用户使用记录中保存的文件夹路径
            // 重新赋值给文件夹路径变量：如果文件夹路径存在，则得到该文件夹路径原值；否则得到C盘根目录
            initialDirectory = Directory.Exists(initialDirectory) ? initialDirectory : "C:" + Path.DirectorySeparatorChar;

            OpenFolderDialog openFolderDialog = new OpenFolderDialog() // 打开文件夹选择对话框
            {
                Multiselect = false, // 禁用多选
                Title = dialogTitle, // 设置对话框标题
                RootDirectory = initialDirectory // 根文件夹路径设为文件夹路径
            };

            if (openFolderDialog.ShowDialog() == true) // 如果对话框返回值为true（点击OK）
            {
                string folderPath = openFolderDialog.FolderName;
                userRecords.LatestFolderPath = folderPath;  // 将文件夹路径赋值给用户使用记录
                return folderPath; // 将文件夹路径赋值给函数返回值
            }
            return null; // 如果上一个if未执行，没有文件夹路径赋给函数返回值，则函数返回值赋值为null

        }

        public static int SelectFunction(List<string> lstOptions, object objRecords, string propertyName)
        {
            // 使用反射方法获取对象属性值
            Type type = objRecords.GetType(); // 获取用户使用记录对象类型
            PropertyInfo? property = type.GetProperty(propertyName); // 获取对象的指定属性
            if (property == null) // 如果对象属性为空，则将-1赋值给函数返回值
            {
                return -1;
            }

            object value = property.GetValue(objRecords) ?? ""; //  获取对象指定属性的值
            string latestOption = (string)value; //将指定属性的值转换成字符串

            InputDialog inputDialog = new InputDialog(question: "Select the Function", options: lstOptions, defaultAnswer: latestOption); //弹出功能选择对话框
            if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则将-1赋值给函数返回值
            {
                return -1;
            }

            string selectedOption = inputDialog.Answer;

            // 使用反射方法设置对象属性值
            if (property!.CanWrite) //  如果对象属性可写
            {
                property.SetValue(objRecords, selectedOption); //将对话框返回的功能选项字符串赋值给用户使用记录对象指定属性
            }

            int functionNum = lstOptions.Contains(selectedOption) ? lstOptions.IndexOf(selectedOption) : -1; //获取对话框返回的功能选项在功能列表中的索引号：如果功能列表包含功能选项，则得到对应的索引号；否则，得到-1
            return functionNum; // 将功能选项索引号赋值给函数返回值
        }

        public static void ShowExceptionMessage(Exception ex)
        {
            MessageDialog messageDialog = new MessageDialog($"{ex.Message}\n{ex.InnerException?.Message ?? ""}");
            messageDialog.ShowDialog();
        }

        public static bool ShowMessage(string message)
        {
            MessageDialog messageDialog = new MessageDialog(message);
            return messageDialog.ShowDialog() ?? false; // 将对话框返回值（点击OK为true，点击Cancel为false）赋值给函数返回值（如果对话框返回null，则为false)
        }

        public static void ShowSuccessMessage()
        {
            MessageDialog messageDialog = new MessageDialog("Operation completed.");
            messageDialog.ShowDialog();
        }

        public static bool IsChineseText(string inText)
        {
            //判断是否为中文文档
            if (inText.Length == 0) return false; // 如果全文长度为0，则将false赋值给函数返回值

            int nonCnCharCount = Regex.Matches(inText, @"[^\u4e00-\u9fa5]").Count; //获取全文非中文字符数量
            double nonCnCharsRatio = nonCnCharCount / inText.Length; // 计算非中文字符占全文的比例
            return nonCnCharsRatio < 0.5 ? true : false; //赋值给函数返回值：如果非中文字符比例小于0.5，得到true；否则，得到false
        }

        public static bool IsModal(this Window window)
        {
            // 使用反射获取 Window 类的私有字段 "_showingAsDialog",该字段是 WPF 内部用于标记窗口是否通过 ShowDialog() 方法显示的布尔值
            // 查找实例成员（非静态）、查找非公开成员（private/internal）
            var field = typeof(Window).GetField("_showingAsDialog", BindingFlags.Instance | BindingFlags.NonPublic); 

            // 返回判断结果：
            // 1. field != null：确保成功获取到字段信息（防止未来版本字段名变更导致异常）
            // 2. (bool)field.GetValue(window)：获取该字段在当前窗口实例中的值并转为 bool
            // 使用短路运算 && ：如果 field 为 null，不会执行 GetValue，避免空引用异常
            return field != null && (bool)field.GetValue(window)!;
        }

        public static void CreateFolder(string targetFolderPath)
        {
            try
            {
                //创建目标文件夹
                if (!Directory.Exists(targetFolderPath))
                {
                    Directory.CreateDirectory(targetFolderPath!);
                }
            }

            catch (Exception)
            {
                throw;
            }
        }

        public static void TrimCellStrings(ExcelWorksheet excelWorksheet, bool covertAllTypesToString = false)
        {
            if (excelWorksheet.Dimension == null) //如果Excel工作表为空，结束本过程
            {
                return;
            }

            foreach (ExcelRangeBase cell in excelWorksheet.Cells[excelWorksheet.Dimension.Address]) //遍历已使用区域的所有单元格
            {
                if (covertAllTypesToString)  //如果“将所有类型的值均转换为字符串”为true，则将当前单元格值移除其首尾空白字符后，重新赋值给单元格
                {
                    cell.Value = cell.Text.Trim();
                }
                //否则，如果当前单元格值为字符串且不含公式，则移除其首尾空白字符后，重新赋值给单元格
                else if (cell.Value is string && string.IsNullOrWhiteSpace(cell.Formula))
                {
                    cell.Value = cell.Text.Trim();
                }
            }

        }

        public static double Val(object? cellValue)
        {
            if (cellValue == null) //如果参数为null，将0赋值给函数返回值
            {
                return 0;
            }

            string cellStr = Convert.ToString(cellValue)!;
            //cellStr = Regex.Replace(cellStr, @"[^\d\.+\-]", ""); //移除单元格值中的非数字、小数点或正负号的字符
            cellStr = Regex.Match(cellStr, @"[+\-]?\d+(?:\.\d*)?")?.Value.ToString() ?? ""; // 正则表达式匹配模式为：正负号至多一个，数字至少一个，（小数点，数字任意个）至多一组；将匹配到的字符串赋值给单元格字符串变量

            //如果将匹配结果转换为double类型成功，则将转换结果赋值给number变量，然后再将number变量值赋值给函数返回值
            if (double.TryParse(cellStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double number))
            {
                return number;
            }

            return 0; //如果以上过程均没有赋值给函数返回值，此处将0赋值给函数返回值
        }

        public static void WriteDataTableIntoExcelWorkbook(List<DataTable> lstDataTable, string filePath)
        {
            FileInfo targetExcelFile = new FileInfo(filePath); //获取目标Excel工作簿文件路径全名信息

            using (ExcelPackage excelPackage = new ExcelPackage()) //新建Excel包，赋值给Excel包变量
            {
                int i = 1;
                foreach (DataTable dataTable in lstDataTable) //遍历数据表列表
                {
                    ExcelWorksheet targetExcelWorksheet = excelPackage.Workbook.Worksheets.Add($"Sheet{i++}"); //新建Excel工作表，赋值给目标工作表变量
                    targetExcelWorksheet.Cells["A1"].LoadFromDataTable(dataTable, true); //将DataTable数据导入目标Excel工作表（true代表将表头赋给第一行）

                    FormatExcelWorksheet(targetExcelWorksheet, 1, 0); //设置目标Excel工作表格式
                }
                excelPackage.SaveAs(targetExcelFile);
            }
        }
    }
}
