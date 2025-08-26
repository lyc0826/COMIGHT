using Microsoft.Office.Interop.Word;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using static COMIGHT.MainWindow;
using static COMIGHT.Methods;
using MSWord = Microsoft.Office.Interop.Word;
using MSWordDocument = Microsoft.Office.Interop.Word.Document;
using MSWordParagraph = Microsoft.Office.Interop.Word.Paragraph;
using MSWordSection = Microsoft.Office.Interop.Word.Section;
using MSWordTable = Microsoft.Office.Interop.Word.Table;
using Task = System.Threading.Tasks.Task;



namespace COMIGHT
{
    public partial class MSOfficeInterop
    {
        public static async Task BatchFormatWordDocumentsAsyncHelper(List<string> filePaths)
        {
            Task task = Task.Run(() => process()); // 创建一个异步任务，执行过程为process()
            void process()
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

                    // 定义大标题正则表达式变量，匹配模式为：从开头开始，不含2个及以上连续的换行符回车符（允许不连续的换行符回车符）、不含“附件/录”、Appendix注释、非“。”分页符的字符1-150个，换行符回车符，后方出现：换行符回车符
                    Regex regExTitle = new Regex(@"(?<=^|\n|\r)(?:(?![\n\r]{2,})(?!(?:附[ |\t]*[件录]|appendix)[^。\f\n\r]{0,3}[\n\r])[^。\f]){1,150}[\n\r](?=[\n\r])", RegexOptions.Multiline | RegexOptions.IgnoreCase);

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
                                if (!matchCnHeading1_2.Groups[1].Success) //如果正则表达式匹配捕获组失败（ 编号开头不含“（” ）
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
                            Regex regExTableTitle = new Regex(tableTitleRegEx, RegexOptions.Multiline | RegexOptions.IgnoreCase);

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

        public static async Task BatchRepairWordDocumentsAsyncHelper(List<string> filePaths)
        {
            Task task = Task.Run(() => process());
            void process()
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

    }
}
