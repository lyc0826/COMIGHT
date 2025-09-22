using GEmojiSharp;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
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
using static COMIGHT.MainWindow;
using Application = System.Windows.Application;
using DataTable = System.Data.DataTable;
using ICell = NPOI.SS.UserModel.ICell;
using Window = System.Windows.Window;

namespace COMIGHT
{
    public static partial class Methods
    {

        // 定义表格标题正则表达式字符串（需要兼顾常规字符串和Word中的文本）
        public static string tableTitleRegEx = @"(?<=^|\n|\r)[^。：:；;\f\n\r]{0,100}(?:表|单|录|册|回执|table|form|list|roll|roster)[\d\.一二三四五六七八九十〇零（）\(\)：:\-| |\t]*[^。：:；;\f\n\r]{0,100}(?:[\n\r]|$)";

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
                            using (FileStream excelStream = File.Create(targetExcelFilePath)) //创建目标Excel工作簿，赋值给Excel工作簿文件流变量
                            {
                                IWorkbook workbook = new XSSFWorkbook(); // 创建Excel工作簿对象，赋值给Excel工作簿变量
                                int wordTableIndex = 0;
                                for (int i = 0; i < wordDocument.BodyElements.Count; i++) // 遍历目标Word文档中的所有元素
                                {
                                    var wordElement = wordDocument.BodyElements[i]; // 获取目标Word文档中当前元素，并赋值给Word元素变量
                                    if (wordElement is XWPFTable wordTable) // 如果当前Word元素是表格类型，则将其赋值给新变量 wordTable，然后：
                                    {
                                        string tableTitle = "Sheet" + (wordTableIndex + 1); // 定义表格标题，默认为“Sheet”与当前word文档表格索引号加1

                                        // 获取表格标题
                                        if (i > 0) // 如果当前Word元素不是0号元素
                                        {
                                            List<string> lstBackupTableTitle = new List<string>();
                                            string preferredTableTitle = string.Empty;
                                            for (int k = 1; k <= 5 && i - k >= 0; k++) // 从当前Word元素开始，向前遍历5个元素，直到0号元素为止）
                                            {
                                                if (wordDocument.BodyElements[i - k] is XWPFParagraph) // 如果前方当前Word元素是Word段落
                                                {
                                                    XWPFParagraph paragraph = (XWPFParagraph)wordDocument.BodyElements[i - k]; // 获取前方当前Word元素，并赋值给段落变量
                                                    // 如果段落文字被表格标题正则表达式匹配成功，将段落文字赋给首选表格标题变量并退出循环
                                                    if (Regex.IsMatch(paragraph.Text, tableTitleRegEx))
                                                    {
                                                        preferredTableTitle = paragraph.Text;
                                                        break;
                                                    }
                                                    // 否则，备选表格标题正则表达式模式设为：开头标记，不含“。；;：:”的字符1-100个，结尾标记；如果段落文字被匹配成功，将被增加到备用表格标题列表中
                                                    else if (Regex.IsMatch(paragraph.Text, @"^[^。；;：:]{1,100}$"))
                                                    {
                                                        lstBackupTableTitle.Add(paragraph.Text);
                                                    }
                                                }
                                            }
                                            // 获取表格标题：如果首选表格标题变量不为空，则得到该变量值；否则，如果备用表格标题列表不为空，则得到其0号（第1个）元素的值；否则，得到表格标题变量原值
                                            tableTitle = !string.IsNullOrWhiteSpace(preferredTableTitle) ? preferredTableTitle : lstBackupTableTitle.Count > 0 ? lstBackupTableTitle[0] : tableTitle;
                                        }

                                        //创建Excel工作表，使用序号加表格标题作为工作表的名称
                                        ISheet worksheet = workbook.CreateSheet(CleanWorksheetName($"{wordTableIndex + 1}_{tableTitle}")); // 创建Excel工作表对象

                                        IRow excelFirstRow = worksheet.CreateRow(0); // 创建Excel 0号（第1）行对象，赋值给Excel第一行变量

                                        int columnCount = wordTable.Rows.Max(r => r.GetTableCells().Count); //获取Word文档表格所有行里包含单元格数量最多的那一行的单元格数量，即Word文档表格列数，赋值给表格列数变量

                                        worksheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, columnCount - 1)); // 合并Excel工作表第一行单元格
                                        excelFirstRow.CreateCell(0).SetCellValue(tableTitle); // 将表格标题赋值给Excel工作表0号（第1）行单元格

                                        int excelRowIndex = 1; // 从Excel工作表1号（第2）行开始写入表格数据
                                        foreach (XWPFTableRow wordTableRow in wordTable.Rows) // 遍历当前Word文档表格中的所有行
                                        {
                                            IRow excelRow = worksheet.CreateRow(excelRowIndex++); // 创建Excel行对象，赋值给Excel行变量
                                            int excelCellIndex = 0;
                                            foreach (XWPFTableCell wordTableCell in wordTableRow.GetTableCells()) // 遍历当前Word文档表格当前行中的所有单元格
                                            {
                                                ICell excelCell = excelRow.CreateCell(excelCellIndex++); // 创建Excel单元格对象，赋值给Excel单元格变量
                                                excelCell.SetCellValue(wordTableCell.GetText()); // 将当前Word文档表格的当前行的当前单元格的文字赋值给当前Excel单元格
                                            }
                                        }
                                        wordTableIndex++; // Word文档表格索引号累加1
                                    }
                                }
                                workbook.Write(excelStream); // 将Excel工作簿文件流写入目标Excel工作簿文件

                                // 格式化目标Excel工作簿中的表格
                                FileInfo targetExcelFile = new FileInfo(targetExcelFilePath); //获取目标Excel文件路径全名信息
                                using (ExcelPackage excelPackage = new ExcelPackage(targetExcelFile)) //打开目标Excel文件，赋值给Excel包变量
                                {
                                    foreach (ExcelWorksheet excelWorksheet in excelPackage.Workbook.Worksheets) //遍历目标Excel工作簿中的所有工作表
                                    {
                                        FormatExcelWorksheet(excelWorksheet, 2, 0); // 格式化表格数据区域（表头为2行）
                                    }
                                    excelPackage.Save(); //保存目标Excel文档
                                }
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

        public static string? GetKeyColumnLetter()
        {
            string latestColumnLetter = latestRecords.LatestKeyColumnLetter; //读取设置中保存的主键列符
            InputDialog inputDialog = new InputDialog(question: "Input the key column letter (e.g. \"A\"）", defaultAnswer: latestColumnLetter); //弹出对话框，输入主键列符
            if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则函数返回值赋值为null
            {
                return null;
            }
            string columnLetter = inputDialog.Answer;
            latestRecords.LatestKeyColumnLetter = columnLetter; // 将对话框返回的列符存入设置

            return columnLetter; //将列符赋值给函数返回值
        }

        public static List<string>? GetWorksheetOperatingRangeAddresses()
        {
            string latestOperatingRangeAddresses = latestRecords.LatestOperatingRangeAddresses; //读取用户使用记录中保存的操作区域
            InputDialog inputDialog = new InputDialog(question: "Input the operating range addresses (separated by a comma, e.g. \"B2:C3,B4:C5\")", defaultAnswer: latestOperatingRangeAddresses); //弹出对话框，输入操作区域
            if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则函数返回值赋值为null
            {
                return null;
            }
            string operatingRangeAddresses = inputDialog.Answer; //获取对话框返回的操作区域
            latestRecords.LatestOperatingRangeAddresses = operatingRangeAddresses; //将对话框返回的操作区域赋值给用户使用记录

            //将操作区域地址拆分为数组，转换成列表，并移除每个元素的首尾空白字符，赋值给函数返回值
            return operatingRangeAddresses.Split(',').ToList().ConvertAll(e => e.Trim());
        }

        public static (int startIndex, int endIndex) GetWorksheetRange()
        {
            string latestExcelWorksheetIndexesStr = latestRecords.LatestExcelWorksheetIndexesStr; //读取用户使用记录中保存的Excel工作表索引号范围字符串
            InputDialog inputDialog = new InputDialog(question: "Input the index number or range of worksheets to be processed (a single number, e.g. \"1\", or 2 numbers separated by a hyphen, e.g. \"1-3\")", defaultAnswer: latestExcelWorksheetIndexesStr); //弹出对话框，输入工作表索引号范围

            if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则工作表索引号范围起始值均-1，赋值给函数返回值
            {
                return (-1, -1);
            }

            string excelWorksheetIndexesStr = inputDialog.Answer;
            latestRecords.LatestExcelWorksheetIndexesStr = excelWorksheetIndexesStr; // 将对话框返回的Excel工作表索引号范围字符串赋值给用户使用记录
            //将Excel工作表索引号字符串拆分成数组，转换成列表，移除每个元素的首尾空白字符，转换成数值，减去1（EPPlus工作表索引号从0开始，Excel从1开始），赋值给Excel工作表索引号列表
            List<int> lstExcelWorksheetIndexesStr = excelWorksheetIndexesStr.Split('-').ToList().ConvertAll(e => Convert.ToInt32(e.Trim())).ConvertAll(e => e - 1);
            int index1 = lstExcelWorksheetIndexesStr[0]; //获取Excel工作表索引号界值1：列表的0号元素的值
            int index2 = lstExcelWorksheetIndexesStr.Count() == 1 ? index1 : lstExcelWorksheetIndexesStr[1]; //获取Excel工作表索引号界值2：如果Excel工作表索引号列表只有一个元素（界值1和2相同），则得到Excel工作表索引号界值1；否则，得到列表的1号元素的值
            return (Math.Min(index1, index2), Math.Max(index1, index2)); // 将Excel工作表索引号的2个界值中较小的和较大的值分别作为起始值和结束值赋值给函数返回值元组
        }

        public static (int headerRowCount, int footerRowCount) GetHeaderAndFooterRowCount()
        {
            string lastestHeaderFooterRowCountStr = latestRecords.LastestHeaderAndFooterRowCountStr; //读取设置中保存的表头表尾行数字符串
            InputDialog inputDialog = new InputDialog(question: "Input the row count of the table header and footer (separated by a comma, e.g. \"2,0\")", defaultAnswer: lastestHeaderFooterRowCountStr); //弹出对话框，输入表头表尾行数
            if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则表头、表尾行数均为-1，赋值给函数返回值元组
            {
                return (-1, -1);
            }

            string headerFooterRowCountStr = inputDialog.Answer; //获取对话框返回的表头、表尾行数字符串
            latestRecords.LastestHeaderAndFooterRowCountStr = headerFooterRowCountStr; // 将对话框返回的表头、表尾行数字符串存入设置

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

        public static void MergeExcelWorksheetHeader(ExcelWorksheet excelWorksheet, int headerRowCount)
        {
            if (excelWorksheet.Dimension == null || headerRowCount < 2) //如果工作表为空或者表头行数小于2，则结束本过程
            {
                return;
            }

            excelWorksheet.Cells[1, 1, headerRowCount, excelWorksheet.Dimension.End.Column].Merge = false; //表头所有单元格的合并状态设为false

            //删除表头行中只含一个有效数据单元格的行（该行没有任何分类意义）
            for (int i = headerRowCount; i >= 1; i--) //遍历表头所有行
            {
                ExcelRange headerRowCells = excelWorksheet.Cells[i, 1, i, excelWorksheet.Dimension.End.Column]; //将当前行所有单元格赋值给表头行单元格变量

                int usedCellCount = headerRowCells.Count(cell => !string.IsNullOrWhiteSpace(cell.Text)); // 计算当前表头行单元格中不为null或全空白字符的单元格数量，赋值给已使用单元格数量变量
                if (usedCellCount <= 1) //如果已使用单元格数量小于等于1
                {
                    excelWorksheet.DeleteRow(i); //删除当前行
                    headerRowCount--; //表头行数减1
                }
            }

            for (int j = 1; j <= excelWorksheet.Dimension.End.Column; j++) //遍历工作表所有列
            {
                List<string> lstFullColumnName = new List<string>(); //定义完整列名称列表
                for (int i = 1; i <= headerRowCount; i++) //遍历表头所有行
                {
                    bool copyLeftCell = false; //“是否复制左侧单元格”赋值为false
                    if (j > 1 && string.IsNullOrWhiteSpace(excelWorksheet.Cells[i, j].Text)) //如果当前列索引号大于1，且当前单元格为null或全空白字符
                    {
                        if (i == 1) //如果当前行是第1行，则“是否复制左侧单元格”赋值为true
                        {
                            copyLeftCell = true;
                        }
                        //否则，如果比当前行索引号小1、列索引号相同（上方）的单元格的值和比当前行索引号小1、比当前列索引号小1（左上方）的单元格相同，则“是否复制左侧单元格”赋值为true
                        else if (excelWorksheet.Cells[i - 1, j].Value == excelWorksheet.Cells[i - 1, j - 1].Value)
                        {
                            copyLeftCell = true;
                        }
                    }
                    //重新赋值给当前行、列的单元格：如果要复制左侧单元格，则得到比当前列索引号小1（左侧）单元格的值；否则得到当前单元格原值
                    excelWorksheet.Cells[i, j].Value = copyLeftCell ? excelWorksheet.Cells[i, j - 1].Text : excelWorksheet.Cells[i, j].Text;
                    lstFullColumnName.Add(excelWorksheet.Cells[i, j].Text); //将当前单元格值添加到完整列名称列表
                }
                //将完整列名称列表中不为null或全空白字符的元素合并（以下划线分隔），赋值给表头最后一行当前列的单元格
                excelWorksheet.Cells[headerRowCount, j].Value = string.Join('_', lstFullColumnName.Where(e => !string.IsNullOrWhiteSpace(e)));

            }
            excelWorksheet.DeleteRow(1, headerRowCount - 1); //删除表头除了最后一行的所有行

        }

        public static DataTable? ReadExcelWorksheetIntoDataTable(string filePath, object worksheetID, int headerRowCount = 1, int footerRowCount = 0)
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
                    if ((excelWorksheet.Dimension?.Rows ?? 0) <= headerRowCount + footerRowCount) //如果Excel工作表已使用行数（如果工作表为空，则为0）小于等于表头表尾行数和，则函数返回值赋值为null
                    {
                        return null;
                    }

                    foreach (ExcelRangeBase cell in excelWorksheet.Cells[excelWorksheet.Dimension!.Address]) //遍历已使用区域的所有单元格
                    {
                        //移除当前单元格文本首尾空白字符后重新赋值给当前单元格（所有单元格均转为文本型）
                        cell.Value = cell.Text.Trim();
                    }

                    MergeExcelWorksheetHeader(excelWorksheet, headerRowCount); //将多行表头合并为单行

                    DataTable dataTable = new DataTable(); // 定义DataTable变量
                    //读取Excel工作表并载入DataTable（第一行为表头，跳过表尾指定行数，将所有错误值视为空值，总是允许无效值）
                    dataTable = excelWorksheet.Cells[excelWorksheet.Dimension.Address].ToDataTable(
                        o =>
                        {
                            o.FirstRowIsColumnNames = true;
                            o.SkipNumberOfRowsEnd = footerRowCount;
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

        public static DataTable RemoveDataTableEmptyRowsAndColumns(DataTable dataTable, bool removeRowsWithSingleValue = false)
        {
            int valueCountThreshold = removeRowsWithSingleValue ? 2 : 1; //获取数据元素计数阈值：如果要移除仅含单个数据的数据行，则为2（每行都必须有2个及以上的数据才不会被移除）；否则为1

            //清除空白数据行
            for (int i = dataTable.Rows.Count - 1; i >= 0; i--) // 遍历DataTable所有数据行
            {
                // 统计当前数据行不为数据库空值且不为null或全空白字符的数据元素的数量
                int valueCount = dataTable.Rows[i].ItemArray.Count(value =>
                    value != DBNull.Value && !string.IsNullOrWhiteSpace(value?.ToString()));

                // 如果以上数据元素的数量小于数据元素计数阈值（该行视为无意义），则删除这一行
                if (valueCount < valueCountThreshold)
                {
                    dataTable.Rows[i].Delete();
                }

                //// 如果当前数据行的所有数据列的值均为数据库空值，或为null或全空白字符，则删除当前数据行
                //if (dataTable.Rows[i].ItemArray.All(value => value == DBNull.Value || string.IsNullOrWhiteSpace(value?.ToString())))
                //{
                //    dataTable.Rows[i].Delete();
                //}

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

        //public static string RemoveEmojis(this string text)
        //{
        //    return Regex.Replace(text, Emoji.RegexPattern, string.Empty); // 正则表达式匹配模式设为所有Emoji字符；将匹配到的字符串替换为空，赋值给函数返回值
        //}
        
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

        public enum EnumFileType { Excel, Word, WordAndExcel, Convertible, Executable, All } //定义文件类型枚举

        public static List<string>? SelectFiles(EnumFileType fileType, bool isMultiselect, string dialogTitle)
        {
            string filter = fileType switch //根据文件类型枚举，返回相应的文件类型和扩展名的过滤项
            {
                EnumFileType.Excel => "Excel Files(*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files(*.*)|*.*",
                EnumFileType.Word => "Word Files(*.docx;*.docm)|*.docx;*.docm|All Files(*.*)|*.*",
                EnumFileType.WordAndExcel => "Word And Excel Files(*.docx;*.xlsx;*.docm;*.xlsm)|*.docx;*.xlsx;*.docm;*.xlsm|All Files(*.*)|*.*",
                EnumFileType.Convertible => "Convertible Files(*.doc;*.xls;*.wps;*.et)|*.doc;*.xls;*.wps;*.et|All Files(*.*)|*.*",
                EnumFileType.Executable => "Executable Files(*.exe)|*.exe|All Files(*.*)|*.*",
                _ => "All Files(*.*)|*.*"
            };

            string initialDirectory = latestRecords.LatestFolderPath; //获取保存在设置中的文件夹路径
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
                latestRecords.LatestFolderPath = Path.GetDirectoryName(openFileDialog.FileNames[0])!; // 将本次选择的文件的文件夹路径保存到设置中

                return openFileDialog.FileNames.ToList(); // 将被选中的文件数组转换成列表，赋给函数返回值
            }
            return null; //如果上一个if未执行，没有文件列表赋给函数返回值，则函数返回值赋值为null
        }

        public static string? SelectFolder(string dialogTitle)
        {
            string initialDirectory = latestRecords.LatestFolderPath; // 读取用户使用记录中保存的文件夹路径
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
                latestRecords.LatestFolderPath = folderPath;  // 将文件夹路径赋值给用户使用记录
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
