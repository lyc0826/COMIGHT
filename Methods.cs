using Microsoft.Win32;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.Style;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Interop;
using static COMIGHT.MainWindow;
using static COMIGHT.MSOfficeInterop;
using static COMIGHT.PublicVariables;
using Application = System.Windows.Application;
using DataTable = System.Data.DataTable;
using Task = System.Threading.Tasks.Task;
using Window = System.Windows.Window;

namespace COMIGHT
{
    public static partial class Methods

    {
        static SettingsManager<AppSettings> settingsManager = new SettingsManager<AppSettings>(settingsJsonFilePath);
        static SettingsManager<LatestRecords> recordsManager = new SettingsManager<LatestRecords>(recordsJsonFilePath);

        public static T Clamp<T>(this T value, T min, T max) where T : IComparable<T> //泛型参数T，T必须实现IComparable<T>接口
        {
            //赋值给函数返回值：如果输入值比最小值小，则得到最小值；如果比最大值大，则得到最大值；否则，得到输入值
            return value.CompareTo(min) < 0 ? min : value.CompareTo(max) > 0 ? max : value;
        }

        public static string CleanFileAndFolderName(string name, int maxLength)
        {
            // 定义文件名和文件夹名中不允许出现的字符，赋值给非法字符变量
            string invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());

            // 将文件名和文件夹名中允许出现的字符提取并形成数组，再转换成字符串，赋值给清理后的名称变量
            string cleanedName = new string(name.Where(c => !invalidChars.Contains(c)).ToArray());

            // 截取到指定长度
            return cleanedName.Length > maxLength ? cleanedName.Substring(0, maxLength) : cleanedName;
        }

        //public static string CleanFileAndFolderName(string inputName, int targetLength)
        //{
        //    string cleanedName = inputName.Trim(); //去除首尾空白字符
        //    //正则表达式匹配模式为：制表符“\/:*?<>|"”换行符回车符等1个及以上（不能用于文件名的字符）；将匹配到的字符串替换为下划线
        //    //在@字符串（逐字字符串字面量）中，双引号只能用双引号转义
        //    cleanedName = Regex.Replace(cleanedName, @"[\t\\/:\*\?\<\>\|""\n\r]+", "_");
        //    //正则表达式匹配模式为：空格2个及以上；将匹配到的字符串替换为一个空格
        //    cleanedName = Regex.Replace(cleanedName, @"[ ]{2,}", " ");
        //    cleanedName = cleanedName[..Math.Min(targetLength, cleanedName.Length)]; //截取目标字数
        //    return cleanedName;
        //}

        public static string CleanWorksheetName(string inputName, int targetLength)
        {
            string cleanedName = inputName.Trim(); //去除首尾空白字符
            //正则表达式匹配模式为：非中文字符、非英文字符、非阿拉伯数字或下划线、非空格等字符1个及以上；将匹配到的字符串替换为空
            cleanedName = Regex.Replace(cleanedName, @"[^\u4e00-\u9fa5\w| ]+", "");
            cleanedName = cleanedName[..Math.Min(targetLength, cleanedName.Length)]; //截取目标字数
            return cleanedName;
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
                    string chineseUnit = arrChineseUnits[n - i - 1]; //获取当前阿拉伯位数字对应的中文单位 （假设是个3位数，从左往右数当i到达第2位（1号）数字时，3-1-1=1，在中文数字单位数组中索引号为1的元素为“十”，依此类推）
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

        public static void FormatDocumentTable(ExcelWorkbook workbook)
        {
            foreach (ExcelWorksheet excelWorksheet in workbook.Worksheets) // 遍历所有Excel工作表
            {
                FormatExcelWorksheet(excelWorksheet, 1, 0); //设置Excel工作表格式

                //设置A-I列列宽（小标题级别、小标题编号、文字、完成时限、责任人、分类）
                excelWorksheet.Cells["A:B"].EntireColumn.Width = 12; //=.Columns[1,2]
                excelWorksheet.Cells["C"].EntireColumn.Width = 80;
                excelWorksheet.Cells["D:F"].EntireColumn.Width = 12;
                excelWorksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; //文字水平左对齐

                if (excelWorksheet.Dimension == null) //如果当前Excel工作表为空，则直接跳过当前循环并进入下一个循环
                {
                    continue;
                }

                if (excelWorksheet.Name == "Body") // 如果当前Excel工作表为“主体”工作表
                {
                    //将A、D、E、F列中所有为null或全空白字符的单元格赋值给空白单元格变量
                    IEnumerable<ExcelRangeBase> emptyCells = excelWorksheet.Cells["A:A,D:D,E:E,F:F"].Where(c => string.IsNullOrWhiteSpace(c.Text));
                    foreach (ExcelRangeBase emptyCell in emptyCells) //遍历所有空白单元格
                    {
                        emptyCell.Value = "-"; // 将当前单元格填充为"-"
                    }

                    // 填加数据验证
                    int lastRowIndex = Math.Max(6, excelWorksheet.Dimension.End.Row); // 获取已使用区域最末行的索引号，如果小于指定值，则将其限定到指定值
                    string rangeStr = "A2:A" + lastRowIndex; // 将A列第2行至最末行单元格区域地址赋值给区域地址字符串变量

                    //在工作表的数据验证集合的ExcelDataValidationList中查找作用区域地址字符串与指定区域地址字符串相同的数据验证列表，从中取出第一个数据验证，将其赋值给existingValidation变量
                    //第一个Address表示数据验证规则所应用的单元格区域地址，第二个Address表示前述单元格区域地址的字符串表达形式，如“A2:Axx”
                    ExcelDataValidationList? existingValidation = excelWorksheet.DataValidations.OfType<ExcelDataValidationList>()
                        .FirstOrDefault(v => v.Address.Address == rangeStr);
                    string[] arrValidations = new string[] { "Lv0", "Lv1", "Lv2", "Lv3", "Lv4", "Enum.", "Itm.", "Immed." }; //将数据验证项赋值给数据验证数组

                    if (existingValidation == null) // 如果不存在数据验证，则添加新的数据验证
                    {
                        IExcelDataValidationList? validation = excelWorksheet.DataValidations.AddListValidation(rangeStr);
                        // 添加数据验证规则
                        foreach (string item in arrValidations)
                        {
                            validation.Formula.Values.Add(item);
                        }
                    }
                    else //否则
                    {
                        // 修改数据验证规则
                        existingValidation.Formula.Values.Clear(); //删除现有数据验证规则
                        foreach (string item in arrValidations)
                        {
                            existingValidation.Formula.Values.Add(item);
                        }
                    }

                    // 纯标题行设置文字加粗
                    for (int i = 2; i <= excelWorksheet.Dimension.End.Row; i++) //遍历Excel工作表从第2行开始到末尾的所有行
                    {
                        int headingCharLimit = IsChineseText(excelWorksheet.Cells[i, 3].Text) ? 50 : 125; // 计算小标题字数上限：如果当前行文字为中文，则得到50；否则，得到125

                        //设置当前行1至3列字体加粗：如果当前行含小标题且文字字数少于小标题字数上限（纯小标题），则加粗；否则不加粗
                        excelWorksheet.Cells[i, 1, i, 3].Style.Font.Bold =
                            excelWorksheet.Cells[i, 1].Text.Contains("Lv") && excelWorksheet.Cells[i, 3].Text.Length < headingCharLimit ? true : false;
                    }
                }
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

            //设置表头格式、自动筛选
            if (headerRowCount >= 1) //如果表头行数大于等于1
            {
                ExcelRange headerRange = excelWorksheet.Cells[1, 1, headerRowCount, excelWorksheet.Dimension.End.Column]; //将表头区域赋值给表头区域变量

                // 设置表头区域字体、对齐
                headerRange.Style.Font.Name = appSettings.WorksheetFontName; // 获取应用程序设置中的字体名称
                headerRange.Style.Font.Size = (float)appSettings.WorksheetFontSize; // 获取应用程序设置中的字体大小
                headerRange.Style.Font.Bold = true; //表头区域字体加粗
                headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; //单元格内容水平居中对齐
                headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center; //单元格内容垂直居中对齐
                headerRange.Style.WrapText = true; //设置文字自动换行

                if (excelWorksheet.AutoFilter.Address == null) // 如果自动筛选区域为null（未开启自动筛选），则将表头最后一行的自动筛选设为true
                {
                    excelWorksheet.Cells[headerRowCount, 1, headerRowCount, excelWorksheet.Dimension.End.Column].AutoFilter = true;
                }

                for (int i = 1; i <= headerRowCount; i++) //遍历表头所有行
                {
                    ExcelRange headerRowCells = excelWorksheet.Cells[i, 1, i, excelWorksheet.Dimension.End.Column]; //将当前行所有单元格赋值给表头行单元格变量

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
            ExcelRange recordRange = excelWorksheet.Cells[headerRowCount + 1, 1, excelWorksheet.Dimension.End.Row - footerRowCount, excelWorksheet.Dimension.End.Column];

            // 设置记录区域字体、对齐
            recordRange.Style.Font.Name = appSettings.WorksheetFontName;
            recordRange.Style.Font.Size = (float)appSettings.WorksheetFontSize;
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
            int lastRefRowIndex = Math.Max(firstRefRowIndex, excelWorksheet.Dimension.End.Row - footerRowCount);

            for (int j = 1; j <= excelWorksheet.Dimension.End.Column; j++) //遍历所有列
            {
                if (!excelWorksheet.Columns[j].Hidden) //如果当前列不为隐藏列
                {
                    ExcelRange columnCells = excelWorksheet.Cells[firstRefRowIndex, j, lastRefRowIndex, j]; //将当前列起始参考行至最末参考行的单元格赋值给列单元格集合变量
                    //从列单元格集合变量中筛选出文本不为null或全空白字符且不是合并的单元格，赋值给合格单元格集合变量
                    IEnumerable<ExcelRangeBase> qualifiedCells = columnCells.Where(cell => !string.IsNullOrWhiteSpace(cell.Text) && !cell.Merge);
                    //计算当前列所有合格单元格的字符数平均值：如果合格单元格集合不为空，则得到所有单元格字符数的平均值，否则得到0
                    double averageCharactersCount = qualifiedCells.Any() ? qualifiedCells.Average(cell => cell.Text.Length) : 0;
                    excelWorksheet.Columns[j].Style.WrapText = false; //设置当前列文字自动换行为false
                    excelWorksheet.Columns[j].AutoFit(); //设置当前列自动调整列宽（在文字不自动换行时，能完整显示文字的最适合列宽）
                    excelWorksheet.Columns[j].Style.WrapText = true; //设置当前列文字自动换行
                    //在当前列最合适列宽、基于单元格字符数平均值计算出的列宽中取较小值（并限制在8-40的范围），赋值给列宽变量
                    double columnWidth = Math.Min(excelWorksheet.Columns[j].Width, averageCharactersCount * 2 + 4).Clamp<double>(8, 40);
                    excelWorksheet.Columns[j].Width = columnWidth; //设置当前列的列宽

                    fullWidth += columnWidth; //将当前列列宽累加至全表格宽度
                }
            }

            //设置记录区域行高
            for (int i = headerRowCount + 1; i <= excelWorksheet.Dimension.End.Row - footerRowCount; i++) //遍历除去表头、表尾的所有行
            {
                if (!excelWorksheet.Rows[i].Hidden)  // 如果当前行没有被隐藏，设置当前行“是否手动调整行高”为false（即为自动）
                {
                    excelWorksheet.Rows[i].CustomHeight = false;
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
            printerSettings.LeftMargin = (decimal)(1.2 / 2.54);
            printerSettings.RightMargin = (decimal)(1.2 / 2.54);
            printerSettings.TopMargin = (decimal)(1.2 / 2.54);
            printerSettings.BottomMargin = (decimal)(1.2 / 2.54);
            printerSettings.HeaderMargin = (decimal)(0.8 / 2.54);
            printerSettings.FooterMargin = (decimal)(0.8 / 2.54);

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
            view.FreezePanes(headerRowCount + 1, 2); // 冻结最上方的行和最左侧的列（参数指定第一个不要冻结的单元格）
            view.PageLayoutView = true; // 将工作表视图设置为页面布局视图
            printerSettings.FitToPage = true; // 启用适应页面的打印设置
            //int printPagesCount = Math.Max(1, (int)Math.Round(fullWidth / 120, 0)); //计算打印页面数：将全表格宽度除以指定最大宽度的商四舍五入取整，如果小于1，则限定为1
            //printerSettings.FitToWidth = printPagesCount;  // 设置缩放为几页宽，1代表即所有列都将打印在一页上
            printerSettings.FitToWidth = 1; // 设置缩放为几页宽，1代表即所有列都将打印在一页上
            printerSettings.FitToHeight = 0; // 设置缩放为几页高，0代表打印页数不受限制，可能会跨越多页
            printerSettings.PageOrder = ePageOrder.OverThenDown; // 将打印顺序设为“先行后列”
            view.PageLayoutView = false; // 将页面布局视图设为false（即普通视图）
        }

        public static string GetHeadingLevel(string heading, bool isChineseText)
        {
            // 定义各级小标题编号正则表达式变量
            // 中文0级小标题编号：从开头开始，空格制表符任意多个，“第”，空格制表符任意多个，阿拉伯数字或中文数字1个及以上，空格制表符任意多个，“部分、篇、章、节”，“：:”空格制表符至少一个
            Regex regExCnHeading0Num = new Regex(@"^[ |\t]*第[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*(?:部分|篇|章|节)[：:| |\t]+", RegexOptions.Multiline);
            // 中文1级小标题编号：从开头开始，空格制表符任意多个，中文数字1个及以上，空格制表符任意多个，“）)、.．，,”，空格制表符任意多个
            Regex regExCnHeading1Num = new Regex(@"^[ |\t]*[一二三四五六七八九十〇零]+[ |\t]*[）\)、\.．，,][ |\t]*", RegexOptions.Multiline);
            // 中文2级小标题编号：从开头开始，空格制表符任意多个，“（(”，空格制表符任意多个，中文数字1个及以上，空格制表符任意多个，“）)、.．，,”，空格制表符任意多个
            Regex regExCnHeading2Num = new Regex(@"^[ |\t]*[（\(][ |\t]*[一二三四五六七八九十〇零]+[ |\t]*[）\)、\.．，,][ |\t]*", RegexOptions.Multiline);
            // 中文3级小标题编号：从开头开始，空格制表符任意多个，阿拉伯数字1个及以上，空格制表符任意多个，“）)、.．，,”，空格制表符任意多个
            Regex regExCnHeading3Num = new Regex(@"^[ |\t]*\d+[ |\t]*[）\)、\.．，,][ |\t]*", RegexOptions.Multiline);
            // 中文4级小标题编号：从开头开始，空格制表符任意多个，“（(”，空格制表符任意多个，阿拉伯数字1个及以上，空格制表符任意多个，“）)、.．，,”，空格制表符任意多个
            Regex regExCnHeading4Num = new Regex(@"^[ |\t]*[（\(][ |\t]*\d+[ |\t]*[）\)、\.．，,][ |\t]*", RegexOptions.Multiline);
            // 中文“X是”编号：从开头开始，空格制表符任意多个，中文数字1个及以上，空格制表符任意多个，“是”，空格制表符任意多个
            Regex regExCnShiNum = new Regex(@"^[ |\t]*[一二三四五六七八九十〇零]+[ |\t]*是[ |\t]*", RegexOptions.Multiline);
            // 中文“第X条”编号：从开头开始，空格制表符任意多个，“第”，空格制表符任意多个，阿拉伯数字或中文数字1个及以上，空格制表符任意多个，“条”，“：:”空格制表符至少一个
            Regex regExCnItemNum = new Regex(@"^[ |\t]*第[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*条[：:| |\t]+", RegexOptions.Multiline);
            // 英文小标题编号，匹配模式为：从开头开始，空格制表符任意多个，“part、charpter、section”标记（至多一个，为捕获组1），【模式为"1./1.2./1.2.3./1.2.3.4."（最末尾可以省略句点），作为捕获组2】，空格制表符至少一个
            Regex regExEnHeadingNum = new Regex(@"^[ |\t]*((?:part|chapter|section)[ |\t]+)?((?:\d+\.?){1,4})[ |\t]+", RegexOptions.Multiline | RegexOptions.IgnoreCase);

            if (isChineseText) // 如果是中文文本
            {
                // 使用正则表达式来匹配小标题编号，并赋值给小标题级别单元格
                if (regExCnHeading0Num.IsMatch(heading)) //如果单元格文本被0级小标题编号正则表达式匹配成功，则将当前行的小标题级别（第1列）单元格赋值为“0级”
                {
                    return "Lv0";
                }
                else if (regExCnHeading1Num.IsMatch(heading))
                {
                    return "Lv1";
                }
                else if (regExCnHeading2Num.IsMatch(heading))
                {
                    return "Lv2";
                }
                else if (regExCnHeading3Num.IsMatch(heading))
                {
                    return "Lv3";
                }
                else if (regExCnHeading4Num.IsMatch(heading))
                {
                    return "Lv4";
                }
                else if (regExCnShiNum.IsMatch(heading))
                {
                    return "Enum.";
                }
                else if (regExCnItemNum.IsMatch(heading))
                {
                    return "Itm.";
                }
                else
                {
                    return "";
                }
            }

            else // 否则（不是中文文本）
            {
                Match matchEnHeadingNum = regExEnHeadingNum.Match(heading); // 获取英文小标题编号正则表达式匹配结果

                if (!matchEnHeadingNum.Success) // 如果英文小标题编号正则表达式匹配失败，则将""赋值给函数返回值
                {
                    return "";
                }

                // 计算英文小标题编号中含有几组数字
                int enHeadingNumsCount = Regex.Split(matchEnHeadingNum.Groups[2].Value, @"\.")
                  .Where(s => !string.IsNullOrWhiteSpace(s)) // 
                  .ToList().Count;

                if (matchEnHeadingNum.Groups[1].Success) // 如果英文小标题编号正则表达式捕获组1匹配成功（以“part、charpter、section”开头），则将"Lv0"赋值给函数返回值
                {
                    return "Lv0";
                }
                else // 否则
                {
                    return "Lv" + enHeadingNumsCount.ToString(); // 将"Lv"和小标题级别合并（编号中有几组数字就为几级标题）并赋值给函数返回值
                }
            }
        }

        public static string? GetKeyColumnLetter()
        {
            string latestColumnLetter = latestRecords.LatestKeyColumnLetter; //读取设置中保存的主键列符
            InputDialog inputDialog = new InputDialog(question: "Input the key column letter (e.g. \"A\"）", defaultAnswer: latestColumnLetter); //弹出对话框，输入主键列符
            if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则函数返回值赋值为null
            {
                return null;
            }
            string columnLetter = inputDialog.Answer;
            latestRecords.LatestKeyColumnLetter = columnLetter; // 将对话框返回的列符存入设置

            return columnLetter; //将列符赋值给函数返回值
        }


        public static (int headerRowCount, int footerRowCount) GetHeaderAndFooterRowCount()
        {
            try
            {
                string lastestHeaderFooterRowCountStr = latestRecords.LastestHeaderAndFooterRowCountStr; //读取设置中保存的表头表尾行数字符串
                InputDialog inputDialog = new InputDialog(question: "Input the row count of the table header and footer (separated by a comma, e.g. \"2,0\")", defaultAnswer: lastestHeaderFooterRowCountStr); //弹出对话框，输入表头表尾行数
                if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则表头、表尾行数均赋值为默认值，并结束本过程
                {
                    return (0, 0);
                }
                string headerFooterRowCountStr = inputDialog.Answer; //获取对话框返回的表头、表尾行数字符串
                latestRecords.LastestHeaderAndFooterRowCountStr = headerFooterRowCountStr; // 将对话框返回的表头、表尾行数字符串存入设置

                //将表头、表尾字符串拆分成数组，转换成列表，移除每个元素的首尾空白字符，转换成数值，如果小于0则限定为0，并赋值给表头表尾行数列表
                List<int> lstHeaderFooterRowCount = headerFooterRowCountStr.Split(',').ToList().ConvertAll(e => Convert.ToInt32(e.Trim())).ConvertAll(e => Math.Max(0, e));
                //将表头表尾行数列表0号、1号元素，赋值给函数返回值
                return (lstHeaderFooterRowCount[0], lstHeaderFooterRowCount[1]);
            }

            catch (Exception ex) // 捕获错误
            {
                ShowExceptionMessage(ex);
                return (0, 0); //表头、表尾行数变量赋值为0
            }
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
                            excelWorksheet = excelPackage.Workbook.Worksheets[index - 1]; //将指定索引号的Excel工作表赋值给Excel工作表变量（Excel工作表索引号从1开始，EPPlus从0开始）
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

            catch (Exception ex) // 捕获错误
            {
                ShowExceptionMessage(ex);
                return null; //函数返回值赋值为null
            }

        }

        public static DataTable RemoveDataTableEmptyRowsAndColumns(DataTable dataTable)
        {
            //清除空白数据行
            for (int i = dataTable.Rows.Count - 1; i >= 0; i--) // 遍历DataTable所有数据行
            {
                //// 统计当前数据行不为数据库空值且不为null或全空白字符的数据元素的数量
                //int nonNullCount = dataTable.Rows[i].ItemArray.Count(value =>
                //    value != DBNull.Value && !string.IsNullOrWhiteSpace(value?.ToString()));

                //// 如果以上数据元素的数量小于等于1（仅含有一个数据的数据行无意义），则删除这一行
                //if (nonNullCount <= 1)
                //{
                //    dataTable.Rows[i].Delete();
                //}

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

        public static string RemoveMarkDownMarks(this string inText)
        {
            string outText = inText;
            // 行首尾空白字符正则表达式匹配模式为：开头标记，不为非空白字符也不为换行符的字符（不为换行符的空白字符）至少一个/或前述字符至少一个，结尾标记；将匹配到的字符串替换为空
            //[^\S\n]+与(?:(?!\n)\s)+等同
            outText = Regex.Replace(outText, @"^[^\S\n]+|[^\S\n]+$", "", RegexOptions.Multiline);

            // 文档分隔线符号正则表达式匹配模式为：开头标记，“*-_”至少一个，结尾标记；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^[\*\-_]+$", "", RegexOptions.Multiline);
            // 表格表头分隔线符号正则表达式匹配模式为：开头标记，“|-:”至少一个，结尾标记；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^[\|\-:]+$", "", RegexOptions.Multiline);

            // 标题符号正则表达式匹配模式为：开头标记，“#”（同行标题标记）至少一个，空格任意多个/或开头标记，“=-”（上一行标题标记）至少一个，结尾标记；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^#+[ ]*|^[=\-]+$", "", RegexOptions.Multiline);
            // 斜体或粗体符号（1个代表斜体，2个代表粗体）正则表达式匹配模式为：开头标记或任意字符任意多个（尽可能少）（捕获组1），“*_”至少一个，任意字符任意多个（尽可能少）（捕获组2），“*_”至少一个，任意字符任意多个（尽可能少）或结尾标记（捕获组3）；将匹配到的字符串替换为3个捕获组合并后的字符串
            outText = Regex.Replace(outText, @"(^|.*?)[\*_]+(.*?)[\*_]+(.*?|$)", "$1$2$3", RegexOptions.Multiline);
            // 引用符号正则表达式匹配模式为：开头标记，“>”；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^>", "", RegexOptions.Multiline);
            // 无序列表符号正则表达式匹配模式为：开头标记，“*-”，空格任意多个；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^[\*-][ ]*", "", RegexOptions.Multiline);

            // 代码引用符号转义符号正则表达式匹配模式为：“`”至少2个，；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"`{2,}", "", RegexOptions.Multiline);
            // 代码引用符号正则表达式匹配模式为：开头标记或任意字符任意多个（尽可能少）（捕获组1），“`”，任意字符任意多个（尽可能少）（捕获组2），“`”，任意字符任意多个（尽可能少）或结尾标记（捕获组3）；将匹配到的字符串替换为3个捕获组用"隔开后的字符串
            outText = Regex.Replace(outText, @"(^|.*?)`(.*?)`(.*?|$)", "$1\"$2\"$3", RegexOptions.Multiline);
            // 删除线符号正则表达式匹配模式为：开头标记或任意字符任意多个（尽可能少）（捕获组1），“~~”，任意字符任意多个（尽可能少）（捕获组2），“~~”，任意字符任意多个（尽可能少）或结尾标记（捕获组3）；将匹配到的字符串替换为3个捕获组合并后的字符串
            outText = Regex.Replace(outText, @"(^|.*?)~~(.*?)~~(.*?|$)", "$1$2$3", RegexOptions.Multiline);

            // 表格行开头和结尾符号正则表达式匹配模式为：开头标记，“|”/或“|”，结尾标记；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"^\||\|$", "", RegexOptions.Multiline);
            // 表格内部多余空白字符正则表达式匹配模式为：前方出现“|”，不为换行符的空白字符至少一个/或前述字符至少一个，后方出现“|”；将匹配到的字符串替换为空
            outText = Regex.Replace(outText, @"(?<=\|)[^\S\n]+|[^\S\n]+(?=\|)", "", RegexOptions.Multiline);

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

        public static string RemoveHeadingNum(string inText)
        {
            // 定义英文标题编号正则表达式字符串：前方出现开头标记或“：:；;”，空格制表符任意多个，“part、charpter、section”标记至多一个，模式为"1./1.2./1.2.3./1.2.3.4."（不限长度，最末尾可以省略句点），空格制表符至少一个
            string enHeadingNumRegEx = @"(?<=^|[：:；;])[ |\t]*(?:(?:part|chapter|section)[ |\t]+)?(?:\d+\.?)*[ |\t]+";

            //定义小标题编号正则表达式字符串：前方出现开头标记或“。：:；;”，空格制表符任意多个，“第（(”至多一个， 空格制表符任意多个，阿拉伯数字或中文数字1个及以上，空格制表符任意多个，“部分、篇、章、节、条”，“：:”空格制表符至少一个/或“、\.，,）)是”，空格制表符任意多个
            string cnHeadingNumRegEx = @"(?<=^|[。：:；;])[ |\t]*[第（\(]?[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*(?:(?:部分|篇|章|节|条)[：:| |\t]+|[、\.，,）\)是])[ |\t]*";

            //定义小标题编号正则表达式变量，匹配模式为：英文小标题编号或中文小标题编号（先按英文小标题编号模式匹配，如果先按中文小标题编号模式匹配，会造成英文2级及以下小标题编号只匹配到第一节段数字，造成替换不全）
            Regex regExHeadingNum = new Regex($"(?:{enHeadingNumRegEx})|(?:{cnHeadingNumRegEx})", RegexOptions.Multiline | RegexOptions.IgnoreCase);
            return regExHeadingNum.Replace(inText, ""); //将输入文字中被小标题编号正则表达式匹配到的字符串替换为空，赋值给函数返回值
        }

        public static string RunCommandline(string arguments)
        {
            try
            {
                // 创建一个进程启动信息对象
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe", // 指定启动的程序为命令提示符
                    Arguments = $"/c \"{arguments}\"", // 指定参数（使用引号包裹命令，确保含空格的命令被正确解析）
                    RedirectStandardOutput = true, // 重定向输出流，以便获取命令的输出
                    RedirectStandardError = true, // 重定向错误流，以便获取命令的错误输出
                    UseShellExecute = false, // 不使用系统外壳程序启动进程
                    CreateNoWindow = true // 不创建新窗口
                };

                // 创建并启动进程
                using (Process process = Process.Start(startInfo)!)
                {
                    // 读取输出结果
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    process.WaitForExit();  // 等待进程退出

                    if (process.ExitCode != 0) // 如果进程退出时返回的代码为0
                    {
                        throw new Exception("Failed to get network status.");
                    }

                    return (output + "\n" + error);
                }

            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
                return "";
            }
        }


        public static List<string>? SelectFiles(FileType fileType, bool isMultiselect, string dialogTitle)
        {
            string filter = fileType switch //根据文件类型枚举，返回相应的文件类型和扩展名的过滤项
            {
                FileType.Excel => "Excel Files(*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files(*.*)|*.*",
                FileType.Word => "Word Files(*.docx;*.docm)|*.docx;*.docm|All Files(*.*)|*.*",
                FileType.WordAndExcel => "Word And Excel Files(*.docx;*.xlsx;*.docm;*.xlsm)|*.docx;*.xlsx;*.docm;*.xlsm|All Files(*.*)|*.*",
                FileType.Convertible => "Convertible Files(*.doc;*.xls;*.wps;*.et)|*.doc;*.xls;*.wps;*.et|All Files(*.*)|*.*",
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

        public static void ShowExceptionMessage(Exception ex)
        {
            MessageDialog messageDialog = new MessageDialog(ex.Message + "\n" + (ex.InnerException?.Message ?? ""));
            messageDialog.ShowDialog();
        }


        public static bool ShowMessage(string message)
        {
            MessageDialog messageDialog = new MessageDialog(message);
            return messageDialog.ShowDialog() ?? false; // 将对话框返回值（点击OK为true，点击Cancel为false)赋值给函数返回值（如果对话框返回null，则为false)
        }

        public static void ShowSuccessMessage()
        {
            MessageDialog messageDialog = new MessageDialog("Operation completed.");
            messageDialog.ShowDialog();
        }

        public static void ProcessParagraphsIntoDocumentTable(List<string>? lstParagraphs, string targetExcelFilePath)
        {
            try
            {
                if ((lstParagraphs?.Count ?? 0) == 0) //如果段落列表元素数为0（如果段落列表为null，则得到0），则结束本过程
                {
                    return;
                }

                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    // 建立目标工作簿和工作表，初始化表头
                    ExcelWorksheet titleWorksheet = excelPackage.Workbook.Worksheets.Add("Title");
                    ExcelWorksheet bodyTextsWorksheet = excelPackage.Workbook.Worksheets.Add("Body");

                    bool isChineseDocument = IsChineseText(lstParagraphs?[0] ?? ""); // 根据段落数组0号（第1个）元素即大标题判断是否为中文文档，并赋值给“是否为中文文档”变量

                    // 定义大标题工作表表头列表
                    List<object[]> lstTitleWorksheetHeader = new List<object[]> { new object[] { "Item", "Index", "Text" } };
                    titleWorksheet.Cells["A1:C1"].LoadFromArrays(lstTitleWorksheetHeader); // 将表头列表加载到大标题工作表

                    // 定义大标题工作表项目列表
                    List<object[]> lstTitleWorksheetItems =
                        new List<object[]>
                            {
                                new object[] { "Title" },
                                new object[] { "Signature" },
                                new object[] { "Date" }
                            };
                    titleWorksheet.Cells["A2:A4"].LoadFromArrays(lstTitleWorksheetItems); // 将项目列表加载到大标题工作表

                    // 定义主体工作表表头列表“Heading Level、Heading Index、Content、Deadline、Responsible Subject、Category”
                    List<object[]> lstBodyWorksheetHeading = new List<object[]> { new object[] { "Heading Level", "Heading Index", "Content", "Remark 1", "Remark 2", "Remark 3" } };
                    bodyTextsWorksheet.Cells["A1:F1"].LoadFromArrays(lstBodyWorksheetHeading);

                    // 将段落数组内容从1号（第2个）元素即正文第一段开始，赋值给“主体”工作表内容列的单元格
                    for (int i = 1; i < lstParagraphs!.Count; i++) //遍历数组所有元素
                    {
                        bodyTextsWorksheet.Cells[i + 1, 3].Value = lstParagraphs[i]; //将当前数组元素赋值给第3列的第i+1行的单元格
                    }

                    // 在“主体”工作表中，判断小标题正文文字的编号级别，赋值给小标题级别单元格，并将小标题正文文字的小标题编号清除，同时更新“小标题”工作表
                    for (int i = 2; i <= bodyTextsWorksheet.Dimension.End.Row; i++) //遍历从第2行开始往下的所有行
                    {
                        string cellText = bodyTextsWorksheet.Cells[i, 3].Text; //将当前行的小标题正文文字（第3列）单元格的文本赋值给单元格文本变量
                        bodyTextsWorksheet.Cells[i, 1].Value = GetHeadingLevel(cellText, isChineseDocument); //获取单元格文本的小标题级别，赋值给当前行的小标题级别单元格
                        bodyTextsWorksheet.Cells[i, 3].Value = RemoveHeadingNum(cellText); //删除单元格文本中的所有小标题编号，赋值给当前行的小标题正文文字单元格
                    }

                    // 在“大标题”工作表中，给大标题、签名、日期单元格赋值
                    titleWorksheet.Cells["C2"].Value = lstParagraphs[0]; // 将段落数组0号（第1个）元素即大标题值赋值给“大标题落款”工作表的“大标题”单元格

                    titleWorksheet.Cells["C3"].Value = isChineseDocument ? "签名" : "Signature"; // 给签名单元格赋值：如果输入文字是中文，则落款为“签名”；否则为“Signature”

                    // 给日期单元格赋值：如果输入文字是中文，则日期为当前日期的“yyyy年M月d日”格式；否则为“MMM-dd yyyy”美国格式
                    titleWorksheet.Cells["C4"].Value = isChineseDocument ? DateTime.Now.ToString("yyyy年M月d日") :
                        DateTime.Now.ToString("MMM-dd yyyy", CultureInfo.CreateSpecificCulture("en-US"));

                    TrimCellStrings(bodyTextsWorksheet); //删除“主体”Excel工作表内所有文本型单元格值的首尾空格
                    RemoveWorksheetEmptyRowsAndColumns(bodyTextsWorksheet); //删除“主体”Excel工作表内所有空白行和空白列

                    FormatDocumentTable(excelPackage.Workbook); //格式化文档表的所有工作表
                    excelPackage.SaveAs(new FileInfo(targetExcelFilePath)); // 保存目标工作簿
                }
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        //public static async Task FormatWordDocumentsAsync(List<string> filePaths)
        //{
        //    Task task = Task.Run(() => process()); // 创建一个异步任务，执行过程为process()
        //    void process()
        //    {
        //        MSWord.Application msWordApp = new MSWord.Application(); //打开Word应用程序并赋值给word应用程序变量
        //        msWordApp.ScreenUpdating = false; //关闭屏幕更新
        //        msWordApp.DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone; //关闭警告
        //        msWordApp.Visible = false; //“程序窗口可见”设为否
        //        MSWordDocument? msWordDocument = null; //定义Word文档变量

        //        try
        //        {
        //            foreach (string filePath in filePaths) //遍历文件路径全名列表所有元素
        //            {
        //                msWordDocument = msWordApp.Documents.Open(filePath); // 打开word文档并赋值给Word文档变量

        //                // 判断是否为空文档
        //                if (msWordDocument.Content.Text.Trim().Length <= 1) // 如果将Word换行符全部删除后，剩下的字符数小于等于1，则结束本过程
        //                {
        //                    return;
        //                }

        //                // 接受并停止修订
        //                msWordDocument.AcceptAllRevisions();
        //                msWordDocument.TrackRevisions = false;
        //                msWordDocument.ShowRevisions = false;

        //                string documentText = msWordDocument.Content.Text; // 全文文字变量赋值

        //                bool isCnDocument = IsChineseText(documentText); // 判断是否为中文文档，赋值给“是否为中文文档”变量

        //                // 根据是否为中文文档，设置Word文档正文、小标题、表格等的样式

        //                // 定义页边距、行距、字体、字号等的值

        //                double topMargin = msWordApp.CentimetersToPoints((float)3.7); // 顶端页边距
        //                double bottomMargin = msWordApp.CentimetersToPoints((float)3.5); // 底端页边距
        //                double leftMargin = msWordApp.CentimetersToPoints((float)2.8); // 左页边距
        //                double rightMargin = msWordApp.CentimetersToPoints((float)2.6); // 右页边距
        //                float lineSpace = (float)(isCnDocument ? appSettings.CnLineSpace : appSettings.EnLineSpace); // 行间距

        //                string titleFontName = isCnDocument ? appSettings.CnTitleFontName : appSettings.EnTitleFontName; // 大标题字体：如果为中文文档，则字体为相应中文字体；否则为英文字体
        //                string bodyFontName = isCnDocument ? appSettings.CnBodyFontName : appSettings.EnBodyFontName; // 正文字体

        //                string cnHeading0FontName = appSettings.CnHeading0FontName; // 中文0级小标题
        //                string cnHeading1FontName = appSettings.CnHeading1FontName; // 中文1级小标题
        //                string cnHeading2FontName = appSettings.CnHeading2FontName;  // 中文2级小标题
        //                string cnHeading3_4FontName = appSettings.CnHeading3_4FontName;  // 通用小标题
        //                string cnItemNumFontName = cnHeading1FontName; // 中文条款项小标题字体

        //                string enHeading0FontName = appSettings.EnHeading0FontName; // 英文0级小标题
        //                string enHeading1FontName = appSettings.EnHeading1FontName;  // 英文1级小标题
        //                string enHeading2FontName = appSettings.EnHeading2FontName; // 英文2级小标题
        //                string enHeading3_4FontName = appSettings.EnHeading3_4FontName; // 英文3-4小标题

        //                string tableTitleFontName = isCnDocument ? cnHeading1FontName : enHeading1FontName; // 表格标题字体
        //                string tableBodyFontName = bodyFontName; // 表格正文字体

        //                string footerFontName = "Times New Roman"; // 页脚字体


        //                float titleFontSize = (float)(isCnDocument ? appSettings.CnTitleFontSize : appSettings.EnTitleFontSize); // 大标题字号：如果为中文文档，则为相应的中文字号；否则，为英文字号
        //                float bodyFontSize = (float)(isCnDocument ? appSettings.CnBodyFontSize : appSettings.EnBodyFontSize); // 正文字号

        //                float cnHeading0FontSize = (float)appSettings.CnHeading0FontSize; // 中文0级小标题
        //                float cnHeading1FontSize = (float)appSettings.CnHeading1FontSize; // 中文1级小标题
        //                float cnHeading2FontSize = (float)appSettings.CnHeading2FontSize; // 中文2级小标题
        //                float cnHeading3_4FontSize = (float)appSettings.CnHeading3_4FontSize; // 中文3-4级小标题
        //                float cnItemNumFontSize = cnHeading1FontSize; // 中文条款项小标题字号

        //                float enHeading0FontSize = (float)appSettings.EnHeading0FontSize; // 英文0级小标题
        //                float enHeading1FontSize = (float)appSettings.EnHeading1FontSize; // 英文1级小标题
        //                float enHeading2FontSize = (float)appSettings.EnHeading2FontSize; // 英文2级小标题
        //                float enHeading3_4FontSize = (float)appSettings.EnHeading3_4FontSize; // 英文3-4级小标题

        //                float tableTitleFontSize = isCnDocument ? cnHeading1FontSize : enHeading1FontSize; // 表格标题字号
        //                float tableBodyFontSize = bodyFontSize - 2; // 表格正文字号

        //                float footerFontSize = 14; // 页脚字号为四号


        //                // 清除原文格式，替换空格、换行符等

        //                // 设置查找模式
        //                MSWord.Selection selection = msWordApp.Selection; //将选区赋值给选区变量
        //                MSWord.Find find = msWordApp.Selection.Find; //将选区查找赋值给查找变量

        //                find.ClearFormatting(); // 清除格式
        //                find.Wrap = WdFindWrap.wdFindStop; // 到文档结尾后停止查找
        //                find.Forward = true; // 正向查找
        //                find.MatchByte = false; // 区分全角半角为False
        //                find.MatchWildcards = false; // 使用通配符为False

        //                // 全文空格替换为半角空格，换行符替换为回车符
        //                selection.WholeStory();

        //                find.Text = " "; // 查找空格
        //                find.Replacement.Text = " "; // 将空格替换为半角空格
        //                find.Execute(Replace: WdReplace.wdReplaceAll);

        //                //find.Text = "\t"; // 查找制表符
        //                //find.Replacement.Text = "    "; // 将制表符替换为4个空格
        //                //find.Execute(Replace: WdReplace.wdReplaceAll);

        //                find.Text = "\v"; // 查找换行符（垂直制表符），^l"
        //                find.Replacement.Text = "\r"; // 将换行符（垂直制表符）替换为回车符
        //                find.Execute(Replace: WdReplace.wdReplaceAll);

        //                // 清除段首、段尾多余空格和制表符，段落自动编号转文本
        //                for (int i = msWordDocument.Paragraphs.Count; i >= 1; i--) // 从末尾往开头遍历所有段落
        //                {
        //                    MSWordParagraph paragraph = msWordDocument.Paragraphs[i];

        //                    //正则表达式匹配模式设为：前方出现开头标记、换行符回车符，空格或制表符；如果段落文字被匹配成功，则继续循环
        //                    while (Regex.IsMatch(paragraph.Range.Text, @"(?<=^|\n|\r)[ |\t]"))
        //                    {
        //                        paragraph.Range.Characters[1].Delete(); // 删除开头第一个字符
        //                    }

        //                    //正则表达式匹配模式设为：空格或制表符，后方出现换行符回车符、结尾标记；如果段落文字被匹配成功，则继续循环
        //                    while (Regex.IsMatch(paragraph.Range.Text, @"[ |\t](?=\n|\r|$)"))
        //                    {
        //                        paragraph.Range.Select();
        //                        selection.EndKey(WdUnits.wdLine); // 光标移动到段落结尾换行符之前
        //                        selection.TypeBackspace(); // 删除前一个字符
        //                    }

        //                    // 如果当前段落不在表格内，且含有自动编号
        //                    if (!paragraph.Range.Information[WdInformation.wdWithInTable] && !string.IsNullOrEmpty(paragraph.Range.ListFormat.ListString))
        //                    {
        //                        paragraph.Range.InsertBefore(paragraph.Range.ListFormat.ListString + (!isCnDocument ? " " : "")); // 在段落文字前添加自动编号（如果不是中文文档，在编号后再加上一个空格）
        //                    }
        //                }


        //                // 清除文首和文末的空白段
        //                while (msWordDocument.Paragraphs[1].Range.Text == "\r") // 如果第1段文字为回车符，则继续循环
        //                {
        //                    msWordDocument.Paragraphs[1].Range.Delete(); // 删除第1段
        //                }

        //                while (msWordDocument.Paragraphs[msWordDocument.Paragraphs.Count].Range.Text == "\r"
        //                    && msWordDocument.Paragraphs[msWordDocument.Paragraphs.Count - 1].Range.Text == "\r") // 如果最后一段和倒数第二段文字均为回车符，则继续循环
        //                {
        //                    msWordDocument.Paragraphs[msWordDocument.Paragraphs.Count].Range.Delete(); // 删除最后一段
        //                }


        //                // 对齐缩进
        //                selection.WholeStory();
        //                selection.ClearFormatting(); // 清除全部格式、样式
        //                MSWord.ParagraphFormat paragraphFormat = msWordApp.Selection.ParagraphFormat; //将选区段落格式赋值给段落格式变量
        //                paragraphFormat.Reset(); // 段落格式清除
        //                paragraphFormat.CharacterUnitFirstLineIndent = (isCnDocument ? 2 : 0); // 设置首行缩进：如果为中文文档，则缩进2个字符；否则为0个字符
        //                paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0pt
        //                paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify; // 对齐方式设为两端对齐
        //                //paragraphFormat.IndentFirstLineCharWidth((short)(isCnDocument ? 3 : 0)); // 设置首行缩进：如果为中文文档，则缩进3个字符；否则为0个字符


        //                // 全文格式初始化
        //                selection.WholeStory(); // 选择word所有文档

        //                MSWord.PageSetup pageSetup = selection.PageSetup; // 将选区页面设置赋值给页面设置变量
        //                pageSetup.PageWidth = msWordApp.CentimetersToPoints((float)21); // 页面宽度设为21cm
        //                pageSetup.PageHeight = msWordApp.CentimetersToPoints((float)29.7); // 页面高度设为29.7cm
        //                pageSetup.TopMargin = (float)topMargin; // 顶端边距设为预设值
        //                pageSetup.BottomMargin = (float)bottomMargin; // 底端边距设为预设值
        //                pageSetup.LeftMargin = (float)leftMargin; // 左边距设为预设值
        //                pageSetup.RightMargin = (float)rightMargin; // 右边距设为预设值

        //                selection.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight; // 突出显示文本取消

        //                MSWord.Paragraphs paragraphs = selection.Paragraphs; // 将选区段落赋值给段落变量
        //                                                                     // '.CharacterUnitFirstLineIndent = 2 '此参数优先级最高，一旦设定，需要再次设置一个绝对值相等的负值或者重置段落格式才能将其归零！
        //                paragraphs.AutoAdjustRightIndent = 0; // 不自动调整右缩进
        //                paragraphs.DisableLineHeightGrid = -1; //取消“如果定义了网格，则对齐到网格”
        //                paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly; // 行距设置为固定值
        //                                                                               // '.LineSpacingRule = wdLineSpace1pt5 '行距固定1.5
        //                paragraphs.LineSpacing = lineSpace; // 行距设为预设值
        //                paragraphs.SpaceBefore = msWordApp.CentimetersToPoints(0); // 段落前间距设为0
        //                paragraphs.SpaceAfter = msWordApp.CentimetersToPoints(0); // 段落后间距设为0

        //                // 基础正文字体设置
        //                MSWord.Font font = msWordApp.Selection.Font; //将选区字体赋值给字体变量
        //                font.Name = bodyFontName; // 正文字体设为预设值
        //                font.Size = bodyFontSize; // 正文字号设为预设值
        //                font.ColorIndex = WdColorIndex.wdBlack; // 颜色设为黑色
        //                font.Bold = 0; // “是否粗体”设为0
        //                font.Kerning = 0; // “为字体调整字符间距”值设为0
        //                font.DisableCharacterSpaceGrid = true;  //取消“如果定义了文档网格,则对齐到网格”，忽略字体的每行字符数

        //                documentText = msWordDocument.Content.Text; // 全文文字变量重赋值（前期对文档进行过处理，内容可能已经改变）

        //                // 文档大标题设置
        //                selection.HomeKey(WdUnits.wdStory);

        //                // 定义大标题正则表达式变量，匹配模式为：从开头开始，不含2个及以上连续的换行符回车符（允许不连续的换行符回车符）、不含“附件/录”、Appendix注释、非“。”分页符的字符1-120个，换行符回车符，后方出现：换行符回车符
        //                Regex regExTitle = new Regex(@"(?<=^|\n|\r)(?:(?![\n\r]{2,})(?!(?:附[ |\t]*[件录]|appendix)[^。\f\n\r]{0,3}[\n\r])[^。\f]){1,120}[\n\r](?=[\n\r])", RegexOptions.Multiline | RegexOptions.IgnoreCase);

        //                // 定义中文发往单位正则表达式变量，匹配模式为：从开头开始，换行符回车符（一个空行），不含“附件/录”注释、不含小标题编号、不含“如下：”、非“。：:；;”分页符换行符回车符的字符1个及以上，“：:”，换行符回车符
        //                Regex regExCnAddressee = new Regex(@"(?<=^|\n|\r)[\n\r](?:(?!附[ |\t]*[件录][^。\f\n\r]{0,3}[\n\r])(?![（\(]?[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*[、\.，,）\)])(?!如下[：:])[^。：:；;\f\n\r]){1,}[：:][\n\r]", RegexOptions.Multiline);

        //                int referencePageNum = 0; //参考页码赋值为0
        //                MatchCollection matchesTitles = regExTitle.Matches(documentText); // 获取全文文字经过大标题正则表达式匹配后的结果

        //                foreach (Match matchTitle in matchesTitles) // 遍历所有匹配到的大标题文字
        //                {
        //                    selection.HomeKey(WdUnits.wdStory);
        //                    find.Text = matchTitle.Value; // 查找大标题
        //                    find.Execute();
        //                    int pageNum = selection.Information[WdInformation.wdActiveEndPageNumber]; // 当前页码变量赋值
        //                    if (!selection.Information[WdInformation.wdWithInTable] && pageNum != referencePageNum) //如果当前大标题不在表格内，且与之前已确定的大标题不在同一页（一页最多一个大标题）
        //                    {
        //                        bool formatTitle = false; // “设置大标题格式”变量赋值为False
        //                        if (pageNum == 1) // 如果大标题候选文字在第一页
        //                        {
        //                            formatTitle = true; // “设置大标题格式”变量赋值为True
        //                        }
        //                        else // 否则
        //                        {
        //                            selection.MoveStart(WdUnits.wdLine, -5); // 将搜索到大标题候选文字选区向上扩展5行
        //                            if (selection.Text.Contains("\f")) // 如果选区内含有分页符，则候选文字判断为大标题，“设置大标题格式”变量赋值为True
        //                            {
        //                                formatTitle = true;
        //                            }
        //                            selection.MoveStart(WdUnits.wdLine, 5); // 选区起点复原
        //                        }

        //                        if (formatTitle) // 如果要设置大标题格式
        //                        {
        //                            paragraphFormat.CharacterUnitFirstLineIndent = 0;
        //                            paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0
        //                            paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 居中对齐
        //                            font.Name = titleFontName; // 设置字体为预设值
        //                            font.Size = titleFontSize; // 设置字号为预设值
        //                            font.Bold = (int)WdConstants.wdToggle; // 字体加粗
        //                            selection.EndKey(WdUnits.wdLine); // 光标移到选区的最后一个字（换行符之前）

        //                            // 中文发往单位设置
        //                            if (isCnDocument) // 如果是中文文档
        //                            {
        //                                selection.MoveDown(WdUnits.wdLine, 1, WdMovementType.wdMove); // 光标下移到下方一行
        //                                selection.Expand(WdUnits.wdLine); // 全选一行
        //                                selection.MoveEnd(WdUnits.wdLine, 5); // 选区向下扩大5行

        //                                MatchCollection matchesCnAddressees = regExCnAddressee.Matches(selection.Text); // 获取选区文字经过中文发往单位正则表达式匹配的结果
        //                                foreach (Match matchCnAddressee in matchesCnAddressees) // 遍历所有匹配到的中文发往单位文字结果
        //                                {
        //                                    find.Text = matchCnAddressee.Value; // 查找发往单位
        //                                    find.Execute(); // 执行查找

        //                                    if (!selection.Information[WdInformation.wdWithInTable]) // 如果找到的文字不在表格内
        //                                    {
        //                                        paragraphFormat.CharacterUnitFirstLineIndent = 0;
        //                                        paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 段落首行缩进为0
        //                                    }
        //                                    selection.Collapse(WdCollapseDirection.wdCollapseEnd); // 将选区折叠到末尾
        //                                }
        //                            }

        //                            referencePageNum = selection.Information[WdInformation.wdActiveEndPageNumber]; // 获取大标题所在页码并赋值给参考页码变量，为以后提供参考
        //                        }
        //                    }
        //                }

        //                if (isCnDocument) // 如果为中文文档
        //                {
        //                    // 中文0级（部分、篇、章、节）小标题设置
        //                    selection.HomeKey(WdUnits.wdStory);

        //                    // 定义中文0级小标题正则表达式变量，匹配模式为：从开头开始，“第”，空格制表符任意多个，阿拉伯数字中文数字1个及以上，空格制表符任意多个，“部分、篇、章、节”，非“。；;”分页符换行符回车符的字符0-40个，换行符回车符
        //                    Regex regExCnHeading0 = new Regex(@"(?<=^|\n|\r)第[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*(?:部分|篇|章|节)[^。；;\f\n\r]{0,40}[\n\r]", RegexOptions.Multiline);
        //                    MatchCollection matchesCnHeading0s = regExCnHeading0.Matches(documentText); // 获取全文文字经过中文0级小标题正则表达式匹配的结果

        //                    foreach (Match matchCnHeading0 in matchesCnHeading0s)
        //                    {
        //                        find.Text = matchCnHeading0.Value;
        //                        find.Execute();
        //                        if (paragraphs[1].Range.Sentences.Count == 1) // 如果中文小标题所在段落只有一句
        //                        {
        //                            paragraphs[1].OutlineLevel = WdOutlineLevel.wdOutlineLevel1; // 将当前中文小标题的大纲级别设为1级
        //                        }
        //                        paragraphFormat.CharacterUnitFirstLineIndent = 0;
        //                        paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进为0
        //                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 居中对齐
        //                        font.Name = cnHeading0FontName;
        //                        font.Size = cnHeading0FontSize;
        //                        font.Bold = 1;
        //                        selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        //                    }

        //                    // 中文1、2级小标题设置
        //                    selection.HomeKey(WdUnits.wdStory);

        //                    // 定义中文1、2级小标题正则表达式变量，匹配模式为：从开头开始，“（(”至多一个（捕获组），中文数字1个及以上，空格制表符任意多个，“、.，,）)”，非“。；;”分页符换行符回车符的字符1-40个，“。”换行符回车符
        //                    Regex regExCnHeading1_2 = new Regex(@"(?<=^|\n|\r)(（|\()?[ |\t]*[一二三四五六七八九十〇零]+[ |\t]*[、\.，,）\)][^。；;\f\n\r]{1,40}[。\n\r]", RegexOptions.Multiline);
        //                    MatchCollection matchesCnHeading1_2s = regExCnHeading1_2.Matches(documentText); // 获取全文文字经过中文1、2级小标题正则表达式匹配的结果

        //                    foreach (Match matchCnHeading1_2 in matchesCnHeading1_2s)
        //                    {
        //                        find.Text = matchCnHeading1_2.Value;
        //                        find.Execute();
        //                        if (paragraphs[1].Range.Sentences.Count == 1)
        //                        {
        //                            if (!matchCnHeading1_2.Groups[1].Success) //如果正则表达式匹配捕获组失败（ 编号开头不含“（” ）
        //                            {
        //                                paragraphs[1].OutlineLevel = WdOutlineLevel.wdOutlineLevel1; // 将当前中文小标题的大纲级别设为1级
        //                                font.Name = cnHeading1FontName;
        //                                font.Size = cnHeading1FontSize;
        //                            }
        //                            else // 否则
        //                            {
        //                                paragraphs[1].OutlineLevel = WdOutlineLevel.wdOutlineLevel2; // 将当前中文小标题的大纲级别设为2级
        //                                font.Name = cnHeading2FontName;
        //                                font.Size = cnHeading2FontSize;
        //                            }
        //                        }

        //                        font.Bold = 1;
        //                        selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        //                    }

        //                    // 中文3、4级小标题设置
        //                    selection.HomeKey(WdUnits.wdStory);

        //                    // 定义中文3、4级小标题正则表达式变量，匹配模式为：从开头开始，“（(”至多一个（捕获组），空格制表符任意多个，阿拉伯数字1个及以上，空格制表符任意多个，“、.，,）)”，非“。；;”分页符换行符回车符的字符1-40个，“。”换行符回车符
        //                    Regex regExCnHeading3_4 = new Regex(@"(?<=^|\n|\r)(（|\()?[ |\t]*\d+[ |\t]*[、\.，,）\)][^。；;\f\n\r]{1,40}[。\n\r]", RegexOptions.Multiline);
        //                    MatchCollection matchesCnHeading3_4s = regExCnHeading3_4.Matches(documentText); // 获取全文文字经过中文3、4级小标题正则表达式匹配的结果

        //                    foreach (Match matchCnHeading3_4 in matchesCnHeading3_4s)
        //                    {
        //                        find.Text = matchCnHeading3_4.Value;
        //                        find.Execute();

        //                        if (paragraphs[1].Range.Sentences.Count == 1)
        //                        {
        //                            paragraphs[1].OutlineLevel = !matchCnHeading3_4.Groups[1].Success ? WdOutlineLevel.wdOutlineLevel3 : WdOutlineLevel.wdOutlineLevel4; //设置小标题所在段落的大纲级别：如果正则表达式匹配捕获组失败（ 编号开头不含“（” ），则设为3级；否则，设为4级  
        //                        }

        //                        font.Name = cnHeading3_4FontName;
        //                        font.Size = cnHeading3_4FontSize;
        //                        font.Bold = 1;
        //                        selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        //                    }

        //                    // 中文“X是”编号设置
        //                    selection.HomeKey(WdUnits.wdStory);

        //                    // 定义中文“X是”编号正则表达式变量，匹配模式为：前方出现换行符回车符“。：:；;，,”，空格制表符任意多个，中文数字1个及以上，空格制表符任意多个，“是”；后方出现非分页符换行符回车符的字符1个及以上
        //                    Regex regExCnShiNum = new Regex(@"(?<=[\n\r。：:；;，,][ |\t]*)[一二三四五六七八九十〇零]+[ |\t]*是(?=[^\f\n\r]{1,})", RegexOptions.Multiline);
        //                    MatchCollection matchesCnShiNums = regExCnShiNum.Matches(documentText); // 获取全文文字经过“X是”编号正则表达式匹配的结果

        //                    foreach (Match matchCnShiNum in matchesCnShiNums)
        //                    {
        //                        find.Text = matchCnShiNum.Value;
        //                        find.Execute();
        //                        font.Bold = 1;
        //                        selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        //                    }

        //                    // 中文“条款项”编号设置
        //                    selection.HomeKey(WdUnits.wdStory);

        //                    // 定义中文“条款项”编号正则表达式变量，匹配模式为：从开头开始，“第”，空格制表符任意多个，阿拉伯数字或中文数字1个及以上，空格制表符任意多个，“条款项”，“：:”空格制表符
        //                    Regex regExCnItemNum = new Regex(@"(?<=^|\n|\r)第[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*[条款项][：:| |\t]", RegexOptions.Multiline); // 将正则匹配模式设为条款项编号
        //                    MatchCollection matchesCnItemNums = regExCnItemNum.Matches(documentText); // 获取全文文字经过条款项编号正则表达式匹配的结果

        //                    foreach (Match matchCnItemNum in matchesCnItemNums)
        //                    {
        //                        find.Text = matchCnItemNum.Value;
        //                        find.Execute();
        //                        font.Name = cnItemNumFontName;
        //                        font.Size = cnItemNumFontSize;
        //                        font.Bold = 1;
        //                        selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        //                    }

        //                    // 中文附件注释设置
        //                    selection.HomeKey(WdUnits.wdStory);

        //                    // 定义中文附件注释正则表达式变量，匹配模式为：从开头开始，“附”，空格制表符任意多个，“件录”，非“。”分页符换行符回车符的字符0-3个，换行符回车符
        //                    Regex regExCnAppendix = new Regex(@"(?<=^|\n|\r)附[ |\t]*[件录][^。\f\n\r]{0,3}[\n\r]", RegexOptions.Multiline);
        //                    MatchCollection matchesCnAppendixes = regExCnAppendix.Matches(documentText); // 获取全文文字经过附件注释正则表达式匹配的结果

        //                    foreach (Match matchCnAppendix in matchesCnAppendixes)
        //                    {
        //                        find.Text = matchCnAppendix.Value;
        //                        find.Execute();
        //                        if (selection.Information[WdInformation.wdWithInTable] == false) // 如果查找结果不在表格内
        //                        {
        //                            paragraphFormat.CharacterUnitFirstLineIndent = 0;
        //                            paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0
        //                            paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; // 左对齐
        //                        }
        //                        selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        //                    }
        //                }

        //                else // 否则（为英文文档）
        //                {
        //                    //设置英文小标题格式
        //                    selection.HomeKey(WdUnits.wdStory);

        //                    // 定义英文小标题正则表达式变量，匹配模式为：从开头开始，“part、charpter、section”标记（捕获组1），【模式为"1./1.2./1.2.3./1.2.3.4."（最末尾可以省略句点），作为捕获组2】，空格制表符至少一个，非“；;”分页符换行符回车符的字符1-10个（尽可能少匹配），英文字符，非“；;”分页符换行符回车符的字符1-100个，换行符回车符
        //                    Regex regExEnHeading = new Regex(@"(?<=^|\n|\r)((?:part|chapter|section)[ |\t]+)?((?:\d+\.?){1,4})[ |\t]+[^；;\f\n\r]{1,10}?[a-zA-Z][^；;\f\n\r]{1,100}[\n\r]", RegexOptions.Multiline | RegexOptions.IgnoreCase);
        //                    MatchCollection matchesEnHeadings = regExEnHeading.Matches(documentText); // 获取全文文字经过英文小标题正则表达式匹配的结果

        //                    foreach (Match matchEnHeading in matchesEnHeadings)
        //                    {
        //                        find.Text = matchEnHeading.Value;
        //                        find.Execute();

        //                        // 计算小标题编号中含有几组数字：使用正则表达式以"."分割小标题编号字符串（捕获组2），筛选出不为null或全空白字符的字符串，转换成列表，并统计列表元素个数
        //                        int enHeadingNumsCount = Regex.Split(matchEnHeading.Groups[2].Value, @"\.")
        //                            .Where(s => !string.IsNullOrWhiteSpace(s)) // 
        //                            .ToList().Count;

        //                        if (matchEnHeading.Groups[1].Success) // 如果英文小标题正则表达式捕获组1匹配成功（以“part、charpter、section”开头），则设置字体为0级标题
        //                        {
        //                            (font.Name, font.Size) = (enHeading0FontName, enHeading0FontSize);
        //                        }
        //                        else // 否则
        //                        {
        //                            // 根据小标题编号中的数字组数设置1-4级小标题字体（有几组数字，就为几级标题）
        //                            (font.Name, font.Size) = enHeadingNumsCount switch
        //                            {
        //                                1 => (enHeading1FontName, enHeading1FontSize),
        //                                2 => (enHeading2FontName, enHeading2FontSize),
        //                                3 => (enHeading3_4FontName, enHeading3_4FontSize),
        //                                _ => (enHeading3_4FontName, enHeading3_4FontSize),
        //                            };
        //                        }
        //                        font.Bold = 1;

        //                        if (paragraphs[1].Range.Text.Length <= 100) // 如果小标题所在段落的长度小于等于100个字符
        //                        {
        //                            paragraphs[1].OutlineLevel = (WdOutlineLevel)enHeadingNumsCount; // 根据小标题编号中数字的组数，设置当前英文小标题的大纲级别（有几组数字就为几级大纲）
        //                        }

        //                        selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        //                    }
        //                }

        //                //将前期被识别为小标题的数字编号清单恢复为正文文字格式

        //                // 定义清单数字编号正则表达式列表变量
        //                List<string> listNums;
        //                if (isCnDocument) // 如果为中文文档，正则表达式列表中的匹配模式为中文1-4级小标题编号
        //                {
        //                    //listNums = new List<string>() { @"[一二三四五六七八九十〇零]+[ |\t]*[、\.，,]" , @"[（\(][ |\t]*[一二三四五六七八九十〇零]+[ |\t]*[）\)]",
        //                    //@"[（\(]?[ |\t]*\d+[ |\t]*[、\.，,）\)]" };
        //                    listNums = new List<string>() { @"[一二三四五六七八九十〇零]+[ |\t]*[、\.，,]", @"[（\(][ |\t]*[一二三四五六七八九十〇零]+[ |\t]*[）\)]",
        //                            @"\d+[ |\t]*[、\.，,）\)]", @"[（\(][ |\t]*\d+[ |\t]*[、\.，,）\)]" };
        //                }
        //                else // 否则，正则表达式列表中的匹配模式为英文1-4级小标题编号
        //                {
        //                    listNums = new List<string>() { @"(?:\d+\.?){1,4}" };
        //                }

        //                foreach (string listNum in listNums)  //遍历清单数字编号正则表达式列表
        //                {
        //                    selection.HomeKey(WdUnits.wdStory);

        //                    // 定义数字编号清单块正则表达式变量，匹配模式为：（从开头开始，数字编号，非分页符换行符回车符的字符至少一个，换行符回车符），以上字符串（捕获组）2个及以上
        //                    Regex regExListBlock = new Regex(@"((?<=^|\n|\r)" + listNum + @"[^\f\n\r]+[\n\r]){2,}", RegexOptions.Multiline);

        //                    MatchCollection matchesListBlocks = regExListBlock.Matches(documentText); // 获取全文文字经过数字编号清单块正则表达式匹配的结果

        //                    foreach (Match matchListBlock in matchesListBlocks) // 遍历数字编号清单块正则表达式匹配结果集合
        //                    {
        //                        //如果数字编号清单块正则表达式匹配到的字符串长度/捕获组匹配数的商（即每个条目的平均字数）大于等于指定数值（中文文档100，英文文档250），则不视为清单条目，直接跳过当前循环并进入下一个循环
        //                        if (matchListBlock.Value.Length / (matchListBlock.Groups[1].Captures.Count) >= (isCnDocument ? 100 : 250))
        //                        {
        //                            continue;
        //                        }

        //                        // 文本片段正则表达式匹配模式设为：含换行符回车符的任意字符的字符1-255个；获取当前数字编号清单块字符串经匹配后的第一个结果（截取前部最多255个字符，避免超出Interop库Find方法的限制）
        //                        Match matchTextSection = Regex.Match(matchListBlock.Value, @"(?:.|[\n\r]){1,255}", RegexOptions.Multiline);

        //                        find.Text = matchTextSection.Value;
        //                        find.Execute();

        //                        selection.MoveEnd(WdUnits.wdCharacter, matchListBlock.Value.Length - matchTextSection.Value.Length); //将搜索结果选区的末尾向后扩展至数字编号清单块的末尾
        //                        paragraphs.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText; // 将选区所在段落的大纲级别设为正文级别

        //                        //将选区设为正文文字格式
        //                        font.Name = bodyFontName;
        //                        font.Size = bodyFontSize;
        //                        font.Bold = 0;
        //                        selection.Collapse(WdCollapseDirection.wdCollapseEnd);

        //                    }

        //                }

        //                // 设置表格格式

        //                // 遍历所有表格
        //                foreach (MSWordTable table in msWordDocument.Tables)
        //                {
        //                    // 表格上方标题、注释设置
        //                    table.Cell(1, 1).Select(); // 选择第1行第1列的单元格
        //                    selection.MoveUp(WdUnits.wdLine, 1, WdMovementType.wdMove); // 光标上移到表格上方一行
        //                    selection.Expand(WdUnits.wdLine); // 全选表格上方一行
        //                    selection.MoveStart(WdUnits.wdLine, -5); // 选区向上扩大5行

        //                    // 定义表格上方标题正则表达式变量，匹配模式为：从开头开始，非“。；;”分页符换行符回车符的字符1-60个，“表单册录执”或“form,table”，中文数字阿拉伯数字“（）()”0-6个，换行符回车符
        //                    Regex regExTableTitle = new Regex(@"(?<=^|\n|\r)[^。；;\f\n\r]{1,60}(?:[表单册录执]|form|table)[ |\t]*[一二三四五六七八九十零〇\d（）\(\)]{0,6}[\n\r]", RegexOptions.Multiline | RegexOptions.IgnoreCase);

        //                    MatchCollection matchesTableTitles = regExTableTitle.Matches(selection.Text); // 获取选区文字经过表格上方标题正则表达式匹配的结果

        //                    if (matchesTableTitles.Count > 0) // 如果匹配到的结果集合元素数大于0
        //                    {
        //                        find.Text = matchesTableTitles[0].Value;
        //                        find.Execute();
        //                        paragraphFormat.CharacterUnitFirstLineIndent = 0;
        //                        paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0);
        //                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //                        font.Name = tableTitleFontName;
        //                        font.Size = tableTitleFontSize;
        //                        font.Bold = 1;
        //                        selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        //                    }

        //                    // 表格设置
        //                    table.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic; // 前景色设为自动
        //                    table.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic; // 背景色设为自动
        //                    table.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight; // 高亮色设为无高亮
        //                    table.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight; // 高亮色设为无高亮

        //                    // 单元格边距
        //                    table.TopPadding = msWordApp.PixelsToPoints(0, true); // 上边距为0
        //                    table.BottomPadding = msWordApp.PixelsToPoints(0, true); // 下边距为0
        //                    table.LeftPadding = msWordApp.PixelsToPoints(0, true); // 左边距为0
        //                    table.RightPadding = msWordApp.PixelsToPoints(0, true); // 右边距为0
        //                    table.Spacing = msWordApp.PixelsToPoints(0, true); // 单元格间距为0
        //                    table.AllowPageBreaks = true; // 允许表格断页
        //                    table.AllowAutoFit = true; // 允许自动重调尺寸

        //                    // 设置边框：内外单线条，0.5磅粗
        //                    //table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle; // 内部线条样式为单线条
        //                    //table.Borders.InsideLineWidth = WdLineWidth.wdLineWidth050pt; // 内部线条粗细为0.5磅
        //                    //table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle; // 外部线条样式为单线条
        //                    //table.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth050pt; // 外部线条粗细为0.5磅

        //                    // 设置行格式
        //                    table.Rows.WrapAroundText = 0; // 取消文字环绕
        //                    table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter; // 表水平居中
        //                    table.Rows.AllowBreakAcrossPages = -1; // 允许行断页
        //                    table.Rows.HeightRule = WdRowHeightRule.wdRowHeightAuto; // 行高设为自动
        //                    table.Rows.LeftIndent = msWordApp.CentimetersToPoints(0); // 左面缩进量为0

        //                    // 设置字体、段落格式
        //                    table.Range.Font.Name = tableBodyFontName; // 字体为仿宋
        //                    table.Range.Font.Color = WdColor.wdColorAutomatic; // 字体颜色设为自动
        //                    table.Range.Font.Size = tableBodyFontSize; // 字号为四号
        //                    table.Range.Font.Kerning = 0; // “为字体调整字符间距”值设为0
        //                    table.Range.Font.DisableCharacterSpaceGrid = true;

        //                    table.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
        //                    table.Range.ParagraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0);
        //                    table.Range.ParagraphFormat.AutoAdjustRightIndent = 0; // 自动调整右缩进为false
        //                    //table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 单元格内容水平居中
        //                    table.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle; // 单倍行距

        //                    table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter; // 单元格内容垂直居中

        //                    // 自动调整表格
        //                    table.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthAuto; // 列宽度设为自动
        //                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent); // 根据内容调整表格
        //                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow); // 根据窗口调整表格
        //                }

        //                // 括号注释设置
        //                selection.HomeKey(WdUnits.wdStory);

        //                // 定义括号注释正则表达式变量，匹配模式为：从开头开始，“（(”，非“（）()。”分页符换行符回车符的字符1-40个，“）)”，换行符回车符
        //                Regex regExBracket = new Regex(@"(?<=^|\n|\r)[（\(][^（）\(\)。\f\n\r]{1,40}[）\)][\n\r]", RegexOptions.Multiline);
        //                MatchCollection matchesBrakets = regExBracket.Matches(documentText); // 获取全文文字经过括号注释正则表达式匹配的结果

        //                foreach (Match matchBraket in matchesBrakets)
        //                {
        //                    find.Text = matchBraket.Value;
        //                    find.Execute();
        //                    if (selection.Information[WdInformation.wdWithInTable] == false) // 如果查找结果不在表格内
        //                    {
        //                        paragraphFormat.CharacterUnitFirstLineIndent = 0;
        //                        paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0
        //                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 居中对齐
        //                    }
        //                    selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        //                }


        //                // 落款设置
        //                selection.HomeKey(WdUnits.wdStory);

        //                // 定义中、英文落款字符串变量，匹配模式为：签名至少1个，最后为日期
        //                string cnSignOff = @"(?:[\u4e00-\u9fa5][^。：:；;，,\f\n\r]{1,}[\n\r])+[12]\d{3}[ |\t]*年[月日期\d：:\.\-/| |\t]{0,10}[\n\r]";
        //                string enSignOff = @"(?:[a-zA-Z][^：:；;，,\f\n\r]{1,}[\n\r])+[a-zA-Z\d：:，,\.\-/| |\t]{0,20}[12]\d{3}[\n\r]";
        //                // 定义中、英文落款正则表达式变量，匹配模式为：前方出现开头符号、换行符回车符，换行符回车符（一个空行），中文落款或英文落款
        //                Regex regExSignOff = new Regex(@$"(?<=^|\n|\r)[\n\r](?:{cnSignOff}|{enSignOff})", RegexOptions.Multiline | RegexOptions.IgnoreCase);
        //                MatchCollection matchesSignOffs = regExSignOff.Matches(documentText); // 获取全文文字经过签名和日期落款正则表达式匹配的结果

        //                foreach (Match matchSignOff in matchesSignOffs)
        //                {
        //                    find.Text = matchSignOff.Value;
        //                    find.Execute();
        //                    if (selection.Information[WdInformation.wdWithInTable] == false) // 如果查找结果不在表格内
        //                    {
        //                        foreach (MSWordParagraph paragraph in selection.Paragraphs) // 遍历所有落款中的段落
        //                        {
        //                            float rightIndentation = Math.Max(0, 10 - paragraph.Range.Text.Length / 2); // 计算右缩进量，如果小于0，则限定为0
        //                            paragraph.Format.CharacterUnitRightIndent = rightIndentation; // 右缩进设为之前计算值
        //                            paragraph.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight; // 右对齐
        //                        }
        //                    }
        //                    selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        //                }


        //                // 页脚页码设置
        //                foreach (MSWordSection section in msWordDocument.Sections) // 遍历所有节
        //                {
        //                    section.PageSetup.DifferentFirstPageHeaderFooter = 0;     // “首页页眉页脚不同”设为否
        //                    section.PageSetup.OddAndEvenPagesHeaderFooter = 0;        // “奇偶页页眉页脚不同”设为否

        //                    section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete(); // 删除页脚中的内容
        //                    // 设置页码
        //                    section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleNumberInDash;  // 页码左右带横线； wdPageNumberStyleArabicFullWidth 阿拉伯数字全宽
        //                    section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;  // 不续前节
        //                    section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = 1;  // 从1开始编号
        //                    section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.Add(WdPageNumberAlignment.wdAlignPageNumberOutside, FirstPage: true); // 页码奇数页靠右，偶数页靠左； wdAlignPageNumberInside  奇左偶右 wdAlignPageNumberCenter 页码居中
        //                    section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name = footerFontName;
        //                    section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Size = footerFontSize;

        //                    section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete(); // 删除页眉中的内容
        //                    section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone; // 段落下边框线设为无
        //                }

        //                msWordDocument.Save(); // 保存Word文档
        //                msWordDocument.Close(); // 关闭Word文档

        //            }

        //        }

        //        catch (Exception ex)
        //        {
        //            ShowExceptionMessage(ex);
        //        }

        //        finally
        //        {
        //            msWordApp.ScreenUpdating = true;
        //            if (msWordDocument != null) Marshal.ReleaseComObject(msWordDocument); // 释放Word文档对象
        //            KillOfficeApps(new object[] { msWordApp }); // 关闭Word应用程序进程
        //        }

        //    }

        //    await task;
        //}

        public static void ConvertDocumentByPandoc(string fromType, string toType, string fromFilePath, string toFilePath)
        {
            try
            {
                string? pandocPath = appSettings.PandocPath; //读取设置中保存的Pandoc程序文件路径全名，赋值给Pandoc程序文件路径全名变量

                ProcessStartInfo startInfo = new ProcessStartInfo //创建ProcessStartInfo对象，包含了启动新进程所需的信息，赋值给启动进程信息变量
                {
                    FileName = pandocPath, // 指定pandoc应用程序的文件路径全名
                                           //指定参数：-s完整独立文件，-f原格式 -t目标格式 -o输出文件路径全名，\"用于确保文件路径（可能包含空格）被视为pandoc命令的单个参数
                    Arguments = $"-s -f {fromType} -t {toType} \"{fromFilePath}\" -o \"{toFilePath}\"",
                    RedirectStandardOutput = true, //设定将外部程序的标准输出重定向到C#程序
                    UseShellExecute = false, //设定使用操作系统shell执行程序为false
                    CreateNoWindow = true, //设定不创建窗口
                };

                //启动新进程
                using (Process process = Process.Start(startInfo)!)
                {
                    process.WaitForExit(); //等待进程结束
                    if (process.ExitCode != 0) //如果进程退出时返回的代码不为0，则抛出异常
                    {
                        throw new Exception("Conversion failed.");
                    }
                }
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }


        public static bool IsChineseText(string inText)
        {
            //判断是否为中文文档
            if (inText.Length == 0) return false; // 如果全文长度为0，则将false赋值给函数返回值

            int nonCnCharCount = Regex.Matches(inText, @"[^\u4e00-\u9fa5]").Count; //获取全文非中文字符数量
            //int nonCnCharsCount = Regex.Matches(inText, @"[a-zA-Z| ]").Count; //获取全文非中文字符数量
            double nonCnCharsRatio = nonCnCharCount / inText.Length; // 计算非中文字符占全文的比例
            return nonCnCharsRatio < 0.5 ? true : false; //赋值给函数返回值：如果非中文字符比例小于0.5，得到true；否则，得到false
        }

        public static bool PreprocessDocumentTexts(ExcelRange range)
        {
            bool contentsChanged = false; // “内容是否改变”变量赋值为false

            foreach (ExcelRangeBase cell in range) // 遍历所有单元格
            {
                if (!cell.EntireRow.Hidden) // 如果当前单元格所在行不是隐藏行
                {
                    //将当前单元格文字按换行符拆分为数组（删除每个元素前后空白字符，并删除空白元素），转换成列表，赋值给拆分后文字列表
                    List<string>? lstSplitTexts = cell.Text.Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .ToList();

                    int lstSplitTextCount = lstSplitTexts!.Count; //获取拆分后文字列表元素个数

                    contentsChanged = lstSplitTextCount > 1 ? true : contentsChanged; // 获取“内容是否改变”变量值：如果拆分后文字列表元素个数大于1，得到true；否则，得到原值

                    for (int i = 0; i < lstSplitTextCount; i++) //遍历拆分后文字列表的所有元素
                    {
                        //将拆分后文字列表当前元素的文字按修订标记字符'^'拆分成数组（删除每个元素前后空白字符，并删除空白元素），转换成列表，移除每个元素的小标题编号，赋值给修订文字列表
                        List<string> lstTextsToRevise = lstSplitTexts[i].Split('^', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                            .ToList().ConvertAll(e => RemoveHeadingNum(e));

                        contentsChanged = lstTextsToRevise.Count > 1 ? true : contentsChanged; // 获取“内容是否改变”变量值：如果修订文字列表元素个数大于1，得到true；否则，得到原值

                        //合并修订文字列表中的所有元素成为完整字符串，重新赋值给拆分后文字列表当前元素
                        lstSplitTexts[i] = MergeRevision(lstTextsToRevise);

                        string MergeRevision(List<string> lstTextsToRevise) //合并修订文字
                        {
                            if ((lstTextsToRevise?.Count ?? 0) == 0) //如果修订文字列表的元素数（如果字符串列表为null，则得到0）为0，则将空字符串赋值给函数返回值
                            {
                                return string.Empty;
                            }

                            if (lstTextsToRevise!.Count == 1) //如果修订文字列表的元素数为1，则将0号元素赋值给函数返回值
                            {
                                return lstTextsToRevise[0];
                            }

                            // 以0号元素中所有的中文句子为基准，逐句比较其他元素中的重复句

                            //定义中文句子正则表达式变量，匹配模式为：非“。；;”字符任意多个，“。；;”
                            Regex regExCnSentence = new Regex(@"[^。；;]*[。；;]");

                            // 获取修订文字列表0号元素经过中文句子正则表达式匹配后的结果集合
                            MatchCollection matchesSentences = regExCnSentence.Matches(lstTextsToRevise[0]);

                            foreach (Match matchSentence in matchesSentences) //遍历所有中文句子正则表达式匹配的结果
                            {
                                int sameSentenceCount = 0;
                                for (int i = 1; i < lstTextsToRevise.Count; i++) //遍历修订文字列表从1号（第2个）元素开始的所有元素
                                {
                                    if (lstTextsToRevise[i].Contains(matchSentence.Value))  //如果修订文字列表当前元素含有当前中文句子（基准句）
                                    {
                                        lstTextsToRevise[i] = lstTextsToRevise[i].Replace(matchSentence.Value, ""); //将修订文字列表当前元素中的当前基准句替换为空（删除重复句）
                                        sameSentenceCount += 1; //相同中文句子计数加1
                                    }
                                }

                                //重新赋值给修订文字列表0号元素：如果相同中文句子计数小于修订文字列表元素数量减1（除0号元素外的其他元素并不都含有当前基准句），则得到将0号元素中的当前基准句替换为空后的字符串（删除非共有句）；否则得到0号元素原值
                                lstTextsToRevise[0] = sameSentenceCount < lstTextsToRevise.Count - 1 ? lstTextsToRevise[0].Replace(matchSentence.Value, "") : lstTextsToRevise[0];
                            }
                            return string.Join("", lstTextsToRevise);  //合并修订文字列表的所有元素，赋值给函数返回值
                        }

                    }

                    if (lstSplitTextCount >= 2) // 如果拆分后文字列表的元素个数大于等于2个
                    {
                        int insertedRowsCount = lstSplitTextCount - 1; // 计算需要插入的行数：拆分后文字列表元素数-1
                        cell.Worksheet.InsertRow(cell.Start.Row + 1, insertedRowsCount); // 从被拆分单元格的下一个单元格开始，插入行
                    }

                    for (int i = 0; i < lstSplitTextCount; i++) //遍历拆分后文字列表的每个元素
                    {
                        cell.Offset(i, 0).Value = lstSplitTexts[i]; //将拆分后文字列表当前元素赋值给当前单元格向下偏移i行的单元格
                        cell.CopyStyles(cell.Offset(i, 0)); //将当前单元格的样式复制到当前单元格向下偏移i行的单元格
                        cell.Offset(i, 0).EntireRow.CustomHeight = false; // 当前单元格向下偏移i行的单元格所在行的手动设置行高设为false（即为自动）   
                    }
                }
            }
            return contentsChanged; // 将“内容是否改变”变量值赋值给函数返回值
        }

        public static async Task ProcessDocumentTableIntoWordAsync(string documentTableFilePath, string targetWordFilePath)
        {
            try
            {
                List<string> lstFullTexts = new List<string>(); //定义全文本列表变量

                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(documentTableFilePath))) //打开结构化文档表Excel工作簿，赋值给Excel包变量
                {
                    ExcelWorksheet titleWorksheet = excelPackage.Workbook.Worksheets["Title"]; //将“大标题”工作表赋值给大标题工作表变量
                    ExcelWorksheet bodyTextsWorksheet = excelPackage.Workbook.Worksheets["Body"]; //将“主体”工作表赋值给“主体”工作表变量
                    RemoveWorksheetEmptyRowsAndColumns(bodyTextsWorksheet); //删除“主体”工作表内所有空白行和空白列
                    if ((bodyTextsWorksheet.Dimension?.Rows ?? 0) <= 1) // 如果“主体”工作表已使用行数小于等于1（如果工作表为空，则为0），只有表头无有效数据，则结束本过程
                    {
                        return;
                    }

                    //在“主体”工作表第2行到最末行（如果工作表为空，则为第2行）的文字（第3）列中，将含有换行符的单元格文字拆分成多段，删除小标题编号，合并修订文字，最后将各段分置于单独的行中。如果此过程中内容已改变，将true赋值给“内容是否改变”变量；否则赋值为false。
                    bool contentChanged = PreprocessDocumentTexts(bodyTextsWorksheet.Cells[2, 3, (bodyTextsWorksheet.Dimension?.End.Row ?? 2), 3]);

                    if (contentChanged) //如果内容已改变，则保存Excel工作簿，并抛出异常
                    {
                        excelPackage.Save();
                        throw new Exception("Document contents changed. Check and re-run.");
                    }

                    //将下方无正文的小标题行设为隐藏：
                    for (int i = 2; i <= bodyTextsWorksheet.Dimension!.End.Row; i++) //遍历“主体”工作表从第2行到最末行的所有行
                    {
                        if (!bodyTextsWorksheet.Rows[i].Hidden) //如果当前行不是隐藏行
                        {
                            int paragraphsCount = 0;
                            if (bodyTextsWorksheet.Cells[i, 1].Text.Contains("Lv") && bodyTextsWorksheet.Cells[i, 3].Text.Length < 50) //如果当前行文字含小标题且字数小于50字（纯小标题行，基准小标题行）
                            {
                                if (i == bodyTextsWorksheet.Dimension.Rows)  //如果当前行（基准小标题行）为最后一行
                                {
                                    bodyTextsWorksheet.Rows[i].Hidden = true; //将当前行（基准小标题行）隐藏
                                }
                                else //否则
                                {
                                    for (int k = i + 1; k <= bodyTextsWorksheet.Dimension.End.Row; k++)  //遍历从基准小标题行的下一行开始直到最后一行的所有行（比较行）
                                    {
                                        if (!bodyTextsWorksheet.Rows[k].Hidden)  //如果当前比较行不是隐藏行
                                        {
                                            //如果当前比较行文字含小标题且小标题级别数小于等于基准小标题行（小标题级别更高或相同），则退出循环
                                            if (bodyTextsWorksheet.Cells[k, 1].Text.Contains("Lv") && Val(bodyTextsWorksheet.Cells[k, 1].Text) <= Val(bodyTextsWorksheet.Cells[i, 1].Text))
                                            {
                                                break;
                                            }
                                            //否则，如果当前比较行文字不含小标题或者字数大于等于50（视为正文），则正文段落计数加1
                                            else if (!bodyTextsWorksheet.Cells[k, 1].Text.Contains("Lv") || bodyTextsWorksheet.Cells[k, 3].Text.Length >= 50)
                                            {
                                                paragraphsCount++;
                                            }
                                        }
                                    }
                                    if (paragraphsCount == 0)
                                    {
                                        bodyTextsWorksheet.Rows[i].Hidden = true; //如果累计的正文段落数为零（基准小标题下方无正文），则将基准小标题行隐藏
                                    }
                                }

                            }
                        }
                    }

                    bool isChineseDocument = IsChineseText(titleWorksheet.Cells["C2"].Text); // 根据大标题工作表中C2单元格文字即大标题文字，判断文档是否为中文文档，赋值给“是否为中文文档”变量

                    //初始化小标题编号变量
                    int heading0Num = 0;
                    int heading1Num = 0;
                    int heading2Num = 0;
                    int heading3Num = 0;
                    int heading4Num = 0;
                    int headingShiNum = 0;
                    int headingItemNum = 0;

                    bodyTextsWorksheet.Cells[2, 2, bodyTextsWorksheet.Dimension.End.Row, 2].Clear(); // 清除第2列旧小标题编号

                    for (int i = 2; i <= bodyTextsWorksheet.Dimension.End.Row; i++) //遍历“主体”工作表第2行开始到最末行的所有行
                    {
                        if (bodyTextsWorksheet.Rows[i].Hidden) //如果当前行是隐藏行
                        {
                            bodyTextsWorksheet.Cells[i, 2].Value = "X"; //将当前行小标题编号单元格赋值为“X”（忽略行）
                        }
                        else //否则
                        {
                            // 给小标题编号
                            bool checkHeadingNecessity = false; // “检查小标题编号必要性”变量初始赋值为False
                            switch (bodyTextsWorksheet.Cells[i, 1].Text) //根据当前行小标题级别进入相应的分支，将对应级别的小标题编号分别赋值给小标题编号单元格
                            {
                                case "Lv0": //如果为0级小标题
                                    heading0Num++; //0级小标题计数加1
                                    heading1Num = 0;
                                    heading2Num = 0;
                                    heading3Num = 0;
                                    heading4Num = 0;
                                    headingShiNum = 0;
                                    bodyTextsWorksheet.Cells[i, 2].Value = isChineseDocument ?
                                        "第" + ConvertArabicNumberIntoChinese(Convert.ToInt32(heading0Num)) + "部分 "
                                        : "Part " + heading0Num + " "; //将0级小标题编号赋值给小标题编号单元格
                                    checkHeadingNecessity = heading0Num == 1 ? true : false; // 获取“检查小标题编号必要性”值：如果编号为1，则得到true；否则，得到false（防止同级编号只有1没有2）

                                    break;

                                case "Lv1":
                                    heading1Num++;
                                    heading2Num = 0;
                                    heading3Num = 0;
                                    heading4Num = 0;
                                    headingShiNum = 0;
                                    bodyTextsWorksheet.Cells[i, 2].Value = isChineseDocument ?
                                        ConvertArabicNumberIntoChinese(Convert.ToInt32(heading1Num)) + "、"
                                        : heading1Num + ". ";
                                    checkHeadingNecessity = heading1Num == 1 ? true : false;

                                    break;

                                case "Lv2":
                                    heading2Num++;
                                    heading3Num = 0;
                                    heading4Num = 0;
                                    headingShiNum = 0;
                                    bodyTextsWorksheet.Cells[i, 2].Value = isChineseDocument ?
                                        "（" + ConvertArabicNumberIntoChinese(Convert.ToInt32(heading2Num)) + "）"
                                        : string.Join(".", new object[] { heading1Num, heading2Num }) + " ";
                                    checkHeadingNecessity = heading2Num == 1 ? true : false;

                                    break;

                                case "Lv3":
                                    heading3Num++;
                                    heading4Num = 0;
                                    headingShiNum = 0;
                                    bodyTextsWorksheet.Cells[i, 2].Style.Numberformat.Format = "@";
                                    bodyTextsWorksheet.Cells[i, 2].Value = isChineseDocument ?
                                        heading3Num + "."
                                        : string.Join(".", new object[] { heading1Num, heading2Num, heading3Num }) + " ";
                                    checkHeadingNecessity = heading3Num == 1 ? true : false;

                                    break;

                                case "Lv4":
                                    heading4Num++;
                                    headingShiNum = 0;
                                    bodyTextsWorksheet.Cells[i, 2].Style.Numberformat.Format = "@";
                                    bodyTextsWorksheet.Cells[i, 2].Value = isChineseDocument ?
                                        "(" + heading4Num + ")"
                                        : string.Join(".", new object[] { heading1Num, heading2Num, heading3Num, heading4Num }) + " ";
                                    checkHeadingNecessity = heading4Num == 1 ? true : false;

                                    break;

                                case "Enum.":
                                    headingShiNum++;
                                    bodyTextsWorksheet.Cells[i, 2].Value = isChineseDocument ?
                                        ConvertArabicNumberIntoChinese(Convert.ToInt32(headingShiNum)) + "是"
                                        : "";
                                    checkHeadingNecessity = headingShiNum == 1 ? true : false;

                                    break;

                                case "Itm.":
                                    headingItemNum++;
                                    bodyTextsWorksheet.Cells[i, 2].Value = isChineseDocument ?
                                        "第" + ConvertArabicNumberIntoChinese(Convert.ToInt32(headingItemNum)) + "条 "
                                        : "";
                                    checkHeadingNecessity = headingItemNum == 1 ? true : false;

                                    break;
                            }

                            //删除多余的小标题编号（如果同级小标题编号只有1没有2，则将编号1删去）
                            if (checkHeadingNecessity) // 如果需要检查小标题编号的必要性（当前小标题的编号为1）
                            {
                                if (i == bodyTextsWorksheet.Dimension.End.Row)  // 如果当前行（基准小标题行）为最后一行
                                {
                                    bodyTextsWorksheet.Cells[i, 2].Value = ""; //将当前行（基准小标题行）的小标题编号单元格清空
                                }
                                else //否则
                                {
                                    int headingsCount = 1;
                                    for (int k = i + 1; k <= bodyTextsWorksheet.Dimension.End.Row; k++)  // 遍历从基准行的下一行开始直到最后一行的所有行（比较行）
                                    {
                                        if (!bodyTextsWorksheet.Rows[k].Hidden)  // 如果当前比较行不是隐藏行
                                        {
                                            // 如果当前比较行文字含小标题且小标题级别数小于基准行（小标题级别更高），则退出循环
                                            if (bodyTextsWorksheet.Cells[k, 1].Text.Contains("Lv") && Val(bodyTextsWorksheet.Cells[k, 1].Text) < Val(bodyTextsWorksheet.Cells[i, 1].Text))
                                            {
                                                break;
                                            }
                                            // 否则，如果当前比较行文字的小标题级别（和类型）与基准行的相同，则基准行同级小标题计数加1
                                            else if (bodyTextsWorksheet.Cells[k, 1].Text == bodyTextsWorksheet.Cells[i, 1].Text)
                                            {
                                                headingsCount++;
                                            }
                                        }
                                    }
                                    if (headingsCount <= 1) // 如果累计的基准行同级小标题计数小于等于1，说明基准行同级小标题只有1没有2，则将基准行小标题编号单元格清空
                                    {
                                        bodyTextsWorksheet.Cells[i, 2].Value = "";
                                    }
                                }

                                // 如果基准行小标题编号单元格为空，且文字字数少于50字（视为多余的纯小标题），则将当前行（基准小标题行）小标题编号单元格赋值为“X”（忽略行）
                                if (bodyTextsWorksheet.Cells[i, 2].Text == "" && bodyTextsWorksheet.Cells[i, 3].Text.Length < 50)
                                {
                                    bodyTextsWorksheet.Cells[i, 2].Value = "X";
                                }
                            }
                        }

                    }

                    ExcelRange titleCells = titleWorksheet.Cells[titleWorksheet.Dimension.Address]; //将“大标题”工作表单元格赋值给大标题工作表单元格变量

                    lstFullTexts.AddRange(new string[] { titleCells["C2"].Text, "" }); //将大标题、空行添加到全文本列表中

                    for (int i = 2; i <= bodyTextsWorksheet.Dimension.End.Row; i++)  // 遍历“主体”工作表第2行到最末行的所有行
                    {

                        if (bodyTextsWorksheet.Cells[i, 2].Text != "X")  // 如果当前行没有"X"标记（非忽略行）
                        {
                            //将当前行的小标题编号和小标题正文文字添加到全文本列表
                            string paragraphText = bodyTextsWorksheet.Cells[i, 2].Text + bodyTextsWorksheet.Cells[i, 3].Text; //将当前行小标题编号和文字合并，赋值给段落文字变量
                            if (bodyTextsWorksheet.Cells[i, 1].Text != "Immed.") //如果当前行没有“接上段”的标记，则将段落文字添加到全文本列表（末尾增加一个元素）
                            {
                                lstFullTexts.Add(paragraphText);
                            }
                            else  //否则，将段落文字累加到全文本列表最后一个元素的文字的末尾
                            {
                                lstFullTexts[lstFullTexts.Count - 1] += paragraphText;
                            }

                            if (!isChineseDocument) lstFullTexts.Add(""); // 如果不是中文文档，则将空行添加到全文本列表中（英文文档段中需要空行）
                        }
                    }

                    // 获取日期单元格的日期值并转换为字符串：如果是中文文档，则转换为“yyyy年M月d日”格式；否则，转换为“MMM-d yyyy”格式
                    string dateStr = titleCells["C4"].GetValue<DateTime>().ToString(isChineseDocument ? "yyyy年M月d日" : "MMM-d yyyy", CultureInfo.CreateSpecificCulture("en-US"));

                    //将空行、签名、日期依次添加到全文本列表中
                    lstFullTexts.AddRange(new string[] { "", titleCells["C3"].Text, dateStr });

                    FormatDocumentTable(excelPackage.Workbook); // 格式化结构化文档表中的所有工作表
                    excelPackage.Save(); //保存Excel工作簿
                }

                using (FileStream fileStream = new FileStream(targetWordFilePath, FileMode.Create, FileAccess.Write)) // 创建文件流，以创建目标Word文档
                {
                    XWPFDocument targetWordDocument = new XWPFDocument(); // 创建Word文档对象，赋值给目标Word文档变量

                    // 遍历全文本列表中的所有元素
                    foreach (string paragraphText in lstFullTexts)
                    {
                        // 插入段落
                        XWPFParagraph paragraph = targetWordDocument.CreateParagraph(); // 创建段落
                        XWPFRun run = paragraph.CreateRun(); // 创建段落文本块
                        run.SetText(paragraphText); // 将当前元素的段落文字插入段落文本块中
                    }
                    targetWordDocument.Write(fileStream); // 写入文件流
                }

                //如果对话框返回值为OK（点击了OK），则对目标Word文档执行排版过程
                if (ShowMessage("Do you want to format the document?"))
                {
                    await taskManager.RunTaskAsync(() => FormatWordDocumentsAsync(new List<string> { targetWordFilePath }));
                }

            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
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


    }
}
