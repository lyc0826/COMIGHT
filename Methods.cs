using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Interop;
using Xceed.Words.NET;
using Application = System.Windows.Application;
using DataTable = System.Data.DataTable;
using MSWord = Microsoft.Office.Interop.Word;
using MSWordDocument = Microsoft.Office.Interop.Word.Document;
using Task = System.Threading.Tasks.Task;
using Window = System.Windows.Window;




namespace COMIGHT
{
    public static partial class Methods
    {
        public static string dataBaseFilePath = Path.Combine(Environment.CurrentDirectory, "Database.xlsx"); //获取数据库Excel工作簿文件路径全名

        //定义小标题文字正则表达式变量，匹配模式为：从开头开始，非“。：:；;分页符换行符回车符”的字符2-40个；后方出现：“。：:”换行符回车符或结尾标记
        public static Regex regExHeadingText = new Regex(@"^[^。：:；;\f\n\r]{2,40}(?=。|：|:|\n|\r|$)", RegexOptions.Multiline);

        public static T Clamp<T>(this T value, T min, T max) where T : IComparable<T> //泛型参数T，T必须实现IComparable<T>接口
        {
            //赋值给函数返回值：如果输入值比最小值小，则得到最小值；如果比最大值大，则得到最大值；否则，得到输入值
            return value.CompareTo(min) < 0 ? min : value.CompareTo(max) > 0 ? max : value;
        }

        public static string CleanName(string inputName, int targetLength)
        {
            string cleanedName = inputName.Trim(); //'去除非打印字符和首尾空格
            //正则表达式匹配模式为：制表符“\/:*?<>|"”换行符回车符等1个及以上（不能用于文件名的字符）；将匹配到的字符串替换为下划线
            //在@字符串（逐字字符串字面量）中，双引号只能用双引号转义
            cleanedName = Regex.Replace(cleanedName, @"[\t\\/:\*\?\<\>\|""\n\r]+", "_");
            //正则表达式匹配模式为：空格2个及以上；将匹配到的字符串替换为一个空格
            cleanedName = Regex.Replace(cleanedName, @"[ ]{2,}", " ");
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
                string chineseUnit = arrChineseUnits[n - i - 1]; //获取当前阿拉伯位数字对应的中文单位 （假设是个3位数，当i到达第2位（1号）数字时，3-1-1=1，在中文数字单位数组中索引号为1的元素为“十”，依此类推）
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

                //设置A-I列列宽（小标题级别、小标题编号、文字、完成时限、责任人、分类、相关度、基准小标题、原文来源）
                excelWorksheet.Cells["A:B"].EntireColumn.Width = 12; //=.Columns[1,6]
                excelWorksheet.Cells["C"].EntireColumn.Width = 80;
                excelWorksheet.Cells["D:G"].EntireColumn.Width = 12;
                excelWorksheet.Cells["H:I"].EntireColumn.Width = 24;
                excelWorksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; //文字水平左对齐

                if (excelWorksheet.Dimension == null) //如果当前Excel工作表为空，则直接跳过当前循环并进入下一个循环
                {
                    continue;
                }

                if (excelWorksheet.Index >= 2) // 如果当前Excel工作表的索引号大于等于2（“主体”工作表或“提取”工作表）
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
                    string[] arrValidations = new string[] { "0级", "1级", "2级", "3级", "4级", "条", "是" }; //将数据验证项赋值给数据验证数组

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

                    for (int i = 2; i <= excelWorksheet.Dimension.End.Row; i++) //遍历Excel工作表从第2行开始到末尾的所有行
                    {
                        //设置当前行1至3列字体加粗：如果当前行不含小标题且文字字数少于50字（纯小标题），则加粗；否则不加粗
                        excelWorksheet.Cells[i, 1, i, 3].Style.Font.Bold =
                            (excelWorksheet.Cells[i, 1].Text.EndsWith("级") && excelWorksheet.Cells[i, 3].Text.Length < 50) ? true : false;
                    }
                }
            }
        }

        public static void FormatExcelWorksheet(ExcelWorksheet excelWorksheet, int headerCount = 0, int footerCount = 0)
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
            if (headerCount >= 1) //如果表头行数大于等于1
            {
                ExcelRange headerRange = excelWorksheet.Cells[1, 1, headerCount, excelWorksheet.Dimension.End.Column]; //将表头区域赋值给表头区域变量
                headerRange.Style.Font.Name = "等线";
                headerRange.Style.Font.Size = 12;
                headerRange.Style.Font.Bold = true; //表头区域字体加粗
                headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; //单元格内容水平居中对齐
                headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center; //单元格内容垂直居中对齐
                headerRange.Style.WrapText = true; //设置文字自动换行

                if (excelWorksheet.AutoFilter.Address == null) // 如果自动筛选区域为null（未开启自动筛选），则将表头最后一行的自动筛选设为true
                {
                    excelWorksheet.Cells[headerCount, 1, headerCount, excelWorksheet.Dimension.End.Column].AutoFilter = true;
                }

                for (int i = 1; i <= headerCount; i++) //遍历表头所有行
                {
                    excelWorksheet.Rows[i].Height = 30; //设置当前行的行高为指定值
                    ExcelRange headerRowCells = excelWorksheet.Cells[i, 1, i, excelWorksheet.Dimension.End.Column]; //将当前行所有单元格赋值给表头行单元格变量
                    
                    int mergedCellsCount = headerRowCells.Count(cell => cell.Merge); // 计算当前表头行单元格中被合并的单元格数量
                    //获取“行单元格是否合并”值：如果被合并的单元格数量占当前行所有单元格的75%以上，得到true；否则得到false
                    bool isRowMerged = mergedCellsCount >= headerRowCells.Count() * 0.75 ? true : false; 
                    //获取边框样式：如果行单元格被合并，则得到无边框样式；否则得到细线边框样式
                    ExcelBorderStyle borderStyle = isRowMerged ? ExcelBorderStyle.None : ExcelBorderStyle.Thin;
                    //获取“是否手动调整行高”值：如果行单元格被合并，则得到true；否则得到false
                    bool customHeight = isRowMerged ? true : false;
                    //获取表格标题字体大小：如果行单元格被合并且当前行为第一行（表格标题行），则得到14；否则得到12
                    int titleFontSize = (isRowMerged && i == 1) ? 14 : 12;
                    
                    //设置当前行所有单元格的边框
                    headerRowCells.Style.Border.BorderAround(borderStyle); //设置当前单元格最外侧的边框为之前获取的边框样式
                    headerRowCells.Style.Border.Top.Style = borderStyle; //设置当前单元格顶部的边框为之前获取的边框样式
                    headerRowCells.Style.Border.Left.Style = borderStyle;
                    headerRowCells.Style.Border.Right.Style = borderStyle;
                    headerRowCells.Style.Border.Bottom.Style = borderStyle;

                    headerRowCells.Style.Font.Size = titleFontSize; //设置当前单元格字体大小为之前获取的值
                    excelWorksheet.Rows[i].CustomHeight = customHeight; //设置当前行“是否手动调整行高”为之前获取的值

                }

            }

            // 将Excel工作表除去表头、表尾的区域赋值给记录区域变量
            ExcelRange recordRange = excelWorksheet.Cells[headerCount + 1, 1, excelWorksheet.Dimension.End.Row - footerCount, excelWorksheet.Dimension.End.Column];

            // 记录区域设置字体、对齐
            recordRange.Style.Font.Name = "等线";
            recordRange.Style.Font.Size = 11;
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

            int firstRefRowIndex = Math.Max(1, headerCount); //获取起始参考行的索引号：等于表头最末行索引号，如果小于1，则限定为1
            //获取最末参考行的索引号：除去表尾后余下行的最后一行的索引号，如果超过指定值（起始参考行至其后30行的索引号范围），则将其限定到指定范围
            int lastRefRowIndex = int.Clamp(excelWorksheet.Dimension.End.Row - footerCount, firstRefRowIndex, firstRefRowIndex + 30 - 1);

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
            for (int i = headerCount + 1; i <= excelWorksheet.Dimension.End.Row - footerCount; i++) //遍历除去表尾的所有行
            {
                if (!excelWorksheet.Rows[i].Hidden)  // 如果当前行没有被隐藏
                {
                    excelWorksheet.Rows[i].CustomHeight = false; //将当前行的手动设置行高设为false（即为自动）
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
            printerSettings.RepeatRows = headerCount >= 1 ? new ExcelAddress($"$1:${headerCount}") : new ExcelAddress("");
            //设定打印左侧重复列为A列
            printerSettings.RepeatColumns = new ExcelAddress($"$A:$A");

            // 设置页脚
            string footerText = "第 &P 页，共 &N 页";
            excelWorksheet.HeaderFooter.OddFooter.CenteredText = footerText; // 设置奇数页页脚
            excelWorksheet.HeaderFooter.EvenFooter.CenteredText = footerText; // 设置偶数页页脚

            // 设置视图和打印版式
            ExcelWorksheetView view = excelWorksheet.View; //将Excel工作表视图设置赋值给视图设置变量
            view.UnFreezePanes(); //取消冻结窗格
            view.FreezePanes(headerCount + 1, 2); // 冻结最上方的行和最左侧的列（参数指定第一个不要冻结的单元格）
            view.PageLayoutView = true; // 将工作表视图设置为页面布局视图
            printerSettings.FitToPage = true; // 启用适应页面的打印设置
            int printPagesCount = Math.Max(1, (int)Math.Round(fullWidth / 120, 0)); //计算打印页面数：将全表格宽度除以指定最大宽度的商四舍五入取整，如果小于1，则限定为1
            printerSettings.FitToWidth = printPagesCount;  // 设置缩放为几页宽，1代表即所有列都将打印在一页上
            printerSettings.FitToHeight = 0; // 设置缩放为几页高，0代表打印页数不受限制，可能会跨越多页
            printerSettings.PageOrder = ePageOrder.OverThenDown; // 将打印顺序设为“先行后列”
            view.PageLayoutView = false; // 将页面布局视图设为false（即普通视图）
        }


        public static string GetTitleLevel(string title)
        {
            //定义正则表达式变量，匹配模式为各级小标题、编号或文字
            //0级小标题编号：从开头开始，空格制表符任意多个，“第”，空格制表符任意多个，阿拉伯数字或中文数字1个及以上，空格制表符任意多个，“部分、篇、章、节”，“：:”空格制表符至少一个
            Regex regExHeading0Num = new Regex(@"^[ |\t]*第[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*(?:部分|篇|章|节)[：:| |\t]+", RegexOptions.Multiline);
            // 1级小标题编号：从开头开始，空格制表符任意多个，中文数字1个及以上，空格制表符任意多个，“）)、.．，,”，空格制表符任意多个
            Regex regExHeading1Num = new Regex(@"^[ |\t]*[一二三四五六七八九十〇零]+[ |\t]*[）\)、\.．，,][ |\t]*", RegexOptions.Multiline);
            // 2级小标题编号：从开头开始，空格制表符任意多个，“（(”，空格制表符任意多个，中文数字1个及以上，空格制表符任意多个，“）)、.．，,”，空格制表符任意多个
            Regex regExHeading2Num = new Regex(@"^[ |\t]*[（\(][ |\t]*[一二三四五六七八九十〇零]+[ |\t]*[）\)、\.．，,][ |\t]*", RegexOptions.Multiline);
            // 3级小标题编号：从开头开始，空格制表符任意多个，阿拉伯数字1个及以上，空格制表符任意多个，“）)、.．，,”，空格制表符任意多个
            Regex regExHeading3Num = new Regex(@"^[ |\t]*\d+[ |\t]*[）\)、\.．，,][ |\t]*", RegexOptions.Multiline);
            // 4级小标题编号：从开头开始，空格制表符任意多个，“（(”，空格制表符任意多个，阿拉伯数字1个及以上，空格制表符任意多个，“）)、.．，,”，空格制表符任意多个
            Regex regExHeading4Num = new Regex(@"^[ |\t]*[（\(][ |\t]*\d+[ |\t]*[）\)、\.．，,][ |\t]*", RegexOptions.Multiline);
            // “X是”编号：从开头开始，空格制表符任意多个，中文数字1个及以上，空格制表符任意多个，“是”，空格制表符任意多个
            Regex regExShiNum = new Regex(@"^[ |\t]*[一二三四五六七八九十〇零]+[ |\t]*是[ |\t]*", RegexOptions.Multiline);
            // “第X条”编号：从开头开始，空格制表符任意多个，“第”，空格制表符任意多个，阿拉伯数字或中文数字1个及以上，空格制表符任意多个，“条”，“：:”空格制表符至少一个
            Regex regExItemNum = new Regex(@"^[ |\t]*第[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*条[：:| |\t]+", RegexOptions.Multiline);


            // 使用正则表达式来匹配小标题编号，并赋值给小标题级别单元格
            if (regExHeading0Num.IsMatch(title)) //如果单元格文本被0级小标题编号正则表达式匹配成功，则将当前行的小标题级别（第1列）单元格赋值为“0级”
            {
                return "0级";
            }
            else if (regExHeading1Num.IsMatch(title))
            {
                return "1级";
            }
            else if (regExHeading2Num.IsMatch(title))
            {
                return "2级";
            }
            else if (regExHeading3Num.IsMatch(title))
            {
                return "3级";
            }
            else if (regExHeading4Num.IsMatch(title))
            {
                return "4级";
            }
            else if (regExShiNum.IsMatch(title))
            {
                return "是";
            }
            else if (regExItemNum.IsMatch(title))
            {
                return "条";
            }
            else
            {
                return "";
            }

        }


        //定义不能用于判断字符串相关性的字符的正则表达式变量，匹配模式为：空白字符，非中文、英文、数字、下划线的任意字符，阿拉伯数字、小数点、下划线
        public static Regex regExUselessChars = new Regex(@"\s|[^\u4e00-\u9fa5\w]|[\d\._]");
        public static Regex? regExStopWords; //定义停止词正则表达式变量

        public static double GetTextRelevance(string str1, string str2)
        {
            if (regExStopWords == null) //如果停止词正则表达式变量为null
            {
                DataTable? stopWordsDataTable = ReadExcelWorksheetIntoDataTable(dataBaseFilePath, "Stop Words"); //读取数据库Excel工作簿的“停止词”工作表，赋值给停止词DataTable变量
                StringBuilder stopWordsStrBu = new StringBuilder(); //定义停止词字符串构建器
                if (stopWordsDataTable != null) //如果停止词DataTable变量不为null
                {
                    for (int j = 0; j < stopWordsDataTable.Columns.Count; j++) //遍历停止词DataTable所有数据列
                    {
                        foreach (DataRow dataRow in stopWordsDataTable.Rows) //遍历停止词DataTable所有数据行
                        {
                            if (dataRow[j] != DBNull.Value) //如果当前数据行j列元素不为空，则在数据末尾添加分隔符'|'后，追加到字符串构建器末尾
                            {
                                stopWordsStrBu.Append(Convert.ToString(dataRow[j]) + '|');
                            }
                            else
                            {
                                break; //否则退出for循环
                            }

                        }
                    }
                }
                string stopWordsRegEx = stopWordsStrBu.ToString().Trim('|'); //将字符串构建器的内容转换成字符串，并删除首尾的'|'字符，赋值给停止词正则表达式字符串变量
                regExStopWords = new Regex(stopWordsRegEx); //定义停止词正则表达式变量，匹配模式为停止词
            }

            //从两个字符串中移除停止词
            str1 = regExUselessChars.Replace(str1, ""); //将字符串中的被停止词正则表达式匹配到的字符串替换为空
            str2 = regExUselessChars.Replace(str2, "");

            str1 = regExStopWords.Replace(str1, "");
            str2 = regExStopWords.Replace(str2, "");

            // 检查是否有一个字符串为空或者两个字符串是否完全相同
            if (str1.Length * str2.Length == 0) //如果两个字符串有一个为空，则将0赋值给函数返回值
            {
                return 0;
            }
            else if (str1.Equals(str2)) //如果两个字符串完全相同，则将1赋值给函数返回值
            {
                return 1;
            }

            // 将较长的字符串和较短的字符串分别赋值给长、短字符串变量
            string shortStr = "", longStr = "";
            shortStr = str1.Length < str2.Length ? str1 : str2; //获取短字符串：如果字符串1的字数小于字符串2，则得到字符串1；否则得到字符串2
            longStr = str1.Length < str2.Length ? str2 : str1; //获取长字符串：如果字符串1的字数小于字符串2，则得到字符串2；否则得到字符串1

            int shortStrLen = shortStr.Length; //获取短字符串字数
            int longStrLen = longStr.Length; //获取长字符串字数
            int longStrRepSum = 0;
            int shortStrRepSum = 0;

            while (shortStr.Length > 0) //当短字符串字数不为0，继续循环
            {
                string firstChar = shortStr[0].ToString(); // 获取短字符串的第一个字符（判断字）
                int shortStrRepCount = shortStr.Length - shortStr.Replace(firstChar, "").Length; // 获取短字符串中判断字的出现次数
                int longStrRepCount = longStr.Length - longStr.Replace(firstChar, "").Length; // 获取长字符串中判断字的出现次数
                longStrRepSum += longStrRepCount; // 将长字符串中判断字的出现次数加到长字符串出现该字符的总数中

                // 重新给短字符串重复字符数合计变量赋值：如果当前判断字在长字符串中出现，则得到现有短字符串重复字符数合计与当前判断字在短字符串中出现次数之和；否则得到短字符串重复字符数合计原值
                shortStrRepSum = longStrRepCount >= 1 ? shortStrRepSum + shortStrRepCount : shortStrRepSum;

                // 从短字符串中移除判断字
                shortStr = shortStr.Replace(firstChar, "");
            }

            //计算字符串相关度：((所有共有字在短字符串中出现次数的总和*所有共有字在长字符串中出现次数的总和)/(短字符串字数*长字符串字数))的平方根，赋值给函数返回值
            return Math.Sqrt(shortStrRepSum * longStrRepSum / (double)(shortStrLen * longStrLen));
        }

        public static string? GetKeyColumnLetter()
        {
            string latestColumnLetter = Properties.Settings.Default.latestSplittingColumnLetter; //读取设置中保存的主键列符
            InputDialog inputDialog = new InputDialog("输入主键列符（如：“A”）", latestColumnLetter); //弹出对话框，输入主键列符
            if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则函数返回值赋值为null
            {
                return null;
            }
            string columnLetter = inputDialog.Answer;
            Properties.Settings.Default.latestSplittingColumnLetter = columnLetter; // 将对话框返回的列符存入设置
            Properties.Settings.Default.Save();
            return columnLetter; //将列符赋值给函数返回值
        }


        public static void GetHeaderAndFooterCount(out int headerCount, out int footerCount)
        {
            try
            {
                string lastestHeaderFooterCountStr = Properties.Settings.Default.lastestHeaderFooterCountStr; //读取设置中保存的表头表尾行数字符串
                InputDialog inputDialog = new InputDialog("输入表头、表尾行数（用英文逗号隔开，如：“2,1”代表表头为2行、表尾为1行）", lastestHeaderFooterCountStr); //弹出对话框，输入表头表尾行数
                if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则表头、表尾行数均赋值为默认值，并结束本过程
                {
                    headerCount = 0;
                    footerCount = 0;
                    return;
                }
                string headerFooterCountStr = inputDialog.Answer; //获取对话框返回的表头、表尾行数字符串
                Properties.Settings.Default.lastestHeaderFooterCountStr = headerFooterCountStr; // 将对话框返回的表头、表尾行数字符串存入设置
                Properties.Settings.Default.Save();
                //将表头、表尾字符串拆分成数组，转换成列表，移除每个元素的首尾空白字符，转换成数值，赋值给表头表尾行数列表
                List<int> lstHeaderFooterCount = headerFooterCountStr.Split(',').ToList().ConvertAll(e => Convert.ToInt32(e.Trim()));
                //获取表头表尾行数列表0号、1号元素，如果小于0则限定为0，然后分别赋值给表头、表尾行数变量（引用型）
                headerCount = Math.Max(0, lstHeaderFooterCount[0]);
                footerCount = Math.Max(0, lstHeaderFooterCount[1]);
            }

            catch (Exception ex) // 捕获错误
            {
                MessageBox.Show(ex.Message, "警告", MessageBoxButton.OK, MessageBoxImage.Information);
                headerCount = 0; footerCount = 0; //表头、表尾行数变量赋值为0
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

        // 定义年份正则表达式变量，匹配模式为：前方不能出现“至到-~”，空格制表符任意多个；“20”，阿拉伯数字2个/或“二[〇零]”，中文数字2个
        public static Regex regExYear = new Regex(@"(?<![至到\-~][ |\t]*)(?:[12]\d{3}|[一二][一二三四五六七八九〇零]{3})");

        public static string GetArabicYear(string inText)
        {
            MatchCollection matchesYears = regExYear.Matches(inText); //获取输入文字经过年份正则表达式匹配后的结果
            //获取年份字符串：如果输入文字经过年份正则表达式匹配的结果集合元素数大于0，则得到最后一个匹配结果；否则得到空字符串
            string yearStr = matchesYears.Count > 0 ? matchesYears[matchesYears.Count - 1].Value : "";
            //将年份字符串中的中文数字替换为阿拉伯数字，移除首尾空白字符，赋值给阿拉伯数字年份变量
            Dictionary<char, char> map = new Dictionary<char, char> //定义字符字典，将中文数字与阿拉伯数字对应
                {
                    {'一', '1'}, {'二', '2'}, {'三', '3'}, {'四', '4'}, {'五', '5'},
                    {'六', '6'}, {'七', '7'}, {'八', '8'}, {'九', '9'}, {'〇', '0'}, {'零', '0'}
                };

            //遍历年份字符串的每个字符，如果在字符字典中找到该字符的键，则得到对应键值；否则得到键本身。将字符组成数组，然后转为字符串，并移除首尾空白字符
            string arabicYearStr = new string(yearStr.Select(c => map.ContainsKey(c) ? map[c] : c).ToArray()).Trim();
            return arabicYearStr;  //将阿拉伯数字年份赋值给函数返回值
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

        public static void MergeExcelWorksheetHeader(ExcelWorksheet excelWorksheet, int headerCount)
        {
            if (excelWorksheet.Dimension == null || headerCount < 2) //如果工作表为空或者表头行数小于2，则结束本过程
            {
                return;
            }

            excelWorksheet.Cells[1, 1, headerCount, excelWorksheet.Dimension.End.Column].Merge = false; //表头所有单元格的合并状态设为false

            for (int j = 1; j <= excelWorksheet.Dimension.End.Column; j++) //遍历工作表所有列
            {
                List<string> lstFullColumnName = new List<string> { }; //定义完整列名称列表
                for (int i = 1; i <= headerCount; i++) //遍历工作表所有行
                {
                    bool copyLeftCell = false; //“是否复制左侧单元格”赋值为false
                    if (j > 1 && string.IsNullOrWhiteSpace(excelWorksheet.Cells[i, j].Text)) //如果当前列索引号大于1，且当前单元格为null或全空白字符
                    {
                        if (i == 1) //如果当前行是第1行，则“是否复制左侧单元格”赋值为true
                        {
                            copyLeftCell = true;
                        }
                        //否则，如果比当前列索引号小1、行索引号相同（上方）的单元格的值和比当前列索引号小1、比当前行索引号小1（左上方）的单元格相同，则“是否复制左侧单元格”赋值为true
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
                excelWorksheet.Cells[headerCount, j].Value = string.Join('_', lstFullColumnName.Where(e => !string.IsNullOrWhiteSpace(e)));

            }
            excelWorksheet.DeleteRow(1, headerCount - 1); //删除表头除了最后一行的所有行

        }

        public static DataTable? ReadExcelWorksheetIntoDataTable(string filePath, object worksheetID, int headerCount = 1, int footerCount = 0)
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
                            throw new Exception("参数错误！");
                    }

                    TrimCellsStrings(excelWorksheet!, true); //删除Excel工作表内所有单元格值的首尾空格，并全部转换为文本型
                    RemoveWorksheetEmptyRowsAndColumns(excelWorksheet!); //删除Excel工作表内所有空白行和空白列
                    if ((excelWorksheet.Dimension?.Rows ?? 0) <= headerCount + footerCount) //如果Excel工作表已使用行数（如果工作表为空，则为0）小于等于表头表尾行数和，则函数返回值赋值为null
                    {
                        return null;
                    }

                    foreach (ExcelRangeBase cell in excelWorksheet.Cells[excelWorksheet.Dimension!.Address]) //遍历已使用区域的所有单元格
                    {
                        //移除当前单元格文本首尾空白字符后重新赋值给当前单元格（所有单元格均转为文本型）
                        cell.Value = cell.Text.Trim();
                    }

                    MergeExcelWorksheetHeader(excelWorksheet, headerCount); //将多行表头合并为单行

                    DataTable dataTable = new DataTable(); // 定义DataTable变量
                    //读取Excel工作表并载入DataTable（第一行为表头，跳过表尾指定行数，将所有错误值视为空值，总是允许无效值）
                    dataTable = excelWorksheet.Cells[excelWorksheet.Dimension.Address].ToDataTable(
                        o =>
                        {
                            o.FirstRowIsColumnNames = true;
                            o.SkipNumberOfRowsEnd = footerCount;
                            o.ExcelErrorParsingStrategy = ExcelErrorParsingStrategy.HandleExcelErrorsAsBlankCells;
                            o.AlwaysAllowNull = true;
                        });
                    return dataTable; //将DataTable赋值给函数返回值
                }
            }

            catch (Exception ex) // 捕获错误
            {
                MessageBox.Show(ex.Message, "警告", MessageBoxButton.OK, MessageBoxImage.Information);
                return null; //函数返回值赋值为null
            }

        }

        public static DataTable RemoveDataTableEmptyRowsAndColumns(DataTable dataTable)
        {
            //清除空白数据行
            for (int i = dataTable.Rows.Count - 1; i >= 0; i--) // 遍历DataTable所有数据行
            {
                //如果当前数据行所有数据列的值均为数据库空值，或为null或全空白字符，则删除当前数据行
                if (dataTable.Rows[i].ItemArray.All(value => value == DBNull.Value || string.IsNullOrWhiteSpace(value?.ToString())))
                {
                    dataTable.Rows[i].Delete();
                }
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

        public static string RemoveHeadingNum(string inText, bool keepLeadingNum = false)
        {
            //定义小标题编号正则表达式字符串：“#*_->~”至少一个，空格至多一个（MD标记捕获组，至多一个），空格制表符任意多个，“第（(”最多一个， 空格制表符任意多个，阿拉伯数字或中文数字1个及以上，空格制表符任意多个，“部分|篇|章|节” “：:”空格至少一个/或“）)、\.．，,是”，空格制表符任意多个
            string headingNumRegEx = @"([#\*_\->~]+[ ]?)?[ |\t]*[第（\(]?[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*(?:(?:部分|篇|章|节)[：:| ]+|[）\)、\.．，,是])[ |\t]*";
            //定义开头标记正则表达式字符串：如果“保留开头小标题编号”为false，则为：前方出现开头标记或“。：:；;”；否则为：前方出现“。：:；;”
            string leadingMarksRegEx = !keepLeadingNum ? @"(?<=^|[。：:；;])" : @"(?<=[。：:；;])";
            //定义小标题编号正则表达式变量，匹配模式为：开头标记和小标题编号两个正则表达式字符串的合并字符串
            Regex regExHeadingNum = new Regex(leadingMarksRegEx + headingNumRegEx, RegexOptions.Multiline);
            return regExHeadingNum.Replace(inText, "$1"); //将输入文字中被小标题编号正则表达式匹配到的字符串替换为捕获组字符串（如有MD标记则替换为之，否则替换为空），赋值给函数返回值
        }


        //定义小标题句正则表达式变量，匹配模式为：开头标记，任意字符0-50个（尽可能少）（捕获组1），非“。：:；;，,”字符任意多个，“。：:”（捕获组2，即小标题句），任意字符任意多个（捕获组3），结尾标记
        //public static Regex regExHeadingSentence = new Regex(@"^(.{0,50}?)([^。：:；;，,]*[。：:])(.*)$", RegexOptions.Multiline);

        //public static string PutHeadingSentenceFirst(this string inText)
        //{
        //    if (string.IsNullOrWhiteSpace(inText)) //如果输入文字为null或全空白字符，则将空字符串赋值给函数返回值
        //    {
        //        return string.Empty;
        //    }

        //    //将输入文本中被小标题句正则表达式匹配到的字符串替换为捕获组2、1、3合并后的字符串（即将小标题句提到最前方），赋值给函数返回值
        //    return regExHeadingSentence.Replace(inText, "$2$1$3");
        //}

        public enum FileType { Excel, Word, WordAndExcel, Convertible } //定义文件类型枚举

        public static List<string>? SelectFiles(FileType fileType, bool isMultiselect, string dialogTitle)
        {
            string filter = fileType switch //根据文件类型枚举，返回相应的文件类型和扩展名的过滤项
            {
                FileType.Excel => "Excel文件(*.xlsx;*.xlsm)|*.xlsx;*.xlsm|所有文件(*.*)|*.*",
                FileType.Word => "Word文件(*.docx;*.docm)|*.docx;*.docm|所有文件(*.*)|*.*",
                FileType.WordAndExcel => "Word或Excel文件(*.docx;*.docm;*.xlsx;*.xlsm)|*.docx;*.docm;*.xlsx;*.xlsm|所有文件(*.*)|*.*",
                FileType.Convertible => "可转换文件(*.doc;*.xls;*.wps;*.et)|*.doc;*.xls;*.wps;*.et|所有文件(*.*)|*.*",
                _ => "所有文件(*.*)|*.*"
            };

            string initialDirectory = Properties.Settings.Default.latestFolderPath; //获取保存在设置中的文件夹路径
            //重新赋值给初始文件夹路径变量：如果初始文件夹路径存在，则得到初始文件夹路径原值；否则得到C盘根目录
            initialDirectory = Directory.Exists(initialDirectory) ? initialDirectory : @"C:\";
            OpenFileDialog openFileDialog = new OpenFileDialog() //打开文件选择对话框
            {
                Multiselect = isMultiselect, //是否可多选
                Title = dialogTitle, //对话框标题
                Filter = filter, //文件类型和相应扩展名的过滤项
                InitialDirectory = initialDirectory //初始文件夹路径
            };

            if (openFileDialog.ShowDialog() == true) //如果对话框返回true（选择了OK）
            {
                Properties.Settings.Default.latestFolderPath = Path.GetDirectoryName(openFileDialog.FileNames[0]); // 将本次选择的文件的文件夹路径保存到设置中
                Properties.Settings.Default.Save(); //

                return openFileDialog.FileNames.ToList(); // 将被选中的文件数组转换成列表，赋给函数返回值
            }
            return null; //如果上一个if未执行，没有文件列表赋给函数返回值，则函数返回值赋值为null
        }

        public static string ProceedToExtractText(string inText, char separator, int targetLength)
        {
            if (targetLength >= inText.Length) //如果目标字数大于等于输入文字字数，则将输入文字赋值给函数返回值
            {
                return inText;
            }

            //将输入文字按换行符拆分为数组（删除每个元素前后空白字符），转换成列表
            List<string> lstParagraphs = inText.Split('\n', StringSplitOptions.TrimEntries).ToList();

            int bodyParagraphCount = Math.Max(1, lstParagraphs.Count(p => p.Length >= 50)); //计算字数大于等于50字的段落数量（正文段落），如果结果小于1则限定为1

            //定义冗余文字正则表达式变量，匹配模式为：前方出现“。；;”，任意字符任意多个（尽可能少），阿拉伯数字至少一个、小数点至多一个、阿拉伯数字任意多个（数字捕获组），任意字符任意多个（尽可能少），“。；;”
            Regex regExRedundantTexts = new Regex(@"(?<=[。；;]).*?(\d+\.?\d*)?.*?[。；;]"); 

            for (int i = lstParagraphs.Count - 1; i >= 0; i--)  //遍历段落列表元素
            {
                bool paragraphIsShortened = false; //“段落是否缩短”变量赋值为false
                MatchCollection matchesRedundantTexts = regExRedundantTexts.Matches(lstParagraphs[i]); //获取当前元素（段落）经过冗余文字正则表达式匹配的结果集合
                //将匹配结果集合转换为单个匹配的枚举集合，颠倒元素顺序，再按捕获组数量从多到少排序，转换成列表，赋值给冗余文字匹配结果列表
                List<Match> lstMatchesRedundantTexts = matchesRedundantTexts.Cast<Match>().Reverse().OrderByDescending(m => m.Groups.Count).ToList();

                //如果当前元素（段落）的字数大于限定至目标字数后平均每个正文段落的字数（正文段落总字数约等于全文总字数的95%），则继续循环
                while (lstParagraphs[i].Length > targetLength * 0.95 / bodyParagraphCount)
                {
                    if (lstMatchesRedundantTexts.Count > 0) //如果冗余文字匹配结果列表元素数大于0
                    {
                        //将段落列表当前元素中的与冗余文字匹配结果列表0号元素（数字捕获组最多的元素）相同的文字替换为空
                        lstParagraphs[i] = lstParagraphs[i].Replace(lstMatchesRedundantTexts[0].Value, "");
                        lstMatchesRedundantTexts.RemoveAt(0); //移除冗余文字匹配结果列表0号元素
                        paragraphIsShortened = true; //“段落是否缩短”变量赋值为true
                    }
                    else //否则，退出循环
                    {
                        break;
                    }
                }
                //重新给当前元素赋值：如果段落被缩短，则得到移除当前元素小标题编号（但保留开头的编号）后的文字；否则得到当前元素原值
                lstParagraphs[i] = paragraphIsShortened ? RemoveHeadingNum(lstParagraphs[i], true) : lstParagraphs[i];
            }

            //如果段落列表所有元素合并后的总字数大于目标字数，则继续循环，删除最后一个元素
            while (string.Join(separator, lstParagraphs).Length > targetLength)
            {
                lstParagraphs.RemoveAt(lstParagraphs.Count - 1);
            }

            return string.Join(separator, lstParagraphs); //将段落列表所有元素合并，赋值给函数返回值

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
                    ExcelWorksheet headingsWorksheet = excelPackage.Workbook.Worksheets.Add("小标题");
                    ExcelWorksheet titleWorksheet = excelPackage.Workbook.Worksheets.Add("大标题首尾");
                    ExcelWorksheet bodyTextsWorksheet = excelPackage.Workbook.Worksheets.Add("主体");

                    titleWorksheet.Cells["A1:C1"].LoadFromArrays(new List<object[]> { new object[] { "项目", "编号", "文字" } });
                    titleWorksheet.Cells["A2:A6"].LoadFromArrays(new List<object[]>
                        {
                            new object[] { "大标题" },
                            new object[] { "首段" },
                            new object[] { "尾段" },
                            new object[] { "落款" },
                            new object[] { "日期" }
                        }); // 初始化表头
                    titleWorksheet.Cells["C5"].Value = "单位名称"; // 将“单位名称”赋值给落款单元格
                    titleWorksheet.Cells["C6"].Value = DateTime.Now.ToString("yyyy年M月d日"); // 将当前日期赋值给日期单元格
                    bodyTextsWorksheet.Cells["A1:F1"].LoadFromArrays(new List<object[]> { new object[] { "小标题级别", "小标题编号", "文字", "完成时限", "责任部门（人）", "分类" } });
                    bodyTextsWorksheet.Cells["A1:F1"].Copy(headingsWorksheet.Cells["A1"]); //将“主体”工作表的表头复制到“小标题”工作表

                    // 将Word文档数组内容赋值给“主体”工作表内容列的单元格
                    for (int i = 0; i < lstParagraphs!.Count; i++) //遍历数组所有元素
                    {
                        bodyTextsWorksheet.Cells[i + 2, 3].Value = lstParagraphs[i]; //将当前数组元素赋值给第3列的第i+2行的单元格
                    }

                    // 在“主体”工作表中，判断小标题正文文字的编号级别，赋值给小标题级别单元格，并将小标题正文文字的小标题编号清除，同时更新“小标题”工作表
                    for (int i = 2; i <= bodyTextsWorksheet.Dimension.End.Row; i++) //遍历从第2行开始往下的所有行
                    {
                        string cellText = bodyTextsWorksheet.Cells[i, 3].Text; //将当前行的小标题正文文字（第3列）单元格的文本赋值给单元格文本变量
                        bodyTextsWorksheet.Cells[i, 1].Value = GetTitleLevel(cellText); //获取单元格文本的小标题级别，赋值给当前行的小标题级别单元格
                        bodyTextsWorksheet.Cells[i, 3].Value = RemoveHeadingNum(cellText); //删除单元格文本中的所有小标题编号，赋值给当前行的小标题正文文字单元格

                        //更新“小标题”工作表
                        if (bodyTextsWorksheet.Cells[i, 1].Text.Contains("级")) // 如果当前行含小标题
                        {
                            MatchCollection matchesHeadingTexts = regExHeadingText.Matches(bodyTextsWorksheet.Cells[i, 3].Text);  // 获取当前行的小标题正文文字经过小标题文字正则表达式匹配的结果
                            string headingText = matchesHeadingTexts.Count > 0 ? matchesHeadingTexts[0].Value : ""; // 如果匹配到的结果集合元素数大于0，则将匹配到的小标题文字赋值给小标题文字变量

                            int lastRowIndex = headingsWorksheet.Dimension?.End.Row ?? 0; //获取“小标题”工作表最末行索引号（如果工作表为空， 则为0）
                            headingsWorksheet.Cells[lastRowIndex + 1, 1, lastRowIndex + 1, 3].Style.Numberformat.Format = "@"; // 将“小标题”工作表第一个空白行第1至3列的单元格的格式设为文本
                            headingsWorksheet.Cells[lastRowIndex + 1, 1].Value = bodyTextsWorksheet.Cells[i, 1].Text; // 将当前行的小标题级别赋值给“小标题”工作表第一个空白行的小标题级别单元格
                            headingsWorksheet.Cells[lastRowIndex + 1, 2].Value = bodyTextsWorksheet.Cells[i, 2].Text; // 将当前行的小标题编号赋值给“小标题”工作表第一个空白行的小标题编号单元格
                            headingsWorksheet.Cells[lastRowIndex + 1, 3].Value = headingText; // 将小标题文字赋值给“小标题”工作表第一个空白行的小标题文字单元格
                        }

                    }

                    // 在“大标题首尾”工作表中，给大标题和首段单元格赋值
                    titleWorksheet.Cells["C2"].Value = bodyTextsWorksheet.Cells["C2"].Value; // 将“主体”工作表第2行“文字”单元格值赋值给“大标题首尾”工作表的“大标题”单元格
                    if (!bodyTextsWorksheet.Cells["A3"].Text.Contains("级")) // 如果“主体”工作表第3行不含小标题（为普通正文）
                    {
                        titleWorksheet.Cells["C3"].Value = bodyTextsWorksheet.Cells["C3"].Value; // 将“主体”工作表第3行“文字”单元格值赋值给“大标题首尾”工作表的“首段”单元格
                        bodyTextsWorksheet.DeleteRow(3); // 删除“主体”工作表第3行（已经被转移的首段）
                    }
                    bodyTextsWorksheet.DeleteRow(2); // 删除“主体”工作表第2行（已经被转移的大标题）

                    TrimCellsStrings(bodyTextsWorksheet); //删除“主体”Excel工作表内所有文本型单元格值的首尾空格
                    RemoveWorksheetEmptyRowsAndColumns(bodyTextsWorksheet); //删除“主体”Excel工作表内所有空白行和空白列

                    FormatDocumentTable(excelPackage.Workbook); //格式化文档表的所有工作表
                    excelPackage.SaveAs(new FileInfo(targetExcelFilePath)); // 保存目标工作簿
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "警告", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        public static async Task FormatWordDocumentsAsync(List<string> filePaths)
        {
            Task task = Task.Run(() => process());
            void process()
            {
                MSWord.Application msWordApp = new MSWord.Application(); //打开Word应用程序并赋值给word应用程序变量
                msWordApp.ScreenUpdating = false;
                msWordApp.Visible = false;
                try
                {
                    foreach (string filePath in filePaths) //遍历文件路径全名列表所有元素
                    {
                        MSWordDocument msWordDocument = msWordApp.Documents.Open(filePath); // 打开word文档并赋值给初始Word文档变量

                        // 判断是否为空文档
                        if (msWordDocument.Content.Text.Trim().Length <= 1) // 如果将 Word 换行符全部删除后，剩下的字符数小于等于1，则结束本过程
                        {
                            return;
                        }

                        // 接受并停止修订
                        msWordDocument.AcceptAllRevisions();
                        msWordDocument.TrackRevisions = false;
                        msWordDocument.ShowRevisions = false;

                        // 设置版式规格
                        double topMargin = msWordApp.CentimetersToPoints((float)3.7); // 顶端页边距
                        double bottomMargin = msWordApp.CentimetersToPoints((float)3.5); // 底端页边距
                        double leftMargin = msWordApp.CentimetersToPoints((float)2.8); // 左页边距
                        double rightMargin = msWordApp.CentimetersToPoints((float)2.6); // 右页边距
                        int lineSpace = 28; // 行间距
                        int titleFontSize = 22; // 大标题字号为二号
                        int bodyTextFontSize = 16; // 正文字号为三号
                        int heading0FontSize = 16; // 0级小标题字号为三号
                        int heading1FontSize = 16; // 1级小标题字号为三号
                        int heading2FontSize = 16; // 2级小标题字号为三号
                        int heading3FontSize = 16; // 3级小标题字号为三号
                        int headingShiFontSize = 16; // “是”语句字号为三号
                        int headingItemFontSize = 16; // “条”编号字号为三号
                        int footerFontSize = 14; // 页脚字号为四号

                        // 设置查找模式
                        MSWord.Selection selection = msWordApp.Selection; //将选区赋值给选区变量
                        MSWord.Find find = msWordApp.Selection.Find; //将选区查找赋值给查找变量

                        find.ClearFormatting(); // 清除格式
                        find.Wrap = WdFindWrap.wdFindStop; // 到文档结尾后停止查找
                        find.Forward = true; // 正向查找
                        find.MatchByte = false; // 区分全角半角为False
                        find.MatchWildcards = false; // 使用通配符为False

                        // 全文空格替换为半角空格，制表符替换为空格，换行符替换为回车符
                        selection.WholeStory();
                        find.Text = " "; // 查找空格
                        find.Replacement.Text = " "; // 将空格替换为半角空格
                        find.Execute(Replace: WdReplace.wdReplaceAll);

                        find.Text = "\t"; // 查找制表符
                        find.Replacement.Text = "    "; // 将制表符替换为4个空格
                        find.Execute(Replace: WdReplace.wdReplaceAll);

                        find.Text = "\v"; // 查找换行符（垂直制表符），^l"
                        find.Replacement.Text = "\r"; // 将换行符（垂直制表符）替换为回车符
                        find.Execute(Replace: WdReplace.wdReplaceAll);

                        // 清除段首、段尾多余空格和制表符，段落自动编号转文本
                        selection.EndKey(WdUnits.wdStory);
                        selection.InsertAfter("\r"); // 在文尾加“保护”换行符，以免在替换最后一段时，造成和倒数第二段错误合并。

                        for (int i = msWordDocument.Paragraphs.Count; i >= 1; i--) // 从末尾往开头遍历所有段落
                        {
                            MSWord.Paragraph paragraph = msWordDocument.Paragraphs[i];

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
                                paragraph.Range.InsertBefore(paragraph.Range.ListFormat.ListString); // 在段落文字前添加自动编号
                            }
                        }

                        // 对齐缩进
                        selection.WholeStory();
                        selection.ClearFormatting(); // 清除全部格式、样式
                        MSWord.ParagraphFormat paragraphFormat = msWordApp.Selection.ParagraphFormat; //将选区段落格式赋值给段落格式变量
                        paragraphFormat.Reset(); // 段落格式清除
                        paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0
                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify; // 对齐方式设为两端对齐
                        paragraphFormat.IndentFirstLineCharWidth(3); // 首行缩进3个字符

                        // 清除文首和文末的空白段
                        while (msWordDocument.Paragraphs[1].Range.Text == "\r") // 如果第1段文字为换行符，则继续循环
                        {
                            msWordDocument.Paragraphs[1].Range.Delete(); // 删除第1段
                        }

                        while (msWordDocument.Paragraphs[msWordDocument.Paragraphs.Count].Range.Text == "\r") // 如果最后一段文字为换行符，则继续循环
                        {
                            msWordDocument.Paragraphs[msWordDocument.Paragraphs.Count].Range.Delete(); // 删除最后一段
                        }

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
                                                                             // '.CharacterUnitFirstLineIndent = 2 '此参数优先级最高，一旦设定，需要再次设置一个绝对值相等的负值或者重置段落格式才能将其归零！
                        paragraphs.AutoAdjustRightIndent = 0; // 不自动调整右缩进
                        paragraphs.DisableLineHeightGrid = -1; //取消“如果定义了网格，则对齐到网格”
                        paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly; // 行距设置为固定值
                                                                                       // '.LineSpacingRule = wdLineSpace1pt5 '行距固定1.5
                        paragraphs.LineSpacing = lineSpace; // 行距设为预设值
                        paragraphs.SpaceBefore = msWordApp.CentimetersToPoints(0); // 段落前间距设为0
                        paragraphs.SpaceAfter = msWordApp.CentimetersToPoints(0); // 段落后间距设为0

                        // 字体设置
                        MSWord.Font font = msWordApp.Selection.Font; //将选区字体赋值给字体变量
                        font.Name = "仿宋"; // 字体名称设为仿宋
                        font.Size = bodyTextFontSize; // 字号设为正文预设字号
                        font.ColorIndex = WdColorIndex.wdBlack; // 颜色设为黑色
                        font.Kerning = 0; // “为字体调整字符间距”值设为0
                        font.DisableCharacterSpaceGrid = true;  //取消“如果定义了文档网格,则对齐到网格”，忽略字体的每行字符数

                        string documentText = msWordDocument.Content.Text; // 全文文字变量赋值

                        selection.HomeKey(WdUnits.wdStory);

                        // 文档大标题设置
                        // 定义大标题正则表达式变量，匹配模式为：从开头开始，不含2个及以上连续的换行符回车符（允许不连续的换行符回车符）、不含“附件/录”注释、非“。”分页符的字符2-60个，换行符回车符，后方出现：换行符回车符
                        Regex regExTitle = new Regex(@"(?<=^|\n|\r)(?:(?![\n\r]{2,})(?!附[ |\t]*[件录][^。\f\n\r]{0,3}[\n\r])[^。\f]){2,60}[\n\r](?=[\n\r])", RegexOptions.Multiline);

                        // 定义发往单位正则表达式变量，匹配模式为：从开头开始，换行符回车符（一个空行），不含“附件/录”注释、不含小标题编号、不含“如下：”、非“。：:；;”分页符换行符回车符的字符2个及以上，换行符回车符
                        Regex regExAddressee = new Regex(@"(?<=^|\n|\r)[\n\r](?:(?!附[ |\t]*[件录][^。\f\n\r]{0,3}[\n\r])(?![（\(]?[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*[）\)、\.．，,])(?!如下[：:])[^。：:；;\f\n\r]){2,}[：:][\n\r]", RegexOptions.Multiline);

                        int referencePageNum = 0; //参考页码赋值为0
                        MatchCollection matchesTitles = regExTitle.Matches(documentText); // 获取全文文字经过大标题正则表达式匹配后的结果

                        foreach (Match matchTitle in matchesTitles) // 遍历所有匹配到的大标题文字
                        {
                            // 文档大标题设置
                            selection.HomeKey(WdUnits.wdStory);
                            find.Text = matchTitle.Value; // 查找大标题
                            find.Execute();
                            int pageNum = selection.Information[WdInformation.wdActiveEndPageNumber]; // 当前页码变量赋值
                            if (!selection.Information[WdInformation.wdWithInTable] && pageNum != referencePageNum) //如果当前大标题不在表格内，且与之前已确定的大标题不在同一页（一页最多一个大标题）
                            {
                                bool formatTitle = false; // “格式化大标题”变量赋值为False
                                if (pageNum == 1) // 如果大标题候选文字在第一页
                                {
                                    formatTitle = true; // “格式化大标题”变量赋值为True
                                }
                                else // 否则
                                {
                                    selection.MoveStart(WdUnits.wdLine, -5); // 将搜索到大标题候选文字选区向上扩展5行
                                    if (selection.Text.Contains("\f")) // 如果选区内含有分页符，则候选文字判断为大标题，“格式化大标题”变量赋值为True
                                    {
                                        formatTitle = true;
                                    }
                                    selection.MoveStart(WdUnits.wdLine, 5); // 选区起点复原
                                }
                                if (formatTitle) // 如大标题需要进行格式化
                                {
                                    paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0
                                    paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 居中对齐
                                    font.Name = "华文中宋"; // 字体为华文中宋
                                    font.Size = titleFontSize; // 字号为大标题预设字号
                                    font.ColorIndex = WdColorIndex.wdBlack; // 颜色为黑色
                                    font.Bold = (int)WdConstants.wdToggle; // 字体加粗
                                    selection.EndKey(WdUnits.wdLine); // 光标一到选区的最后一个字（换行符之前）

                                    // 发往单位设置
                                    selection.MoveDown(WdUnits.wdLine, 1, WdMovementType.wdMove); // 光标下移到下方一行
                                    selection.Expand(WdUnits.wdLine); // 全选一行
                                    selection.MoveEnd(WdUnits.wdLine, 5); // 选区向下扩大5行

                                    MatchCollection matchesAddressees = regExAddressee.Matches(selection.Text); // 获取选区文字经过发往单位正则表达式匹配的结果
                                    foreach (Match matchAddressee in matchesAddressees) // 遍历所有匹配到的发往单位文字结果
                                    {
                                        find.Text = matchAddressee.Value; // 查找发往单位
                                        find.Execute(); // 执行查找

                                        if (!selection.Information[WdInformation.wdWithInTable]) // 如果找到的文字不在表格内
                                        {
                                            paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 段落首行缩进为0
                                        }
                                        selection.Collapse(WdCollapseDirection.wdCollapseEnd); // 将选区折叠到末尾
                                    }
                                    referencePageNum = selection.Information[WdInformation.wdActiveEndPageNumber]; // 获取大标题所在页码并赋值给相应变量，为以后提供参考
                                }
                            }
                        }

                        int outlineLevelOffset = 0; // 大纲级别偏移量赋值为0

                        // 0级（部分、篇、章、节）小标题设置
                        selection.HomeKey(WdUnits.wdStory);

                        // 定义0级小标题正则表达式变量，匹配模式为：从开头开始，“第”，空格制表符任意多个，阿拉伯数字中文数字1个及以上，空格制表符任意多个，“部分、篇、章、节”，非“。：:；;”分页符换行符回车符的字符0-40个，换行符回车符
                        Regex regExHeading0 = new Regex(@"(?<=^|\n|\r)第[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*(?:部分|篇|章|节)[^。：:；;\f\n\r]{0,40}[\n\r]", RegexOptions.Multiline);
                        MatchCollection matchesHeading0s = regExHeading0.Matches(documentText); // 获取全文文字经过0级小标题正则表达式匹配的结果

                        foreach (Match matchHeading0 in matchesHeading0s)
                        {
                            find.Text = matchHeading0.Value;
                            find.Execute();
                            if (paragraphs[1].Range.Sentences.Count == 1) // 如果找到的小标题所在段落只有一句
                            {
                                paragraphs[1].OutlineLevel = WdOutlineLevel.wdOutlineLevel1; // 将当前小标题的大纲级别设为1级
                                outlineLevelOffset = 1; // 大纲级别偏移量设为1（后续一、二、三级小标题的大纲级别相应推后至二、三、四级）
                            }
                            paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进为0
                            paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 居中对齐
                            font.Name = "黑体";
                            font.Size = heading0FontSize;
                            font.ColorIndex = WdColorIndex.wdBlack;
                            font.Bold = 1;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 一级小标题设置
                        selection.HomeKey(WdUnits.wdStory);

                        // 定义1级小标题正则表达式变量，匹配模式为：从开头开始，中文数字1个及以上，空格制表符任意多个，“、.．，,”，非“。：:；;”分页符换行符回车符的字符2-40个，“。：:”换行符回车符
                        Regex regExHeading1 = new Regex(@"(?<=^|\n|\r)[一二三四五六七八九十〇零]+[ |\t]*[、\.．，,][^。：:；;\f\n\r]{2,40}[。：:\n\r]", RegexOptions.Multiline);
                        MatchCollection matchesHeading1s = regExHeading1.Matches(documentText); // 获取全文文字经过1级小标题正则表达式匹配的结果

                        foreach (Match matchHeading1 in matchesHeading1s)
                        {
                            find.Text = matchHeading1.Value;
                            find.Execute();

                            if (paragraphs[1].Range.Sentences.Count == 1)
                            {
                                paragraphs[1].OutlineLevel = WdOutlineLevel.wdOutlineLevel1 + outlineLevelOffset; // 将当前小标题的大纲级别设为1级加大纲级别偏移量
                            }
                            font.Name = "黑体";
                            font.Size = heading1FontSize;
                            font.ColorIndex = WdColorIndex.wdBlack;
                            font.Bold = 1;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 二级小标题设置
                        selection.HomeKey(WdUnits.wdStory);

                        // 定义2级小标题正则表达式变量，匹配模式为：从开头开始，“（(”，空格制表符任意多个，中文数字1个及以上，空格制表符任意多个，“）)”，非“。：:；;”分页符换行符回车符的字符2-40个，“。：:”换行符回车符
                        Regex regExHeading2 = new Regex(@"(?<=^|\n|\r)[（\(][ |\t]*[一二三四五六七八九十〇零]+[ |\t]*[）\)][^。：:；;\f\n\r]{2,40}[。：:\n\r]", RegexOptions.Multiline);
                        MatchCollection matchesHeading2s = regExHeading2.Matches(documentText); // 获取全文文字经过2级小标题正则表达式匹配的结果

                        foreach (Match matchHeading2 in matchesHeading2s)
                        {
                            find.Text = matchHeading2.Value;
                            find.Execute();
                            if (selection.Paragraphs[1].Range.Sentences.Count == 1)
                            {
                                selection.Paragraphs[1].OutlineLevel = WdOutlineLevel.wdOutlineLevel2 + outlineLevelOffset;
                            }
                            font.Name = "楷体";
                            font.Size = heading2FontSize;
                            font.ColorIndex = WdColorIndex.wdBlack;
                            font.Bold = 1;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 三级及以下小标题设置
                        selection.HomeKey(WdUnits.wdStory);

                        // 定义3级及以下小标题正则表达式变量，匹配模式为：从开头开始，“（(”，空格制表符任意多个，阿拉伯数字1个及以上，空格制表符任意多个，“）)、.．，,”，非“。：:；;”分页符换行符回车符的字符2-40个，“。：:”换行符回车符，换行符回车符至多1个；后方不可出现一级、二级、三级标题编号；但要出现非分页符换行符回车符的字符2个及以上
                        Regex regExHeading3Below = new Regex(@"(?<=^|\n|\r)[（\(]?[ |\t]*\d+[ |\t]*[）\)、\.．，,][^。：:；;\f\n\r]{2,40}[。：:\n\r]", RegexOptions.Multiline);  //[\n\r]?(?![（\(]?[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*[）\)、\.．，,])(?=[^\f\n\r]{2,})", RegexOptions.Multiline);
                        MatchCollection matchesHeading3Belows = regExHeading3Below.Matches(documentText); // 获取全文文字经过3级及以下小标题正则表达式匹配的结果

                        foreach (Match matchHeading3Below in matchesHeading3Belows)
                        {
                            find.Text = matchHeading3Below.Value;
                            find.Execute();
                            if (selection.Paragraphs[1].Range.Sentences.Count == 1)
                            {
                                //正则表达式匹配模式设为：前方出现开头标记、换行符回车符，阿拉伯数字一个及以上；如果选区文字匹配成功（为三级小标题），则将当前小标题的大纲级别设为3级加大纲级别偏移量
                                if (Regex.IsMatch(selection.Range.Text, @"(?<=^|\n|\r)\d+")) 
                                {
                                    selection.Paragraphs[1].OutlineLevel = WdOutlineLevel.wdOutlineLevel3 + outlineLevelOffset;
                                }
                            }
                            font.Name = "仿宋";
                            font.Size = heading3FontSize;
                            font.ColorIndex = WdColorIndex.wdBlack;
                            font.Bold = 1;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // “X是”编号设置
                        selection.HomeKey(WdUnits.wdStory);

                        // 定义“X是”编号正则表达式变量，匹配模式为：换行符回车符“。：:；;，,”，空格制表符任意多个，中文数字1个及以上，空格制表符任意多个，“是”；后方出现非分页符换行符回车符的字符2个及以上
                        Regex regExShiNum = new Regex(@"[\n\r。：:；;，,][ |\t]*[一二三四五六七八九十〇零]+[ |\t]*是(?=[^\f\n\r]{2,})", RegexOptions.Multiline);
                        MatchCollection matchesHeadingShis = regExShiNum.Matches(documentText); // 获取全文文字经过“X是”标记正则表达式匹配的结果

                        foreach (Match matchHeadingShi in matchesHeadingShis)
                        {
                            find.Text = matchHeadingShi.Value;
                            find.Execute();
                            selection.MoveStart(WdUnits.wdCharacter, 1); // 将选区的开头向后移动一个字符，避开前方的换行符回车符或标点
                            font.Name = "仿宋";
                            font.Size = headingShiFontSize;
                            font.ColorIndex = WdColorIndex.wdBlack;
                            font.Bold = 1;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // “条款项”编号设置
                        selection.HomeKey(WdUnits.wdStory);

                        // 定义“条款项”编号正则表达式变量，匹配模式为：从开头开始，“第”，空格制表符任意多个，阿拉伯数字或中文数字1个及以上，空格制表符任意多个，“条款项”，“：:”空格制表符1个及以上
                        Regex regExItemNum = new Regex(@"(?<=^|\n|\r)第[ |\t]*[\d一二三四五六七八九十〇零]+[ |\t]*[条款项][：:| |\t]", RegexOptions.Multiline); // 将正则匹配模式设为条款项编号
                        MatchCollection matchesHeadingItems = regExItemNum.Matches(documentText); // 获取全文文字经过条款项编号正则表达式匹配的结果

                        foreach (Match matchHeadingItem in matchesHeadingItems)
                        {
                            find.Text = matchHeadingItem.Value;
                            find.Execute();
                            font.Name = "黑体";
                            font.Size = headingItemFontSize;
                            font.ColorIndex = WdColorIndex.wdBlack;
                            font.Bold = 0;
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 遍历所有表格
                        foreach (Table table in msWordDocument.Tables)
                        {
                            // 表格上方标题、注释设置
                            table.Cell(1, 1).Select(); // 选择第1行第1列的单元格
                            selection.MoveUp(WdUnits.wdLine, 1, WdMovementType.wdMove); // 光标上移到表格上方一行
                            selection.Expand(WdUnits.wdLine); // 全选表格上方一行
                            selection.MoveStart(WdUnits.wdLine, -5); // 选区向上扩大5行

                            // 定义表格上方标题正则表达式变量，匹配模式为：从开头开始，非“。；;”分页符换行符回车符的字符1-30个，“表单册录执”，非“。；;”分页符换行符回车符的字符0-10个，换行符回车符
                            Regex regExTableHeadings = new Regex(@"(?<=^|\n|\r)[^。；;\f\n\r]{1,30}[表单册录执][^。；;\f\n\r]{0,10}[\n\r]", RegexOptions.Multiline);

                            MatchCollection matchesTableHeadings = regExTableHeadings.Matches(selection.Text); // 获取选区文字经过表格上方标题正则表达式匹配的结果

                            if (matchesTableHeadings.Count > 0) // 如果匹配到的结果集合元素数大于0
                            {
                                find.Text = matchesTableHeadings[0].Value;
                                find.Execute();
                                paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0);
                                paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                font.Name = "黑体";
                                font.Size = 16;
                                font.ColorIndex = WdColorIndex.wdBlack;
                                font.Bold = 0;
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
                            table.Range.Font.Name = "仿宋"; // 字体为仿宋
                            table.Range.Font.Color = WdColor.wdColorAutomatic; // 字体颜色设为自动
                            table.Range.Font.Size = 14; // 字号为四号
                            table.Range.Font.Kerning = 0; // “为字体调整字符间距”值设为0
                            table.Range.Font.DisableCharacterSpaceGrid = true;

                            table.Range.ParagraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0);
                            table.Range.ParagraphFormat.AutoAdjustRightIndent = 0; // 自动调整右缩进为false
                            table.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle; // 单倍行距

                            table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter; // 单元格内容垂直居中

                            // 自动调整表格
                            table.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthAuto; // 列宽度设为自动
                            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent); // 根据内容调整表格
                            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow); // 根据窗口调整表格
                        }

                        // 括号注释设置
                        selection.HomeKey(WdUnits.wdStory);

                        // 定义括号注释正则表达式变量，匹配模式为：从开头开始，“（(”，非“。”分页符换行符回车符的字符1-12个，“）)”，换行符回车符
                        Regex regExBrackets = new Regex(@"(?<=^|\n|\r)[（\(][^。\f\n\r]{1,12}[）\)][\n\r]", RegexOptions.Multiline);
                        MatchCollection matchesBrakets = regExBrackets.Matches(documentText); // 获取全文文字经过括号注释正则表达式匹配的结果

                        foreach (Match matchBraket in matchesBrakets)
                        {
                            find.Text = matchBraket.Value;
                            find.Execute();
                            if (selection.Information[WdInformation.wdWithInTable] == false) // 如果查找结果不在表格内
                            {
                                paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0
                                paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 居中对齐
                            }
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 附件注释设置
                        selection.HomeKey(WdUnits.wdStory);

                        // 定义附件注释正则表达式变量，匹配模式为：从开头开始，“附”，空格制表符任意多个，“件录”，非“。”分页符换行符回车符的字符0-3个，换行符回车符
                        Regex regExAppendixes = new Regex(@"(?<=^|\n|\r)附[ |\t]*[件录][^。\f\n\r]{0,3}[\n\r]", RegexOptions.Multiline);
                        MatchCollection matchesAppendixes = regExAppendixes.Matches(documentText); // 获取全文文字经过附件注释正则表达式匹配的结果

                        foreach (Match matchAppendix in matchesAppendixes)
                        {
                            find.Text = matchAppendix.Value;
                            find.Execute();
                            if (selection.Information[WdInformation.wdWithInTable] == false) // 如果查找结果不在表格内
                            {
                                paragraphFormat.FirstLineIndent = msWordApp.CentimetersToPoints(0); // 首行缩进设为0
                                paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; // 左对齐
                            }
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 单位和日期落款设置
                        selection.HomeKey(WdUnits.wdStory);

                        // 定义单位和日期落款正则表达式变量，匹配模式为：前方出现开头符号、换行符回车符，换行符回车符（一个空行），单位字符串1个及以上，最后为日期（如果日期都有明确数字，则可以用非中文符号分隔，否则只能用“年月日”标明）
                        Regex regExInscriptions = new Regex(@"(?<=^|\n|\r)[\n\r](?:[\u4E00-\u9FA5\w、：:（）\(\)| |\t]{2,}[\n\r])+(?:(?:(?:[12]\d{3}|[一二][一二三四五六七八九〇零]{3})[ |\t]*[年\.．\-/][ |\t]*"
                              + @"[\d一二三四五六七八九十元]{1,2}[\.．\-/\u4E00-\u9FA5\w（）\(\)| |\t]*)|(?:[ |\t]*年[ |\t]*月[ |\t]*日?))[\n\r]", RegexOptions.Multiline);
                        MatchCollection matchesInscriptions = regExInscriptions.Matches(documentText); // 获取全文文字经过单位和日期落款正则表达式匹配的结果

                        foreach (Match matchInscription in matchesInscriptions)
                        {
                            find.Text = matchInscription.Value;
                            find.Execute();
                            if (selection.Information[WdInformation.wdWithInTable] == false) // 如果查找结果不在表格内
                            {
                                foreach (Paragraph paragraph in selection.Paragraphs) // 遍历所有落款中的段落
                                {
                                    float rightIndentation = Math.Max(0, 10 - paragraph.Range.Text.Length / 2); // 计算右缩进量，如果右缩进量小于0，则限定为0
                                    paragraph.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight; // 右对齐
                                    paragraph.Format.CharacterUnitRightIndent = rightIndentation; // 右缩进设为之前计算值
                                    paragraph.Range.Font.Name = "仿宋";
                                    paragraph.Range.Font.Size = bodyTextFontSize;
                                    paragraph.Range.Font.ColorIndex = WdColorIndex.wdBlack;
                                    paragraph.Range.Font.Bold = 0;
                                }
                            }
                            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }

                        // 页脚页码设置
                        foreach (Section section in msWordDocument.Sections) // 遍历所有节
                        {
                            section.PageSetup.DifferentFirstPageHeaderFooter = 0;     // “首页页眉页脚不同”设为否
                            section.PageSetup.OddAndEvenPagesHeaderFooter = 0;        // “奇偶页页眉页脚不同”设为否

                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete(); // 删除页脚中的内容
                            // 设置页码
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleNumberInDash;  // 页码左右带横线； wdPageNumberStyleArabicFullWidth 阿拉伯数字全宽
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;  // 不续前节
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = 1;  // 从1开始编号
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.Add(WdPageNumberAlignment.wdAlignPageNumberOutside, FirstPage: true); // 页码奇数页靠右，偶数页靠左； wdAlignPageNumberInside  奇左偶右 wdAlignPageNumberCenter 页码居中
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name = "Times New Roman";
                            section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Size = footerFontSize;

                            section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete(); // 删除页眉中的内容
                            section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone; // 段落下边框线设为无
                        }

                        msWordDocument.Save(); // 保存Word文档
                        msWordDocument.Close(); // 关闭Word文档
                          
                    }

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "警告", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                finally
                {
                    msWordApp.ScreenUpdating = true;
                    KillOfficeApps(new object[] { msWordApp });
                }

            }
            await task;
        }

        //定义句子正则表达式变量，匹配模式为：非“。；;”字符任意多个，“。；;”
        public static Regex regExSentence = new Regex(@"[^。；;]*[。；;]");

        public static void PreprocessDocumentTexts(ExcelRange range)
        {
            foreach (ExcelRangeBase cell in range) // 遍历所有单元格
            {
                if (!cell.EntireRow.Hidden) // 如果当前单元格所在行不是隐藏行
                {
                    //将当前单元格文字按换行符拆分为数组（删除每个元素前后空白字符，并删除空白元素），转换成列表，赋值给拆分后文字列表
                    List<string>? lstSplittedTexts = cell.Text.Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .ToList();
                    int lstSplittedTextsCount = lstSplittedTexts!.Count; //获取拆分后文字列表元素个数

                    for (int i = 0; i < lstSplittedTextsCount; i++) //遍历拆分后文字列表的所有元素
                    {
                        //将拆分后文字列表当前元素的文字按修订标记字符拆分成数组（删除每个元素前后空白字符，并删除空白元素），转换成列表，移除每个元素的小标题编号，赋值给修订文字列表
                        List<string> lstRevisedTexts = lstSplittedTexts[i].Split('^', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                            .ToList().ConvertAll(e => RemoveHeadingNum(e));
                        
                        //合并修订文字列表中的所有元素成为完整字符串，重新赋值给拆分后文字列表当前元素
                        lstSplittedTexts[i] = MergeRevision(lstRevisedTexts);

                        string MergeRevision(List<string> lstStrs) //合并修订文字
                        {
                            if ((lstStrs?.Count ?? 0) == 0) //如果字符串列表的元素数（如果字符串列表为null，则得到0）为0，则将空字符串赋值给函数返回值
                            {
                                return string.Empty;
                            }

                            if (lstStrs!.Count == 1) //如果字符串列表的元素数为1，则将0号元素赋值给函数返回值
                            {
                                return lstStrs[0];
                            }

                            //获取字符串列表0号元素经过句子正则表达式匹配后的结果集合
                            MatchCollection matchesSentences = regExSentence.Matches(lstStrs[0]);

                            foreach (Match matchSentence in matchesSentences) //遍历所有句子正则表达式匹配的结果
                            {
                                int sameSentenceCount = 0;
                                for (int i = 1; i < lstStrs.Count; i++) //遍历字符串列表从1号开始的所有元素
                                {
                                    if (lstStrs[i].Contains(matchSentence.Value))  //如果字符串列表当前元素含有当前句子
                                    {
                                        lstStrs[i] = lstStrs[i].Replace(matchSentence.Value, ""); //将字符串列表当前元素中的当前句子替换为空（删除重复句）
                                        sameSentenceCount += 1; //相同句子计数加1
                                    }
                                }
                                //重新赋值给字符串列表0号元素：如果相同句子计数小于字符串列表元素数量减1（除0号元素外的其他元素并不都含有当前句子），则得到将0号元素中的当前句子替换为空后的字符串（删除非共有句）；否则得到0号元素原值
                                lstStrs[0] = sameSentenceCount < lstStrs.Count - 1 ? lstStrs[0].Replace(matchSentence.Value, "") : lstStrs[0];
                            }
                            return string.Join("", lstStrs);  //合并字符串列表的所有元素，赋值给函数返回值
                        }

                    }

                    if (lstSplittedTextsCount >= 2) // 如果拆分后文字列表的元素个数大于等于2个
                    {
                        int insertedRowsCount = lstSplittedTextsCount - 1; // 计算需要插入的行数：列表元素数-1
                        cell.Worksheet.InsertRow(cell.Start.Row + 1, insertedRowsCount); // 从被拆分单元格的下一个单元格开始，插入行
                    }

                    for (int i = 0; i < lstSplittedTextsCount; i++) //遍历拆分后文字列表的每个元素
                    {
                        cell.Offset(i, 0).Value = lstSplittedTexts[i]; //将拆分后文字列表当前元素赋值给当前单元格向下偏移i行的单元格
                        cell.CopyStyles(cell.Offset(i, 0)); //将当前单元格的样式复制到当前单元格向下偏移i行的单元格
                        cell.Offset(i, 0).EntireRow.CustomHeight = false; // 当前单元格向下偏移i行的单元格所在行的手动设置行高设为false（即为自动）   
                    }
                }
            }

        }

        public static async Task ProcessDocumentTableIntoWordAsync(string documentTableFilePath, string targetWordFilePath)
        {
            try
            {
                List<string> lstFullTexts = new List<string> { }; //定义完整文章列表变量

                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(documentTableFilePath))) //打开结构化文档表Excel工作簿，赋值给Excel包变量
                {

                    ExcelWorksheet bodyTextsWorksheet = excelPackage.Workbook.Worksheets[2]; //将“主体”（第3张，2号）工作表赋值给“主体”工作表变量
                    RemoveWorksheetEmptyRowsAndColumns(bodyTextsWorksheet); //删除“主体”工作表内所有空白行和空白列
                    if ((bodyTextsWorksheet.Dimension?.Rows ?? 0) <= 1) // 如果“主体”工作表已使用行数小于等于1（如果工作表为空，则为0），只有表头无有效数据，则结束本过程
                    {
                        return;
                    }
                    
                    ExcelWorksheet headingsWorksheet = excelPackage.Workbook.Worksheets[0]; // 将“小标题”（第1张，0号）工作表赋值给“小标题”工作表变量
                    // 删除工作表中的所有行
                    for (int i = headingsWorksheet.Dimension.End.Row; i >= 2; i--)
                    {
                        headingsWorksheet.DeleteRow(i);
                    }

                    while (excelPackage.Workbook.Worksheets.Count > 3) //当Excel工作簿中的工作表大于3张，则继续循环，删除最后一张
                    {
                        excelPackage.Workbook.Worksheets.Delete(excelPackage.Workbook.Worksheets.Count - 1);
                    }

                    excelPackage.Workbook.Worksheets.Copy(bodyTextsWorksheet.Name, $"备份{new Random().Next(1000, 10000)}"); //将“主体”Excel工作表复制为“备份”工作表
                    bodyTextsWorksheet.Select();

                    //在“主体”工作表第2行到最末行（如果工作表为空，则为第2行）的文字（第3）列中，将含有换行符的单元格文字拆分成多段，删除小标题编号，合并修订文字，将小标题句提到最前方，最后将各段分置于单独的行中
                    PreprocessDocumentTexts(bodyTextsWorksheet.Cells[2, 3, (bodyTextsWorksheet.Dimension?.End.Row ?? 2), 3]);

                    //将下方无正文的小标题行设为隐藏：
                    for (int i = 2; i <= bodyTextsWorksheet.Dimension!.End.Row; i++)
                    {
                        if (!bodyTextsWorksheet.Rows[i].Hidden) //如果当前行不是隐藏行
                        {
                            int paragraphsCount = 0;
                            if (bodyTextsWorksheet.Cells[i, 1].Text.Contains("级") && bodyTextsWorksheet.Cells[i, 3].Text.Length < 50) //如果当前行文字含小标题且字数小于50字（纯小标题行，基准小标题行）
                            {
                                if (i < bodyTextsWorksheet.Dimension.Rows)  //如果基准小标题行不为最后一行
                                {
                                    for (int k = i + 1; k <= bodyTextsWorksheet.Dimension.End.Row; k++)  //遍历从基准小标题行的下一行开始直到最后一行的所有行（比较行）
                                    {
                                        if (!bodyTextsWorksheet.Rows[k].Hidden)  //如果当前比较行不是隐藏行
                                        {
                                            //如果当前比较行文字含小标题且小标题级别数小于等于基准小标题行（小标题级别更高或相同），则退出循环
                                            if (bodyTextsWorksheet.Cells[k, 1].Text.Contains("级") && Val(bodyTextsWorksheet.Cells[k, 1].Text) <= Val(bodyTextsWorksheet.Cells[i, 1].Text))
                                            {
                                                break;
                                            }
                                            //否则，如果当前比较行文字不含小标题或者字数大于等于50（视为正文），则正文段落计数加1
                                            else if (!bodyTextsWorksheet.Cells[k, 1].Text.Contains("级") || bodyTextsWorksheet.Cells[k, 3].Text.Length >= 50)
                                            {
                                                paragraphsCount++;
                                            }
                                        }
                                    }
                                    if (paragraphsCount == 0) bodyTextsWorksheet.Rows[i].Hidden = true; //如果累计的正文段落数为零（基准小标题下方无正文），则将基准小标题行隐藏
                                }
                                else //否则，则将当前行（基准小标题行）隐藏
                                {
                                    bodyTextsWorksheet.Rows[i].Hidden = true;
                                }
                            }
                        }
                    }

                    //初始化小标题编号变量
                    int heading0Num = 1;
                    int heading1Num = 1;
                    int heading2Num = 1;
                    int heading3Num = 1;
                    int heading4Num = 1;
                    int headingShiNum = 1;
                    int headingItemNum = 1;

                    bodyTextsWorksheet.Cells[2, 2, bodyTextsWorksheet.Dimension.End.Row, 2].Clear(); // 清除第2列旧小标题编号

                    for (int i = 2; i <= bodyTextsWorksheet.Dimension.End.Row; i++) //遍历“主体”工作表第2行开始到最末行的所有行
                    {
                        if (!bodyTextsWorksheet.Rows[i].Hidden) // 如果当前行不是隐藏行
                        {
                            // 给小标题编号
                            bool checkHeadingNecessity = false; // “检查小标题编号必要性”变量初始赋值为False
                            switch (bodyTextsWorksheet.Cells[i, 1].Text) //根据当前行小标题级别进入相应的分支，将对应级别的小标题编号分别赋值给小标题编号单元格
                            {
                                case "0级": //如果为0级小标题
                                    bodyTextsWorksheet.Cells[i, 2].Value = "第" + ConvertArabicNumberIntoChinese(Convert.ToInt32(heading0Num)) + "部分 "; //将0级小标题编号赋值给小标题编号单元格
                                    checkHeadingNecessity = heading0Num == 1 ? true : false; // 获取“检查小标题编号必要性”值：如果编号为1，则得到true；否则，得到false（防止同级编号只有1没有2）
                                    heading0Num++; //0级小标题计数加1
                                    heading1Num = 1;
                                    heading2Num = 1;
                                    heading3Num = 1;
                                    heading4Num = 1;
                                    headingShiNum = 1;
                                    break;
                                case "1级":
                                    bodyTextsWorksheet.Cells[i, 2].Value = ConvertArabicNumberIntoChinese(Convert.ToInt32(heading1Num)) + "、";
                                    checkHeadingNecessity = heading1Num == 1 ? true : false;
                                    heading1Num++;
                                    heading2Num = 1;
                                    heading3Num = 1;
                                    heading4Num = 1;
                                    headingShiNum = 1;
                                    break;
                                case "2级":
                                    bodyTextsWorksheet.Cells[i, 2].Value = "（" + ConvertArabicNumberIntoChinese(Convert.ToInt32(heading2Num)) + "）";
                                    checkHeadingNecessity = heading2Num == 1 ? true : false;
                                    heading2Num++;
                                    heading3Num = 1;
                                    heading4Num = 1;
                                    headingShiNum = 1;
                                    break;
                                case "3级":
                                    bodyTextsWorksheet.Cells[i, 2].Style.Numberformat.Format = "@";
                                    bodyTextsWorksheet.Cells[i, 2].Value = heading3Num + ".";
                                    checkHeadingNecessity = heading3Num == 1 ? true : false;
                                    heading3Num++;
                                    heading4Num = 1;
                                    headingShiNum = 1;
                                    break;
                                case "4级":
                                    bodyTextsWorksheet.Cells[i, 2].Style.Numberformat.Format = "@";
                                    bodyTextsWorksheet.Cells[i, 2].Value = "（" + heading4Num + "）";
                                    checkHeadingNecessity = heading4Num == 1 ? true : false;
                                    heading4Num++;
                                    headingShiNum = 1;
                                    break;
                                case "是":
                                    bodyTextsWorksheet.Cells[i, 2].Value = ConvertArabicNumberIntoChinese(Convert.ToInt32(headingShiNum)) + "是";
                                    checkHeadingNecessity = headingShiNum == 1 ? true : false;
                                    headingShiNum++;
                                    break;
                                case "条":
                                    bodyTextsWorksheet.Cells[i, 2].Value = "第" + ConvertArabicNumberIntoChinese(Convert.ToInt32(headingItemNum)) + "条 ";
                                    checkHeadingNecessity = headingItemNum == 1 ? true : false;
                                    headingItemNum++;
                                    break;
                            }

                            //删除多余的小标题编号（如果同级小标题编号只有1没有2，则将编号1删去）
                            if (checkHeadingNecessity) // 如果需要检查小标题编号的必要性（当前小标题的编号为1）
                            {
                                int headingsCount = 1;
                                if (i < bodyTextsWorksheet.Dimension.End.Row)  // 如果当前行（基准小标题行）不为最后一行
                                {
                                    for (int k = i + 1; k <= bodyTextsWorksheet.Dimension.End.Row; k++)  // 遍历从基准行的下一行开始直到最后一行的所有行（比较行）
                                    {
                                        if (!bodyTextsWorksheet.Rows[k].Hidden)  // 如果当前比较行不是隐藏行
                                        {
                                            // 如果当前比较行文字含小标题且小标题级别数小于基准行（小标题级别更高），则退出循环
                                            if (bodyTextsWorksheet.Cells[k, 1].Text.Contains("级") && Val(bodyTextsWorksheet.Cells[k, 1].Text) < Val(bodyTextsWorksheet.Cells[i, 1].Text))
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
                                else // 否则，将当前行（基准小标题行）的小标题编号单元格清空
                                {
                                    bodyTextsWorksheet.Cells[i, 2].Value = "";
                                }
                                // 如果基准行小标题编号单元格为空，且文字字数少于50字（视为多余的纯小标题）
                                if (bodyTextsWorksheet.Cells[i, 2].Text == "" && bodyTextsWorksheet.Cells[i, 3].Text.Length < 50)
                                {
                                    bodyTextsWorksheet.Cells[i, 2].Value = "X"; // 将当前行（基准小标题行）小标题编号单元格赋值为“X”（忽略行）
                                }
                            }
                        }
                        else // 否则，则将小标题编号单元格赋值为“X”（忽略行）
                        {
                            bodyTextsWorksheet.Cells[i, 2].Value = "X";
                        }
                    }

                    ExcelWorksheet titleWorksheet = excelPackage.Workbook.Worksheets[1]; //将“大标题首尾”工作表（第2张，1号）赋值给大标题首尾工作表变量
                    ExcelRange titleCells = titleWorksheet.Cells; //将“大标题首尾”工作表单元格赋值给大标题首尾工作表单元格变量

                    lstFullTexts.AddRange(new string[] { titleCells["C2"].Text, "", titleCells["C3"].Text }); //将大标题、空行、首段添加到完整文章列表中

                    for (int i = 2; i <= bodyTextsWorksheet.Dimension.End.Row; i++)  // 遍历“主体”工作表第2行到最末行的所有行
                    {
                        string headingText = ""; // 小标题文字变量赋值为空
                        if (bodyTextsWorksheet.Cells[i, 2].Text != "X")  // 如果当前行没有"X"标记（非忽略行）
                        {
                            //更新“小标题”工作表
                            if (bodyTextsWorksheet.Cells[i, 1].Text.Contains("级")) // 如果当前行含小标题
                            {
                                MatchCollection matchesHeadingTexts = regExHeadingText.Matches(bodyTextsWorksheet.Cells[i, 3].Text);  // 获取当前行的小标题正文文字经过小标题文字正则表达式匹配的结果
                                // 获取小标题文字：如果正则表达式匹配到的结果集合元素数大于0，得到匹配到的小标题文字；否则得到空字符串
                                headingText = matchesHeadingTexts.Count > 0 ? matchesHeadingTexts[0].Value : "";

                                // 更新“小标题”工作表内容
                                int lastRowIndex = headingsWorksheet.Dimension?.End.Row ?? 0; //获取“小标题”工作表最末行的索引号（如果工作表为空，则为0）
                                //将“小标题”工作表第一个空白行第1至6列的单元格赋值给小标题单元格组变量
                                ExcelRange headingsCells = headingsWorksheet.Cells[lastRowIndex + 1, 1, lastRowIndex + 1, 6];
                                headingsCells.Style.Numberformat.Format = "@"; // 将小标题单元格组的格式设为文本
                                //将当前行的小标题级别、编号、正文、完成时限、责任部门、分类赋值给小标题单元格组
                                headingsCells.LoadFromArrays(new List<object[]> { new object[]
                                    {bodyTextsWorksheet.Cells[i, 1].Text, bodyTextsWorksheet.Cells[i, 2].Text, headingText,
                                     bodyTextsWorksheet.Cells[i, 4].Text, bodyTextsWorksheet.Cells[i, 5].Text, bodyTextsWorksheet.Cells[i, 6].Text} });

                            }

                            //将当前行的小标题编号和小标题正文文字添加到完整文章列表
                            string paragraphText = bodyTextsWorksheet.Cells[i, 2].Text + bodyTextsWorksheet.Cells[i, 3].Text; //将当前行小标题编号和文字合并，赋值给段落文字变量
                            lstFullTexts.Add(paragraphText); //将段落文字添加到完整文章列表
                        }
                    }

                    // 获取日期单元格的日期值并转换为字符串
                    string dateStr = titleCells["C6"].GetValue<DateTime>().ToString("yyyy年M月d日"); // 将日期值转换为字符串

                    //将尾段、空行、落款单位、日期依次添加到完整文章列表中
                    lstFullTexts.AddRange(new string[] { titleCells["C4"].Text, "", titleCells["C5"].Text, dateStr });

                    FormatDocumentTable(excelPackage.Workbook); // 格式化结构化文档表中的所有工作表
                    excelPackage.Save(); //保存Excel工作簿
                }

                //获取目标缩短版Word文档文件全名路径
                string targetShortWordFilePath = Path.Combine(Path.GetDirectoryName(targetWordFilePath)!, $"摘要_{Path.GetFileName(targetWordFilePath)}");

                DocX targetWordDocument = DocX.Create(targetWordFilePath); //新建Word文档，赋值给目标Word文档变量
                DocX targetShortWordDocument = DocX.Create(targetShortWordFilePath); //新建缩短版Word文档，赋值给目标缩短版Word文档变量

                for (int i = 0; i < lstFullTexts.Count; i++)  //遍历完整文章列表中的所有元素
                {
                    targetWordDocument.InsertParagraph(lstFullTexts[i]); //将当前元素的段落文字插入目标Word文档
                    targetShortWordDocument.InsertParagraph(ProceedToExtractText(lstFullTexts[i], '\0', 100)); //将当前元素的段落文字缩短至目标字数，插入目标缩短版Word文档
                }

                targetWordDocument.Save(); //保存目标Word文档
                targetShortWordDocument.Save(); //保存目标缩短版Word文档
                targetWordDocument.Dispose(); //关闭目标Word文档
                targetShortWordDocument.Dispose(); //关闭目标缩短版Word文档

                //如果对话框返回值为OK（点击了OK），则对目标Word文档执行排版过程
                if (MessageBox.Show("是否需要排版？", "询问", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    await FormatWordDocumentsAsync(new List<string> { targetWordFilePath });
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "警告", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        public static void SetPandocPath()
        {
            try 
            {
                string currentPandocPath = Properties.Settings.Default.pandocPath; //读取设置中保存的Pandoc程序文件路径全名，赋值给当前Pandoc程序文件路径全名变量
                InputDialog inputDialog = new InputDialog("输入Pandoc.exe程序文件路径", currentPandocPath); //弹出对话框，输入Pandoc程序文件路径全名
                if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则结束本过程
                {
                    return;
                }

                currentPandocPath = inputDialog.Answer; //获取对话框返回的当前Pandoc程序文件路径全名
                //如果当前Pandoc程序文件路径全名的文件不存在或转换为小写后不以“.exe”结尾，则抛出异常
                if (!File.Exists(currentPandocPath) || !currentPandocPath.ToLower().EndsWith(".exe"))
                {
                    throw new Exception("程序文件不存在或不合法！");
                }
                Properties.Settings.Default.pandocPath = currentPandocPath; // 将对话框返回的当前Pandoc程序文件路径全名存入设置
                Properties.Settings.Default.Save();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "警告", MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }

        public static void TrimCellsStrings(ExcelWorksheet excelWorksheet, bool covertAllTypesToString = false)
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

        public static Regex regExNumber = new Regex(@"\d+\.?\d*"); //定义数字正则表达式变量，匹配模式为：阿拉伯数字一个及以上，小数点至多一个，阿拉伯数字任意多个

        public static double Val(object? cellValue)
        {
            if (cellValue == null) //如果参数为null，将0赋值给函数返回值
            {
                return 0;
            }

            string cellStr = Convert.ToString(cellValue)!;
            Match matchNumVal = regExNumber.Match(cellStr); //获取字符串经过数字正则表达式匹配的第一个结果

            if (matchNumVal.Success) //如果被正则表达式匹配成功
            {
                //如果将匹配结果转换为double类型成功，则将转换结果赋值给number变量，然后再将number变量值赋值给函数返回值
                if (double.TryParse(matchNumVal.Value, out double number))
                {
                    return number;
                }
            }
            return 0; //如果以上过程均没有赋值给函数返回值，此处将0赋值给函数返回值
        }

    }
}
