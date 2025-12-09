using DocSharp.Binary.DocFileFormat;
using DocSharp.Binary.OpenXmlLib;
using DocSharp.Binary.Spreadsheet.XlsFileFormat;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Markdown;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using static COMIGHT.Methods;
using DataTable = System.Data.DataTable;
using DocSharpSpreadsheetMapping = DocSharp.Binary.SpreadsheetMLMapping;
using DocSharpWordMapping = DocSharp.Binary.WordprocessingMLMapping;
using ITextDocument = iText.Layout.Document;
using ITextParagraph = iText.Layout.Element.Paragraph;
using SpreadsheetDocument = DocSharp.Binary.OpenXmlLib.SpreadsheetML.SpreadsheetDocument;
using Task = System.Threading.Tasks.Task;
using Window = System.Windows.Window;
using WordprocessingDocument = DocSharp.Binary.OpenXmlLib.WordprocessingML.WordprocessingDocument;
using static COMIGHT.UniversalObjects;

namespace COMIGHT
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    /// 

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            ExcelPackage.License.SetNonCommercialPersonal("Yuechen Lou"); //定义EPPlus库许可证类型为非商用

            appSettings = settingsManager.GetSettings(); // 从应用设置管理器中读取应用设置，赋值给应用设置对象变量
            latestRecords = recordsManager.GetSettings(); // 从用户使用记录管理器中读取用户使用记录，赋值给用户使用记录对象变量

        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Top = 50.0;
            this.Left = SystemParameters.WorkArea.Width - this.Width - 150.0;

            lblStatus.DataContext = taskManager; // 将状态标签控件的数据环境设为任务管理器对象
            lblIntro.Content = $"For Better Productivity. © Yuechen Lou 2022-{DateTime.Now:yyyy}";

            try
            {
                CreateFolder(appSettings.SavingFolderPath); // 创建保存文件夹
            }
            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }

        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            settingsManager.SaveSettings(appSettings); // 保存应用程序设置到Json文件
            recordsManager.SaveSettings(latestRecords);  // 保存最近记录到Json文件

            Environment.Exit(0); // 退出程序，关闭所有窗口
        }


        private void MnuBatchConvertOfficeFileTypes_Click(object sender, RoutedEventArgs e)
        {
            BatchConvertOfficeFileTypes();
        }

        private async void MnuBatchCreateFolders_Click(object sender, RoutedEventArgs e)
        {
            await BatchCreateFolders();
        }

        private void MnuBatchCreatePlaceCards_Click(object sender, RoutedEventArgs e)
        {
            BatchCreatePlaceCards();
        }

        private void MnuBatchExtractTablesFromWord_Click(object sender, RoutedEventArgs e)
        {
            BatchExtractTablesFromWord();
        }

        private async void MnuBatchFormatWordDocuments_Click(object sender, RoutedEventArgs e)
        {
            await BatchFormatWordDocumentsAsync();
        }

        private void MnuBatchProcessExcelWorksheets_Click(object sender, RoutedEventArgs e)
        {
            BatchProcessExcelWorksheets();
        }

        private async void MnuBatchRepairWordDocuments_Click(object sender, RoutedEventArgs e)
        {
            await BatchRepairWordDocumentsAsync();
        }

        private void MnuBatchUnhideExcelWorksheets_Click(object sender, RoutedEventArgs e)
        {
            BatchUnhideExcelWorksheets();
        }

        //private void MnuBrowser_Click(object sender, RoutedEventArgs e)
        //{
        //    if (GetInstanceCountByHandle<BrowserWindow>() < 3) //如果被打开的浏览器窗口数量小于3个，则新建一个浏览器窗口实例并显示
        //    {
        //        BrowserWindow browserWindow = new BrowserWindow();
        //        browserWindow.Show();
        //    }
        //}

        private void MnuConvertMarkdownIntoWord_Click(object sender, RoutedEventArgs e)
        {
            ConvertMarkdownIntoWord();
        }

        private void MnuCreateFileList_Click(object sender, RoutedEventArgs e)
        {
            CreateFileList();
        }

        private void MnuExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void MnuHelp_Click(object sender, RoutedEventArgs e)
        {
            ShowHelpWindow();
        }

        private void MnuMergeDataIntoDocument_Click(object sender, RoutedEventArgs e)
        {
            MergeDataIntoDocument();
        }

        private void MnuOpenSavingFolder_Click(object sender, RoutedEventArgs e)
        {
            OpenSavingFolder();
        }

        private void MnuRemoveMarkdownMarksInCopiedText_Click(object sender, RoutedEventArgs e)
        {
            RemoveMarkdownMarksInCopiedText();
        }

        private void MnuSettings_Click(object sender, RoutedEventArgs e)
        {
            SettingsWindow settingDialog = new SettingsWindow();
            settingDialog.ShowDialog();
        }

        private void MnuBatchDisassembleExcelWorkbooks_Click(object sender, RoutedEventArgs e)
        {
            BatchDisassembleExcelWorkbooks();
        }

        public void BatchConvertOfficeFileTypes()
        {
            try
            {
                List<string>? filePaths = SelectFiles(EnumFileType.Convertible, true, "Select Old Version Office or WPS Files"); //获取所选文件列表

                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                string folderPath = Path.GetDirectoryName(filePaths[0])!; //获取保存转换文件的文件夹路径

                //定义可用Excel打开的文件正则表达式变量，匹配模式为: "xls"或"et"，结尾标记，忽略大小写
                Regex regExExcelFile = new Regex(@"(?:xls|et)$", RegexOptions.IgnoreCase);
                //定义可用Word打开的文件正则表达式，匹配模式为: "doc"或"wps"，结尾标记，忽略大小写
                Regex regExWordFile = new Regex(@"(?:doc|wps)$", RegexOptions.IgnoreCase);

                foreach (string filePath in filePaths) //遍历所有文件
                {
                    if (new FileInfo(filePath).Length == 0) //如果当前文件大小为0，则直接跳过当前循环并进入下一个循环
                    {
                        continue;
                    }

                    if (regExExcelFile.IsMatch(filePath)) //如果当前文件名被可用Excel打开的文件正则表达式匹配成功
                    {
                        // 获取目标Excel文件路径全名
                        string targetFilePath = Path.Combine(folderPath, $"{Path.GetFileNameWithoutExtension(filePath)}.xlsx"); //获取目标文件路径全名
                        //获取目标文件路径全名：如果目标文件不存在，则得到原目标文件路径全名；否则，在原目标文件主名后添加4位随机数，得到新目标文件路径全名
                        targetFilePath = !File.Exists(targetFilePath) ? targetFilePath :
                            Path.Combine(folderPath, $"{Path.GetFileNameWithoutExtension(filePath)}{new Random().Next(1000, 10000)}.xlsx");

                        using (StructuredStorageReader reader = new StructuredStorageReader(filePath)) //使用结构化存储读取器读取当前文件
                        {
                            SpreadsheetDocumentType outputType = SpreadsheetDocumentType.Workbook; // 定义输出文件类型为Workbook
                            XlsDocument xls = new XlsDocument(reader); // 创建Xls对象
                            using (SpreadsheetDocument xlsx = SpreadsheetDocument.Create(targetFilePath, outputType)) //  创建xlsx目标文件
                            {
                                DocSharpSpreadsheetMapping.Converter.Convert(xls, xlsx); // 将xls文件转换为xlsx文件
                            }
                        }
                    }

                    else if (regExWordFile.IsMatch(filePath)) //如果当前文件名被可用Word打开的文件正则表达式匹配成功
                    {
                        string targetFilePath = Path.Combine(folderPath, $"{Path.GetFileNameWithoutExtension(filePath)}.docx"); //获取目标Word文件路径全名
                        //获取目标文件路径全名：如果目标文件不存在，则得到原目标文件路径全名；否则，在原目标文件主名后添加4位随机数，得到新目标文件路径全名
                        targetFilePath = !File.Exists(targetFilePath) ? targetFilePath :
                            Path.Combine(folderPath, $"{Path.GetFileNameWithoutExtension(filePath)}{new Random().Next(1000, 10000)}.docx");

                        using (StructuredStorageReader reader = new StructuredStorageReader(filePath)) // 使用结构化存储读取器读取当前文件
                        {
                            WordprocessingDocumentType outputType = WordprocessingDocumentType.Document; // 定义输出文件类型为Document
                            WordDocument doc = new WordDocument(reader); //  创建Word对象
                            using (WordprocessingDocument docx = WordprocessingDocument.Create(targetFilePath, outputType)) // 创建docx目标文件
                            {
                                DocSharpWordMapping.Converter.Convert(doc, docx); // 将doc文件转换为docx文件
                            }
                        }

                    }
                    File.Delete(filePath); //删除当前文件
                }

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        // 定义批量拆分工作簿功能选项枚举
        public enum EnumDissambleFunction
        {
            Cancel = 0,
            SplitByColumnIntoWorkbooks = 1,
            SplitByColumnIntoWorksheets = 2,
            DissembleByWorksheets = 3
        }

        public void BatchDisassembleExcelWorkbooks()
        {
            try
            {
                // 定义功能选项列表
                List<string> lstFunctions = new List<string> { "0-Cancel", "1-Split by a Column into Workbooks", "2-Split by a Column into Worksheets", "3-Dissemble by Worksheets" };

                //获取功能选项
                int functionNum = SelectFunction(lstOptions: lstFunctions, objRecords: latestRecords, propertyName: nameof(latestRecords.LatestBatchDisassembleWorkbooksOption));
                if (functionNum <= 0) //如果功能选项小于等于0（选择“Cancel”或不在设定范围），则结束本过程
                {
                    return;
                }

                EnumDissambleFunction function = (EnumDissambleFunction)functionNum; // 将功能选项枚举的整数值转换为枚举值

                List<string>? filePaths = SelectFiles(EnumFileType.Excel, true, "Select the Excel Files"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                // 定义Excel工作簿索引范围、表头表尾行数、列符号变量（初始值均为非法数值或空）
                int excelWorksheetStartIndex = -1; int excelWorksheetEndIndex = -1;
                int headerRowCount = -1; int footerRowCount = -1;
                string? columnLetter = null;
                bool createDataDict = false; // 定义“创建数据字典”变量，初始值为false

                (excelWorksheetStartIndex, excelWorksheetEndIndex) = GetWorksheetRange(); // 获取Excel工作表索引范围
                if (excelWorksheetStartIndex < 0 || excelWorksheetEndIndex < 0) // 如果获取到的Excel工作表索引号有一个小于0（范围无效），则结束本过程
                {
                    return;
                }

                switch (function) // 根据功能选项进入相应分支
                {
                    case EnumDissambleFunction.SplitByColumnIntoWorkbooks: // 按列拆分为Excel工作簿
                    case EnumDissambleFunction.SplitByColumnIntoWorksheets: // 按列拆分为Excel工作表

                        (headerRowCount, footerRowCount) = GetHeaderAndFooterRowCount(); //获取表头、表尾行数; 
                        if (headerRowCount < 0 || footerRowCount < 0) //如果获取到的表头、表尾行数有一个小于0（范围无效），则结束本过程
                        {
                            return;
                        }

                        columnLetter = GetKeyColumnLetter(); //获取主键列符
                        if (columnLetter == null) //如果主键列符为null，则结束本过程
                        {
                            return;
                        }

                        createDataDict = true; // “是否创建数据字典”赋值为True

                        break;

                    default:
                        createDataDict = false; // “是否创建数据字典”赋值为False

                        break;
                }

                Dictionary<string, List<ExcelRangeBase>> dataDict = new Dictionary<string, List<ExcelRangeBase>>(); // 定义数据字典（保存按列拆分的数据）

                string targetFolderPath = appSettings.SavingFolderPath; // 获取保存文件夹路径

                //定义集合工作簿变量、集合工作表计数变量和集合工作簿文件主名变量（功能4“集合工作簿”时使用）
                ExcelPackage assembledExcelPackage = new ExcelPackage(); // 新建Excel包，赋值给目标Excel包变量（为当前工作表创建一个新工作簿）
                string assembledExcelFileMainName = Path.GetFileNameWithoutExtension(filePaths[0]); // 获取集合Excel文件主名

                foreach (string filePath in filePaths) // 遍历文件列表
                {

                    string excelWorkbookFileMainName = Path.GetFileNameWithoutExtension(filePath); //获取当前Excel工作簿文件主名

                    // 创建目标文件夹（为每个工作簿创建一个独立文件夹）
                    targetFolderPath = Path.Combine(appSettings.SavingFolderPath, excelWorkbookFileMainName);
                    CreateFolder(targetFolderPath);

                    using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath))) // 打开当前Excel工作簿，赋值给Excel包变量
                    {
                        ExcelWorkbook excelWorkbook = excelPackage.Workbook;

                        //获取被处理Excel工作表索引号的起始值和结束值，如果大于工作表数量-1，则限定为工作表数量-1 (EPPlus工作表索引号从0开始，Excel工作表索引号从1开始)
                        excelWorksheetStartIndex = Math.Min(excelWorksheetStartIndex, excelWorkbook.Worksheets.Count - 1);
                        excelWorksheetEndIndex = Math.Min(excelWorksheetEndIndex, excelWorkbook.Worksheets.Count - 1);

                        for (int i = excelWorksheetStartIndex; i <= excelWorksheetEndIndex; i++) //遍历所有指定范围的Excel工作表
                        {
                            ExcelWorksheet excelWorksheet = excelWorkbook.Worksheets[i]; // 将当前索引号的Excel工作表赋值给Excel工作表变量

                            TrimCellStrings(excelWorksheet); //删除Excel工作表内所有文本型单元格值的首尾空格
                            RemoveWorksheetEmptyRowsAndColumns(excelWorksheet); //删除Excel工作表内所有空白行和空白列

                            if (createDataDict) // 如果需要创建数据字典（按列拆分时）
                            {
                                dataDict.Clear(); // 清空数据字典
                                if ((excelWorksheet.Dimension?.Rows ?? 0) <= headerRowCount + footerRowCount) //如果当前Excel工作表已使用行数（如果工作表为空， 则为0）小于等于表头表尾行数和，则直接跳过并进入下一个过程（结束当前工作表，跳至下一个工作表）
                                {
                                    continue;
                                }

                                for (int j = headerRowCount + 1; j <= excelWorksheet.Dimension!.End.Row - footerRowCount; j++) // 遍历Excel工作表除去表头、表尾的每一行
                                {
                                    string key = !string.IsNullOrWhiteSpace(excelWorksheet.Cells[columnLetter + j.ToString()].Text) ?
                                        excelWorksheet.Cells[columnLetter + j.ToString()].Text : "-Blank-"; //将当前行拆分基准列的值赋值给键值变量：如果当前行单元格文字不为空，则得到得到单元格文字，否则得到"-Blank-"
                                    if (dataDict.ContainsKey(key)) // 如果字典中已经有这个键，就将当前行添加到对应的列表中
                                    {
                                        dataDict[key].Add(excelWorksheet.Cells[j, 1, j, excelWorksheet.Dimension.End.Column]);
                                    }
                                    else // 否则，定义一个列表并向其中添加当前行，而后将列表并添加到字典中
                                    {
                                        dataDict[key] = new List<ExcelRangeBase> { excelWorksheet.Cells[j, 1, j, excelWorksheet.Dimension.End.Column] };
                                    }
                                }
                            }

                            switch (function) //根据功能序号进入相应的分支
                            {
                                case EnumDissambleFunction.SplitByColumnIntoWorkbooks: //按列拆分为Excel工作簿

                                    foreach (KeyValuePair<string, List<ExcelRangeBase>> pair in dataDict) // 遍历字典中的每一个键值对
                                    {
                                        using (ExcelPackage targetExcelPackage = new ExcelPackage()) //新建Excel包，赋值给目标Excel包变量（为当前工作表的每个键值对创建一个新工作簿）
                                        {
                                            ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add("Sheet1"); // 新建Excel工作表，赋值给目标工作表变量

                                            // 将表头复制到目标Excel工作表
                                            if (headerRowCount >= 1) //如果表头行数大于等于1，复制表头
                                            {
                                                excelWorksheet.Cells[1, 1, headerRowCount, excelWorksheet.Dimension.End.Column].CopyStyles(targetExcelWorksheet.Cells["A1"]);
                                                excelWorksheet.Cells[1, 1, headerRowCount, excelWorksheet.Dimension.End.Column].Copy(targetExcelWorksheet.Cells["A1"]);
                                            }

                                            // 将字典中的每一行复制到目标Excel工作表
                                            foreach (ExcelRangeBase dictRow in pair.Value) //遍历所有字典数据
                                            {
                                                //获取目标Excel工作表最末行索引号（如果工作表为空， 则为0）
                                                int lastRowIndex = targetExcelWorksheet.Dimension?.End.Row ?? 0;
                                                dictRow.CopyStyles(targetExcelWorksheet.Cells[lastRowIndex + 1, 1]); //将当前行的样式复制到目标Excel工作表的第一个非空白行
                                                dictRow.Copy(targetExcelWorksheet.Cells[lastRowIndex + 1, 1]); //将当前行的数据复制到目标Excel工作表的第一个非空白行
                                            }

                                            FormatExcelWorksheet(targetExcelWorksheet, headerRowCount, 0); //设置目标Excel工作表格式

                                            // 保存目标Excel工作簿文件
                                            FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"{pair.Key}_{excelWorksheet.Name}_{excelWorkbookFileMainName}")}.xlsx")); //获取目标Excel工作簿文件路径全名信息
                                            targetExcelPackage.SaveAs(targetExcelFile);
                                        }
                                    }

                                    break;

                                case EnumDissambleFunction.SplitByColumnIntoWorksheets:  //按列拆分为Excel工作表

                                    using (ExcelPackage targetExcelPackage = new ExcelPackage()) // 新建Excel包，赋值给目标Excel包变量（为当前工作表创建一个新工作簿）
                                    {

                                        foreach (KeyValuePair<string, List<ExcelRangeBase>> pair in dataDict) //遍历所有字典数据
                                        {
                                            // 新建Excel工作表，赋值给目标工作表变量
                                            ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add(CleanWorksheetName($"{pair.Key}"));

                                            // 将表头复制到目标Excel工作表
                                            if (headerRowCount >= 1) //如果表头行数大于等于1，复制表头
                                            {
                                                excelWorksheet.Cells[1, 1, headerRowCount, excelWorksheet.Dimension.End.Column].CopyStyles(targetExcelWorksheet.Cells["A1"]);
                                                excelWorksheet.Cells[1, 1, headerRowCount, excelWorksheet.Dimension.End.Column].Copy(targetExcelWorksheet.Cells["A1"]);
                                            }

                                            // 将字典中的每一行复制到目标Excel工作表
                                            foreach (ExcelRangeBase dictRow in pair.Value) //遍历所有字典数据
                                            {
                                                //获取目标Excel工作表最末行索引号（如果工作表为空， 则为0）
                                                int lastRowIndex = targetExcelWorksheet.Dimension?.End.Row ?? 0;
                                                dictRow.CopyStyles(targetExcelWorksheet.Cells[lastRowIndex + 1, 1]); //将当前行的样式复制到目标Excel工作表的第一个非空白行
                                                dictRow.Copy(targetExcelWorksheet.Cells[lastRowIndex + 1, 1]); //将当前行的数据复制到目标Excel工作表的第一个非空白行
                                            }

                                            FormatExcelWorksheet(targetExcelWorksheet, headerRowCount, 0); //设置目标Excel工作表格式

                                        }

                                        // 保存目标Excel工作簿文件
                                        FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"{excelWorksheet.Name}_{excelWorkbookFileMainName}")}.xlsx")); //获取目标Excel工作簿文件路径全名信息
                                        targetExcelPackage.SaveAs(targetExcelFile);

                                    }

                                    break;

                                case EnumDissambleFunction.DissembleByWorksheets: // 拆分工作表到独立工作簿

                                    using (ExcelPackage targetExcelPackage = new ExcelPackage()) // 新建Excel包，赋值给目标Excel包变量（为当前工作表创建一个新工作簿）
                                    {
                                        ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add("Sheet1");  // 新建Excel工作表

                                        // 将当前整个工作表复制到目标Excel工作表
                                        excelWorksheet.Cells[excelWorksheet.Dimension.Address].CopyStyles(targetExcelWorksheet.Cells["A1"]);
                                        excelWorksheet.Cells[excelWorksheet.Dimension.Address].Copy(targetExcelWorksheet.Cells["A1"]);

                                        FormatExcelWorksheet(targetExcelWorksheet, 0, 0); //设置目标Excel工作表格式

                                        // 保存目标Excel工作簿文件
                                        FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"{excelWorksheet.Name}_{excelWorkbookFileMainName}")}.xlsx")); //获取目标Excel工作簿文件路径全名信息
                                        targetExcelPackage.SaveAs(targetExcelFile);

                                    }

                                    break;
                            }

                        }
                    }

                }

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        private void BatchExtractTablesFromWord()
        {
            try
            {
                List<string>? filePaths = SelectFiles(EnumFileType.Word, true, "Select Word Files"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                foreach (string filePath in filePaths) // 遍历所有文件
                {
                    if (new FileInfo(filePath).Length == 0) //如果当前文件大小为0，则直接跳过并进入下一个循环
                    {
                        continue;
                    }

                    string targetExcelFilePath = Path.Combine(appSettings.SavingFolderPath, $"{CleanFileAndFolderName($"Tbl_{Path.GetFileNameWithoutExtension(filePath)}")}.xlsx"); // 获取目标Excel文件路径全名
                    ExtractTablesFromWordToExcel(filePath, targetExcelFilePath); // 从Word文档中提取表格并保存为目标Excel工作簿
                }

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        private async Task BatchFormatWordDocumentsAsync()
        {
            try
            {
                List<string>? filePaths = SelectFiles(EnumFileType.Word, true, "Select Word Files"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                await taskManager.RunTaskAsync(() => BatchFormatWordDocumentsHelperAsync(filePaths)); // 调用任务管理器执行批量格式化Word文档的方法
                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }

        }

        private async Task BatchRepairWordDocumentsAsync()
        {
            try
            {
                List<string>? filePaths = SelectFiles(EnumFileType.Word, true, "Select Word Files"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                await taskManager.RunTaskAsync(() => BatchRepairWordDocumentsHelperAsync(filePaths)); // 调用任务管理器执行批量修复Word文档的方法
                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }

        }

        private void BatchUnhideExcelWorksheets()
        {
            try
            {
                List<string>? filePaths = SelectFiles(EnumFileType.Excel, true, "Select Excel Files"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                int fileNum = 0;
                foreach (string filePath in filePaths) //遍历所有文件
                {
                    int hiddenExcelWorksheetCount = 0;
                    using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath))) //打开当前Excel工作簿，赋值给Excel包变量
                    {
                        foreach (ExcelWorksheet excelWorksheet in excelPackage.Workbook.Worksheets) //遍历所有Excel工作表
                        {
                            if (excelWorksheet.Hidden != eWorkSheetHidden.Visible) //如果当前Excel工作表不可见，则将其设为可见，隐藏工作表计数器加一
                            {
                                excelWorksheet.Hidden = eWorkSheetHidden.Visible;
                                hiddenExcelWorksheetCount++;
                            }
                        }
                        if (hiddenExcelWorksheetCount > 0) //如果隐藏Excel工作表数量大于0
                        {
                            fileNum++;  //文件计数器加一
                        }
                        excelPackage.Save(); //保存Excel工作簿
                    }
                }
                ShowMessage($"{fileNum} files processed.");
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        // 定义批处理工作表功能选项枚举
        private enum EnumProcessFunctions
        {
            Cancel = 0,
            MergeRecords = 1,
            AccumulateValues = 2,
            ExtractCellData = 3,
            ConvertTextualNumbers = 4,
            AdjustForPrinting = 5
        }

        private void BatchProcessExcelWorksheets()
        {
            string currentFilePath = "";
            try
            {
                // 定义功能选项列表
                List<string> lstFunctions = new List<string> { "0-Cancel", "1-Merge Records", "2-Accumulate Values", "3-Extract Cell Data", "4-Convert Textual Numbers into Numeric", "5-Adjust Worksheet Format for Printing" };
                //  获取功能选项
                int functionNum = SelectFunction(lstOptions: lstFunctions, objRecords: latestRecords, propertyName: nameof(latestRecords.LatestBatchProcessWorkbooksOption));

                if (functionNum <= 0) //如果功能选项索引号小于等于0（选择“Cancel”或不在设定范围），则结束本过程
                {
                    return;
                }

                EnumProcessFunctions function = (EnumProcessFunctions)functionNum; // 将功能选项枚举的整数值转换为枚举值

                //获取所选文件列表
                List<string>? filePaths = SelectFiles(EnumFileType.Excel, true, "Select Excel Files");
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                // 获取Excel工作表索引范围
                (int excelWorksheetStartIndex, int excelWorksheetEndIndex) = GetWorksheetRange();
                if (excelWorksheetStartIndex < 0 || excelWorksheetEndIndex < 0) // 如果获取到的Excel工作表索引号有一个小于0（范围无效），则结束本过程
                {
                    return;
                }

                int headerRowCount = 0;
                int footerRowCount = 0;
                List<string>? lstOperatingRangeAddresses = null;

                switch (function) //根据功能序号进入相应的分支
                {
                    case EnumProcessFunctions.MergeRecords: //记录合并
                    case EnumProcessFunctions.AdjustForPrinting: //调整工作表打印版式
                        (headerRowCount, footerRowCount) = GetHeaderAndFooterRowCount(); //获取表头、表尾行数
                        if (headerRowCount < 0 || footerRowCount < 0) //如果获取到的表头、表尾行数有一个小于0（范围无效），则结束本过程
                        {
                            return;
                        }

                        break;

                    case EnumProcessFunctions.AccumulateValues:
                    case EnumProcessFunctions.ExtractCellData:
                    case EnumProcessFunctions.ConvertTextualNumbers: //数值累加, 提取单元格数据, 文本型数字转数值型
                        lstOperatingRangeAddresses = GetWorksheetOperatingRangeAddresses();
                        if (lstOperatingRangeAddresses == null) //如果获取到的操作范围地址列表为null，则结束本过程
                        {
                            return;
                        }

                        break;

                }

                string? excelFileName = null; //定义被处理Excel工作簿文件名变量

                ExcelPackage targetExcelPackage = new ExcelPackage(); //新建Excel包，赋值给目标Excel包变量
                ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add("Sheet1"); //在目标Excel工作簿中添加一个工作表，赋值给目标工作表变量
                string? targetFolderPath = appSettings.SavingFolderPath; //获取目标文件夹路径
                string? targetFileMainName = Path.GetFileNameWithoutExtension(filePaths[0]); ; //获取目标文件主名

                // 定义数据表、数据行（功能3“提取单元格数据”时使用）
                DataTable? dataTable = null; //定义DataTable变量
                DataRow? dataRow = null; //定义DataTable行变量

                int fileCount = 1;

                foreach (string excelFilePath in filePaths) //遍历所有文件
                {
                    currentFilePath = excelFilePath; //将当前Excel文件路径全名赋值给当前文件路径全名变量
                    List<string> lstPrefixes = new List<string>(); //定义文件名前缀列表（给Excel文件名加前缀用）

                    using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelFilePath))) //打开当前Excel工作簿，赋值给Excel包变量
                    {
                        ExcelWorkbook excelWorkbook = excelPackage.Workbook; //将工作簿赋值给工作簿变量
                        excelFileName = Path.GetFileName(excelFilePath); //获取当前被处理Excel工作簿文件名

                        //获取被处理Excel工作表索引号的起始值和结束值，如果大于工作表数量-1，则限定为工作表数量-1 (EPPlus工作表索引号从0开始，Excel工作表索引号从1开始)
                        excelWorksheetStartIndex = Math.Min(excelWorksheetStartIndex, excelWorkbook.Worksheets.Count - 1);
                        excelWorksheetEndIndex = Math.Min(excelWorksheetEndIndex, excelWorkbook.Worksheets.Count - 1);

                        for (int i = excelWorksheetStartIndex; i <= excelWorksheetEndIndex; i++) //遍历指定范围内的所有Excel工作表
                        {
                            ExcelWorksheet excelWorksheet = excelWorkbook.Worksheets[i];
                            //如果当前Excel工作表为隐藏且使用工作表索引号，则抛出异常
                            if (excelWorksheet.Hidden != eWorkSheetHidden.Visible)
                            {
                                throw new Exception("Hidden worksheets found. Operation aborted.");
                            }

                            if (excelWorksheet.Dimension == null) //如果当前Excel工作表为空，则直接跳过当前循环并进入下一个循环
                            {
                                continue;
                            }

                            switch (function) //根据功能序号进入相应的分支
                            {

                                case EnumProcessFunctions.MergeRecords: //记录合并

                                    TrimCellStrings(excelWorksheet); //删除当前Excel工作表内所有文本型单元格值的首尾空格
                                    RemoveWorksheetEmptyRowsAndColumns(excelWorksheet); //删除当前Excel工作表内所有空白行和空白列
                                    //如果当前被处理Excel工作表的已使用行数（如果工作表为空，则为0）小于等于表头表尾行数之和，只有表头表尾无有效数据，则直接跳过当前循环并进入下一个循环
                                    if ((excelWorksheet.Dimension?.Rows ?? 0) <= headerRowCount + footerRowCount)
                                    {
                                        continue;
                                    }

                                    int sourceStartRowIndex = (fileCount == 1 && i == excelWorksheetStartIndex) ? 1 : headerRowCount + 1; //获取被处理工作表起始行索引号：如果当前是第一个Excel工作簿的第一个工作表，则得到1；否则得到表头行数+1
                                    int sourceEndRowIndex = excelWorksheet.Dimension!.End.Row - footerRowCount; //获取被处理工作表末尾行索引号：已使用区域最末行的索引号-表尾行数
                                    int targetStartRowIndex = (targetExcelWorksheet.Dimension?.End.Row ?? 0) + 1; //获取目标工作表起始行索引号：已使用区域最末行的索引号（如果工作表为空，则为0）+1

                                    //将当前被处理Excel工作表从指定起始行到指定末尾行的区域的格式，复制到目标工作表从指定起始行第3列开始的区域
                                    excelWorksheet.Cells[sourceStartRowIndex, 1, sourceEndRowIndex, excelWorksheet.Dimension.End.Column].CopyStyles(targetExcelWorksheet.Cells[targetStartRowIndex, 3]);
                                    //将当前被处理Excel工作表从指定起始行到指定末尾行的区域的数据，复制到目标工作表从指定起始行第3列开始的区域
                                    excelWorksheet.Cells[sourceStartRowIndex, 1, sourceEndRowIndex, excelWorksheet.Dimension.End.Column].Copy(targetExcelWorksheet.Cells[targetStartRowIndex, 3]);
                                    //将当前被处理Excel工作簿文件名赋值给目标工作表从指定起始行到已使用区域最末行的工作簿文件名（第1列）的单元格
                                    targetExcelWorksheet.Cells[targetStartRowIndex, 1, targetExcelWorksheet.Dimension!.End.Row, 1].Value = excelFileName;
                                    //将当前被处理Excel工作表名赋值给目标工作表从指定起始行到已使用区域最末行的工作表名（第2列）的单元格
                                    targetExcelWorksheet.Cells[targetStartRowIndex, 2, targetExcelWorksheet.Dimension.End.Row, 2].Value = excelWorksheet.Name;

                                    if (headerRowCount >= 1) //如果表头大于等于1行
                                    {
                                        targetExcelWorksheet.Cells[1, 1, headerRowCount, 2].Value = string.Empty; //将目标工作表的表头第1、2列的数据清空
                                        //在目标工作表的表头最末行的第1、2列单元格分别添加"工作簿名", "工作表名"的列名
                                        targetExcelWorksheet.Cells[headerRowCount, 1, headerRowCount, 2].LoadFromArrays(new List<object[]> { new object[] { "Source Workbook", "Source Worksheet" } });
                                    }

                                    break;

                                case EnumProcessFunctions.AccumulateValues: //数值累加

                                    if (fileCount == 1 && i == excelWorksheetStartIndex) // 如果是第一个文件的第一个Excel工作表
                                    {
                                        // 整体复制粘贴到目标Excel工作表
                                        excelWorksheet.Cells[excelWorksheet.Dimension.Address].CopyStyles(targetExcelWorksheet.Cells["A1"]); //将被处理Excel工作表的已使用区域的格式复制到目标工作表
                                        excelWorksheet.Cells[excelWorksheet.Dimension.Address].Copy(targetExcelWorksheet.Cells["A1"]); //将被处理Excel工作表的已使用区域的数据复制到目标工作表

                                        // 清除操作区域数据
                                        foreach (string anOperatingRange in lstOperatingRangeAddresses!) //遍历所有操作区域
                                        {
                                            targetExcelWorksheet.Cells[anOperatingRange].Clear(); // 清除目标Excel工作表当前操作区域数据
                                        }
                                    }

                                    // 累加操作区域数值
                                    foreach (string anOperatingRange in lstOperatingRangeAddresses!) //遍历所有操作区域
                                    {
                                        for (int k = 0; k < targetExcelWorksheet.Cells[anOperatingRange].Rows; k++) //遍历目标Excel工作表操作区域行偏移值（第1行相对第1行的偏移值为0，最后一行相对第1行的偏移值为区域总行数-1）
                                        {
                                            for (int l = 0; l < targetExcelWorksheet.Cells[anOperatingRange].Columns; l++) //遍历目标Excel工作表操作区域列偏移值（第1列相对第1列的偏移值为0，最后一列相对第1列的偏移值为区域总列数-1）
                                            {
                                                string cellStr1 = targetExcelWorksheet.Cells[anOperatingRange].Offset(k, l, 1, 1).Text; //将目标Excel工作表操作区域第1行第1列的单元格向右、向下偏移k、l个单位的单元格值转换成字符串
                                                string cellStr2 = excelWorksheet.Cells[anOperatingRange].Offset(k, l, 1, 1).Text; //将被处理Excel工作表操作区域第1行第1列的单元格向右、向下偏移k、l个单位的单元格值转换成字符串
                                                double cellNumVal1 = 0, cellNumVal2 = 0;
                                                double.TryParse(cellStr1, NumberStyles.Any, CultureInfo.InvariantCulture, out cellNumVal1); //将单元格字符串1转换成数值，如果成功则将转换后的数值赋值给单元格数值1变量
                                                double.TryParse(cellStr2, NumberStyles.Any, CultureInfo.InvariantCulture, out cellNumVal2); //将单元格字符串2转换成数值，如果成功则将转换后的数值赋值给单元格数值2变量
                                                //将转换结果值之和赋值给目标Excel工作表操作区域第1行第1列的单元格向右、向下偏移k、l个单位的单元格
                                                targetExcelWorksheet.Cells[anOperatingRange].Offset(k, l, 1, 1).Value = cellNumVal1 + cellNumVal2;
                                            }
                                        }
                                    }

                                    break;

                                case EnumProcessFunctions.ExtractCellData: //提取单元格数据

                                    if (fileCount == 1 && i == excelWorksheetStartIndex) //如果是第一个文件的第一个Excel工作表
                                    {
                                        dataTable = new DataTable(); //定义DataTable
                                        dataTable.Columns.Add("Source Workbook"); //添加列
                                        dataTable.Columns.Add("Source Worksheet");

                                        foreach (string anOperatingRange in lstOperatingRangeAddresses!)
                                        {
                                            for (int k = 0; k < excelWorksheet.Cells[anOperatingRange].Rows; k++) //遍历目标Excel工作表操作区域行偏移值（第1行相对第1行的偏移值为0，最后一行相对第1行的偏移值为区域总行数-1）
                                            {
                                                for (int l = 0; l < excelWorksheet.Cells[anOperatingRange].Columns; l++) //遍历目标Excel工作表操作区域列偏移值（第1列相对第1列的偏移值为0，最后一列相对第1列的偏移值为区域总列数-1）
                                                {
                                                    dataTable.Columns.Add(excelWorksheet.Cells[anOperatingRange].Offset(k, l, 1, 1).Address.ToString()); //将被处理Excel工作表操作区域第1行第1列的单元格向右、向下偏移k、l个单位的单元格地址作为数据列加入DataTable
                                                }
                                            }
                                        }
                                    }

                                    dataRow = dataTable!.NewRow(); //定义DataTable新数据行
                                    dataRow["Source Workbook"] = excelFileName;
                                    dataRow["Source Worksheet"] = excelWorksheet.Name;
                                    foreach (string anOperatingRange in lstOperatingRangeAddresses!) //遍历所有操作区域
                                    {
                                        for (int k = 0; k < excelWorksheet.Cells[anOperatingRange].Rows; k++) //遍历目标Excel工作表操作区域行偏移值（第1行相对第1行的偏移值为0，最后一行相对第1行的偏移值为区域总行数-1）
                                        {
                                            for (int l = 0; l < excelWorksheet.Cells[anOperatingRange].Columns; l++) //遍历目标Excel工作表操作区域列偏移值（第1列相对第1列的偏移值为0，最后一列相对第1列的偏移值为区域总列数-1）
                                            {
                                                //将被处理Excel工作表操作区域第1行第1列的单元格向右、向下偏移k、l个单位的单元格值赋值给DataTable数据行中对应单元格地址的数据列元素中
                                                ExcelRangeBase cell = excelWorksheet.Cells[anOperatingRange].Offset(k, l, 1, 1);
                                                dataRow[cell.Address.ToString()] = cell?.Value;
                                            }
                                        }
                                    }

                                    dataTable.Rows.Add(dataRow); //向DataTable添加数据行

                                    //如果当前文件是文件列表中的最后一个，且当前Excel工作表也是最后一个，且DataTable的行数和列数均不为0，则将DataTable写入目标工作表
                                    if (fileCount == filePaths.Count && i == excelWorksheetEndIndex
                                        && dataTable!.Rows.Count * dataTable.Columns.Count > 0)
                                    {
                                        targetExcelWorksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
                                    }
                                    break;

                                case EnumProcessFunctions.ConvertTextualNumbers: //文本型数字转数值型

                                    foreach (string anOperatingRange in lstOperatingRangeAddresses!) // 遍历所有操作区域
                                    {
                                        for (int k = 0; k < excelWorksheet.Cells[anOperatingRange].Rows; k++) //遍历Excel工作表操作区域行偏移值（第1行相对第1行的偏移值为0，最后一行相对第1行的偏移值为区域总行数-1）
                                        {
                                            for (int l = 0; l < excelWorksheet.Cells[anOperatingRange].Columns; l++) //遍历Excel工作表操作区域列偏移值（第1列相对第1列的偏移值为0，最后一列相对第1列的偏移值为区域总列数-1）
                                            {
                                                //将被处理Excel工作表操作区域第1行第1列的单元格向右、向下偏移k、l个单位的单元格数据转换成数值型
                                                double cellNumVal;
                                                ExcelRangeBase cell = excelWorksheet.Cells[anOperatingRange].Offset(k, l, 1, 1);
                                                if (double.TryParse(cell.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out cellNumVal)) //将当前单元格转换为数值，如果成功则将转换得到的数值赋值给单元格数值变量，然后：
                                                {
                                                    cell.Style.Numberformat.Format = ""; //将当前单元格的格式设为常规
                                                    cell.Value = cellNumVal; //将转换得到的数值赋值给当前单元格
                                                }
                                            }
                                        }
                                    }
                                    if (i == excelWorksheetEndIndex) //如果当前Excel工作表是最后一个，则保存当前被处理Excel工作簿
                                    {
                                        excelPackage.Save();
                                    }

                                    break;

                                case EnumProcessFunctions.AdjustForPrinting: //调整工作表打印版式
                                    TrimCellStrings(excelWorksheet); //删除当前Excel工作表内所有文本型单元格值的首尾空格
                                    RemoveWorksheetEmptyRowsAndColumns(excelWorksheet); //删除当前Excel工作表内所有空白行和空白列
                                    FormatExcelWorksheet(excelWorksheet, headerRowCount, footerRowCount); //设置当前Excel工作表格式

                                    if (i == excelWorksheetEndIndex) //如果当前Excel工作表是最后一个，则保存当前被处理Excel工作簿
                                    {
                                        excelPackage.Save();
                                    }

                                    break;

                            }
                        }

                    }

                    fileCount++; //文件计数器加1
                }

                string? targetExcelWorkbookPrefix = function switch  //根据功能序号返回相应的目标Excel工作簿前缀
                {
                    EnumProcessFunctions.MergeRecords => "Mrg", //记录合并
                    EnumProcessFunctions.AccumulateValues => "Accu", //数值累加
                    EnumProcessFunctions.ExtractCellData => "Extr", //提取单元格数据
                    _ => null
                };

                if (targetExcelWorkbookPrefix != null)  //如果目标Excel工作簿前缀不为null（执行功能1-3时，将生成新工作簿并保存）
                {
                    //根据功能序号返回相应的目标工作表表头行数
                    int targetHeaderRowCount = function switch
                    {
                        EnumProcessFunctions.MergeRecords => headerRowCount,  //记录合并 - 输出记录合并后的汇总表，表头行数为源数据表格的表头行数
                        EnumProcessFunctions.ExtractCellData => 1,  //提取单元格数据 - 输出提取单元格值后的汇总表，表头行数为1
                        _ => 0  //其余情况 - 表头行数为0
                    };

                    FormatExcelWorksheet(targetExcelWorksheet, targetHeaderRowCount, 0); //设置目标工作表格式

                    FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath!, $"{CleanFileAndFolderName($"{targetExcelWorkbookPrefix}_{targetFileMainName}")}.xlsx")); //获取目标Excel工作簿文件路径全名信息
                    targetExcelPackage.SaveAs(targetExcelFile);
                    targetExcelPackage.Dispose(); //关闭目标Excel工作簿
                }

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowMessage($"{ex.Message} at {currentFilePath}.");
            }

        }

        private void ConvertMarkdownIntoWord()
        {
            try
            {
                InputDialog inputDialog = new InputDialog(question: "Input the text to be converted", defaultAnswer: "", answerboxHeight: 300, acceptsReturn: true); //弹出对话框，输入功能选项
                if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则结束本过程
                {
                    return;
                }

                string mdText = inputDialog.Answer; //获取对话框返回的文本，赋值给Markdown文本变量
                mdText = appSettings.KeepEmojisInMarkdown ? mdText : mdText.RemoveEmojis(); //获取Markdown文本变量：如果程序设置允许Office文件中存在Emoji字符，则得到Markdown文本变量原值；否则，得到删除Markdown文本中Emoji后的值

                //将导出文本框的文字按换行符拆分为数组（删除每个元素前后空白字符，并删除空白元素），转换成列表
                List<string> lstParagraphs = mdText
                    .Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();

                if (lstParagraphs.Count == 0) //如果段落列表元素数为0，则抛出异常
                {
                    throw new Exception("No valid text found.");
                }

                string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径
                // 获取目标文件主名：将段落列表所有元素的Markdown标记和不能作为文件名的字符删除后，将不为null或空的字符串的元素的第一个，作为目标文件主名
                string targetFileMainName = lstParagraphs.ConvertAll(e => CleanFileAndFolderName(e.RemoveMarkdownMarks())).Where(e => !string.IsNullOrWhiteSpace(e)).First();

                //将目标Markdown文档转换为目标Word文档
                string targetWordFilePath = Path.Combine(targetFolderPath, $"{targetFileMainName}.docx"); //获取目标Word文档文件路径全名

                MarkdownSource markdown = MarkdownSource.FromMarkdownString(mdText); // 创建Markdown源对象
                MarkdownConverter converter = new MarkdownConverter() //  创建Markdown转换器对象
                {
                    //ImagesBaseUri = Path.GetDirectoryName(targetMDFilePath)  // 设置图片的路径
                };
                converter.ToDocx(markdown, targetWordFilePath, append: false); // 将Markdown文档转换成Word文档

                // 提取目标Word文档中的表格并转存为目标Excel文档
                string targetExcelFilePath = Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"Tbl_{targetFileMainName}")}.xlsx"); //获取目标Excel文件路径全名

                ExtractTablesFromWordToExcel(targetWordFilePath, targetExcelFilePath); // 提取目标Word文档中的表格并转存为目标Excel文档

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        public void BatchCreatePlaceCards()
        {
            try
            {
                List<string>? filePaths = SelectFiles(EnumFileType.Excel, false, "Select the Excel File Containing the Name List"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                using (ExcelPackage sourceExcelPackage = new ExcelPackage(new FileInfo(filePaths[0]))) //打开源数据Excel工作簿，赋值给源数据Excel包变量（源数据Excel工作簿）
                using (ExcelPackage targetExcelPackage = new ExcelPackage()) //新建Excel包，赋值给目标Excel包变量（目标Excel工作簿）
                {
                    ExcelWorksheet sourceExcelWorksheet = sourceExcelPackage.Workbook.Worksheets[0]; //将工作表1（0号）赋值给源工作表变量

                    TrimCellStrings(sourceExcelWorksheet); //删除源数据Excel工作表内所有文本型单元格值的首尾空格
                    RemoveWorksheetEmptyRowsAndColumns(sourceExcelWorksheet); //删除源数据Excel工作表内所有空白行和空白列

                    if ((sourceExcelWorksheet.Dimension?.End.Row ?? 0) <= 1) //如果工作表最末行序号（如果工作表为null，则为0）小于等于1，则抛出异常
                    {
                        throw new Exception("No valid data found.");
                    }

                    for (int i = 2; i <= sourceExcelWorksheet.Dimension!.End.Row; i++) //从第2行开始遍历源数据工作表所有行
                    {
                        string name = sourceExcelWorksheet.Cells[i, 1].Text; // 将A列当前行单元格的文字赋值给名称变量

                        // 在目标工作簿中添加一个工作表，赋值给目标Excel工作表变量
                        ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add(CleanWorksheetName($"{i - 1}_{name}"));

                        // 在目标工作表中插入名称并设置样式
                        targetExcelWorksheet.Cells["A1:A2"].Merge = true; //合并A1、A2单元格
                        targetExcelWorksheet.Cells["B1:B2"].Merge = true;
                        targetExcelWorksheet.Columns[1, 2].Width = 48; //设置1、2列的列宽
                        targetExcelWorksheet.Rows[1, 2].Height = 275; //设置1、2行的行高

                        ExcelRange cellA = targetExcelWorksheet.Cells["A1"];  //将A1单元格赋值给单元格A变量
                        ExcelRange cellB = targetExcelWorksheet.Cells["B1"];  //将B1单元格赋值给单元格B变量
                        cellA.Value = name; //将当前名称赋值给单元格A
                        cellB.Formula = "=A1"; //将公式赋值给单元格B，使之与A1相等
                        cellA.Style.TextRotation = 180; //设定单元格A文字角度：从X轴开始顺时针旋转，旋转角度为负值，设定值等于90°-旋转角度（最多不超过-90°）
                        cellB.Style.TextRotation = 90; //设定单元格B文字角度：从X轴开始逆时针旋转，旋转角度为正值，设定值等于旋转角度（最多不超过90°）

                        ExcelStyle cellABStyle = targetExcelWorksheet.Cells["A1:B1"].Style; //将单元格A、B样式赋值给单元格A、B样式变量
                        cellABStyle.Font.Name = appSettings.NameCardFontName; // 获取应用程序设置中的字体名称，设置单元格A、B字体
                        int charLimit = IsChineseText(name) ? 8 : 16; // 计算字符上限：如果是中文名称，则得到8；否则得到16
                        cellABStyle.Font.Size = (float)((!name.Contains('\n') ? 160 : 90)
                            * (1 - (name.Length - charLimit) * 0.04).Clamp(0.5, 1)); //设置字体大小：如果单元格文字不含换行符，为160；否则为90。再乘以一个缩小字体的因子
                        cellABStyle.HorizontalAlignment = ExcelHorizontalAlignment.Center; //单元格内容水平居中对齐
                        cellABStyle.VerticalAlignment = ExcelVerticalAlignment.Center; //单元格内容垂直居中对齐
                        cellABStyle.ShrinkToFit = !name.Contains('\n') ? true : false; //缩小字体填充：如果单元格文字不含换行符，为true；否则为false
                        cellABStyle.WrapText = name.Contains('\n') ? true : false; //文字自动换行：如果单元格文字含换行符，为true，否则为false

                    }

                    foreach (ExcelWorksheet excelWorksheet in targetExcelPackage.Workbook.Worksheets)
                    {
                        //设置纸张、方向、对齐
                        ExcelPrinterSettings printerSettings = excelWorksheet.PrinterSettings; //将当前Excel工作表打印设置赋值给工作表打印设置变量
                        printerSettings.PaperSize = ePaperSize.A4; // 纸张设置为A4
                        printerSettings.Orientation = eOrientation.Landscape; //方向为横向
                        printerSettings.HorizontalCentered = false; //表格水平居中为false
                        printerSettings.VerticalCentered = false; //表格垂直居中为false

                        //设置页边距
                        printerSettings.TopMargin = 0.4 / 2.54; // 边距0.4cm转inch
                        printerSettings.BottomMargin = 0.4 / 2.54;
                        printerSettings.LeftMargin = 0.4 / 2.54;
                        printerSettings.RightMargin = 0.4 / 2.54;

                        //设置视图和打印版式
                        ExcelWorksheetView view = excelWorksheet.View; //将Excel工作表视图设置赋值给视图设置变量
                        view.PageLayoutView = true; // 将工作表视图设置为页面布局视图
                        printerSettings.FitToPage = true; // 启用适应页面的打印设置
                        printerSettings.FitToWidth = 0; // 设置缩放为几页宽，0代表打印页数不受限制，可能会跨越多页
                        printerSettings.FitToHeight = 1; // 设置缩放为几页高，1代表所有行都将打印到一页上
                        view.PageLayoutView = false; // 将页面布局视图设为false（即普通视图）
                    }

                    // 保存目标工作簿
                    string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径
                    string targetFilePath = Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"PlCd_{Path.GetFileNameWithoutExtension(filePaths[0])}")}.xlsx"); //获取目标Excel工作簿文件路径全名
                    targetExcelPackage.SaveAs(new FileInfo(targetFilePath)); //保存目标Excel工作簿
                    ShowSuccessMessage();
                }

            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        private async Task BatchCreateFolders()
        {
            try
            {
                List<string>? filePaths = SelectFiles(EnumFileType.Excel, false, "Select the Excel File Containing the Directory Tree Data"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                DataTable? dataTable = await ReadExcelWorksheetIntoDataTableAsync(filePaths[0], 0); //读取Excel工作簿的第1张（0号）工作表，赋值给DataTable变量

                if (dataTable == null) //如果DataTable为null，则抛出异常
                {
                    throw new Exception("No valid data found!");
                }

                for (int i = 0; i < dataTable!.Rows.Count; i++) //遍历DataTable所有数据行
                {
                    string newPathStr = ""; //每下移一个数据行，新文件夹路径字符串变量清零
                    for (int j = 0; j < dataTable.Columns.Count; j++) //遍历所有数据列
                    {
                        newPathStr += Convert.ToString(dataTable.Rows[i][j]); //将DataTable当前数据行当前数据列元素的文件夹名累加到新文件夹路径字符串上
                        if (i >= 1 && newPathStr == "") //如果当前数据行索引号大于等于1（从第2个记录行起），且新文件夹路径字符串变量为空字符串（当前元素及左侧所有元素均为空字符串），则将DataTable当前数据行当前数据列的元素填充为上一行同数据列的文件夹名
                        {
                            dataTable.Rows[i][j] = Convert.ToString(dataTable.Rows[i - 1][j]);
                        }
                    }
                }

                // 创建目标文件夹路径
                string targetFolderPath = Path.Combine(appSettings.SavingFolderPath, CleanFileAndFolderName($"Dir_{Path.GetFileNameWithoutExtension(filePaths[0])}")); //获取目标文件夹路径

                CreateFolder(targetFolderPath);

                // 创建各级文件夹路径
                for (int i = 0; i < dataTable.Rows.Count; i++) //遍历DataTable所有数据行
                {
                    string newPath = targetFolderPath; //将目标文件夹路径赋值给新文件夹路径
                    for (int j = 0; j < dataTable.Columns.Count; j++) //遍历DataTable所有数据列
                    {
                        if (dataTable.Rows[i][j] != null) //如果当前数据行当前数据列的数据不为空
                        {
                            newPath = Path.Combine(newPath, CleanFileAndFolderName(Convert.ToString(dataTable.Rows[i][j])!)); //将现有新文件夹路径和当前数据行当前数据列的文件夹名合并，重新赋值给自身
                            CreateFolder(newPath);
                        }
                    }
                }

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }

        }

        private void CreateFileList()
        {
            try
            {
                string? folderPath = SelectFolder("Select the Folder"); //获取所选文件夹路径
                if (folderPath == null)  // 如果所选文件夹路径为null，则结束本过程
                {
                    return;
                }

                int latestSubpathDepth = latestRecords.LatestSubpathDepth;  // 读取用户使用记录中保存的子路径深度
                // 弹出功能选择对话框，提示用户输入子路径深度
                InputDialog inputDialog = new InputDialog(question: "Input the depth(level) of subdirectories", defaultAnswer: latestSubpathDepth.ToString());

                if (inputDialog.ShowDialog() == false) // 如果对话框返回值为false（点击了Cancel），则结束本过程
                {
                    return;
                }

                int subpathDepth = Convert.ToInt32(inputDialog.Answer); // 获取对话框返回的子路径深度
                latestRecords.LatestSubpathDepth = subpathDepth; // 将子路径深度赋值给用户使用记录

                DataTable dataTable = new DataTable(); // 定义DataTable，赋值给DataTable变量

                dataTable.Columns.AddRange(new DataColumn[]
                    {
                        new DataColumn("Index"),
                        new DataColumn("Path"),
                        new DataColumn("Subpath"),
                        new DataColumn("Item"),
                        new DataColumn("Type"),
                        new DataColumn("Date", typeof(DateTime))
                    });  // 向DataTable添加列

                Stack<(string FolderPath, int Depth)> stack = new Stack<(string, int)>(); // 创建栈，用于存储待处理的文件夹路径及其相对于第一级文件夹路径的子路径深度

                stack.Push((folderPath, 0)); // 将第一级文件夹路径及其子路径深度0压入栈

                while (stack.Count > 0) // 当栈不为空时，继续循环
                {

                    (string currentFolderPath, int currentSubpathDepth) = stack.Pop(); // 从栈中弹出一个文件夹路径及其相对于第一级文件夹路径的子路径深度

                    if (currentSubpathDepth > subpathDepth) // 如果当前文件夹路径路径相对于第一级文件夹路径的子路径深度超过指定的子路径深度，则直接跳过进入下一个循环
                    {
                        continue;
                    }

                    DirectoryInfo directories = new DirectoryInfo(currentFolderPath); // 获取当前文件夹的信息
                    FileInfo[] files = directories.GetFiles(); // 获取当前文件夹中的所有文件信息

                    foreach (FileInfo file in files) // 遍历每个文件信息
                    {
                        FileAttributes attributes = File.GetAttributes(file.FullName); // 获取文件的属性
                        // 如果当前文件不是隐藏或临时文件
                        if ((attributes & FileAttributes.Hidden) != FileAttributes.Hidden &&
                            (attributes & FileAttributes.Temporary) != FileAttributes.Temporary)
                        {

                            DataRow dataRow = dataTable.NewRow(); // 创建一个新的数据行

                            DateTime fileSystemDate = file.CreationTime < file.LastWriteTime ? file.CreationTime.Date : file.LastWriteTime.Date; // 获取文件的系统日期：：如果创建日期小于最后修改时间，则得到创建日期；否则，得到最后修改日期
                            dataRow["Path"] = file.FullName; // 将文件路径赋值给数据行的Path列
                            dataRow["Item"] = Path.GetFileNameWithoutExtension(file.Name); // 将文件主名赋值给数据行的Item列
                            dataRow["Type"] = file.Extension; // 将文件扩展名赋值给数据行的Type列
                            dataRow["Date"] = fileSystemDate; // 将文件日期赋值给数据行的Date列
                            dataTable.Rows.Add(dataRow); // 将数据行添加到 DataTable 中
                        }
                    }

                    DirectoryInfo[] subdirectories = directories.GetDirectories(); // 获取当前文件夹中的所有子文件夹信息
                    foreach (DirectoryInfo subdirectory in subdirectories) // 遍历每个子文件夹信息
                    {

                        FileAttributes attributes = File.GetAttributes(subdirectory.FullName); // 获取子文件夹的属性
                        // 如果当前子文件夹不是隐藏或临时文件夹
                        if ((attributes & FileAttributes.Hidden) != FileAttributes.Hidden &&
                            (attributes & FileAttributes.Temporary) != FileAttributes.Temporary)
                        {
                            DataRow dataRow = dataTable.NewRow();

                            DateTime subdirectorySystemDate = subdirectory.CreationTime < subdirectory.LastWriteTime ? subdirectory.CreationTime.Date : subdirectory.LastWriteTime.Date; // 获取子文件夹的系统日期：如果创建日期小于最后修改时间，则得到创建日期；否则，得到最后修改日期
                            dataRow["Path"] = subdirectory.FullName; // 将子文件夹路径赋值给数据行的Path列
                            dataRow["Item"] = subdirectory.Name; // 将子文件夹名赋值给数据行的Item列
                            dataRow["Type"] = "Directory"; // 将"Directory"赋值给数据行的Type列
                            dataRow["Date"] = subdirectorySystemDate; // 将子文件夹日期赋值给数据行的Date列  
                            dataTable.Rows.Add(dataRow); // 将数据行添加到 DataTable 中

                            stack.Push((subdirectory.FullName, currentSubpathDepth + 1)); // 将当前子文件夹路径及其相对于第一级路径的子路径深度累加1后的数值压入栈
                        }
                    }
                }

                if (dataTable.Rows.Count * dataTable.Columns.Count == 0) //如果DataTable的行数或列数有一个为0，则抛出异常
                {
                    throw new Exception("No valid files or directories found.");
                }

                using (ExcelPackage targetExcelPackage = new ExcelPackage()) //新建Excel包，赋值给目标Excel包变量
                {
                    ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add("Sheet1"); //新建“文件列表”Excel工作表，赋值给目标Excel工作表变量
                    targetExcelWorksheet.Cells["A1"].LoadFromDataTable(dataTable, true); //将DataTable数据导入目标工作表（true代表将表头赋给第一行）
                    int endRowIndex = targetExcelWorksheet.Dimension.End.Row; //获取目标Excel工作表最末行的行索引号
                    int dateColumnIndex = dataTable.Columns["Date"]!.Ordinal + 1; //获取目标Excel工作表日期列的索引号（工作表列索引号从1开始，DataTable从0开始）
                    //将目标Excel工作表时间列从第2行到最末行所有单元格的数据格式设为“年-月-日”
                    targetExcelWorksheet.Cells[2, dateColumnIndex, endRowIndex, dateColumnIndex].Style.Numberformat.Format = "yyyy-m-d";

                    for (int i = 2; i <= targetExcelWorksheet.Dimension.End.Row; i++) //遍历目标Excel工作表从第2行开始到末尾的所有行
                    {
                        targetExcelWorksheet.Cells[i, 1].Formula = "= ROW() - 1"; //将当前行的序号（第1）列单元格的公式设置为行索引号减1

                        ExcelRange pathCell = targetExcelWorksheet.Cells[i, 2]; //将当前行路径（第2）列单元格赋值给路径单元格变量
                        pathCell.Hyperlink = new Uri($"file:///{pathCell.Text}"); //将当前行路径单元格的超链接设定为单元格内的路径（使用file://协议）
                        pathCell.Style.Font.UnderLine = true; //将当前行路径单元格文字加下划线
                        pathCell.Style.Font.Color.SetColor(Color.Blue); //将当前行路径单元格文字颜色设为蓝色

                        //将当前行路径单元格中第一级文件夹路径替换为空，去除首尾路径分隔符，剩下的部分以路径分隔符为分隔拆分成数组，转换成列表，赋值给子路径列表
                        List<string> lstSubPath = pathCell.Text.Replace(folderPath, "").Trim(Path.DirectorySeparatorChar).Split(Path.DirectorySeparatorChar).ToList();
                        lstSubPath.RemoveAt(lstSubPath.Count - 1); //删去子路径列表中最末一个元素（最末级文件夹名或文件名）
                        targetExcelWorksheet.Cells[i, 3].Value = string.Join(Path.DirectorySeparatorChar, lstSubPath); //将子路径列表所有元素以路径分隔符为分隔合并成字符串，赋值给当前行的子路径（第3）列单元格

                    }

                    FormatExcelWorksheet(targetExcelWorksheet, 1, 0); //设置目标Excel工作表格式

                    string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径
                    FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"Lst_{folderPath}")}.xlsx")); //获取目标Excel工作簿文件路径全名信息
                    targetExcelPackage.SaveAs(targetExcelFile); //保存目标Excel工作簿文件
                }

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        private void MergeDataIntoDocument()
        {
            try
            {
                List<string>? filePaths = SelectFiles(EnumFileType.DocumentAndTable, true, "Select Document and Table Files"); //获取所选文件路径全名列表

                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                if (filePaths.Count < 2) //如果所选文件少于2个，则抛出异常
                {
                    throw new Exception("2 or more files need to be selected.");
                }

                List<string> pdfToMergeFilePaths = new List<string>(); // 建立待合并PDF文件路径全名列表

                List<string> lstTextToMerge = new List<string>(); // 建立待合并文本列表
                StringBuilder tableRowStringBuilder = new StringBuilder(); // 定义表格行数据字符串构建器

                foreach (string filePath in filePaths) //遍历列表中的所有文件
                {
                    if (new FileInfo(filePath).Length == 0) //如果当前文件大小为0，则直接跳过当前循环并进入下一个循环
                    {
                        continue;
                    }

                    string fileName = Path.GetFileName(filePath); // 获取当前文件的全名
                    string fileExtension = Path.GetExtension(filePath); // 获取当前文件的扩展名

                    if (fileExtension.Contains("xls", StringComparison.InvariantCultureIgnoreCase)) // 如果当前文件扩展名含有“xls”（Excel文件，xlsx、xlsm）
                    {
                        using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath))) //打开当前Excel工作簿，赋值给Excel包变量
                        {
                            foreach (ExcelWorksheet excelWorksheet in excelPackage.Workbook.Worksheets) // 遍历所有Excel工作表
                            {
                                TrimCellStrings(excelWorksheet); //删除当前Excel工作表内所有文本型单元格值的首尾空格
                                RemoveWorksheetEmptyRowsAndColumns(excelWorksheet); //删除当前Excel工作表内所有空白行和空白列
                                if (excelWorksheet.Dimension == null) //如果当前Excel工作表为空，则直接跳过当前循环并进入下一个循环
                                {
                                    continue;
                                }

                                lstTextToMerge.Add($"{fileName}: {excelWorksheet.Name}"); //待合并文本列表中追加当前Excel文件主名和当前工作表名

                                for (int i = 1; i <= excelWorksheet.Dimension.End.Row; i++) // 遍历Excel工作表所有行
                                {
                                    for (int j = 1; j <= excelWorksheet.Dimension.End.Column; j++) // 遍历Excel工作表所有列
                                    {
                                        tableRowStringBuilder.Append(excelWorksheet.Cells[i, j].Text.Replace('|', ';')); // 将当前单元格文字中的表格分隔符替换成分号，并追加到字符串构建器中
                                        tableRowStringBuilder.Append('|'); //追加表格分隔符到字符串构建器中
                                    }
                                    lstTextToMerge.Add(tableRowStringBuilder.ToString().TrimEnd('|')); //将字符串构建器中当前行数据转换成字符串，移除尾部的分隔符，并追加到待合并文本列表中
                                    tableRowStringBuilder.Clear(); //清空字符串构建器
                                }

                                lstTextToMerge.Add(""); //当前Excel工作表的所有行遍历完后，到了工作表末尾，在待合并文本列表最后追加一个空字符串元素
                            }
                        }
                    }

                    else if (fileExtension.Contains("doc", StringComparison.InvariantCultureIgnoreCase)) // 如果当前文件扩展名含有“doc”（Word文件，docx、docm）
                    {
                        using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read)) // 打开Word文档，赋值给文件流变量
                        {
                            XWPFDocument wordDocument = new XWPFDocument(fileStream); // 打开Word文档文件流，赋值给Word文档变量

                            lstTextToMerge.Add($"{fileName}"); // 待合并文本列表中追加当前Word文件主名

                            foreach (IBodyElement element in wordDocument.BodyElements) // 遍历Word文档所有元素
                            {
                                switch (element) // 根据元素类型进行操作
                                {
                                    case XWPFParagraph paragraph: // 如果当前元素是段落
                                        string paragraphText = paragraph.Text;
                                        if (!string.IsNullOrWhiteSpace(paragraphText)) // 如果当前段落不为空，则将当前段落文字追加到待合并文本列表中
                                        {
                                            lstTextToMerge.Add(paragraphText);
                                        }
                                        break;

                                    case XWPFTable table: // 如果当前元素是表格
                                        foreach (XWPFTableRow row in table.Rows) // 遍历表格所有行
                                        {
                                            foreach (XWPFTableCell cell in row.GetTableCells()) // 遍历当前行的所有列
                                            {
                                                string cellText = string.Join(" ", cell.Paragraphs.Select(p => p.Text.Trim())).Replace('|', ';'); // 提取单元格内的所有段落文本并连接起来（当中用空格隔开），再将表格分隔符替换成分号
                                                tableRowStringBuilder.Append(cellText); // 将当前单元格文字追加到字符串构建器中
                                                tableRowStringBuilder.Append('|'); // 追加表格分隔符到字符串构建器中
                                            }
                                            lstTextToMerge.Add(tableRowStringBuilder.ToString().TrimEnd('|')); // 将字符串构建器中当前行数据转换成字符串，移除尾部的分隔符，并追加到待合并文本列表中
                                            tableRowStringBuilder.Clear(); // 清空字符串构建器
                                        }
                                        break;

                                    default:
                                        // 忽略其他类型的元素
                                        break;
                                }
                            }

                            lstTextToMerge.Add(""); // 当前Word文档的所有段落行遍历完后，到了文档末尾，在待合并文本列表最后追加一个空字符串元素
                        }
                    }

                    // 如果当前文件扩展名含有“pdf”（PDF文件），则将当前文件路径全名追加到待合并PDF文件路径全名列表中
                    else if (fileExtension.Contains("pdf", StringComparison.InvariantCultureIgnoreCase))
                    {
                        pdfToMergeFilePaths.Add(filePath);
                    }

                }

                string targetFileMainName = Path.GetFileNameWithoutExtension(filePaths[0]); //获取列表中第一个（0号）文件的主名，赋值给目标文件主名变量
                string targetFolderPath = appSettings.SavingFolderPath; //获取目标文件夹路径
                string targetTxtFilePath = Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"Mrg_{targetFileMainName}")}.txt"); // 获取目标文本文件的路径全名
                string targetPdfFilePath = Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"Mrg_{targetFileMainName}")}.pdf"); // 获取目标PDF文件的路径全名


                // 合并待合并文本列表中的内容（来自Word、Excel文件内容），输出TXT和PDF文件

                if (lstTextToMerge.Count > 0) // 如果待合并文本列表中元素数量大于0
                {
                    // 写入目标txt文档
                    using (StreamWriter writer = new StreamWriter(targetTxtFilePath, false, Encoding.UTF8)) // 创建文本写入器对象（新建或覆盖目标文件，编码为UTF-8），赋值给文本写入器对象
                    {
                        foreach (string paragraphText in lstTextToMerge) // 遍历待合并文本列表的所有元素
                        {
                            writer.WriteLine(paragraphText); // 将当前元素的段落文字写入文本文件中，并换行
                        }
                    }

                    // 写入目标PDF文档
                    using (PdfWriter writer = new PdfWriter(targetPdfFilePath)) // 创建PDF写入器对象，赋值给PDF写入器对象
                    using (PdfDocument pdf = new PdfDocument(writer)) // 创建PDF文档对象，赋值给PDF文档对象
                    using (ITextDocument document = new ITextDocument(pdf)) // 创建文档对象，赋值给文档对象
                    {
                        PdfFont font = PdfFontFactory.CreateFont("STSong-Light", "UniGB-UCS2-H", PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED); // 创建pdf字体对象：中文宋体，Adobe-GB1符集UCS-2编码，水平书写，优先嵌入字体

                        // 遍历待合并文本列表
                        foreach (string textToMerge in lstTextToMerge)
                        {
                            ITextParagraph paragraph = new ITextParagraph(textToMerge).SetFont(font); // 为当前字符串创建一个段落，使用已定义的字体
                            document.Add(paragraph); // 将段落添加到文档中
                        }
                    }

                    pdfToMergeFilePaths.Add(targetPdfFilePath); // 将目标PDF文件路径全名添加到待合并PDF文件路径全名列表中
                }


                // 合并所有PDF文件（含原选定文件中的PDF文件和之前由Word、Excel文件合并生成的PDF文件）

                if (pdfToMergeFilePaths.Count > 1) // 如果待合并PDF文件多于1个
                {
                    // 获取最终PDF文件路径
                    string finalPdfFilePath = Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"Fnl_{targetFileMainName}")}.pdf");

                    using (PdfWriter writer = new PdfWriter(finalPdfFilePath))
                    using (PdfDocument pdf = new PdfDocument(writer))
                    {
                        PdfMerger merger = new PdfMerger(pdf); // 创建PDF合并对象，赋值给PDF合并对象

                        foreach (string pdfToMergeFilePath in pdfToMergeFilePaths) // 遍历待合并PDF文件列表
                        {
                            using PdfDocument pdfToMerge = new PdfDocument(new PdfReader(pdfToMergeFilePath)); // 打开当前待合并pdf源文件
                            {
                                merger.Merge(pdfToMerge, 1, pdfToMerge.GetNumberOfPages()); // 将源pdf文件中的全部页添加到PDF合并对象中
                            }
                        }
                    }

                    // 删除由Word、Excel文件合并而成的文本文件和PDF文件（过程性文件，如果存在的话）
                    if (File.Exists(targetTxtFilePath))
                    {
                        File.Delete(targetTxtFilePath);
                    }

                    if (File.Exists(targetPdfFilePath))
                    {
                        File.Delete(targetPdfFilePath);
                    }
                }

                //await taskManager.RunTaskAsync(() => MergeDataIntoDocumentHelperAsync(filePaths));

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }

        }

        private static void OpenSavingFolder()
        {
            try
            {
                string savingFolderPath = appSettings.SavingFolderPath; // 获取目标保存文件夹路径
                if (!Directory.Exists(savingFolderPath)) // 如果目标文件夹不存在，则抛出异常
                {
                    throw new Exception("Folder doesn't exist.");
                }

                // 创建ProcessStartInfo对象，包含了启动新进程所需的信息，赋值给启动进程信息变量
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = $"\"{savingFolderPath}\"", //指定需要打开的文件夹路径
                    UseShellExecute = true //设定使用操作系统shell执行程序
                };
                //启动新的进程
                Process.Start(startInfo);
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        private void RemoveMarkdownMarksInCopiedText()
        {
            try
            {
                if (!Clipboard.ContainsText()) //  如果剪贴板中不包含文本，则抛出异常
                {
                    ShowExceptionMessage(new Exception("No text in clipboard."));
                    return;
                }

                string originalText = Clipboard.GetText(); // 从剪贴板获取文本
                string cleanedText = originalText.RemoveMarkdownMarks(); // 清除文本中的Markdown标记
                cleanedText = appSettings.KeepEmojisInMarkdown ? cleanedText : cleanedText.RemoveEmojis();
                Clipboard.SetDataObject(cleanedText, true); // 将清理后的文本放回剪贴板

            }

            catch // 捕获异常（即使因剪贴板访问异常，内容也更新成功，故不处理）
            {

            }

            finally
            {
                ShowSuccessMessage();
            }
        }

        private static void ShowHelpWindow()
        {
            try
            {
                //创建ProcessStartInfo对象，包含了启动新进程所需的信息，赋值给启动进程信息变量
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = appSettings.UserManualUrl, //指定需要打开的网址（用户手册网址）
                    UseShellExecute = true //设定使用操作系统shell执行程序
                };
                //启动新的进程
                Process.Start(startInfo);
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        

        private void MnuTest_Click(object sender, RoutedEventArgs e)
        {
            //InputDialog inputDialog = new InputDialog(question:"Number", defaultAnswer:"1000"); //弹出功能选择对话框
            //if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则结束本过程
            //{
            //    return;
            //}
            //int numbers = Convert.ToInt32(inputDialog.Answer); //获取对话框返回的功能选项
            //string result = ConvertArabicNumberIntoChinese(numbers);
            //MessageBox.Show("转换后的中文数字为：" + result);

            //InputDialog inputDialog = new InputDialog(question: "Number", defaultAnswer: "1000"); //弹出功能选择对话框
            //if (inputDialog.ShowDialog() == false) //如果对话框返回false（点击了Cancel），则结束本过程
            //{
            //    return;
            //}
            ////获取对话框返回的功能选项
            //double result = Val(inputDialog.Answer);
            //ShowMessage("提取后的数字为：" + result.ToString());

            //InputDialog inputDialog = new InputDialog(question: "Markdown", defaultAnswer: "ABC", acceptsReturn: true); //弹出功能选择对话框
            //if (inputDialog.ShowDialog() == false) //如果对话框返回false（点击了Cancel），则结束本过程
            //{
            //    return;
            //}
            ////获取对话框返回的功能选项
            //string result = inputDialog.Answer.RemoveMarkdownMarks();
            //ShowMessage($"转换后的文字为：\n\n{result}");

            string userProfile = appSettings.UserProfile.ToString(); // 枚举转字符串
            bool isAdmin = appSettings.UserProfile == EnumUserProfile.Profile1;
            MessageBox.Show($"当前用户为：{userProfile}，是否管理员：{isAdmin.ToString()}");
        }

        private void MnuCreateQRCode_Click(object sender, RoutedEventArgs e)
        {
            QRCodeWindow qRCodeWindow = new QRCodeWindow();
            qRCodeWindow.Show();
        }
    }

}
