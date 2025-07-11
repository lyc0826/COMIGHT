﻿using DocSharp.Binary.DocFileFormat;
using DocSharp.Binary.OpenXmlLib;
using DocSharp.Binary.Spreadsheet.XlsFileFormat;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Markdown;
using Hardware.Info;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
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
using static COMIGHT.MSOfficeInterop;
using DataTable = System.Data.DataTable;
using DocSharpSpreadsheetMapping = DocSharp.Binary.SpreadsheetMLMapping;
using DocSharpWordMapping = DocSharp.Binary.WordprocessingMLMapping;
using ITextDocument = iText.Layout.Document;
using ITextParagraph = iText.Layout.Element.Paragraph;
using SpreadsheetDocument = DocSharp.Binary.OpenXmlLib.SpreadsheetML.SpreadsheetDocument;
using Task = System.Threading.Tasks.Task;
using Window = System.Windows.Window;
using WordprocessingDocument = DocSharp.Binary.OpenXmlLib.WordprocessingML.WordprocessingDocument;


namespace COMIGHT
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    /// 

    public partial class MainWindow : Window
    {

        // 获取路径
        public static string appPath = AppDomain.CurrentDomain.BaseDirectory; //获取程序所在文件夹路径
        public static string websitesJsonFilePath = Path.Combine(appPath, "Websites.json"); //获取网址Json文件路径全名
        public static string settingsJsonFilePath = Path.Combine(appPath, "Settings.json"); //获取应用程序设置Json文件路径全名
        public static string recordsJsonFilePath = Path.Combine(appPath, "Records.json"); //获取用户使用记录Json文件路径全名

        // 定义应用设置类
        public class AppSettings
        {
            public string SavingFolderPath { get; set; } = string.Empty;
            public string PandocPath { get; set; } = string.Empty;
            public string UserManualUrl { get; set; } = string.Empty;
            public string CnTitleFontName { get; set; } = string.Empty;
            public double CnTitleFontSize { get; set; }
            public string CnBodyFontName { get; set; } = string.Empty;
            public double CnBodyFontSize { get; set; }
            public string CnHeading0FontName { get; set; } = string.Empty;
            public double CnHeading0FontSize { get; set; }
            public string CnHeading1FontName { get; set; } = string.Empty;
            public double CnHeading1FontSize { get; set; }
            public string CnHeading2FontName { get; set; } = string.Empty;
            public double CnHeading2FontSize { get; set; }
            public string CnHeading3_4FontName { get; set; } = string.Empty;
            public double CnHeading3_4FontSize { get; set; }
            public double CnLineSpace { get; set; }
            public string EnTitleFontName { get; set; } = string.Empty;
            public double EnTitleFontSize { get; set; }
            public string EnBodyFontName { get; set; } = string.Empty;
            public double EnBodyFontSize { get; set; }
            public string EnHeading0FontName { get; set; } = string.Empty;
            public double EnHeading0FontSize { get; set; }
            public string EnHeading1FontName { get; set; } = string.Empty;
            public double EnHeading1FontSize { get; set; }
            public string EnHeading2FontName { get; set; } = string.Empty;
            public double EnHeading2FontSize { get; set; }
            public string EnHeading3_4FontName { get; set; } = string.Empty;
            public double EnHeading3_4FontSize { get; set; }
            public double EnLineSpace { get; set; }
            public string WorksheetFontName { get; set; } = string.Empty;
            public double WorksheetFontSize { get; set; }
            public string NameCardFontName { get; set; } = string.Empty;
            public bool KeepEmojisInMarkdown { get; set; } = false;
        }

        //定义用户使用记录类
        public class LatestRecords
        {
            public string LatestFolderPath { get; set; } = string.Empty;
            public string LastestHeaderAndFooterRowCountStr { get; set; } = string.Empty;
            public string LatestKeyColumnLetter { get; set; } = string.Empty;
            public string LatestExcelWorksheetIndexesStr { get; set; } = string.Empty;
            public string LatestOperatingRangeAddresses { get; set; } = string.Empty;
            public int LatestSubpathDepth { get; set; }
            public string LatestBatchProcessWorkbookOption { get; set; } = string.Empty;
            public string LatestSplitWorksheetOption { get; set; } = string.Empty;
            public string LatestUrl { get; set; } = string.Empty;

        }


        // 定义应用设置管理器、用户使用记录管理器对象，应用设置类、用户使用记录类对象，用于读取、保存应用设置和用户使用记录
        public static SettingsManager<AppSettings> settingsManager = new SettingsManager<AppSettings>(settingsJsonFilePath);
        public static SettingsManager<LatestRecords> recordsManager = new SettingsManager<LatestRecords>(recordsJsonFilePath);
        public static AppSettings appSettings = new AppSettings();
        public static LatestRecords latestRecords = new LatestRecords();

        public static TaskManager taskManager = new TaskManager(); //定义任务管理器对象变量，用于执行异步任务，并提供任务执行状态数据


        public MainWindow()
        {
            InitializeComponent();

            ExcelPackage.License.SetNonCommercialPersonal("Yuechen Lou"); //定义EPPlus库许可证类型为非商用

            this.Title = $"COMIGHT Assistant {DateTime.Now:yyyy}";

            lblStatus.DataContext = taskManager; // 将状态标签控件的数据环境设为任务管理器对象
            lblIntro.Content = $"For Better Productivity. © Yuechen Lou 2022-{DateTime.Now:yyyy}";

            appSettings = settingsManager.GetSettings(); // 从应用设置管理器中读取应用设置，赋值给应用设置对象变量
            latestRecords = recordsManager.GetSettings(); // 从用户使用记录管理器中读取用户使用记录，赋值给用户使用记录对象变量

            CreateFolder(appSettings.SavingFolderPath); // 创建保存文件夹
        }

        private void MnuBatchConvertOfficeFileTypes_Click(object sender, RoutedEventArgs e)
        {
            BatchConvertOfficeFileTypes();
        }

        private void MnuBatchCreateFolders_Click(object sender, RoutedEventArgs e)
        {
            BatchCreateFolders();
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

        private void MnuBrowser_Click(object sender, RoutedEventArgs e)
        {
            if (GetInstanceCountByHandle<BrowserWindow>() < 3) //如果被打开的浏览器窗口数量小于3个，则新建一个浏览器窗口实例并显示
            {
                BrowserWindow browserWindow = new BrowserWindow();
                browserWindow.Show();
            }
        }

        private void MnuCompareExcelWorksheets_Click(object sender, RoutedEventArgs e)
        {
            CompareExcelWorksheets();
        }

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

        private async void MnuExportDocumentTableIntoWord_Click(object sender, RoutedEventArgs e)
        {
            await ExportDocumentTableIntoWordAsync();
        }

        private void MnuImportTextIntoDocumentTable_Click(object sender, RoutedEventArgs e)
        {
            ImportTextIntoDocumentTable();
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
            SettingsDialog settingDialog = new SettingsDialog();
            settingDialog.ShowDialog();
        }

        private async void MnuSystemInfo_Click(object sender, RoutedEventArgs e)
        {
            await ShowSystemInfoAsync();
        }

        private void MnuBatchDisassembleAssembleExcelWorkbooks_Click(object sender, RoutedEventArgs e)
        {
            BatchDisassembleAssembleExcelWorkbooks();
        }

        public void BatchConvertOfficeFileTypes()
        {
            try
            {
                List<string>? filePaths = SelectFiles(FileType.Convertible, true, "Select Old Version Office or WPS Files"); //获取所选文件列表
                
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

                            // 用EPPlus库将转换后的工作簿的每一张工作表复制到新工作簿中，并保存覆盖原工作簿（此过程用来修正DocSharp转换时的错误）
                            using (ExcelPackage xlsPackage = new ExcelPackage(new FileInfo(targetFilePath)))
                            using (ExcelPackage xlsxPackage = new ExcelPackage())
                            {
                                // 获取源工作簿
                                ExcelWorkbook xlsWorkbook = xlsPackage.Workbook; // 定义源工作簿对象
                                // 获取目标工作簿
                                ExcelWorkbook xlsxWorkbook = xlsxPackage.Workbook; // 定义目标工作簿对象
                                foreach (ExcelWorksheet xlsWorksheet in xlsWorkbook.Worksheets) // 遍历源工作簿中的每一张工作表
                                {
                                    if (xlsWorksheet == null) //  如果当前工作表为空，则直接跳过进入下一个工作表
                                    {
                                        continue;
                                    }
                                    // 使用 Copy 方法将工作表复制到目标工作簿
                                    xlsxWorkbook.Worksheets.Add(xlsWorksheet.Name, xlsWorksheet); // 将当前工作表复制到目标工作簿
                                }

                                // 保存目标工作簿，覆盖原文件
                                xlsxPackage.SaveAs(targetFilePath);
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

        //public async Task BatchConvertOfficeFileTypes()
        //{
        //    try
        //    {
        //        List<string>? filePaths = SelectFiles(FileType.Convertible, true, "Select Old Version Office or WPS Files"); //获取所选文件列表
        //        if (filePaths == null) //如果文件列表为null，则结束本过程
        //        {
        //            return;
        //        }

        //        await taskManager.RunTaskAsync(() => BatchConvertOfficeFileTypesAsyncHelper(filePaths)); // 调用任务管理器执行批量转换Office文件类型的方法
        //        ShowSuccessMessage();
        //    }

        //    catch (Exception ex)
        //    {
        //        ShowExceptionMessage(ex);
        //    }
        //}

        public void BatchDisassembleAssembleExcelWorkbooks()
        {
            try
            {
                // 定义功能选项列表
                List<string> lstFunctions = new List<string> { "0-Cancel", "1-Split by a Column into Workbooks", "2-Split by a Column into Worksheets", "3-Disassemble Workbooks", "4-Assemble Workbooks" };

                //获取功能选项
                int functionNum = SelectFunction(options: lstFunctions, objRecords: latestRecords, propertyName: "LatestSplitWorksheetOption");
                if (functionNum <= 0) //如果功能选项小于等于0（选择“Cancel”或不在设定范围），则结束本过程
                {
                    return;
                }

                List<string>? filePaths = SelectFiles(FileType.Excel, true, "Select the Excel Files"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                // 定义Excel工作簿索引范围、表头表尾行数、列符号变量（初始值均为“非法”数值）
                int excelWorksheetStartIndex = -1; int excelWorksheetEndIndex = -1;
                int headerRowCount = -1; int footerRowCount = -1;
                string? columnLetter = null;

                (excelWorksheetStartIndex, excelWorksheetEndIndex) = GetWorksheetRange(); // 获取Excel工作表索引范围
                if (excelWorksheetStartIndex < 0 || excelWorksheetEndIndex < 0) // 如果获取到的Excel工作表索引号有一个小于0（范围无效），则结束本过程
                {
                    return;
                }

                switch (functionNum) // 根据功能选项进入相应分支
                {
                    case 1: // 按列拆分为Excel工作簿
                    case 2: // 按列拆分为Excel工作表

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

                        break;
                }

                bool createDataDict = true; // 定义“是否创建数据字典”变量（默认赋值为True）
                Dictionary<string, List<ExcelRangeBase>> dataDict = new Dictionary<string, List<ExcelRangeBase>>(); // 定义数据字典（保存按列拆分的数据）

                bool createFolderForEachWorkbook = true; // 定义“是否为每个工作簿创建一个文件夹”变量（默认赋值为true）

                // 根据功能选项，给“是否创建数据字典”和“是否创建数据字典”变量赋值
                (createDataDict, createFolderForEachWorkbook) = functionNum switch
                {
                    1 => (true, true), // 按列拆分为Excel工作簿
                    2 => (true, true),  // 按列拆分为Excel工作表
                    3 => (false, true), // 拆分工作表到独立工作簿
                    4 => (false, false), // 集合工作簿
                    _ => (false, false)
                };

                string targetFolderPath = appSettings.SavingFolderPath; // 获取保存文件夹路径

                //定义集合工作簿变量、集合工作表计数变量和集合工作簿文件主名变量（功能4“集合工作簿”时使用）
                ExcelPackage assembledExcelPackage = new ExcelPackage(); // 新建Excel包，赋值给目标Excel包变量（为当前工作表创建一个新工作簿）
                int collectedWorksheetCount = 1;
                string assembledExcelFileMainName = Path.GetFileNameWithoutExtension(filePaths[0]); // 获取集合Excel文件主名

                foreach (string filePath in filePaths) // 遍历文件列表
                {

                    string excelWorkbookFileMainName = Path.GetFileNameWithoutExtension(filePath); //获取当前Excel工作簿文件主名

                    if (createFolderForEachWorkbook) //  如果要为每个工作簿创建一个独立文件夹
                    {
                        // 创建目标文件夹（为每个工作簿创建一个独立文件夹）
                        targetFolderPath = Path.Combine(appSettings.SavingFolderPath, excelWorkbookFileMainName);
                        CreateFolder(targetFolderPath);
                    }

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

                            switch (functionNum) //根据功能序号进入相应的分支
                            {
                                case 1: //按列拆分为Excel工作簿

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

                                case 2:  //按列拆分为Excel工作表

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

                                case 3: // 拆分工作表到独立工作簿

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

                                case 4: //集合工作簿

                                    ExcelWorksheet collectedExcelWorksheet = assembledExcelPackage.Workbook.Worksheets.Add(CleanWorksheetName($"{collectedWorksheetCount++}_{excelWorksheet.Name}"));  // 新建收集Excel工作表
                                    excelWorksheet.Cells[excelWorksheet.Dimension.Address].CopyStyles(collectedExcelWorksheet.Cells["A1"]);
                                    excelWorksheet.Cells[excelWorksheet.Dimension.Address].Copy(collectedExcelWorksheet.Cells["A1"]);

                                    FormatExcelWorksheet(collectedExcelWorksheet, 0, 0); //设置目标Excel工作表格式

                                    // 如果当前Excel工作簿是最后一个工作簿，并且当前工作表是最后一个工作表，则将数据写入集合Excel工作表
                                    if (filePaths.IndexOf(filePath) == filePaths.Count - 1 && i == excelWorksheetEndIndex)
                                    {
                                        // 保存目标Excel工作簿文件
                                        FileInfo assembledExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"Coll_{assembledExcelFileMainName}")}.xlsx")); //获取目标Excel工作簿文件路径全名信息
                                        assembledExcelPackage.SaveAs(assembledExcelFile);
                                        assembledExcelPackage.Dispose();
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


        //public void BatchDisassembleAssembleExcelWorkbooks()
        //{
        //    try
        //    {
        //        // 定义功能选项列表
        //        List<string> lstFunctions = new List<string> { "0-Cancel", "1-Split by a Column into Workbooks", "2-Split by a Column into Worksheets", "3-Disassemble Workbooks", "4-Assemble Workbooks" };

        //        //获取功能选项
        //        int functionNum = SelectFunction(options: lstFunctions, objRecords: latestRecords, propertyName: "LatestSplitWorksheetOption");
        //        if (functionNum <= 0) //如果功能选项小于等于0（选择“Cancel”或不在设定范围），则结束本过程
        //        {
        //            return;
        //        }

        //        List<string>? filePaths = SelectFiles(FileType.Excel, true, "Select the Excel Files"); //获取所选文件列表
        //        if (filePaths == null) //如果文件列表为null，则结束本过程
        //        {
        //            return;
        //        }

        //        // 定义Excel工作簿索引范围、表头表尾行数、列符号变量（初始值均为“非法”数值）
        //        int excelWorksheetStartIndex = -1; int excelWorksheetEndIndex = -1;
        //        int headerRowCount = -1; int footerRowCount = -1;
        //        string? columnLetter = null;

        //        (excelWorksheetStartIndex, excelWorksheetEndIndex) = GetWorksheetRange(); // 获取Excel工作表索引范围
        //        if (excelWorksheetStartIndex < 0 || excelWorksheetEndIndex < 0) // 如果获取到的Excel工作表索引号有一个小于0（范围无效），则结束本过程
        //        {
        //            return;
        //        }

        //        switch (functionNum) // 根据功能选项进入相应分支
        //        {
        //            case 1: // 按列拆分为Excel工作簿
        //            case 2: // 按列拆分为Excel工作表

        //                (headerRowCount, footerRowCount) = GetHeaderAndFooterRowCount(); //获取表头、表尾行数; 
        //                if (headerRowCount < 0 || footerRowCount < 0) //如果获取到的表头、表尾行数有一个小于0（范围无效），则结束本过程
        //                {
        //                    return;
        //                }

        //                columnLetter = GetKeyColumnLetter(); //获取主键列符
        //                if (columnLetter == null) //如果主键列符为null，则结束本过程
        //                {
        //                    return;
        //                }

        //                break;
        //        }

        //        bool createDataDict = true; // 定义“是否创建数据字典”变量（默认赋值为True）
        //        Dictionary<string, List<ExcelRangeBase>> dataDict = new Dictionary<string, List<ExcelRangeBase>>(); // 定义数据字典（保存按列拆分的数据）

        //        bool createFolderForEachWorkbook = true; // 定义“是否为每个工作簿创建一个文件夹”变量（默认赋值为true）

        //        // 根据功能选项，给“是否创建数据字典”和“是否创建数据字典”变量赋值
        //        (createDataDict, createFolderForEachWorkbook) = functionNum switch
        //        {
        //            1 => (true, true), // 按列拆分为Excel工作簿
        //            2 => (true, true),  // 按列拆分为Excel工作表
        //            3 => (false, true), // 拆分工作表到独立工作簿
        //            4 => (false, false), // 集合工作簿
        //            _ => (false, false)
        //        };

        //        string targetFolderPath = appSettings.SavingFolderPath; // 获取保存文件夹路径

        //        //定义集合工作簿变量、集合工作表计数变量和集合工作簿文件主名变量（功能4“集合工作簿”时使用）
        //        ExcelPackage assembledExcelPackage = new ExcelPackage(); // 新建Excel包，赋值给目标Excel包变量（为当前工作表创建一个新工作簿）
        //        int collectedWorksheetCount = 1;
        //        string assembledExcelFileMainName = Path.GetFileNameWithoutExtension(filePaths[0]); // 获取集合Excel文件主名

        //        foreach (string filePath in filePaths) // 遍历文件列表
        //        {

        //            string excelWorkbookFileMainName = Path.GetFileNameWithoutExtension(filePath); //获取当前Excel工作簿文件主名

        //            if (createFolderForEachWorkbook) //  如果要为每个工作簿创建一个独立文件夹
        //            {
        //                // 创建目标文件夹（为每个工作簿创建一个独立文件夹）
        //                targetFolderPath = Path.Combine(appSettings.SavingFolderPath, excelWorkbookFileMainName);
        //                CreateFolder(targetFolderPath);
        //            }

        //            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath))) // 打开当前Excel工作簿，赋值给Excel包变量
        //            {
        //                ExcelWorkbook excelWorkbook = excelPackage.Workbook;

        //                //获取被处理Excel工作表索引号的起始值和结束值，如果大于工作表数量-1，则限定为工作表数量-1 (EPPlus工作表索引号从0开始，Excel工作表索引号从1开始)
        //                excelWorksheetStartIndex = Math.Min(excelWorksheetStartIndex, excelWorkbook.Worksheets.Count - 1);
        //                excelWorksheetEndIndex = Math.Min(excelWorksheetEndIndex, excelWorkbook.Worksheets.Count - 1);

        //                int workbookCount = 1;

        //                for (int i = excelWorksheetStartIndex; i <= excelWorksheetEndIndex; i++) //遍历所有指定范围的Excel工作表
        //                {
        //                    int worksheetCount = 1;

        //                    ExcelWorksheet excelWorksheet = excelWorkbook.Worksheets[i]; // 将当前索引号的Excel工作表赋值给Excel工作表变量

        //                    TrimCellStrings(excelWorksheet); //删除Excel工作表内所有文本型单元格值的首尾空格
        //                    RemoveWorksheetEmptyRowsAndColumns(excelWorksheet); //删除Excel工作表内所有空白行和空白列

        //                    if (createDataDict) // 如果需要创建数据字典（按列拆分时）
        //                    {
        //                        dataDict.Clear(); // 清空数据字典
        //                        if ((excelWorksheet.Dimension?.Rows ?? 0) <= headerRowCount + footerRowCount) //如果当前Excel工作表已使用行数（如果工作表为空， 则为0）小于等于表头表尾行数和，则直接跳过并进入下一个过程（结束当前工作表，跳至下一个工作表）
        //                        {
        //                            continue;
        //                        }

        //                        for (int j = headerRowCount + 1; j <= excelWorksheet.Dimension!.End.Row - footerRowCount; j++) // 遍历Excel工作表除去表头、表尾的每一行
        //                        {
        //                            string key = !string.IsNullOrWhiteSpace(excelWorksheet.Cells[columnLetter + j.ToString()].Text) ?
        //                                excelWorksheet.Cells[columnLetter + j.ToString()].Text : "-Blank-"; //将当前行拆分基准列的值赋值给键值变量：如果当前行单元格文字不为空，则得到得到单元格文字，否则得到"-Blank-"
        //                            if (dataDict.ContainsKey(key)) // 如果字典中已经有这个键，就将当前行添加到对应的列表中
        //                            {
        //                                dataDict[key].Add(excelWorksheet.Cells[j, 1, j, excelWorksheet.Dimension.End.Column]);
        //                            }
        //                            else // 否则，定义一个列表并向其中添加当前行，而后将列表并添加到字典中
        //                            {
        //                                dataDict[key] = new List<ExcelRangeBase> { excelWorksheet.Cells[j, 1, j, excelWorksheet.Dimension.End.Column] };
        //                            }
        //                        }
        //                    }

        //                    switch (functionNum) //根据功能序号进入相应的分支
        //                    {
        //                        case 1: //按列拆分为Excel工作簿

        //                            foreach (KeyValuePair<string, List<ExcelRangeBase>> pair in dataDict) // 遍历字典中的每一个键值对
        //                            {
        //                                using (ExcelPackage targetExcelPackage = new ExcelPackage()) //新建Excel包，赋值给目标Excel包变量（为当前工作表的每个键值对创建一个新工作簿）
        //                                {
        //                                    ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add("Sheet1"); // 新建Excel工作表，赋值给目标工作表变量

        //                                    // 将表头复制到目标Excel工作表
        //                                    if (headerRowCount >= 1) //如果表头行数大于等于1，复制表头
        //                                    {
        //                                        excelWorksheet.Cells[1, 1, headerRowCount, excelWorksheet.Dimension.End.Column].CopyStyles(targetExcelWorksheet.Cells["A1"]);
        //                                        excelWorksheet.Cells[1, 1, headerRowCount, excelWorksheet.Dimension.End.Column].Copy(targetExcelWorksheet.Cells["A1"]);
        //                                    }

        //                                    // 将字典中的每一行复制到目标Excel工作表
        //                                    foreach (ExcelRangeBase dictRow in pair.Value) //遍历所有字典数据
        //                                    {
        //                                        //获取目标Excel工作表最末行索引号（如果工作表为空， 则为0）
        //                                        int lastRowIndex = targetExcelWorksheet.Dimension?.End.Row ?? 0;
        //                                        dictRow.CopyStyles(targetExcelWorksheet.Cells[lastRowIndex + 1, 1]); //将当前行的样式复制到目标Excel工作表的第一个非空白行
        //                                        dictRow.Copy(targetExcelWorksheet.Cells[lastRowIndex + 1, 1]); //将当前行的数据复制到目标Excel工作表的第一个非空白行
        //                                    }

        //                                    FormatExcelWorksheet(targetExcelWorksheet, headerRowCount, 0); //设置目标Excel工作表格式

        //                                    // 保存目标Excel工作簿文件
        //                                    FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"{workbookCount++}_{pair.Key}_{excelWorksheet.Name}_{excelWorkbookFileMainName}")}.xlsx")); //获取目标Excel工作簿文件路径全名信息
        //                                    targetExcelPackage.SaveAs(targetExcelFile);
        //                                }
        //                            }

        //                            break;

        //                        case 2:  //按列拆分为Excel工作表

        //                            using (ExcelPackage targetExcelPackage = new ExcelPackage()) // 新建Excel包，赋值给目标Excel包变量（为当前工作表创建一个新工作簿）
        //                            {

        //                                foreach (KeyValuePair<string, List<ExcelRangeBase>> pair in dataDict) //遍历所有字典数据
        //                                {
        //                                    // 新建Excel工作表，赋值给目标工作表变量
        //                                    ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add(CleanWorksheetName($"{worksheetCount++}_{pair.Key}"));

        //                                    // 将表头复制到目标Excel工作表
        //                                    if (headerRowCount >= 1) //如果表头行数大于等于1，复制表头
        //                                    {
        //                                        excelWorksheet.Cells[1, 1, headerRowCount, excelWorksheet.Dimension.End.Column].CopyStyles(targetExcelWorksheet.Cells["A1"]);
        //                                        excelWorksheet.Cells[1, 1, headerRowCount, excelWorksheet.Dimension.End.Column].Copy(targetExcelWorksheet.Cells["A1"]);
        //                                    }

        //                                    // 将字典中的每一行复制到目标Excel工作表
        //                                    foreach (ExcelRangeBase dictRow in pair.Value) //遍历所有字典数据
        //                                    {
        //                                        //获取目标Excel工作表最末行索引号（如果工作表为空， 则为0）
        //                                        int lastRowIndex = targetExcelWorksheet.Dimension?.End.Row ?? 0;
        //                                        dictRow.CopyStyles(targetExcelWorksheet.Cells[lastRowIndex + 1, 1]); //将当前行的样式复制到目标Excel工作表的第一个非空白行
        //                                        dictRow.Copy(targetExcelWorksheet.Cells[lastRowIndex + 1, 1]); //将当前行的数据复制到目标Excel工作表的第一个非空白行
        //                                    }

        //                                    FormatExcelWorksheet(targetExcelWorksheet, headerRowCount, 0); //设置目标Excel工作表格式

        //                                }

        //                                // 保存目标Excel工作簿文件
        //                                FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"{workbookCount++}_{excelWorksheet.Name}_{excelWorkbookFileMainName}")}.xlsx")); //获取目标Excel工作簿文件路径全名信息
        //                                targetExcelPackage.SaveAs(targetExcelFile);

        //                            }

        //                            break;

        //                        case 3: // 拆分工作表到独立工作簿

        //                            using (ExcelPackage targetExcelPackage = new ExcelPackage()) // 新建Excel包，赋值给目标Excel包变量（为当前工作表创建一个新工作簿）
        //                            {
        //                                ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add("Sheet1");  // 新建Excel工作表

        //                                // 将当前整个工作表复制到目标Excel工作表
        //                                excelWorksheet.Cells[excelWorksheet.Dimension.Address].CopyStyles(targetExcelWorksheet.Cells["A1"]);
        //                                excelWorksheet.Cells[excelWorksheet.Dimension.Address].Copy(targetExcelWorksheet.Cells["A1"]);

        //                                FormatExcelWorksheet(targetExcelWorksheet, 0, 0); //设置目标Excel工作表格式

        //                                // 保存目标Excel工作簿文件
        //                                FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"{workbookCount++}_{excelWorksheet.Name}_{excelWorkbookFileMainName}")}.xlsx")); //获取目标Excel工作簿文件路径全名信息
        //                                targetExcelPackage.SaveAs(targetExcelFile);

        //                            }

        //                            break;

        //                        case 4: //集合工作簿

        //                            ExcelWorksheet collectedExcelWorksheet = assembledExcelPackage.Workbook.Worksheets.Add(CleanWorksheetName($"{collectedWorksheetCount++}_{excelWorksheet.Name}"));  // 新建收集Excel工作表
        //                            excelWorksheet.Cells[excelWorksheet.Dimension.Address].CopyStyles(collectedExcelWorksheet.Cells["A1"]);
        //                            excelWorksheet.Cells[excelWorksheet.Dimension.Address].Copy(collectedExcelWorksheet.Cells["A1"]);

        //                            FormatExcelWorksheet(collectedExcelWorksheet, 0, 0); //设置目标Excel工作表格式

        //                            // 如果当前Excel工作簿是最后一个工作簿，并且当前工作表是最后一个工作表，则将数据写入集合Excel工作表
        //                            if (filePaths.IndexOf(filePath) == filePaths.Count - 1 && i == excelWorksheetEndIndex)
        //                            {
        //                                // 保存目标Excel工作簿文件
        //                                FileInfo assembledExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"Coll_{assembledExcelFileMainName}")}.xlsx")); //获取目标Excel工作簿文件路径全名信息
        //                                assembledExcelPackage.SaveAs(assembledExcelFile);
        //                                assembledExcelPackage.Dispose();
        //                            }

        //                            break;
        //                    }

        //                }
        //            }

        //        }

        //        ShowSuccessMessage();
        //    }

        //    catch (Exception ex)
        //    {
        //        ShowExceptionMessage(ex);
        //    }

        //}

        private void BatchExtractTablesFromWord()
        {
            try
            {
                List<string>? filePaths = SelectFiles(FileType.Word, true, "Select Word Files"); //获取所选文件列表
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
                List<string>? filePaths = SelectFiles(FileType.Word, true, "Select Word Files"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                await taskManager.RunTaskAsync(() => BatchFormatWordDocumentsAsyncHelper(filePaths)); // 调用任务管理器执行批量格式化Word文档的方法
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
                List<string>? filePaths = SelectFiles(FileType.Word, true, "Select Word Files"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                await taskManager.RunTaskAsync(() => BatchRepairWordDocumentsAsyncHelper(filePaths)); // 调用任务管理器执行批量修复Word文档的方法
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
                List<string>? filePaths = SelectFiles(FileType.Excel, true, "Select Excel Files"); //获取所选文件列表
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

        private void BatchProcessExcelWorksheets()
        {
            string currentFilePath = "";
            try
            {
                // 定义功能选项列表
                List<string> lstFunctions = new List<string> {"0-Cancel", "1-Merge Records", "2-Accumulate Values", "3-Extract Cell Data", "4-Convert Textual Numbers into Numeric",
                    "5-Copy Formula to Multiple Worksheets", "6-Adjust Worksheet Format for Printing"};
                //  获取功能选项
                int functionNum = SelectFunction(options: lstFunctions, objRecords: latestRecords, propertyName: "LatestBatchProcessWorkbookOption");

                if (functionNum <= 0) //如果功能选项索引号小于等于0（选择“Cancel”或不在设定范围），则结束本过程
                {
                    return;
                }

                //获取所选文件列表
                List<string>? filePaths = SelectFiles(FileType.Excel, true, "Select Excel Files");
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

                switch (functionNum) //根据功能序号进入相应的分支
                {
                    case 1: //记录合并
                    case 6: //调整工作表打印版式
                        (headerRowCount, footerRowCount) = GetHeaderAndFooterRowCount(); //获取表头、表尾行数
                        if (headerRowCount < 0 || footerRowCount < 0) //如果获取到的表头、表尾行数有一个小于0（范围无效），则结束本过程
                        {
                            return;
                        }

                        break;

                    case 2:
                    case 3:
                    case 4:
                    case 5: //2-数值累加, 3-提取单元格数据, 4-文本型数字转数值型, 5-复制公式到多Excel工作表
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

                //定义模板Excel文件列表、模板Excel包和模板Excel工作表（功能5“复制公式到多Excel工作表”时使用）
                List<string>? templateExcelFilePaths = null; //定义模板Excel文件列表变量
                ExcelPackage? templateExcelPackage = null; //定义模板Excel包变量
                ExcelWorksheet? templateExcelWorksheet = null; //定义模板Excel工作表变量

                int fileCount = 1;

                foreach (string excelFilePath in filePaths) //遍历所有文件
                {
                    currentFilePath = excelFilePath; //将当前Excel文件路径全名赋值给当前文件路径全名变量
                    List<string> lstPrefixes = new List<string>(); //定义文件名前缀列表（给Excel文件名加前缀用）

                    using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelFilePath))) //打开当前Excel工作簿，赋值给Excel包变量
                    {
                        ExcelWorkbook excelWorkbook = excelPackage.Workbook; //将工作簿赋值给工作簿变量
                        excelFileName = Path.GetFileName(excelFilePath); //获取当前被处理Excel工作簿文件主名

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

                            switch (functionNum) //根据功能序号进入相应的分支
                            {

                                case 1: //记录合并

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

                                case 2: //数值累加

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

                                case 3: //提取单元格数据

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
                                                //将被处理Excel工作表操作区域第1行第1列的单元格向右、向下偏移k、l个单位的单元格值赋值给DataTable数据行中对应单元格地址数据列的元素中
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

                                case 4: //文本型数字转数值型

                                    foreach (string anOperatingRange in lstOperatingRangeAddresses!) // 遍历所有操作区域
                                    {
                                        for (int k = 0; k < excelWorksheet.Cells[anOperatingRange].Rows; k++) //遍历Excel工作表操作区域行偏移值（第1行相对第1行的偏移值为0，最后一行相对第1行的偏移值为区域总行数-1）
                                        {
                                            for (int l = 0; l < excelWorksheet.Cells[anOperatingRange].Columns; l++) //遍历Excel工作表操作区域列偏移值（第1列相对第1列的偏移值为0，最后一列相对第1列的偏移值为区域总列数-1）
                                            {
                                                //将被处理Excel工作表操作区域第1行第1列的单元格向右、向下偏移k、l个单位的单元格数据转换成数值型
                                                double cellNumVal;
                                                ExcelRangeBase cell = excelWorksheet.Cells[anOperatingRange].Offset(k, l, 1, 1);
                                                if (double.TryParse(cell.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out cellNumVal)) //将当前单元格转换为数值，如果成功则将转换得到的数值赋值给单元格数值变量
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

                                case 5: //复制公式到多Excel工作表

                                    if (fileCount == 1 && i == excelWorksheetStartIndex) // 如果是第一个文件的第一个Excel工作表
                                    {
                                        templateExcelFilePaths = SelectFiles(FileType.Excel, false, "Select the Template Excel File"); //选择模板文件
                                        if (templateExcelFilePaths == null) //如果文件为null，结束本过程
                                        {
                                            return;
                                        }
                                        templateExcelPackage = new ExcelPackage(new FileInfo(templateExcelFilePaths[0])); //打开模板Excel工作簿，赋值给模板Excel包变量
                                        templateExcelWorksheet = templateExcelPackage.Workbook.Worksheets[0]; //将模板Excel工作簿第一个（0号）工作表赋值给模板工作表变量
                                    }

                                    foreach (string anOperatingRange in lstOperatingRangeAddresses!) // 遍历所有操作区域
                                    {
                                        templateExcelWorksheet?.Cells[anOperatingRange].Copy(excelWorksheet.Cells[anOperatingRange]); //将模板Excel工作表指定区域的公式复制到当前工作表中
                                    }

                                    if (i == excelWorksheetEndIndex) //如果当前Excel工作表是最后一个，则保存当前被处理Excel工作簿
                                    {
                                        excelPackage.Save();

                                        // 如果当前Excel文件是最后一个，则关闭模板工作簿
                                        if (filePaths.IndexOf(excelFilePath) == filePaths.Count - 1)
                                        {
                                            templateExcelPackage?.Dispose();
                                        }
                                    }

                                    break;

                                case 6: //调整工作表打印版式
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

                string? targetExcelWorkbookPrefix = functionNum switch  //根据功能序号返回相应的目标Excel工作簿前缀
                {
                    1 => "Mrg", //记录合并
                    2 => "Accu", //数值累加
                    3 => "Extr", //提取单元格数据
                    _ => null
                };

                if (targetExcelWorkbookPrefix != null)  //如果目标Excel工作簿前缀不为null（执行功能1-3时，将生成新工作簿并保存）
                {
                    //获取目标工作表表头行数
                    //根据功能序号返回相应的目标工作表表头行数
                    int targetHeaderRowCount = functionNum switch
                    {
                        1 => headerRowCount,  //记录合并 - 输出记录合并后的汇总表，表头行数为源数据表格的表头行数
                        3 => 1,  //提取单元格数据 - 输出提取单元格值后的汇总表，表头行数为1
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

        private void CompareExcelWorksheets()
        {
            try
            {
                List<string>? startFilePaths = SelectFiles(FileType.Excel, false, "Select the Excel File Containing the Start Data"); //获取所选起始数据文件列表
                List<string>? endFilePaths = SelectFiles(FileType.Excel, false, "Select the Excel File Containing the End Data"); //获取所选终点数据文件列表

                if (startFilePaths == null || endFilePaths == null) //如果起始数据或终点数据文件列表有一个为null，则结束本过程
                {
                    return;
                }

                (int worksheetStartIndex, int worksheetEndIndex) = GetWorksheetRange();
                if (worksheetStartIndex < 0 || worksheetEndIndex < 0) //如果获取到的工作表起始和结束索引号有一个小于0（范围无效），则结束本过程
                {
                    return;
                }

                (int headerRowCount, int footerRowCount) = GetHeaderAndFooterRowCount(); //获取表头、表尾行数
                if (headerRowCount < 0 || footerRowCount < 0) //如果获取到的表头、表尾行数有一个小于0（范围无效），则结束本过程
                {
                    return;
                }

                string? keyColumnLetter = GetKeyColumnLetter(); //获取主键列符
                if (keyColumnLetter == null) //如果获取到的主键列符为null，则结束本过程
                {
                    return;
                }

                List<DataTable> lstDataTable = new List<DataTable>(); // 新建数据表列表变量
                using (ExcelPackage excelPackage = new ExcelPackage()) //新建Excel包，赋值给Excel包变量
                {

                    for (int i = worksheetStartIndex; i <= worksheetEndIndex; i++)
                    {

                        DataTable? startDataTable = ReadExcelWorksheetIntoDataTable(startFilePaths[0], i, headerRowCount, footerRowCount); //读取起始数据Excel工作簿的第1张工作表，赋值给起始DataTable变量
                        DataTable? endDataTable = ReadExcelWorksheetIntoDataTable(endFilePaths[0], i, headerRowCount, footerRowCount); //读取终点数据Excel工作簿的第1张工作表，赋值给终点DataTable变量

                        if (startDataTable == null || endDataTable == null) //如果起始DataTable或终点DataTable有一个为null，则直接退出循环
                        {
                            break;
                        }

                        //获取Excel工作表的主键列对应的DataTable主键数据列的名称（工作表列索引号从1开始，DataTable从0开始）
                        string keyDataColumnName = endDataTable.Columns[ConvertColumnLettersIntoIndex(keyColumnLetter) - 1].ColumnName;

                        List<string> lstRecordKeys = new List<string>(); //定义记录主键列表
                        List<string> lstDataColumnNames = new List<string>(); //定义数据列名称列表

                        //将起始和终点DataTable的所有记录的主键数据列的值，和所有数据列名称分别添加到记录主键列表和数据列名称列表中
                        foreach (DataRow endDataRow in endDataTable.Rows) //遍历终点DataTable的每一数据行
                        {
                            lstRecordKeys.Add(Convert.ToString(endDataRow[keyDataColumnName])!); //将当前数据行主键数据列的值添加到记录主键列表中
                        }

                        foreach (DataColumn endDataColumn in endDataTable.Columns) //遍历终点DataTable的每一数据列
                        {
                            lstDataColumnNames.Add(endDataColumn.ColumnName); //将当前数据列名称添加到数据列名称列表中
                        }

                        foreach (DataRow startDataRow in startDataTable.Rows) //遍历起点DataTable的每一数据行
                        {
                            string key = Convert.ToString(startDataRow[keyDataColumnName])!; //获取当前数据行主键数据列的值
                            if (!lstRecordKeys.Contains(key)) //如果记录主键列表不含当前数据行的主键数据列的值，则将该值添加到记录主键列表中
                            {
                                lstRecordKeys.Add(key);
                            }
                        }

                        foreach (DataColumn startDataColumn in startDataTable.Columns) //遍历起点DataTable的每一数据列
                        {
                            if (!lstDataColumnNames.Contains(startDataColumn.ColumnName))  //如果数据列名称列表不含当前数据列名称，则将该数据列名称添加到数据列名称列表中
                            {
                                lstDataColumnNames.Add(startDataColumn.ColumnName);
                            }
                        }

                        DataTable differenceDataTable = new DataTable(); //定义差异DataTable，赋值给差异DataTable变量
                        foreach (string dataColumnName in lstDataColumnNames) //遍历数据列名称列表的所有元素
                        {
                            differenceDataTable.Columns.Add(dataColumnName, typeof(string)); //将当前数据列名称作为新数据列添加到差异DataTable中，数据类型为string
                        }

                        foreach (string recordKey in lstRecordKeys) //遍历记录主键列表的所有元素
                        {
                            DataRow differenceDataRow = differenceDataTable.NewRow(); //定义差异DataTable新数据行，赋值给差异DataTable数据行变量
                            differenceDataRow[keyDataColumnName] = recordKey; //将当前记录主键赋值给差异DataTable新数据行的主键数据列
                            differenceDataTable.Rows.Add(differenceDataRow); //向差异DataTable添加该新数据行

                            //从起始和终点DataTable中筛选出主键数据列的值为当前主键的行，并取其中第一个，分别赋值给起始数据行和终点数据行
                            DataRow? startDataRow = startDataTable.AsEnumerable().Where(dataRow => Convert.ToString(dataRow[keyDataColumnName]) == recordKey).FirstOrDefault();
                            DataRow? endDataRow = endDataTable.AsEnumerable().Where(dataRow => Convert.ToString(dataRow[keyDataColumnName]) == recordKey).FirstOrDefault();

                            foreach (string dataColumnName in lstDataColumnNames) //遍历数据列名称列表的所有元素
                            {
                                if (dataColumnName == keyDataColumnName) //如果当前数据列名称等于主键列名称，则直接跳过进入下一个循环
                                {
                                    continue;
                                }

                                //获取起始和终点数据字符串：如果起始（终点）数据行不为null且起始（终点）DataTable含有当前数据列，则得到起始（终点）数据行当前数据列的数据字符串；否则得到空字符串
                                string startDataStr = startDataRow != null && startDataTable.Columns.Contains(dataColumnName) ?
                                        Convert.ToString(startDataRow[dataColumnName])! : "";
                                string endDataStr = endDataRow != null && endDataTable.Columns.Contains(dataColumnName) ?
                                        Convert.ToString(endDataRow[dataColumnName])! : "";

                                string? result;
                                if ((startDataStr == endDataStr) && endDataStr != "") //如果起始数据字符串与终点数据字符串相同且不为空字符串，结果变量赋值为null
                                {
                                    result = null;
                                }
                                else //否则
                                {
                                    double startDataValue, endDataValue;
                                    //将起始和终点数据字符串转换成数值，如果成功则将转换结果赋值给各自的数据数值变量并将true赋值给各自的“数据为数值”变量；否则将false赋值给各自的“数据为数值”变量
                                    bool startDataIsNumeric = double.TryParse(startDataStr, NumberStyles.Any, CultureInfo.InvariantCulture, out startDataValue);
                                    bool endDataIsNumeric = double.TryParse(endDataStr, NumberStyles.Any, CultureInfo.InvariantCulture, out endDataValue);

                                    //如果起始或终点数据字符串之中有一个没有被成功地转换为数值，则将起始和终点数据字符串结果合并后赋值给结果变量
                                    if (!startDataIsNumeric || !endDataIsNumeric)
                                    {
                                        result = $"Start: {startDataStr}\nEnd: {endDataStr}";
                                    }
                                    else //否则
                                    {
                                        double difference = endDataValue - startDataValue; //计算终点和起始数据的差值
                                        double diffRate = startDataValue != 0 ? difference / startDataValue : double.NaN; //获取终点和起始数据的变化率：如果起始数值不为零，得到变化率；否则得到NaN
                                        result = $"Start: {startDataValue}\nEnd: {endDataValue}\nDiff: {difference}({diffRate.ToString("P2", CultureInfo.InvariantCulture)})"; //将起始和终点数据数值、差值和变化率合并后赋值给结果变量
                                    }
                                }
                                differenceDataRow[dataColumnName] = result; //将结果赋值给差异DataTable当前新数据行的当前数据列
                            }
                        }

                        differenceDataTable = RemoveDataTableEmptyRowsAndColumns(differenceDataTable, true); // 移除差异DataTable中的空数据行和空数据列

                        if (differenceDataTable.Rows.Count * differenceDataTable.Columns.Count == 0) //如果差异DataTable的数据行数或列数有一个为0，则直接跳过进入下一个循环
                        {
                            continue;
                        }

                        lstDataTable.Add(differenceDataTable); //将差异DataTable添加到差异DataTable列表中

                    }
                }

                string targetFolderPath = appSettings.SavingFolderPath; //获取目标文件夹路径

                string targetExcelFile = Path.Combine(targetFolderPath!, $"{CleanFileAndFolderName($"Comp_{Path.GetFileNameWithoutExtension(endFilePaths[0])}")}.xlsx"); //获取目标Excel工作簿文件路径全名信息

                WriteDataTableIntoExcelWorkbook(lstDataTable, targetExcelFile); //将所有差异DataTable列表写入目标Excel工作簿

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

        private void ConvertMarkdownIntoWord()
        {
            try
            {
                InputDialog inputDialog = new InputDialog(question: "Input the text to be converted", defaultAnswer: "", textboxHeight: 300, acceptsReturn: true); //弹出对话框，输入功能选项
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
                converter.ToDocx(markdown, targetWordFilePath); // 将Markdown文档转换成Word文档

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


        //private void ConvertMarkdownIntoWord()
        //{
        //    try
        //    {
        //        InputDialog inputDialog = new InputDialog(question: "Input the text to be converted", defaultAnswer: "", textboxHeight: 300, acceptsReturn: true); //弹出对话框，输入功能选项
        //        if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则结束本过程
        //        {
        //            return;
        //        }

        //        string mdText = inputDialog.Answer; //获取对话框返回的文本，赋值给Markdown文本变量
        //        mdText = appSettings.KeepEmojisInMarkdown ? mdText : mdText.RemoveEmojis(); //获取Markdown文本变量：如果程序设置允许Office文件中存在Emoji字符，则得到Markdown文本变量原值；否则，得到删除Markdown文本中Emoji后的值

        //        //将导出文本框的文字按换行符拆分为数组（删除每个元素前后空白字符，并删除空白元素），转换成列表
        //        List<string> lstParagraphs = mdText
        //            .Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();

        //        if (lstParagraphs.Count == 0) //如果段落列表元素数为0，则抛出异常
        //        {
        //            throw new Exception("No valid text found.");
        //        }

        //        string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径
        //        // 获取目标文件主名：将段落列表0号元素（一般为标题）删除Markdown标记，截取前40个字符
        //        string targetFileMainName = CleanFileAndFolderName(lstParagraphs[0].RemoveMarkdownMarks());

        //        //导入目标Markdown文档
        //        string targetMDFilePath = Path.Combine(targetFolderPath, $"{targetFileMainName}.md"); //获取目标Markdown文档文件路径全名
        //        File.WriteAllText(targetMDFilePath, mdText); //将导出文本框内的markdown文字导入目标Markdown文档

        //        //将目标Markdown文档转换为目标Word文档
        //        string targetWordFilePath = Path.Combine(targetFolderPath, $"{targetFileMainName}.docx"); //获取目标Word文档文件路径全名
        //        ConvertDocumentByPandoc("docx", targetMDFilePath, targetWordFilePath); // 将目标Markdown文档转换为目标Word文档

        //        File.Delete(targetMDFilePath); //删除Markdown文件

        //        // 提取目标Word文档中的表格并转存为目标Excel文档
        //        string targetExcelFilePath = Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"Tbl_{targetFileMainName}")}.xlsx"); //获取目标Excel文件路径全名

        //        ExtractTablesFromWordToExcel(targetWordFilePath, targetExcelFilePath); // 提取目标Word文档中的表格并转存为目标Excel文档

        //        ShowSuccessMessage();
        //    }

        //    catch (Exception ex)
        //    {
        //        ShowExceptionMessage(ex);
        //    }
        //}

        public void BatchCreatePlaceCards()
        {
            try
            {
                List<string>? filePaths = SelectFiles(FileType.Excel, false, "Select the Excel File Containing the Name List"); //获取所选文件列表
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

        private void ImportTextIntoDocumentTable()
        {
            try
            {
                InputDialog inputDialog = new InputDialog(question: "Input the text to be imported", defaultAnswer: "", textboxHeight: 300, acceptsReturn: true); //弹出对话框，输入功能选项
                if (inputDialog.ShowDialog() == false) //如果对话框返回值为false（点击了Cancel），则结束本过程
                {
                    return;
                }

                //将导出文本框的文字按换行符拆分为数组（删除每个元素前后空白字符，并删除空白元素），转换成列表
                List<string> lstParagraphs = inputDialog.Answer
                    .Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();

                if (lstParagraphs.Count == 0) //如果段落列表元素数为0，则抛出异常
                {
                    throw new Exception("No valid text found!");
                }

                string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径

                //获取目标结构化文档表文件路径全名（移除段落列表0号元素中不能作为文件名的字符，截取前40个字符，作为目标文件主名）
                string targetExcelFilePath = Path.Combine(targetFolderPath, $"{CleanFileAndFolderName(lstParagraphs[0])}.xlsx");
                ImportParagraphListIntoDocumentTable(lstParagraphs, targetExcelFilePath); //将段落列表内容导入目标结构化文档表

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }

        }

        private void BatchCreateFolders()
        {
            try
            {
                List<string>? filePaths = SelectFiles(FileType.Excel, false, "Select the Excel File Containing the Directory Tree Data"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                DataTable? dataTable = ReadExcelWorksheetIntoDataTable(filePaths[0], 0); //读取Excel工作簿的第1张（0号）工作表，赋值给DataTable变量

                if (dataTable == null) //如果DataTable为null，则抛出异常
                {
                    throw new Exception("No valid data found!");
                }

                for (int i = 0; i < dataTable!.Rows.Count; i++) //遍历DataTable所有数据行
                {
                    string newPathStr = ""; //每下移一个数据行，新文件夹路径字符串变量清零
                    for (int j = 0; j < dataTable.Columns.Count; j++) //遍历所有数据列
                    {
                        //dataTable.Rows[i][j] = Convert.ToString(dataTable.Rows[i][j])!; //去除DataTable当前数据行当前数据列数据的文件夹名中不可用于文件夹名的字符
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
                            //newPath = Path.Combine(newPath, Convert.ToString(dataTable.Rows[i][j])!); //将现有新文件夹路径和当前数据行当前数据列的文件夹名合并，重新赋值给自身
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

        public async Task ExportDocumentTableIntoWordAsync()
        {
            try
            {
                List<string>? filePaths = SelectFiles(FileType.Excel, false, "Select the Document Table File"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }
                string targetFolderPath = appSettings.SavingFolderPath; //获取目标文件夹路径

                string targetWordFilePath = Path.Combine(targetFolderPath, $"{Path.GetFileNameWithoutExtension(filePaths[0])}.docx"); //获取目标Word文件路径全名
                await ExportDocumentTableIntoWordAsyncHelper(filePaths[0], targetWordFilePath); //将结构化文档表导出为目标Word文档

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
                List<string>? filePaths = SelectFiles(FileType.WordAndExcel, true, "Select Word and Excel Files"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }


                List<string> lstFullText = new List<string>(); //建立全文本列表
                StringBuilder tableRowStringBuilder = new StringBuilder(); // 定义表格行数据字符串构建器

                foreach (string filePath in filePaths) //遍历所有列表中的文件
                {
                    if (new FileInfo(filePath).Length == 0) //如果当前文件大小为0，则直接跳过当前循环并进入下一个循环
                    {
                        continue;
                    }

                    string fileName = Path.GetFileName(filePath); // 获取当前文件的全名
                    string fileExtension = Path.GetExtension(filePath); // 获取当前文件的扩展名

                    if (fileExtension.ToLower().Contains("xls")) // 如果当前文件扩展名转换为小写后含有“xls”（Excel文件，xlsx、xlsm）
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

                                lstFullText.Add($"{fileName}: {excelWorksheet.Name}"); //全文本列表中追加当前Excel文件主名和当前工作表名

                                for (int i = 1; i <= excelWorksheet.Dimension.End.Row; i++) // 遍历Excel工作表所有行
                                {
                                    for (int j = 1; j <= excelWorksheet.Dimension.End.Column; j++) // 遍历Excel工作表所有列
                                    {
                                        tableRowStringBuilder.Append(excelWorksheet.Cells[i, j].Text); // 将当前单元格文字追加到字符串构建器中
                                        tableRowStringBuilder.Append('|'); //追加表格分隔符到字符串构建器中
                                    }
                                    lstFullText.Add(tableRowStringBuilder.ToString().TrimEnd('|')); //将字符串构建器中当前行数据转换成字符串，移除尾部的分隔符，并追加到全文本列表中
                                    tableRowStringBuilder.Clear(); //清空字符串构建器
                                }
                                lstFullText.AddRange(new string[] { "(The End)", "" }); //当前Excel工作表的所有行遍历完后，到了工作表末尾，在全文本列表最后追加一个"(The End)"元素和一个空字符串元素
                            }
                        }
                    }

                    else if (fileExtension.ToLower().Contains("doc")) // 如果当前文件扩展名转换为小写后含有“doc”（Word文件，docx、docm）
                    {
                        using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read)) // 打开Word文档，赋值给文件流变量
                        {
                            XWPFDocument wordDocument = new XWPFDocument(fileStream); // 打开Word文档文件流，赋值给Word文档变量

                            lstFullText.Add($"{fileName}"); // 全文本列表中追加当前Word文件主名

                            foreach (IBodyElement element in wordDocument.BodyElements) // 遍历Word文档所有元素
                            {
                                switch (element) // 根据元素类型进行操作
                                {
                                    case XWPFParagraph paragraph: // 如果当前元素是段落
                                        string paragraphText = paragraph.Text;
                                        if (!string.IsNullOrWhiteSpace(paragraphText)) // 如果当前段落不为空，则将当前段落文字追加到全文本列表中
                                        {
                                            lstFullText.Add(paragraphText);
                                        }
                                        break;

                                    case XWPFTable table: // 如果当前元素是表格
                                        foreach (XWPFTableRow row in table.Rows) // 遍历表格所有行
                                        {
                                            foreach (XWPFTableCell cell in row.GetTableCells()) // 遍历当前行的所有列
                                            {
                                                string cellText = string.Join(" ", cell.Paragraphs.Select(p => p.Text.Trim())); // 提取单元格内的所有段落文本并连接起来
                                                tableRowStringBuilder.Append(cellText); // 将当前单元格文字追加到字符串构建器中
                                                tableRowStringBuilder.Append('|'); // 追加表格分隔符到字符串构建器中
                                            }
                                            lstFullText.Add(tableRowStringBuilder.ToString().TrimEnd('|')); // 将字符串构建器中当前行数据转换成字符串，移除尾部的分隔符，并追加到全文本列表中
                                            tableRowStringBuilder.Clear(); // 清空字符串构建器
                                        }
                                        break;

                                    default:
                                        // 忽略其他类型的元素
                                        break;
                                }
                            }

                            lstFullText.AddRange(new string[] { "(The End)", "" }); // 当前Word文档的所有段落行遍历完后，到了文档末尾，在全文本列表最后追加一个"(The End)"元素和一个空字符串元素
                        }
                    }
                }

                string targetFileMainName = Path.GetFileNameWithoutExtension(filePaths[0]); //获取列表中第一个（0号）文件的主名，赋值给目标文件主名变量
                string targetFolderPath = appSettings.SavingFolderPath; //获取目标文件夹路径

                ////写入目标Word文档
                //string targetWordFilePath = Path.Combine(targetFolderPath, $"Mrg_{excelWorkbookFileMainName}.docx"); //获取目标Word文件的路径全名

                //using (FileStream fileStream = new FileStream(targetWordFilePath, FileMode.Create, FileAccess.Write)) // 创建文件流，以创建目标Word文档，赋值给文件流变量
                //{
                //    XWPFDocument targetWordDocument = new XWPFDocument(); // 创建Word文件对象，赋值给目标Word文档变量

                //    foreach (string paragraphText in lstFullText) // 遍历全文本列表的所有元素
                //    {
                //        XWPFParagraph paragraph = targetWordDocument.CreateParagraph(); // 创建段落
                //        XWPFRun run = paragraph.CreateRun(); // 创建段落文本块
                //        run.SetText(paragraphText); // 将当前元素的段落文字插入段落文本块中
                //    }
                //    targetWordDocument.Write(fileStream); // 写入文件流
                //}

                // 写入目标文本文档
                string targetTxtFilePath = Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"Mrg_{targetFileMainName}")}.txt"); // 获取目标文本文件的路径全名

                using (StreamWriter writer = new StreamWriter(targetTxtFilePath, false, Encoding.UTF8)) // 创建文本写入器对象（新建或覆盖目标文件，编码为UTF-8），赋值给文本写入器对象
                {
                    foreach (string paragraphText in lstFullText) // 遍历全文本列表的所有元素
                    {
                        writer.WriteLine(paragraphText); // 将当前元素的段落文字写入文本文件中，并换行
                    }
                }

                // 写入目标PDF文档
                string targetPdfFilePath = Path.Combine(targetFolderPath, $"{CleanFileAndFolderName($"Mrg_{targetFileMainName}")}.pdf"); // 获取目标文本文件的路径全名

                using (PdfWriter writer = new PdfWriter(targetPdfFilePath)) // 创建PDF写入器对象，赋值给PDF写入器对象
                using (PdfDocument pdf = new PdfDocument(writer)) // 创建PDF文档对象，赋值给PDF文档对象
                using (ITextDocument document = new ITextDocument(pdf)) // 创建文档对象，赋值给文档对象
                {
                    PdfFont font = PdfFontFactory.CreateFont("STSong-Light", "UniGB-UCS2-H", PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED); // 创建pdf字体对象：中文宋体，Adobe-GB1符集UCS-2编码，水平书写，优先嵌入字体
                    // 遍历字符串列表，
                    foreach (string paragraphText in lstFullText)
                    {
                        ITextParagraph paragraph = new ITextParagraph(paragraphText).SetFont(font); // 为当前字符串创建一个段落，使用已定义的字体
                        document.Add(paragraph); // 将段落添加到文档中
                    }
                }

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

        public async Task ShowSystemInfoAsync()
        {
            try
            {
                HardwareInfo hardwareInfo = new HardwareInfo();

                // 定义异步委托方法，用于刷新硬件信息
                async Task RefreshHardwareInfoAsync()
                {
                    await Task.Run(() => hardwareInfo.RefreshAll());
                }
                await taskManager.RunTaskAsync(RefreshHardwareInfoAsync);

                // 创建 DataTable
                DataTable systemInfoTable = new DataTable("System Information");
                systemInfoTable.Columns.Add("Hardware Name", typeof(string));
                systemInfoTable.Columns.Add("Hardware Info", typeof(string));

                int i;

                // 操作系统
                systemInfoTable.Rows.Add($"Operating System", hardwareInfo.OperatingSystem.ToString());

                // 计算机系统
                i = 1;
                foreach (var computerSystem in hardwareInfo.ComputerSystemList)
                {
                    systemInfoTable.Rows.Add($"Computer System", computerSystem.ToString());
                }

                // BIOS
                i = 1;
                foreach (var bios in hardwareInfo.BiosList)
                {
                    systemInfoTable.Rows.Add($"BIOS", bios.ToString());
                }

                // 主板
                i = 1;
                foreach (var motherboard in hardwareInfo.MotherboardList)
                {
                    systemInfoTable.Rows.Add($"Motherboard", motherboard.ToString());
                }

                // CPU
                i = 1;
                foreach (var cpu in hardwareInfo.CpuList)
                {
                    systemInfoTable.Rows.Add($"CPU {i++}", cpu.ToString());
                }

                // 内存
                i = 1;
                long totalMemCapacity = 0;
                foreach (var memory in hardwareInfo.MemoryList)
                {
                    systemInfoTable.Rows.Add($"Memory {i++}", memory.ToString());
                    totalMemCapacity += (long)(Convert.ToInt64(memory.Capacity) / Math.Pow(1024, 3));  // 将容量从Byte换算到GB
                }
                systemInfoTable.Rows.Add("Total Memory Capacity", $"{totalMemCapacity.ToString()} GB");

                // 硬盘
                i = 1;
                foreach (var disk in hardwareInfo.DriveList)
                {
                    long diskSize = (long)(Convert.ToInt64(disk.Size) / Math.Pow(1024, 3)); // 将硬盘容量转换为GB
                    string diskInfo = $"{disk.ToString()}\nDisk Size: {diskSize.ToString()} GB";
                    systemInfoTable.Rows.Add($"Disk {i++}", diskInfo);
                }

                // 视频控制器
                i = 1;
                foreach (var videoController in hardwareInfo.VideoControllerList)
                {
                    systemInfoTable.Rows.Add($"Video Controller {i++}", videoController.ToString());
                }

                // 音频适配器
                i = 1;
                foreach (var soundDevice in hardwareInfo.SoundDeviceList)
                {
                    systemInfoTable.Rows.Add($"Sound Device {i++}", soundDevice.ToString());
                }

                // 网络适配器
                i = 1;
                foreach (var networkAdapter in hardwareInfo.NetworkAdapterList)
                {
                    List<string> lstIPAddressInfos = new List<string>();
                    foreach (var ipAddress in networkAdapter.IPAddressList)
                    {
                        lstIPAddressInfos.Add(ipAddress.ToString());
                    }
                    string ipAddressInfo = "IP Address: " + string.Join("; ", lstIPAddressInfos);
                    string networkAdapterInfo = $"{networkAdapter.ToString()}\n{ipAddressInfo}";
                    systemInfoTable.Rows.Add($"Network Adapter {i++}", networkAdapterInfo);
                }

                // 显示 DataTable
                if (GetInstanceCountByHandle<DataGridWindow>() < 1) //如果被打开的浏览器窗口数量小于1个，则新建一个浏览器窗口实例并显示
                {
                    DataGridWindow systemInfoWindow = new DataGridWindow("System Info", systemInfoTable);
                    systemInfoWindow.Show();
                }

            }
            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }




        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Top = 50.0;
            this.Left = SystemParameters.WorkArea.Width - this.Width - 150.0;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            settingsManager.SaveSettings(appSettings); // 保存应用程序设置到Json文件
            recordsManager.SaveSettings(latestRecords);  // 保存最近记录到Json文件

            Environment.Exit(0); // 退出程序，关闭所有窗口
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

            InputDialog inputDialog = new InputDialog(question: "Markdown", defaultAnswer: "ABC", acceptsReturn: true); //弹出功能选择对话框
            if (inputDialog.ShowDialog() == false) //如果对话框返回false（点击了Cancel），则结束本过程
            {
                return;
            }
            //获取对话框返回的功能选项
            string result = inputDialog.Answer.RemoveMarkdownMarks();
            ShowMessage($"转换后的文字为：\n\n{result}");
        }




    }

}
