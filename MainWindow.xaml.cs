using iText.Kernel.Font;
using iText.Kernel.Pdf;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using NPOI.HSSF.Record.CF;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using static COMIGHT.Methods;
using static COMIGHT.PublicVariables;
using DataTable = System.Data.DataTable;
using Document = iText.Layout.Document;
using MSExcel = Microsoft.Office.Interop.Excel;
using MSExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using MSWord = Microsoft.Office.Interop.Word;
using MSWordDocument = Microsoft.Office.Interop.Word.Document;
using Paragraph = iText.Layout.Element.Paragraph;
using Task = System.Threading.Tasks.Task;
using Window = System.Windows.Window;


namespace COMIGHT
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

        public static TaskManager taskManager = new TaskManager(); //定义任务管理器对象变量

        // 定义应用设置管理器、用户使用记录管理器对象，以及应用设置类、用户使用记录类，用于读取和保存设置
        public static SettingsManager<AppSettings> settingsManager = new SettingsManager<AppSettings>(settingsJsonFilePath);
        public static SettingsManager<LatestRecords> recordsManager = new SettingsManager<LatestRecords>(recordsJsonFilePath);
        public static AppSettings appSettings = new AppSettings();
        public static LatestRecords latestRecords = new LatestRecords();

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  //定义EPPlus库许可证类型为非商用！！！

            this.Title = $"COMIGHT Assistant {DateTime.Now:yyyy}";

            lblStatus.DataContext = taskManager; // 将状态标签控件的数据环境设为任务管理器对象
            lblIntro.Content = $"For Better Productivity. © Yuechen Lou 2022-{DateTime.Now:yyyy}";

            appSettings = settingsManager.GetSettings(); // 从应用设置管理器中读取应用设置
            latestRecords = recordsManager.GetSettings(); // 从用户使用记录管理器中读取用户使用记录
        }

        private async void MnuBatchConvertOfficeFilesTypes_Click(object sender, RoutedEventArgs e)
        {
            await BatchConvertOfficeFilesTypes();
        }

        private async void MnuBatchFormatWordDocuments_Click(object sender, RoutedEventArgs e)
        {
            await BatchFormatWordDocumentsAsync();
        }

        private void MnuBatchProcessExcelWorksheets_Click(object sender, RoutedEventArgs e)
        {
            BatchProcessExcelWorksheets();
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

        private void mnuConvertMarkdownIntoWord_Click(object sender, RoutedEventArgs e)
        {
            ConvertMarkDownIntoWord();
        }

        private void MnuCreateNameCards_Click(object sender, RoutedEventArgs e)
        {
            CreateNameCards();
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
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void MnuMakeFileList_Click(object sender, RoutedEventArgs e)
        {
            MakeFileList();
        }

        private void MnuMakeFolders_Click(object sender, RoutedEventArgs e)
        {
            MakeFolders();
        }

        private void MnuMergeDocumentsAndTables_Click(object sender, RoutedEventArgs e)
        {
            MergeDocumentsAndTables();
        }

        private void mnuOpenSavingFolder_Click(object sender, RoutedEventArgs e)
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
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void MnuScreenStocks_Click(object sender, RoutedEventArgs e)
        {
            ScreenStocks();
        }

        private void MnuSettings_Click(object sender, RoutedEventArgs e)
        {
            SettingsDialog settingDialog = new SettingsDialog();
            settingDialog.ShowDialog();
        }

        private void MnuSplitExcelWorksheet_Click(object sender, RoutedEventArgs e)
        {
            SplitExcelWorksheet();
        }

        private void MnuSubConverter_Click(object sender, RoutedEventArgs e)
        {
            if (GetInstanceCountByHandle<SubConverterWindow>() < 1) //如果被打开的浏览器窗口数量小于1个，则新建一个浏览器窗口实例并显示
            {
                SubConverterWindow subConverterWindow = new SubConverterWindow();
                subConverterWindow.Show();
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

                await taskManager.RunTaskAsync(() => FormatWordDocumentsAsync(filePaths)); // 创建一个任务管理器实例，并使用RunTaskAsync方法异步执行FormatWordDocumentsAsync过程
                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
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
                MessageBox.Show($"{fileNum} files processed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BatchProcessExcelWorksheets()
        {
            string currentFilePath = "";
            try
            {
                List<string> lstFunctions = new List<string> {"0-Cancel", "1-Merge Records", "2-Accumulate Values", "3-Extract Cell Data", "4-Convert Textual Numbers into Numeric",
                    "5-Copy Formula to Multiple Worksheets", "6-Prefix Workbook Filenames with Cell Data", "7-Adjust Worksheet Format for Printing"};

                string latestBatchProcessWorkbookOption = latestRecords.LatestBatchProcessWorkbookOption; //读取用户使用记录中保存的批量处理Excel工作簿功能选项字符串
                InputDialog inputDialog = new InputDialog(question: "Select the function", options: lstFunctions, defaultAnswer: latestBatchProcessWorkbookOption); //弹出功能选择对话框
                if (inputDialog.ShowDialog() == false) //如果对话框返回false（点击了Cancel），则结束本过程
                {
                    return;
                }
                string batchProcessWorkbookOption = inputDialog.Answer;
                latestRecords.LatestBatchProcessWorkbookOption = batchProcessWorkbookOption; //将对话框返回的批量处理Excel工作簿功能选项字符串赋值给用户使用记录
                recordsManager.SaveSettings(latestRecords); //保存用户使用记录

                int functionNum = lstFunctions.Contains(batchProcessWorkbookOption) ? lstFunctions.IndexOf(batchProcessWorkbookOption) : -1; //获取对话框返回的功能选项在功能列表中的索引号：如果功能列表包含功能选项，则得到对应的索引号；否则，得到-1

                if (functionNum < 1 || functionNum > 7) //如果功能选项索引号不在设定范围，则结束本过程
                {
                    return;
                }

                List<string>? filePaths = SelectFiles(FileType.Excel, true, "Select Excel Files"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                int fileNum;
                int excelWorksheetIndexLower = 0;
                int excelWorksheetIndexUpper = 0;
                string? excelWorksheetName = null;
                bool useExcelWorksheetIndex = true;
                int headerRowCount = 0;
                int footerRowCount = 0;
                List<string>? lstOperatingRangeAddresses = null;

                string latestExcelWorksheetIndexesStr = latestRecords.LatestExcelWorksheetIndexesStr; //读取用户使用记录中保存的Excel工作表索引号范围字符串
                inputDialog = new InputDialog(question: "Input the indexes range of worksheets to be processed (separated by a hyphen, e.g. \"1-3\"); Leave blank to designate the worksheet name", defaultAnswer: latestExcelWorksheetIndexesStr); //弹出对话框，输入工作表索引号范围
                if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则结束本过程
                {
                    return;
                }

                string excelWorksheetIndexesStr = inputDialog.Answer; //获取对话框返回的Excel工作表索引号范围字符串
                if (!string.IsNullOrWhiteSpace(excelWorksheetIndexesStr)) //如果Excel工作表索引号范围字符串不为null或全空白字符
                {
                    latestRecords.LatestExcelWorksheetIndexesStr = excelWorksheetIndexesStr; // 将对话框返回的Excel工作表索引号范围字符串赋值给用户使用记录
                    recordsManager.SaveSettings(latestRecords); //保存用户使用记录
                    //将Excel索引号字符串拆分成数组，转换成列表，并移除每个元素的首尾空白字符
                    List<string> lstExcelWorksheetIndexesStr = excelWorksheetIndexesStr.Split('-').ToList().ConvertAll(e => e.Trim());
                    excelWorksheetIndexLower = Convert.ToInt32(lstExcelWorksheetIndexesStr[0]) - 1; //获取Excel工作表索引号范围起始值（Excel工作表索引号从1开始，EPPlus从0开始）
                    excelWorksheetIndexUpper = Convert.ToInt32(lstExcelWorksheetIndexesStr[1]) - 1; //获取Excel工作表索引号范围结束值
                    useExcelWorksheetIndex = true; //“使用工作表索引号”变量赋值为true
                }
                else
                {
                    string latestExcelWorksheetName = latestRecords.LatestExcelWorksheetName; //读取用户使用记录中保存的Excel工作表名称
                    inputDialog = new InputDialog(question: "Input the worksheet name (one worksheet per operation)", defaultAnswer: latestExcelWorksheetName); //弹出对话框，输入工作表名称
                    if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则结束本过程
                    {
                        return;
                    }
                    excelWorksheetName = inputDialog.Answer;
                    latestRecords.LatestExcelWorksheetName = excelWorksheetName; // 将对话框返回的Excel工作表名称赋值给用户使用记录
                    recordsManager.SaveSettings(latestRecords); //保存用户使用记录
                    useExcelWorksheetIndex = false; //“使用工作表索引号”变量赋值为false
                }

                switch (functionNum) //根据功能序号进入相应的分支
                {
                    case 1: //记录合并
                    case 7: //调整工作表打印版式
                        GetHeaderAndFooterRowCount(out headerRowCount, out footerRowCount); //获取表头、表尾行数
                        break;

                    case 2:
                    case 3:
                    case 4:
                    case 5:
                    case 6: //2-数值累加, 3-提取单元格数据, 4-文本型数字转数值型, 5-复制公式到多Excel工作表, 6-提取单元格数据给工作簿文件名加前缀
                        string latestOperatingRangeAddresses = latestRecords.LatestOperatingRangeAddresses; //读取用户使用记录中保存的操作区域
                        inputDialog = new InputDialog(question: "Input the operating range addresses (separated by a comma, e.g. \"B2:C3,B4:C5\")", defaultAnswer: latestOperatingRangeAddresses); //弹出对话框，输入操作区域
                        if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则结束本过程
                        {
                            return;
                        }
                        string operatingRangeAddresses = inputDialog.Answer; //获取对话框返回的操作区域
                        latestRecords.LatestOperatingRangeAddresses = operatingRangeAddresses; //将对话框返回的操作区域赋值给用户使用记录
                        recordsManager.SaveSettings(latestRecords); //保存用户使用记录
                        //将操作区域地址拆分为数组，转换成列表，并移除每个元素的首尾空白字符
                        lstOperatingRangeAddresses = operatingRangeAddresses.Split(',').ToList().ConvertAll(e => e.Trim());
                        break;

                }

                ExcelPackage targetExcelPackage = new ExcelPackage(); //新建Excel包，赋值给目标Excel包变量
                ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add("Sheet1"); //在目标Excel工作簿中添加一个工作表，赋值给目标工作表变量
                string? excelFileName = null; //定义被处理Excel工作簿文件名变量
                string? targetFolderPath = null; //定义目标文件夹路径变量
                string? targetFileMainName = null; //定义目标文件主名变量
                string? targetExcelWorkbookPrefix = null; //定义目标Excel工作簿前缀变量
                DataTable? dataTable = null; //定义DataTable变量
                DataRow? dataRow = null; //定义DataTable行变量
                List<string>? templateExcelFilePaths = null; //定义模板Excel文件列表
                ExcelPackage? templateExcelPackage = null; //定义模板Excel包变量
                ExcelWorksheet? templateExcelWorksheet = null; //定义模板Excel工作表变量

                fileNum = 1;
                foreach (string excelFilePath in filePaths) //遍历所有文件
                {
                    currentFilePath = excelFilePath; //将当前Excel文件路径全名赋值给当前文件路径全名变量
                    List<string> lstPrefixes = new List<string>(); //定义文件名前缀列表（给Excel文件名加前缀用）

                    using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelFilePath))) //打开当前Excel工作簿，赋值给Excel包变量
                    {
                        ExcelWorkbook excelWorkbook = excelPackage.Workbook; //将工作簿赋值给工作簿变量
                        excelFileName = Path.GetFileName(excelFilePath); //获取当前被处理Excel工作簿文件主名

                        if (fileNum == 1) //如果当前是第一个Excel工作簿文件
                        {
                            targetFileMainName = Path.GetFileNameWithoutExtension(excelFileName); //获取目标文件的文件主名
                            targetFolderPath = appSettings.SavingFolderPath; //获取目标文件的文件夹路径
                        }

                        //获取被处理Excel工作表索引号范围
                        if (useExcelWorksheetIndex) //如果使用Excel工作表索引号
                        {
                            //获取被处理Excel工作表索引号上下限，如果大于工作表数量-1，则限定为工作表数量-1
                            excelWorksheetIndexLower = Math.Min(excelWorksheetIndexLower, excelWorkbook.Worksheets.Count - 1);
                            excelWorksheetIndexUpper = Math.Min(excelWorksheetIndexUpper, excelWorkbook.Worksheets.Count - 1);
                        }
                        else //否则（使用Excel工作表名称）
                        {
                            //如果当前Excel工作簿没有指定名称的工作表，则直接跳过当前循环进入下一个循环
                            if (!excelWorkbook.Worksheets.Any(sheet => sheet.Name.Trim() == excelWorksheetName))
                            {
                                continue;
                            }
                            //获取被处理Excel工作表索引号上下限：筛选出工作表名称移除首尾空白字符后与指定名称相同的工作表，将其中第一个的索引号作为下限；上限与下限相同
                            excelWorksheetIndexLower = excelWorkbook.Worksheets.Where(sheet => sheet.Name.Trim() == excelWorksheetName)
                                .FirstOrDefault()!.Index;
                            excelWorksheetIndexUpper = excelWorksheetIndexLower;
                        }

                        for (int i = excelWorksheetIndexLower; i <= excelWorksheetIndexUpper; i++) //遍历指定范围内的所有Excel工作表
                        {
                            ExcelWorksheet excelWorksheet = excelWorkbook.Worksheets[i];
                            //如果当前Excel工作表为隐藏且使用工作表索引号，则抛出异常
                            if (excelWorksheet.Hidden != eWorkSheetHidden.Visible && useExcelWorksheetIndex)
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

                                    targetExcelWorkbookPrefix = "Mrg"; //目标Excel工作簿类型变量赋值为“合并”

                                    TrimCellsStrings(excelWorksheet); //删除当前Excel工作表内所有文本型单元格值的首尾空格
                                    RemoveWorksheetEmptyRowsAndColumns(excelWorksheet); //删除当前Excel工作表内所有空白行和空白列
                                    //如果当前被处理Excel工作表的已使用行数（如果工作表为空，则为0）小于等于表头表尾行数之和，只有表头表尾无有效数据，则直接跳过当前循环并进入下一个循环
                                    if ((excelWorksheet.Dimension?.Rows ?? 0) <= headerRowCount + footerRowCount)
                                    {
                                        continue;
                                    }

                                    int sourceStartRowIndex = (fileNum == 1 && i == excelWorksheetIndexLower) ? 1 : headerRowCount + 1; //获取被处理工作表起始行索引号：如果当前是第一个Excel工作簿的第一个工作表，则得到1；否则得到表头行数+1
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
                                        //在目标工作表的表头最末行的第1、2列单元格分别添加"工作簿文件名", "工作表名"的列名
                                        targetExcelWorksheet.Cells[headerRowCount, 1, headerRowCount, 2].LoadFromArrays(new List<object[]> { new object[] { "Source Workbook", "Source Worksheet" } });
                                    }

                                    break;

                                case 2: //数值累加

                                    targetExcelWorkbookPrefix = "Accum"; //目标Excel工作簿类型变量赋值为“累加”

                                    if (fileNum == 1 && i == excelWorksheetIndexLower) // 如果是第一个文件的第一个Excel工作表
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

                                    targetExcelWorkbookPrefix = "Extcd"; //目标Excel工作簿类型变量赋值为“提取”

                                    if (fileNum == 1 && i == excelWorksheetIndexLower) //如果是第一个文件的第一个Excel工作表
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
                                    if (fileNum == filePaths.Count && i == excelWorksheetIndexUpper
                                        && dataTable!.Rows.Count * dataTable.Columns.Count > 0)
                                    {
                                        targetExcelWorksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
                                    }
                                    break;

                                case 4: //文本型数字转数值型

                                    targetExcelWorkbookPrefix = "Fail"; //目标Excel工作簿类型变量赋值为“失败”

                                    if (fileNum == 1 && i == excelWorksheetIndexLower) // 如果是第一个文件的第一个Excel工作表
                                    {
                                        dataTable = new DataTable(); //定义DataTable
                                        dataTable.Columns.AddRange(new DataColumn[]
                                            {
                                                new DataColumn("Source Workbook"),
                                                new DataColumn("Source Worksheet"),
                                                new DataColumn("Unconverted Address"),
                                                new DataColumn("Unconverted Value")
                                            }); //向DataTable添加列

                                    }

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
                                                else //否则
                                                {
                                                    dataRow = dataTable!.NewRow(); //定义DataTable新数据行
                                                    //将相关数据填入对应的数据列
                                                    dataRow["Source Workbook"] = excelFileName;
                                                    dataRow["Source Worksheet"] = excelWorksheet.Name;
                                                    dataRow["Unconverted Address"] = cell.Address;
                                                    dataRow["Unconverted Value"] = cell.Value;
                                                    dataTable.Rows.Add(dataRow); //向DataTable添加数据行
                                                }
                                            }
                                        }
                                    }
                                    if (i == excelWorksheetIndexUpper) //如果当前Excel工作表是最后一个，则保存当前被处理Excel工作簿
                                    {
                                        excelPackage.Save();
                                    }

                                    //如果当前文件是文件列表中的最后一个，且当前Excel工作表也是最后一个，且DataTable的行数和列数均不为0，则将DataTable写入目标工作表
                                    if (fileNum == filePaths.Count && i == excelWorksheetIndexUpper
                                        && dataTable!.Rows.Count * dataTable.Columns.Count > 0)
                                    {
                                        targetExcelWorksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
                                    }

                                    break;

                                case 5: //复制公式到多Excel工作表

                                    if (fileNum == 1 && i == excelWorksheetIndexLower) // 如果是第一个文件的第一个Excel工作表
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

                                    if (i == excelWorksheetIndexUpper) //如果当前Excel工作表是最后一个，则保存当前被处理Excel工作簿
                                    {
                                        excelPackage.Save();
                                    }

                                    break;

                                case 6: //提取单元格数据给工作簿文件名加前缀
                                    foreach (string anOperatingRange in lstOperatingRangeAddresses!) // '遍历所有操作区域
                                    {
                                        for (int k = 0; k < excelWorksheet.Cells[anOperatingRange].Rows; k++) //遍历Excel工作表操作区域行偏移值（第1行相对第1行的偏移值为0，最后一行相对第1行的偏移值为区域总行数-1）
                                        {
                                            for (int l = 0; l < excelWorksheet.Cells[anOperatingRange].Columns; l++) //遍历Excel工作表操作区域列偏移值（第1列相对第1列的偏移值为0，最后一列相对第1列的偏移值为区域总列数-1）
                                            {
                                                lstPrefixes.Add(excelWorksheet.Cells[anOperatingRange].Offset(k, l, 1, 1).Text); //将被处理Excel工作表操作区域第1行第1列的单元格向右、向下偏移k、l个单位的单元格数据转换成文本并追加到前缀列表
                                            }
                                        }

                                    }

                                    if (i == excelWorksheetIndexUpper) //如果当前Excel工作表是最后一个
                                    {
                                        string prefixes = string.Join(' ', lstPrefixes); //合并前缀列表中的字符串，当中用空格分隔
                                        excelPackage.Dispose(); //关闭当前被处理Excel工作簿
                                                                //获取新文件名：将前缀加到当前文件主名之前，清除不能作为文件名的字符并截取指定数量的字符，再加上当前文件扩展名
                                        string renamedExcelFileName = CleanFileAndFolderName($"{prefixes}_{Path.GetFileNameWithoutExtension(excelFilePath)}", 40) + Path.GetExtension(excelFilePath);
                                        string renamedExcelFilePath = Path.Combine(Path.GetDirectoryName(excelFilePath)!, renamedExcelFileName); //获取新文件路径全名
                                        File.Move(excelFilePath, renamedExcelFilePath); //将当前Excel工作簿文件更名
                                    }

                                    break;

                                case 7: //调整工作表打印版式
                                    TrimCellsStrings(excelWorksheet); //删除当前Excel工作表内所有文本型单元格值的首尾空格
                                    RemoveWorksheetEmptyRowsAndColumns(excelWorksheet); //删除当前Excel工作表内所有空白行和空白列
                                    FormatExcelWorksheet(excelWorksheet, headerRowCount, footerRowCount); //设置当前Excel工作表格式
                                    if (i == excelWorksheetIndexUpper) //如果当前Excel工作表是最后一个，则保存当前被处理Excel工作簿
                                    {
                                        excelPackage.Save();
                                    }

                                    break;
                            }
                        }

                    }
                    fileNum++; //文件计数器加1
                }

                if (targetExcelWorkbookPrefix != null)  //如果目标Excel工作簿前缀不为null（执行功能1-4时，将生成新工作簿并保存）
                {
                    //创建目标文件夹
                    if (!Directory.Exists(targetFolderPath))
                    {
                        Directory.CreateDirectory(targetFolderPath!);
                    }

                    //获取目标工作表表头行数
                    //根据功能序号返回相应的目标工作表表头行数
                    int targetHeaderRowCount = functionNum switch
                    {
                        1 => headerRowCount,  //记录合并，输出记录合并后的汇总表，表头行数为源数据表格的表头行数
                        3 => 1,  //提取单元格数据，输出提取单元格值后的汇总表，表头行数为1
                        4 => 1,  //文本型数字转数值型，输出未能转换为数值的单元格地址和值的汇总表，表头行数为1
                        _ => 0  //其余情况，表头行数为0
                    };

                    FormatExcelWorksheet(targetExcelWorksheet, targetHeaderRowCount, 0); //设置目标Excel工作表格式
                    FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath!, $"{targetExcelWorkbookPrefix}_{targetFileMainName}.xlsx")); //获取目标Excel工作簿文件路径全名信息
                    targetExcelPackage.SaveAs(targetExcelFile);
                    targetExcelPackage.Dispose(); //关闭目标Excel工作簿
                }
                templateExcelPackage?.Dispose(); //关闭模板Excel工作簿（仅在模板工作簿已打开的情况下）
                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message} at {currentFilePath}。", "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
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

                GetHeaderAndFooterRowCount(out int headerRowCount, out int footerRowCount); //获取表头、表尾行数

                string? columnLetter = GetKeyColumnLetter(); //获取主键列符
                if (columnLetter == null) //如果主键列符为null，则结束本过程
                {
                    return;
                }

                DataTable? startDataTable = ReadExcelWorksheetIntoDataTable(startFilePaths[0], 1, headerRowCount, footerRowCount); //读取起始数据Excel工作簿的第1张工作表，赋值给起始DataTable变量
                DataTable? endDataTable = ReadExcelWorksheetIntoDataTable(endFilePaths[0], 1, headerRowCount, footerRowCount); //读取终点数据Excel工作簿的第1张工作表，赋值给终点DataTable变量

                if (startDataTable == null || endDataTable == null) //如果起始DataTable或终点DataTable有一个为null，则结束本过程
                {
                    return;
                }

                //获取Excel工作表的主键列对应的DataTable主键数据列的名称（工作表列索引号从1开始，DataTable从0开始）
                string keyDataColumnName = endDataTable.Columns[ConvertColumnLettersIntoIndex(columnLetter) - 1].ColumnName;

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

                differenceDataTable = RemoveDataTableEmptyRowsAndColumns(differenceDataTable); // 移除差异DataTable中的空数据行和空数据列

                if (differenceDataTable.Rows.Count * differenceDataTable.Columns.Count == 0) //如果差异DataTable的数据行数或列数有一个为0，则抛出异常
                {
                    throw new Exception("No difference detected.");
                }

                string targetFolderPath = appSettings.SavingFolderPath; //获取目标文件夹路径

                //创建目标文件夹
                if (!Directory.Exists(targetFolderPath))
                {
                    Directory.CreateDirectory(targetFolderPath!);
                }
                FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath!, $"Comp_{Path.GetFileNameWithoutExtension(endFilePaths[0])}.xlsx")); //获取目标Excel工作簿文件路径全名信息

                using (ExcelPackage excelPackage = new ExcelPackage()) //新建Excel包，赋值给Excel包变量
                {
                    ExcelWorksheet targetExcelWorksheet = excelPackage.Workbook.Worksheets.Add($"Sheet1"); //新建“数据比较”Excel工作表，赋值给目标工作表变量
                    targetExcelWorksheet.Cells["A1"].LoadFromDataTable(differenceDataTable, true); //将DataTable数据导入目标Excel工作表（true代表将表头赋给第一行）

                    FormatExcelWorksheet(targetExcelWorksheet, 1, 0); //设置目标Excel工作表格式
                    excelPackage.SaveAs(targetExcelFile);
                }
                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private static void ConvertDocumentByPandoc(string fromType, string toType, string fromFilePath, string toFilePath)
        {
            try
            {
                string? pandocPath = appSettings.PandocPath; //读取设置中保存的Pandoc程序文件路径全名，赋值给Pandoc程序文件路径全名变量

                ProcessStartInfo startInfo = new ProcessStartInfo //创建ProcessStartInfo对象，包含了启动新进程所需的信息，赋值给启动进程信息变量
                {
                    FileName = pandocPath, // 指定pandoc应用程序的文件路径全名
                                           //指定参数，-f从markdown -t转换为docx -o输出文件路径全名，\"用于确保文件路径（可能包含空格）被视为pandoc命令的单个参数
                    Arguments = $"-f {fromType} -t {toType} \"{fromFilePath}\" -o \"{toFilePath}\"",
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
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ConvertMarkDownIntoWord()
        {
            try
            {
                InputDialog inputDialog = new InputDialog(question: "Input the text to be converted", defaultAnswer: "", textboxHeight: 300, acceptReturn: true); //弹出对话框，输入功能选项
                if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则结束本过程
                {
                    return;
                }

                string MDText = inputDialog.Answer;
                //将导出文本框的文字按换行符拆分为数组（删除每个元素前后空白字符，并删除空白元素），转换成列表
                List<string> lstParagraphs = MDText
                    .Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();

                if (lstParagraphs.Count == 0) //如果段落列表元素数为0，则抛出异常
                {
                    throw new Exception("No valid text found.");
                }

                string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径
                // 获取目标文件主名：将段落列表0号元素（一般为标题）删除Markdown标记，截取前40个字符
                string targetFileMainName = CleanFileAndFolderName(lstParagraphs[0].RemoveMarkDownMarks(), 40);

                //创建目标文件夹
                if (!Directory.Exists(targetFolderPath))
                {
                    Directory.CreateDirectory(targetFolderPath);
                }
                //导入目标Markdown文档
                string targetMDFilePath = Path.Combine(targetFolderPath, $"{targetFileMainName}.md"); //获取目标Markdown文档文件路径全名
                File.WriteAllText(targetMDFilePath, MDText); //将导出文本框内的markdown文字导入目标Markdown文档

                //将目标Markdown文档转换为目标Word文档
                string targetWordFilePath = Path.Combine(targetFolderPath, $"{targetFileMainName}.docx"); //获取目标Word文档文件路径全名

                ConvertDocumentByPandoc("markdown", "docx", targetMDFilePath, targetWordFilePath); // 将目标Markdown文档转换为目标Word文档
                File.Delete(targetMDFilePath); //删除Markdown文件

                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }

        public async Task BatchConvertOfficeFilesTypes()
        {
            MSExcel.Application? msExcelApp = null;
            MSWord.Application? msWordApp = null;
            try
            {
                List<string>? filePaths = SelectFiles(FileType.Convertible, true, "Select Old Version Office or WPS Files"); //打开文件选择对话框，选择文件
                if (filePaths == null) // 如果文件列表为null，则结束本过程
                {
                    return;
                }

                string folderPath = Path.GetDirectoryName(filePaths[0])!; //获取保存转换文件的文件夹路径

                //定义可用Excel打开的文件正则表达式变量，匹配模式为: "xls"或"et"，结尾标记，忽略大小写
                Regex regExExcelFile = new Regex(@"(?:xls|et)$", RegexOptions.IgnoreCase);
                //定义可用Word打开的文件正则表达式，匹配模式为: "doc"或"wps"，结尾标记，忽略大小写
                Regex regExWordFile = new Regex(@"(?:doc|wps)$", RegexOptions.IgnoreCase);

                Task task = Task.Run(() => process());
                void process()
                {
                    if (filePaths.Any(f => regExExcelFile.IsMatch(f))) //如果文件列表中有任一文件被可用Excel打开的文件正则表达式匹配成功
                    {
                        msExcelApp = new MSExcel.Application(); //打开Excel应用程序，赋值给Excel应用程序变量
                        msExcelApp.Visible = false;
                        msExcelApp.DisplayAlerts = false;
                    }

                    if (filePaths.Any(f => regExWordFile.IsMatch(f))) //如果文件列表中有任一文件被可用Word打开的文件正则表达式匹配成功
                    {
                        msWordApp = new MSWord.Application(); //打开Word应用程序，赋值给Word应用程序变量
                        msWordApp.Visible = false;
                        msWordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    }

                    foreach (string filePath in filePaths) //遍历所有文件
                    {
                        if (regExExcelFile.IsMatch(filePath)) //如果当前文件名被可用Excel打开的文件正则表达式匹配成功
                        {
                            MSExcelWorkbook msExcelWorkbook = msExcelApp!.Workbooks.Open(filePath); //打开当前Excel工作簿，赋值给Excel工作簿变量
                            string targetFilePath = Path.Combine(folderPath, $"{Path.GetFileNameWithoutExtension(filePath)}.xlsx"); //获取目标文件路径全名
                            //获取目标文件路径全名：如果目标文件不存在，则得到原目标文件路径全名；否则，在原目标文件主名后添加4位随机数，得到新目标文件路径全名
                            targetFilePath = !File.Exists(targetFilePath) ? targetFilePath :
                                Path.Combine(folderPath, $"{Path.GetFileNameWithoutExtension(filePath)}{new Random().Next(1000, 10000)}.xlsx"); //DateTime.Now.ToString("ssfff")
                            msExcelWorkbook.SaveAs(Filename: targetFilePath, FileFormat: XlFileFormat.xlWorkbookDefault); //目标Excel工作簿另存为xlsx格式
                            msExcelWorkbook.Close(); //关闭当前Excel工作簿
                        }
                        else if (regExWordFile.IsMatch(filePath)) //如果当前文件名被可用Word打开的文件正则表达式匹配成功
                        {
                            MSWordDocument msWordDocument = msWordApp!.Documents.Open(filePath); //打开当前Word文档，赋值给Word文档变量
                            string targetFilePath = Path.Combine(folderPath, $"{Path.GetFileNameWithoutExtension(filePath)}.docx"); //获取目标Word文件路径全名
                            //获取目标文件路径全名：如果目标文件不存在，则得到原目标文件路径全名；否则，在原目标文件主名后添加4位随机数，得到新目标文件路径全名
                            targetFilePath = !File.Exists(targetFilePath) ? targetFilePath :
                                Path.Combine(folderPath, $"{Path.GetFileNameWithoutExtension(filePath)}{new Random().Next(1000, 10000)}.docx");
                            //目标Word文件另存为docx格式，使用最新Word版本兼容模式
                            msWordDocument.SaveAs2(FileName: targetFilePath, FileFormat: WdSaveFormat.wdFormatDocumentDefault, CompatibilityMode: WdCompatibilityMode.wdCurrent);
                            msWordDocument.Close(); //关闭当前Word文件
                        }
                        File.Delete(filePath); //删除当前文件
                    }
                }
                await task;

                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            finally
            {
                KillOfficeApps(new object[] { msExcelApp!, msWordApp! }); //结束Office应用程序进程
            }

        }

        public void CreateNameCards()
        {
            try
            {
                List<string>? filePaths = SelectFiles(FileType.Excel, false, "Select the Excel File of Name List"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                //获取已安装的字体名称：读取系统中已安装的字体，赋值给字体名称列表变量
                InstalledFontCollection installedFontCollention = new InstalledFontCollection();
                List<string> lstFontNames = installedFontCollention.Families.Select(f => f.Name).ToList();

                string latestFontName = latestRecords.LatestNameCardFontName; //读取用户使用记录中保存的字体名称
                InputDialog inputDialog = new InputDialog(question: "Select the font", options: lstFontNames, defaultAnswer: latestFontName); //弹出对话框，输入字体名称

                if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则结束本过程
                {
                    return;
                }
                string fontName = inputDialog.Answer;
                latestRecords.LatestNameCardFontName = fontName; // 将对话框返回的字体名称赋值给用户使用记录
                recordsManager.SaveSettings(latestRecords); //保存用户使用记录

                using (ExcelPackage sourceExcelPackage = new ExcelPackage(new FileInfo(filePaths[0]))) //打开源数据Excel工作簿，赋值给源数据Excel包变量（源数据Excel工作簿）
                using (ExcelPackage targetExcelPackage = new ExcelPackage()) //新建Excel包，赋值给目标Excel包变量（目标Excel工作簿）
                {
                    ExcelWorksheet sourceExcelWorksheet = sourceExcelPackage.Workbook.Worksheets[0]; //将工作表1（0号）赋值给源工作表变量

                    TrimCellsStrings(sourceExcelWorksheet); //删除源数据Excel工作表内所有文本型单元格值的首尾空格
                    RemoveWorksheetEmptyRowsAndColumns(sourceExcelWorksheet); //删除源数据Excel工作表内所有空白行和空白列
                    if (sourceExcelWorksheet.Dimension == null) //如果工作表为空，则抛出异常
                    {
                        throw new Exception("No valid data found.");
                    }

                    for (int i = 1; i <= sourceExcelWorksheet.Dimension.End.Row; i++) //遍历源数据工作表所有行
                    {
                        string name = sourceExcelWorksheet.Cells[i, 1].Text; // 将A列当前行单元格的文字赋值给名称变量

                        // 在目标工作簿中添加一个工作表，表名为编号i加名称后截取前10个字符，赋值给目标Excel工作表变量
                        ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add(CleanFileAndFolderName(i.ToString() + name, 10));

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
                        cellABStyle.Font.Name = fontName; //设置字体

                        int charLimit = IsChineseText(name) ? 8 : 16; // 计算字符上限：如果是中文名称，则得到8；否则得到16
                        cellABStyle.Font.Size = (float)((!name.Contains('\n') ? 160 : 90)
                            * (1 - (name.Length - charLimit) * 0.04).Clamp(0.5, 1)); //设置字体大小：如果单元格文字不含换行符，为160；否则为90。再乘以一个缩小字体的因子
                        cellABStyle.HorizontalAlignment = ExcelHorizontalAlignment.Center; //单元格内容水平居中对齐
                        cellABStyle.VerticalAlignment = ExcelVerticalAlignment.Center; //单元格内容垂直居中对齐
                        cellABStyle.ShrinkToFit = !name.Contains('\n') ? true : false; //缩小字体填充：如果单元格文字不含换行符，为true；否则为false
                        cellABStyle.WrapText = name.Contains('\n') ? true : false; //文字自动换行：如果单元格文字含换行符，为true，否则为false

                    }

                    // 设置页面为A4，横向，页边距为0.4cm
                    foreach (ExcelWorksheet excelWorksheet in targetExcelPackage.Workbook.Worksheets)
                    {
                        ExcelPrinterSettings printerSettings = excelWorksheet.PrinterSettings; //将当前Excel工作表打印设置赋值给工作表打印设置变量
                        printerSettings.PaperSize = ePaperSize.A4; // 纸张设置为A4
                        printerSettings.Orientation = eOrientation.Landscape; //方向为横向
                        printerSettings.HorizontalCentered = false; //表格水平居中为false
                        printerSettings.VerticalCentered = false; //表格垂直居中为false
                        printerSettings.TopMargin = (decimal)(0.4 / 2.54); // 边距0.4cm转inch
                        printerSettings.BottomMargin = (decimal)(0.4 / 2.54);
                        printerSettings.LeftMargin = (decimal)(0.4 / 2.54);
                        printerSettings.RightMargin = (decimal)(0.4 / 2.54);
                    }

                    // 保存目标工作簿
                    string targetFolderPath = appSettings.SavingFolderPath; // 获取目标文件夹路径
                    string targetFilePath = Path.Combine(targetFolderPath, $"Cards_{Path.GetFileNameWithoutExtension(filePaths[0])}.xlsx"); //获取目标Excel工作簿文件路径全名
                    targetExcelPackage.SaveAs(new FileInfo(targetFilePath)); //保存目标Excel工作簿
                    MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ImportTextIntoDocumentTable()
        {
            try
            {
                InputDialog inputDialog = new InputDialog(question: "Input the text to be imported", defaultAnswer: "", textboxHeight: 300, acceptReturn: true); //弹出对话框，输入功能选项
                if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则结束本过程
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

                //创建目标文件夹
                if (!Directory.Exists(targetFolderPath))
                {
                    Directory.CreateDirectory(targetFolderPath);
                }

                //获取目标结构化文档表文件路径全名（移除段落列表0号元素中不能作为文件名的字符，截取前40个字符，作为目标文件主名）
                string targetExcelFilePath = Path.Combine(targetFolderPath, $"{CleanFileAndFolderName(lstParagraphs[0], 40)}.xlsx");
                ProcessParagraphsIntoDocumentTable(lstParagraphs, targetExcelFilePath); //将段落列表内容导入目标结构化文档表

                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }

        private void MakeFolders()
        {
            try
            {
                List<string>? filePaths = SelectFiles(FileType.Excel, false, "Select the Excel File Containing Directory Tree Data"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                DataTable? dataTable = ReadExcelWorksheetIntoDataTable(filePaths[0], 1); //读取Excel工作簿的第1张工作表，赋值给DataTable变量

                if (dataTable == null) //如果DataTable为null，则抛出异常
                {
                    throw new Exception("No valid data found!");
                }

                // 创建目标文件夹路径
                string targetFolderPath = Path.Combine(appSettings.SavingFolderPath, $"Dir_{Path.GetFileNameWithoutExtension(filePaths[0])}"); //获取目标文件夹路径
                if (!Directory.Exists(targetFolderPath))
                {
                    Directory.CreateDirectory(targetFolderPath);
                }

                for (int i = 0; i < dataTable!.Rows.Count; i++) //遍历DataTable所有数据行
                {
                    string newPathStr = ""; //每下移一个数据行，新文件夹路径字符串变量清零
                    for (int j = 0; j < dataTable.Columns.Count; j++) //遍历所有数据列
                    {
                        dataTable.Rows[i][j] = CleanFileAndFolderName(Convert.ToString(dataTable.Rows[i][j])!, 40); //去除DataTable当前数据行当前数据列数据的文件夹名中不可用于文件夹名的字符，截取指定数量的字符
                        newPathStr = newPathStr + Convert.ToString(dataTable.Rows[i][j]); //每右移一个数据列，新文件夹路径字符串延长一级（包含当前文件夹名和所有上级文件夹名），赋值给新文件夹路径字符串变量
                        if (i >= 1 && newPathStr == "") //如果当前数据行索引号大于等于1（从第2个记录行起），且新文件夹路径字符串变量为空，则将DataTable当前数据行当前数据列的元素填充为上一行同数据列的文件夹名
                        {
                            dataTable.Rows[i][j] = Convert.ToString(dataTable.Rows[i - 1][j]);
                        }
                    }
                }

                // 创建各级文件夹路径
                for (int i = 0; i < dataTable.Rows.Count; i++) //遍历DataTable所有数据行
                {
                    string newPath = targetFolderPath; //将目标文件夹路径赋值给新文件夹路径
                    for (int j = 0; j < dataTable.Columns.Count; j++) //遍历DataTable所有数据列
                    {
                        if (dataTable.Rows[i][j] != null) //如果当前数据行当前数据列的数据不为空
                        {
                            newPath = Path.Combine(newPath, Convert.ToString(dataTable.Rows[i][j])!); //将现有新文件夹路径和当前数据行当前数据列的文件夹名合并，重新赋值给自身

                            if (!Directory.Exists(newPath)) //如果新建文件夹路径不存在，则建立此文件夹路径
                            {
                                Directory.CreateDirectory(newPath);
                            }
                        }
                    }
                }
                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
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

                //创建目标文件夹
                if (!Directory.Exists(targetFolderPath)) //如果目标文件夹路径不存在，则建立该文件夹路径
                {
                    Directory.CreateDirectory(targetFolderPath);
                }

                string targetWordFilePath = Path.Combine(targetFolderPath, $"{Path.GetFileNameWithoutExtension(filePaths[0])}.docx"); //获取目标Word文件路径全名
                await ProcessDocumentTableIntoWordAsync(filePaths[0], targetWordFilePath); //将结构化文档表导出为目标Word文档

                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }


        private void MakeFileList()
        {
            try
            {
                string initialDirectory = latestRecords.LatestFolderPath; //读取用户使用记录中保存的文件夹路径
                //重新赋值给初始文件夹路径变量：如果初始文件夹路径存在，则得到初始文件夹路径原值；否则得到C盘根目录
                initialDirectory = Directory.Exists(initialDirectory) ? initialDirectory : "C:" + Path.DirectorySeparatorChar;
                OpenFolderDialog openFolderDialog = new OpenFolderDialog() //定义文件夹选择对话框
                {
                    Multiselect = false,
                    Title = "Select the Directory",
                    RootDirectory = initialDirectory //根文件夹路径设为设置中保存的文件夹路径
                };
                if (openFolderDialog.ShowDialog() == false) //如果对话框返回值为false（点击Cancel），则结束本过程
                {
                    return;
                }
                string folderPath = openFolderDialog.FolderName; //将选择的文件夹路径赋值给第一级文件夹路径变量
                latestRecords.LatestFolderPath = folderPath; //将第一级文件夹路径赋值给用户使用记录
                recordsManager.SaveSettings(latestRecords); //保存用户使用记录


                int latestSubpathDepth = latestRecords.LatestSubpathDepth; //读取用户使用记录中保存的子路径深度
                InputDialog inputDialog = new InputDialog(question: "Input the depth(level) of subdirectories", defaultAnswer: latestSubpathDepth.ToString()); //弹出功能选择对话框
                if (inputDialog.ShowDialog() == false) //如果对话框返回false（点击了Cancel），则结束本过程
                {
                    return;
                }
                int subpathDepth = Convert.ToInt32(inputDialog.Answer); //获取对话框返回的子路径深度
                latestRecords.LatestSubpathDepth = subpathDepth; // 将子路径深度赋值给用户使用记录
                recordsManager.SaveSettings(latestRecords);


                DataTable dataTable = new DataTable(); //定义DataTable，赋值给DataTable变量

                dataTable.Columns.AddRange(new DataColumn[]
                    {new DataColumn("Index"), new DataColumn("Path"), new DataColumn("Subpath"),
                    new DataColumn("Item"), new DataColumn("Type"), new DataColumn("Date", typeof(DateTime)) }); //向DataTable添加列

                //计算指定级数文件夹路径的总分隔符数：将第一级文件夹路径中的路径分隔符计数加上子路径的路径分隔符计数（总分隔符计数等于文件夹路径的级数）
                //（逐一比较每个字符是否为路径分隔符，如果是则"c=>"Lambda表达式返回true，用Count方法计数true的数量，即得到当前字符串中一共包含多少个路径分隔符）
                int separatorsCount = folderPath.Count(c => c == Path.DirectorySeparatorChar) + subpathDepth;

                GetFolderFiles(folderPath, separatorsCount, dataTable); //获取文件夹内的文件和下级文件夹信息，并存入DataTable

                void GetFolderFiles(string folderPath, int separatorsCount, DataTable dataTable) // 定义方法，获取文件夹内的文件和下级文件夹信息
                {
                    //如果输入文件夹路径所包含的路径分隔符数大于指定总路径分隔符数（文件夹路径级数大于指定级数），则结束本过程
                    if (folderPath.Count(c => c == Path.DirectorySeparatorChar) > separatorsCount)
                    {
                        return;
                    }

                    DirectoryInfo directories = new DirectoryInfo(folderPath); //将第一级文件夹路径内的目录信息赋值给目录信息变量

                    FileInfo[] files = directories.GetFiles(); //将第一级文件夹路径的目录信息中的所有文件信息赋值给文件信息集合变量
                    foreach (FileInfo file in files) //遍历文件信息集合中的所有文件
                    {
                        FileAttributes attributes = File.GetAttributes(file.FullName); //获取当前文件的属性
                        if ((attributes & FileAttributes.Hidden) != FileAttributes.Hidden &&
                            (attributes & FileAttributes.Temporary) != FileAttributes.Temporary) //如果当前文件的属性不为隐藏也不为临时
                        {
                            DataRow dataRow = dataTable.NewRow(); //定义DataTable新数据行

                            //获取当前文件系统日期：如果文件创建时间小于最后修改时间，则得到创建日期；否则得到最后修改日期
                            DateTime fileSystemDate = file.CreationTime < file.LastWriteTime ? file.CreationTime.Date : file.LastWriteTime.Date;

                            dataRow["Path"] = file.FullName; //将当前文件路径全名赋值给DataTable的新数据行的"路径"列
                            dataRow["Item"] = Path.GetFileNameWithoutExtension(file.Name); ; //将当前文件主名赋值给DataTable的新数据行的"名称"列
                            dataRow["Type"] = file.Extension; //将当前文件扩展名赋值给DataTable的新数据行的"类型"列
                            dataRow["Date"] = fileSystemDate; //将当前文件系统日期赋值给DataTable的新数据行的"日期"列
                            dataTable.Rows.Add(dataRow); //向DataTable中添加新数据行
                        }
                    }

                    DirectoryInfo[] subdirectories = directories.GetDirectories(); //将第一级文件夹路径的目录信息中的所有子文件夹信息赋值给子文件夹信息集合变量
                    foreach (DirectoryInfo subdirectory in subdirectories) //遍历子文件夹信息集合中所有的子路径
                    {
                        FileAttributes attributes = File.GetAttributes(subdirectory.FullName); //获取当前子文件夹的属性
                        if ((attributes & FileAttributes.Hidden) != FileAttributes.Hidden &&
                            (attributes & FileAttributes.Temporary) != FileAttributes.Temporary) //如果当前子文件夹的属性不为隐藏也不为临时
                        {
                            DataRow dataRow = dataTable.NewRow();

                            //获取当前子文件夹系统日期：如果子文件夹创建时间小于最后修改时间，则得到创建日期；否则得到最后修改日期
                            DateTime subdirectorySystemDate = subdirectory.CreationTime < subdirectory.LastWriteTime ? subdirectory.CreationTime.Date : subdirectory.LastWriteTime.Date;

                            dataRow["Path"] = subdirectory.FullName; //将当前子文件夹路径赋值给DataTable的新数据行的"路径"列
                            dataRow["Item"] = subdirectory.Name; //将当前子文件夹名赋值给DataTable的新数据行的"名称"列
                            dataRow["Type"] = "Directory"; //将"文件夹"字符串赋值给DataTable的新数据行的"类型"列
                            dataRow["Date"] = subdirectorySystemDate; //将当前子文件夹系统日期赋值给DataTable的新数据行的"日期"列
                            dataTable.Rows.Add(dataRow);

                            GetFolderFiles(subdirectory.FullName, separatorsCount, dataTable); //递归调用自身过程，将当前子路径作为参数传入
                        }
                    }
                }

                if (dataTable.Rows.Count * dataTable.Columns.Count == 0) //如果DataTable的行数或列数有一个为0，则抛出异常
                {
                    throw new Exception("No valid files or directories found.");
                }

                using (ExcelPackage targetExcelPackage = new ExcelPackage()) //新建Excel包，赋值给目标Excel包变量
                {
                    ExcelWorksheet targetExcelWorksheet = targetExcelPackage.Workbook.Worksheets.Add("Sheet1"); //新建“文件列表”Excel工作表，赋值给目标工作表变量
                    targetExcelWorksheet.Cells["A1"].LoadFromDataTable(dataTable, true); //将DataTable数据导入目标工作表（true代表将表头赋给第一行）
                    int endRowIndex = targetExcelWorksheet.Dimension.End.Row; //获取目标Excel工作表最末行的行索引号
                    int dateColumnIndex = dataTable.Columns["Date"]!.Ordinal + 1; //获取目标Excel工作表日期列的索引号（工作表列索引号从1开始，DataTable从0开始）
                    //将Excel工作表最末列（时间列）的数据格式设为“年-月-日”
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
                    FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"List_{CleanFileAndFolderName(folderPath, 40)}.xlsx")); //获取目标Excel工作簿文件路径全名信息
                    targetExcelPackage.SaveAs(targetExcelFile); //保存目标Excel工作簿文件
                }

                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }


        private void MergeDocumentsAndTables()
        {
            try
            {
                List<string>? filePaths = SelectFiles(FileType.WordAndExcel, true, "选择Word文档或Excel工作簿"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }


                List<string> lstFullText = new List<string>(); //建立全文本列表

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
                                TrimCellsStrings(excelWorksheet); //删除当前Excel工作表内所有文本型单元格值的首尾空格
                                RemoveWorksheetEmptyRowsAndColumns(excelWorksheet); //删除当前Excel工作表内所有空白行和空白列
                                if (excelWorksheet.Dimension == null) //如果当前Excel工作表为空，则直接跳过当前循环并进入下一个循环
                                {
                                    continue;
                                }

                                lstFullText.Add($"{fileName}: {excelWorksheet.Name}"); //全文本列表中追加当前Excel文件主名和当前工作表名

                                for (int i = 1; i <= excelWorksheet.Dimension.End.Row; i++) // 遍历Excel工作表所有行
                                {
                                    StringBuilder tableRowStringBuilder = new StringBuilder(); // 定义表格行数据字符串构建器
                                    for (int j = 1; j <= excelWorksheet.Dimension.End.Column; j++) // 遍历Excel工作表所有列
                                    {
                                        tableRowStringBuilder.Append(excelWorksheet.Cells[i, j].Text); // 将当前单元格文字追加到字符串构建器中
                                        tableRowStringBuilder.Append('|'); //追加表格分隔符到字符串构建器中
                                    }
                                    lstFullText.Add(tableRowStringBuilder.ToString().TrimEnd('|')); //将字符串构建器中当前行数据转换成字符串，移除尾部的分隔符，并追加到全文本列表中
                                }
                                lstFullText.AddRange(new string[] { "(The End)", "" }); //当前Excel工作表的所有行遍历完后，到了工作表末尾，在全文本列表最后追加一个"(The End)"元素和一个空字符串元素
                            }
                        }
                    }

                    else if (fileExtension.ToLower().Contains("doc")) // 如果当前文件扩展名转换为小写后含有“doc”（Word文件，docx、docm）
                    {
                        using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read)) // 创建文件流，以打开Word文档，赋值给文件流变量
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
                                            StringBuilder rowTextStringBuilder = new StringBuilder(); // 定义表格行数据字符串构建器
                                            foreach (XWPFTableCell cell in row.GetTableCells()) // 遍历表格所有列
                                            {
                                                string cellText = string.Join(" ", cell.Paragraphs.Select(p => p.Text.Trim())); // 提取单元格内的所有段落文本并连接起来
                                                rowTextStringBuilder.Append(cellText); // 将当前单元格文字追加到字符串构建器中
                                                rowTextStringBuilder.Append('|'); // 追加表格分隔符到字符串构建器中
                                            }
                                            lstFullText.Add(rowTextStringBuilder.ToString().TrimEnd('|')); // 将字符串构建器中当前行数据转换成字符串，移除尾部的分隔符，并追加到全文本列表中
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

                //创建目标文件夹
                if (!Directory.Exists(targetFolderPath)) //如果目标文件夹路径不存在，则建立该文件夹路径
                {
                    Directory.CreateDirectory(targetFolderPath);
                }

                ////写入目标Word文档
                //string targetWordFilePath = Path.Combine(targetFolderPath, $"Mrg_{targetFileMainName}.docx"); //获取目标Word文件的路径全名

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
                string targetTxtFilePath = Path.Combine(targetFolderPath, $"Mrg_{targetFileMainName}.txt"); // 获取目标文本文件的路径全名

                using (StreamWriter writer = new StreamWriter(targetTxtFilePath, false, Encoding.UTF8)) // 创建文本写入器对象（新建或覆盖目标文件，编码为UTF-8），赋值给文本写入器对象
                {
                    foreach (string paragraphText in lstFullText) // 遍历全文本列表的所有元素
                    {
                        writer.WriteLine(paragraphText); // 将当前元素的段落文字写入文本文件中，并换行
                    }
                }

                // 写入目标PDF文档
                string targetPdfFilePath = Path.Combine(targetFolderPath, $"Mrg_{targetFileMainName}.pdf"); // 获取目标文本文件的路径全名

                using (PdfWriter writer = new PdfWriter(targetPdfFilePath)) // 创建PDF写入器对象，赋值给PDF写入器对象
                using (PdfDocument pdf = new PdfDocument(writer)) // 创建PDF文档对象，赋值给PDF文档对象
                using (Document document = new Document(pdf)) // 创建文档对象，赋值给文档对象
                {
                    PdfFont font = PdfFontFactory.CreateFont("STSong-Light", "UniGB-UCS2-H", PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED); // 创建pdf字体对象：中文宋体，Adobe-GB1符集UCS-2编码，水平书写，优先嵌入字体
                    // 遍历字符串列表，
                    foreach (string paragraphText in lstFullText)
                    {
                        Paragraph paragraph = new Paragraph(paragraphText).SetFont(font); // 为当前字符串创建一个段落，使用已定义的字体
                        document.Add(paragraph); // 将段落添加到文档中
                    }
                }

                MessageBox.Show("Operation Completed", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }

        private void ScreenStocks()
        {
            try
            {
                List<string>? filePaths = SelectFiles(FileType.Excel, false, "Select the Excel File Containing Stocks Data"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                string latestStockDataColumnNamesStr = latestRecords.LatestStockDataColumnNamesStr; //读取用户使用记录中保存的列名称字符串
                InputDialog inputDialog = new InputDialog(question: "Input the column name of Stock Symbol, Name, Sector, Price, PB, PE and Total Shares (separated by commas)", defaultAnswer: latestStockDataColumnNamesStr); //弹出对话框，输入列名称
                if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则结束本过程
                {
                    return;
                }
                string dataColumnNamesStr = inputDialog.Answer; //获取对话框返回的列名称字符串
                latestRecords.LatestStockDataColumnNamesStr = dataColumnNamesStr; // 将对话框返回的列名称字符串赋值给用户使用记录
                recordsManager.SaveSettings(latestRecords);

                //将列名称字符串拆分成数组，转换成列表，然后移除每个元素的首尾空白字符
                List<string> lstDataColumnNamesStr = dataColumnNamesStr.Split(',').ToList().ConvertAll(e => e.Trim());
                //将各指标的列名称赋值给数据列名变量
                string codeDataColumnName = lstDataColumnNamesStr[0];
                string nameDataColumnName = lstDataColumnNamesStr[1];
                string sectorDataColumnName = lstDataColumnNamesStr[2];
                string prDataColumnName = lstDataColumnNamesStr[3];
                string pbDataColumnName = lstDataColumnNamesStr[4];
                string peDataColumnName = lstDataColumnNamesStr[5];
                string totalSharesDataColumnName = lstDataColumnNamesStr[6];

                DataTable? dataTable = ReadExcelWorksheetIntoDataTable(filePaths[0], 1); //读取Excel工作簿的第1张工作表，赋值给DataTable变量
                if (dataTable == null) //如果DataTable变量为null，则抛出异常
                {
                    throw new Exception("No valid data found!");
                }

                List<string> lstDataColumnNames = new List<string>
                    { codeDataColumnName, nameDataColumnName, sectorDataColumnName, prDataColumnName, pbDataColumnName, peDataColumnName, totalSharesDataColumnName }; //将各指标的数据列名称赋值给数据列名列表

                for (int i = dataTable.Columns.Count - 1; i >= 0; i--) // 遍历DataTable的所有数据列
                {
                    // 如果当前数据列的列名不在需要保留的数据列名列表中，则删除该列
                    if (!lstDataColumnNames.Contains(dataTable.Columns[i].ColumnName))
                    {
                        dataTable.Columns.RemoveAt(i);
                    }
                }

                dataTable.Columns.Add("Prem%", typeof(double)); //在DataTable中增加“溢价百分比”数据列

                //计算每个股票的溢价百分比
                foreach (DataRow dataRow in dataTable.Rows) //遍历DataTable每个数据行
                {
                    double pb = -1, pe = -1; //PB、PE初始赋值为-1（默认为缺失、无效/或亏损状态）
                    pb = Val(dataRow[pbDataColumnName]).Clamp<double>(0, double.MaxValue); //将当前数据行的PB数据列数据转换成数值型（限定为不小于指定值），赋值给PB变量
                    pe = Val(dataRow[peDataColumnName]); //将当前数据行的PE数据列数据转换成数值型，赋值给PE变量
                    double peThreshold = pb / (Math.Log(pb) / 4.3006); //计算PE阈值
                    dataRow["Prem%"] = Math.Round((pe - peThreshold) / peThreshold * 100, 2);  //计算溢价百分比，保留2位小数，赋值给当前行的“溢价百分比”数据列
                }

                //筛选低估值股票
                DataTable targetDataTable = dataTable.AsEnumerable().Where(
                    dataRow =>
                    {
                        double pr = -1; //现价初始赋值为-1（默认为缺失、无效）
                        pr = Val(dataRow[prDataColumnName]); //将当前数据行的现价数据列数据转换成数值型，赋值给现价变量
                        double peRelativePercentage = Convert.ToDouble(dataRow["Prem%"]);  //将当前数据行的溢价百分比数据列数据赋值给溢价百分比变量
                        //筛选溢价百分比大于-100小于-10，现价大于等于10的记录（此时"dataRow =>"lambda表达式函数返回true）
                        //当PE超过PE阈值（估值过高）时，溢价百分比会大于0；当PE为负（业绩亏损）时，溢价百分比会小于-100；因此溢价百分比仅在-100~0之间时估值较合理（为留有余量，将溢价百分比限定在-100~-10之间）
                        return (peRelativePercentage > -100 && peRelativePercentage < -10) && pr >= 10;
                    }).CopyToDataTable();  //将筛选出的数据行复制到目标DataTable

                if (targetDataTable.Rows.Count * targetDataTable.Columns.Count == 0) //如果目标DataTable的数据行数或列数有一个为0，则抛出异常
                {
                    throw new Exception("No qualified stocks found.");
                }

                targetDataTable.DefaultView.Sort = sectorDataColumnName + " ASC," + "Prem% ASC"; //按行业升序、溢价百分比升序对数据排序
                targetDataTable = targetDataTable.DefaultView.ToTable(); //将排序后的目标DataTable重新赋值给自身

                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePaths[0]))) //打开股票数据Excel工作簿，赋值给Excel包变量
                {
                    while (excelPackage.Workbook.Worksheets.Count > 1) //当Excel工作簿中的工作表大于1张，则继续循环，删除最后一张
                    {
                        excelPackage.Workbook.Worksheets.Delete(excelPackage.Workbook.Worksheets.Count - 1);
                    }

                    ExcelWorksheet targetWorksheet = excelPackage.Workbook.Worksheets.Add($"Results{new Random().Next(1000, 10000)}"); //在Excel工作簿中添加一个筛选结果工作表，赋值给目标工作表变量
                    targetWorksheet.Cells["A1"].LoadFromDataTable(targetDataTable, true); //将目标DataTable的数据导入目标Excel工作表（true代表将表头赋给第一行，或使用“c => c.PrintHeaders = true”）

                    foreach (ExcelRangeBase cell in targetWorksheet.Cells[targetWorksheet.Dimension.Address]) //遍历目标Excel工作表已使用区域的所有单元格
                    {
                        //重新赋值给当前单元格：将单元格文本值转换成数值，如果成功则赋值给单元格数值变量，然后单元格将得到该数值；否则，得到单元格原值
                        cell.Value = double.TryParse(cell.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double cellNumVal) ? cellNumVal : cell.Value;
                    }

                    //将目标Excel工作表第2行至最末行所有列单元格的数值格式设为保留两位小数
                    targetWorksheet.Cells[2, 1, targetWorksheet.Dimension.End.Row, targetWorksheet.Dimension.End.Column].Style.Numberformat.Format = "0.00";

                    FormatExcelWorksheet(targetWorksheet, 1, 0); //设置目标Excel工作表格式
                    excelPackage.Save(); //保存目标Excel工作簿文件
                }


                // 分别计算市场平均市净率、市盈率（个股市净率、市盈率以总股本为权重的加权平均数）

                // 定义统计计算方法，计算指定列的平均值和标准差
                (double, double) CalculateStatistics(DataTable dataTable, string columnName)
                {
                    // 获取指定列的数据行，并将它们转换为double类型的列表
                    List<double> values = dataTable.AsEnumerable()
                                          .Select(row => Val(row[columnName]))
                                          .ToList();

                    double mean = values.Average(); // 计算均值
                    double variance = values.Sum(value => Math.Pow(value - mean, 2)) / (values.Count - 1);  // 计算方差（每个数值与均值之差的平方的平均）

                    double standardDeviation = Math.Sqrt(variance); // 计算标准差（方差的平方根）

                    return (mean, standardDeviation); // 将均值和标准差赋值给函数返回值元组

                }

                // 计算所有股票市净率、市盈率的算数均数和标准差
                (double pbMean, double pbStd) = CalculateStatistics(dataTable, pbDataColumnName);
                (double peMean, double peStd) = CalculateStatistics(dataTable, peDataColumnName);

                double stdThreshold = 0.6745; // 设定标准差阈值，以便利用正态分布筛选有效数据

                EnumerableRowCollection<DataRow> validRows = dataTable.AsEnumerable()
                    .Where(row =>
                    {
                        double pb = Val(row[pbDataColumnName]);
                        double pe = Val(row[peDataColumnName]);
                        return pb > 0 && pe > 0 &&
                            pb >= pbMean - stdThreshold * pbStd && pb <= pbMean + stdThreshold * pbStd &&
                            pe >= peMean - stdThreshold * peStd && pe <= peMean + stdThreshold * peStd;
                    }); // 筛选有效数据行，即指市净率、市盈率都大于0，且市净率与市盈率均值差在正态分布的样本数占比范围内，作为计算市场平均指标的样本股

                double weight = validRows.Sum(row => Val(row[totalSharesDataColumnName])); // 计算样本股的总股本

                double marketPB = (validRows.Sum(row => Val(row[pbDataColumnName]) * Val(row[totalSharesDataColumnName])) / weight); // 计算市场平均PB（以个股总股本为权重，进行加权平均）

                double marketPE = validRows.Sum(row => Val(row[peDataColumnName]) * Val(row[totalSharesDataColumnName])) / weight; // 计算市场平均PE

                double marketPEThreshold = marketPB / (Math.Log(marketPB) / 4.3006); // 计算市场平均PE阈值
                double marketPremiumRate = (marketPE - marketPEThreshold) / marketPEThreshold; // 计算市场溢价率

                string marketIndicators = $"Market Average PB: {marketPB.ToString("0.00", CultureInfo.InvariantCulture)}\n" +
                    $"Market Average PE: {marketPE.ToString("0.00", CultureInfo.InvariantCulture)}\n" +
                    $"Market Average PE Threshold: {marketPEThreshold.ToString("0.00", CultureInfo.InvariantCulture)}\n" +
                    $"Market Premium Rate：{marketPremiumRate.ToString("P2", CultureInfo.InvariantCulture)}"; // 生成市场平均指标字符串
                
                MessageBox.Show(marketIndicators, "Result", MessageBoxButton.OK, MessageBoxImage.Information);

                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }

        public void SplitExcelWorksheet()
        {
            try
            {
                InputDialog inputDialog;

                List<string> lstFunctions = new List<string> { "0-Cancel", "1-Split into Workbooks", "2-Split into Worksheets" };

                string latestSplitWorksheetOption = latestRecords.LatestSplitWorksheetOption; //读取用户使用记录中保存的拆分Excel工作表功能选项字符串
                inputDialog = new InputDialog(question: "Select the function", options: lstFunctions, defaultAnswer: latestSplitWorksheetOption); //弹出功能选择对话框
                if (inputDialog.ShowDialog() == false) //如果对话框返回false（点击了Cancel），则结束本过程
                {
                    return;
                }
                string splitWorksheetOption = inputDialog.Answer; // 获取对话框返回的拆分Excel工作表功能选项字符串
                latestRecords.LatestSplitWorksheetOption = splitWorksheetOption; //将对话框返回的拆分Excel工作表功能选项字符串赋值给用户使用记录
                recordsManager.SaveSettings(latestRecords);

                int functionNum = lstFunctions.Contains(splitWorksheetOption) ? lstFunctions.IndexOf(splitWorksheetOption) : -1; //获取对话框返回的功能选项在功能列表中的索引号：如果功能列表包含功能选项，则得到对应的索引号；否则，得到-1

                if (functionNum < 1 || functionNum > 2) //如果功能选项不在设定范围，则结束本过程
                {
                    return;
                }

                List<string>? filePaths = SelectFiles(FileType.Excel, false, "Select the Excel File"); //获取所选文件列表
                if (filePaths == null) //如果文件列表为null，则结束本过程
                {
                    return;
                }

                GetHeaderAndFooterRowCount(out int headerRowCount, out int footerRowCount); //获取表头、表尾行数

                string? columnLetter = GetKeyColumnLetter(); //获取主键列符
                if (columnLetter == null) //如果主键列符为null，则结束本过程
                {
                    return;
                }

                inputDialog = new InputDialog(question: "Input the filename of target workbooks", defaultAnswer: Path.GetFileNameWithoutExtension(filePaths[0])); //弹出对话框，输入拆分后Excel工作簿文件主名
                if (inputDialog.ShowDialog() == false) //如果对话框返回为false（点击了Cancel），则结束本过程
                {
                    return;
                }
                string targetFileMainName = CleanFileAndFolderName(inputDialog.Answer, 40); //获取对话框返回的目标Excel工作簿文件主名

                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePaths[0]))) // 打开Excel工作簿，赋值给Excel包变量
                {
                    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.First(); // 将第一张Excel工作表赋值给Excel工作表变量

                    TrimCellsStrings(excelWorksheet); //删除Excel工作表内所有文本型单元格值的首尾空格
                    RemoveWorksheetEmptyRowsAndColumns(excelWorksheet); //删除Excel工作表内所有空白行和空白列
                    if ((excelWorksheet.Dimension?.Rows ?? 0) <= headerRowCount + footerRowCount) //如果当前Excel工作表已使用行数（如果工作表为空， 则为0）小于等于表头表尾行数和，则抛出异常
                    {
                        throw new Exception("No valid data found!");
                    }

                    Dictionary<string, List<ExcelRangeBase>> dataDict = new Dictionary<string, List<ExcelRangeBase>>(); // 定义一个字典来保存拆分的数据

                    for (int i = headerRowCount + 1; i <= excelWorksheet.Dimension!.End.Row - footerRowCount; i++) // 遍历Excel工作表除去表头、表尾的每一行
                    {
                        string key = excelWorksheet.Cells[columnLetter + i.ToString()].Text != "" ?
                            excelWorksheet.Cells[columnLetter + i.ToString()].Text : "-Blank-"; //将当前行拆分基准列的值赋值给键值变量：如果当前行单元格文字不为空，则得到得到单元格文字，否则得到"-Blank-"
                        if (dataDict.ContainsKey(key)) // 如果字典中已经有这个键，就将当前行添加到对应的列表中
                        {
                            dataDict[key].Add(excelWorksheet.Cells[i, 1, i, excelWorksheet.Dimension.End.Column]);
                        }
                        else // 否则，定义一个列表并向其中添加当前行，而后将列表并添加到字典中
                        {
                            dataDict[key] = new List<ExcelRangeBase> { excelWorksheet.Cells[i, 1, i, excelWorksheet.Dimension.End.Column] };
                        }
                    }

                    // 创建目标文件夹
                    string targetFolderPath = Path.Combine(appSettings.SavingFolderPath, $"Splt_{Path.GetFileNameWithoutExtension(filePaths[0])}");
                    if (!Directory.Exists(targetFolderPath))
                    {
                        Directory.CreateDirectory(targetFolderPath);
                    }

                    switch (functionNum) //根据功能序号进入相应的分支
                    {
                        case 1: //拆分为Excel工作簿
                            foreach (KeyValuePair<string, List<ExcelRangeBase>> pair in dataDict) // 遍历字典中的每一个键值对
                            {
                                using (ExcelPackage targetExcelPackage = new ExcelPackage()) //新建Excel包，赋值给目标Excel包变量
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
                                    FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"Splt_{targetFileMainName}_{pair.Key}.xlsx")); //获取目标Excel工作簿文件路径全名信息
                                    targetExcelPackage.SaveAs(targetExcelFile);
                                }
                            }
                            break;

                        case 2:  //拆分为Excel工作表
                            using (ExcelPackage targetExcelPackage = new ExcelPackage()) // 新建Excel包，赋值给目标Excel包变量
                            {
                                ExcelWorkbook targetExcelWorkbook = targetExcelPackage.Workbook; // 将Excel工作簿赋值给目标Excel工作簿变量
                                foreach (KeyValuePair<string, List<ExcelRangeBase>> pair in dataDict) //遍历所有字典数据
                                {
                                    // 新建Excel工作表，表名为键名去掉不能作为工作表名的字符并截取指定数量字符后的字符串，赋值给目标工作表变量
                                    ExcelWorksheet targetExcelWorksheet = targetExcelWorkbook.Worksheets.Add(CleanFileAndFolderName(pair.Key, 40));

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
                                FileInfo targetExcelFile = new FileInfo(Path.Combine(targetFolderPath, $"Splt_{targetFileMainName}.xlsx")); //获取目标Excel工作簿文件路径全名信息
                                targetExcelPackage.SaveAs(targetExcelFile);

                            }
                            break;
                    }
                }
                MessageBox.Show("Operation completed.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Top = 50.0;
            this.Left = SystemParameters.WorkArea.Width - this.Width - 150.0;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Environment.Exit(0);
        }

        private void MnuTest_Click(object sender, RoutedEventArgs e)
        {
            //InputDialog inputDialog = new InputDialog(question:"Number", defaultAnswer:"1000"); //弹出功能选择对话框
            //if (inputDialog.ShowDialog() == false) //如果对话框返回false（点击了Cancel），则结束本过程
            //{
            //    return;
            //}
            //int numbers = Convert.ToInt32(inputDialog.Answer); //获取对话框返回的功能选项
            //string result = ConvertArabicNumberIntoChinese(numbers);
            //MessageBox.Show("转换后的中文数字为：" + result);

            InputDialog inputDialog = new InputDialog(question: "Number", defaultAnswer: "1000"); //弹出功能选择对话框
            if (inputDialog.ShowDialog() == false) //如果对话框返回false（点击了Cancel），则结束本过程
            {
                return;
            }
            //获取对话框返回的功能选项
            double result = Val(inputDialog.Answer);
            MessageBox.Show("提取后的数字为：" + result.ToString());
        }


    }
}
