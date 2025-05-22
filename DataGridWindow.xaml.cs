using System.Data; 
using System.Windows;
using static COMIGHT.Methods;
using GEmojiSharp;
using iText.IO.Source;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Interop;
using static COMIGHT.MainWindow;
using static COMIGHT.MSOfficeInterop;
using Application = System.Windows.Application;
using DataTable = System.Data.DataTable;
using ICell = NPOI.SS.UserModel.ICell;
using Task = System.Threading.Tasks.Task;
using Window = System.Windows.Window;
using Microsoft.Web.WebView2.Core;


namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class DataGridWindow : Window
    {
        private DataTable currentDataTable; // 使用私有字段存储当前DataGrid绑定的DataTable

        // 默认构造函数（可选，如果总是通过参数传入，可以移除）
        public DataGridWindow()
        {
            InitializeComponent();
            // 可以在这里加载一个默认的空 DataTable 或者不加载任何数据
            currentDataTable = new DataTable(); // 初始化一个空的DataTable
            myDataGrid.ItemsSource = currentDataTable.DefaultView;
        }

        // 新增的构造函数，允许传入 DataTable
        public DataGridWindow(string title, DataTable initialDataTable) : this() // 调用默认构造函数进行初始化
        {
            this.Title = title;
            if (initialDataTable != null)
            {
                SetDataTable(initialDataTable);
            }
        }

        // 新增的公共方法，用于外部设置 DataTable
        public void SetDataTable(DataTable dataTable)
        {
            if (dataTable != null)
            {
                currentDataTable = dataTable;
                myDataGrid.ItemsSource = currentDataTable.DefaultView;
            }
            else
            {
                // 如果传入null，可以清空DataGrid
                currentDataTable = new DataTable();
                myDataGrid.ItemsSource = currentDataTable.DefaultView;
            }
        }

        private void BtnExportData_Click(object sender, RoutedEventArgs e)
        {
            ExportData(); // 调用ExportData方法，导出数据
        }

        private void ExportData()
        {
            try
            {
                string targetFolderPath = appSettings.SavingFolderPath; //获取目标文件夹路径
                string targetExcelFile = Path.Combine(targetFolderPath!, $"{CleanFileAndFolderName(this.Title, 40)}.xlsx"); //获取目标Excel工作簿文件路径全名信息
                WriteDataTableIntoExcelWorkbook(new List<DataTable>() { currentDataTable }, targetExcelFile);

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }
    }
}

