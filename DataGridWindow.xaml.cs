using System.IO;
using System.Windows;
using static COMIGHT.MainWindow;
using static COMIGHT.Methods;
using DataTable = System.Data.DataTable;
using Window = System.Windows.Window;


namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class DataGridWindow : Window
    {
        private readonly DataTable dataTable; // 定义DataTable变量
        private readonly string windowTitle = string.Empty; // 定义窗口标题变量

        // 默认构造函数
        public DataGridWindow()
        {
            InitializeComponent();
            dataTable = new DataTable(); // 初始化一个空的DataTable
            myDataGrid.ItemsSource = dataTable.DefaultView; //  将DataTable绑定到DataGrid
        }

        // 新增的构造函数，允许传入DataTable
        public DataGridWindow(string windowTitle, DataTable dataTable) : this() // 调用默认构造函数进行初始化
        {
            this.Title = this.windowTitle = windowTitle; //  设置窗口标题，并赋值给windowTitle字段
            this.dataTable = dataTable != null ? dataTable : new DataTable(); // 获取参数传入的DataTable，如果为null则创建一个空的DataTable
            myDataGrid.ItemsSource = this.dataTable.DefaultView; // 将DataTable绑定到DataGrid
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
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
                string targetExcelFile = Path.Combine(targetFolderPath!, $"{CleanFileAndFolderName(windowTitle)}.xlsx"); //获取目标Excel工作簿文件路径全名信息
                WriteDataTableIntoExcelWorkbook(new List<DataTable>() { dataTable }, targetExcelFile); //将DataTable数据写入Excel工作簿

                ShowSuccessMessage();
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }

    }
}

