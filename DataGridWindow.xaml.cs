using System.Data; // 引入 System.Data 命名空间
using System.Windows;


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

         //------------- 以下是原有的事件处理方法，稍作修改以使用 currentDataTable -------------

        private void ClearData_Click(object sender, RoutedEventArgs e)
        {
            currentDataTable.Clear(); // 使用_currentDataTable
        }
    }
}

