using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Drawing.Text;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using static COMIGHT.MainWindow;
using static COMIGHT.Methods;


namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for SettingDialog.xaml
    /// </summary>
    /// 

    public partial class SettingsDialog : Window
    {

        public SettingsDialog()
        {
            InitializeComponent();

            //this.Loaded += SettingsDialog_Loaded; // 窗口加载完成后，执行SettingsDialog_Loaded过程
            //this.Closing += SettingsDialog_Closing; // 窗口关闭前，执行SettingsDialog_Closing过程

        }

        private void SettingsDialog_Loaded(object sender, RoutedEventArgs e)
        {
            InstalledFontCollection installedFontCollention = new InstalledFontCollection();
            List<string> lstFontNames = installedFontCollention.Families.Select(f => f.Name).ToList(); //读取系统中已安装的字体，赋值给字体名称列表变量
            var listItemsSource = (ListItemsSource)this.Resources["ListItemsSource"]; // 将窗体资源中的ListItemsSource对象赋值给listItemsSource对象
            listItemsSource.FontList = lstFontNames; // 将字体名称列表赋值给listItemsSource对象中的字体列表属性

            this.DataContext = appSettings; // 将应用设置窗口的数据环境设为应用设置对象
        }

        private void SettingsDialog_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                CreateFolder(appSettings.SavingFolderPath); // 创建保存文件夹
            }

            catch (Exception ex)
            {
                ShowExceptionMessage(ex);
            }
        }


        private void BtnDialogOK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnSavingFolderSelector_Click(object sender, RoutedEventArgs e)
        {
            string? savingFolderPath = SelectFolder("Select the Saving Folder"); // 选择保存文件夹路径，赋值给保存文件夹路径变量：如果用户选择了文件夹路径并点击了OK，则返回选择的文件夹路径；否则，返回null

            txtbxSavingFolder.Text = savingFolderPath ?? txtbxSavingFolder.Text; // 将保存文件夹路径赋值给对应的文本框（如果保存文件夹路径变量为null，则得到文本框原值）
        }

        private void FillComboBoxes(DependencyObject root, IEnumerable<string> items, string? nameContains = null)
        {
            Stack<DependencyObject> stack = new Stack<DependencyObject>(); // 创建一个栈，用于存储依赖项对象
            stack.Push(root); // 将根依赖项对象压入栈中

            while (stack.Count > 0) // 当栈元素数量不为0，循环执行以下代码
            {
                DependencyObject current = stack.Pop(); // 弹出栈顶元素

                // 如果当前控件是ComboBox且：未指定名称必须包含的字符串，或控件名称包含了指定的字符串，则设置其数据源为相应枚举对象
                if (current is ComboBox comboBox && (nameContains == null || comboBox.Name.Contains(nameContains, StringComparison.CurrentCultureIgnoreCase))) 
                {
                    comboBox.ItemsSource = items;
                }

                if (current is FrameworkElement fe) // 如果是 FrameworkElement
                {
                    // 遍历FrameworkElement中所有DependencyObject元素
                    foreach (var child in LogicalTreeHelper.GetChildren(fe).OfType<DependencyObject>().Reverse())
                    {
                        stack.Push(child); // 将子元素压入栈中
                    }
                }
                else if (current is FrameworkContentElement fce) // 如果是 FrameworkContentElement
                {
                    // 遍历FrameworkContentElement中所有DependencyObject元素
                    foreach (var child in LogicalTreeHelper.GetChildren(fce).OfType<DependencyObject>().Reverse())
                    {
                        stack.Push(child); // 将子元素压入栈中
                    }
                }
            }
        }

        private void TextBoxLostFocus(object sender, RoutedEventArgs e)
        {
            TextBox? textBox = e.Source as TextBox; // 将事件源对象转换为TextBox类型
            if (textBox != null) // 如果事件源对象不为空
            {
                textBox.Text = textBox.Text.Trim(); // 去除文本框中的空格

                // 当文本框失去焦点时，绑定的数据源在Trim之前已经更新，因此Trim以后需要再次强制更新绑定源！！
                BindingExpression binding = textBox.GetBindingExpression(TextBox.TextProperty); // 获取TextBox控件文本属性的绑定表达式
                if (binding != null) // 如果绑定表达式不为空
                {
                    binding.UpdateSource(); // 强制触发绑定源更新
                }
            }
        }

    }

}
