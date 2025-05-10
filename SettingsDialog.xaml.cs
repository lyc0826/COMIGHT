using System.Data;
using System.Drawing.Text;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
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

            this.DataContext = appSettings; // 将应用设置窗口的数据环境设为应用设置对象
        }

        private void OnLostFocus(object sender, RoutedEventArgs e)
        {
            TextBox? textBox = e.Source as TextBox;
            if (textBox != null)
            {
                textBox.Text = textBox.Text.Trim();

                BindingExpression binding = textBox.GetBindingExpression(TextBox.TextProperty); // 获取TextBox控件文本属性的绑定表达式
                if (binding != null) // 如果绑定表达式不为空
                {
                    binding.UpdateSource(); // 强制触发绑定源更新
                }
            }
        }

        private void btnDialogCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnDialogOK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnShowFonts_Click(object sender, RoutedEventArgs e)
        {
            InstalledFontCollection installedFontCollention = new InstalledFontCollection();
            List<string> lstFontNames = installedFontCollention.Families.Select(f => f.Name).ToList(); //读取系统中已安装的字体，赋值给字体名称列表变量
            ShowMessage(string.Join('\n', lstFontNames)); 
        }

    }
}
