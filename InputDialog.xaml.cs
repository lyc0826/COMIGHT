
using System.Windows;
using System.Windows.Input;


namespace COMIGHT
{
    /// <summary>
    /// InputDialog.xaml 的交互逻辑
    /// </summary>
    public partial class InputDialog : Window
    {

        public InputDialog(string question, List<string>? options = null, string defaultAnswer = "", double textboxHeight = 30, bool acceptReturn = false )
        {
            InitializeComponent();
            txtbxQuestion.Text = question; //将问题值赋值给问题文本块
            if (options != null) //如果选项列表不为null
            { 
                cmbbxOptions.ItemsSource = options; // 将选项列表赋值给选项组合框
            }
            else
            {
                cmbbxOptions.Visibility = Visibility.Collapsed; // 隐藏选项组合框
            }
            txtbxAnswer.Text = defaultAnswer; //将默认答案值赋值给答案文本框
            txtbxAnswer.Height = textboxHeight; //将答案文本框的高度设为输入的高度，默认为30
            txtbxAnswer.AcceptsReturn = acceptReturn; //将答案文本框是否接受回车键设为输入的值，默认为false
        }

        private void btnDialogOk_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true; //对话框返回值设为true
        }

        private void btnDialogCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false; //对话框返回值设为false
        }

        private void cmbbxOptions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            txtbxAnswer.Text = cmbbxOptions.SelectedItem.ToString(); // 将选项值赋值给答案文本框
        }

        private void txtbxAnswer_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            //弹出对话框，如果返回true（点击了OK），则清除“输入文字”文本框
            if (MessageBox.Show("Do you want to clear the content?", "Inquiry", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                txtbxAnswer.Text = "";
            }
        }

        private void InputDialog_ContentRendered(object sender, EventArgs e)
        {
            txtbxAnswer.SelectAll(); //全选答案文本框文字
            txtbxAnswer.Focus(); //答案文本框获取焦点
        }

        public string Answer
        {
            get { return txtbxAnswer.Text.Trim(); } //移除答案文本框的文字的首尾空白字符，赋值给答案属性
        }

        public int SelectedIndex
        {
            get { return cmbbxOptions.SelectedIndex; } //获取选项组合框的选中项索引，赋值给选中索引属性
        }
    }
}