using System.Windows;
using System.Windows.Input;
using static COMIGHT.Methods;


namespace COMIGHT
{
    /// <summary>
    /// InputDialog.xaml 的交互逻辑
    /// </summary>
    public partial class InputDialog : Window
    {

        public InputDialog(string question, List<string>? options = null, string defaultAnswer = "", double textboxHeight = 30, bool acceptsReturn = false)
        {
            InitializeComponent();

            txtbxQuestion.Text = question; //将问题值赋值给问题文本块

            if (options != null) //如果选项列表不为null
            {
                txtbxAnswer.Visibility = Visibility.Collapsed; // 隐藏文本框

                cmbbxOptions.ItemsSource = options; // 将选项列表赋值给选项组合框
                // 设置选项组合框的选定项：如果选项组合框列表包含默认答案值，则选定与之相符的选项；否则，将选定项设为null
                cmbbxOptions.SelectedItem = cmbbxOptions.Items.Contains(defaultAnswer) ? defaultAnswer : null;
            }
            else //否则
            {
                cmbbxOptions.Visibility = Visibility.Collapsed; // 隐藏功能选项组合框

                txtbxAnswer.Text = defaultAnswer; //将默认答案值赋值给答案文本框
                txtbxAnswer.Height = textboxHeight; //将答案文本框的高度设为输入参数的高度，默认为30
                txtbxAnswer.AcceptsReturn = acceptsReturn; //将答案文本框是否接受回车键设为输入参数的值，默认为false
            }

        }

        private void btnDialogOk_Click(object sender, RoutedEventArgs e)
        {
            if (this.DialogResult.HasValue)
            {
                this.DialogResult = true; //对话框返回值设为false
            }
            this.Close();
        }

        private void btnDialogCancel_Click(object sender, RoutedEventArgs e)
        {
            if (this.DialogResult.HasValue)
            {
                this.DialogResult = false; //对话框返回值设为false
            }
            this.Close();
        }

        private void txtbxAnswer_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            //弹出对话框，如果返回true（点击了OK），则清除“输入文字”文本框
            if (ShowMessage("Do you want to clear the content?"))
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
            get
            {
                //如果选项组合框的选定项不为null，则将该项的值赋给函数返回值；否则，将答案文本框的文本去除首尾空白字符后赋给函数返回值
                return cmbbxOptions.SelectedItem != null ? cmbbxOptions.SelectedItem.ToString()! : txtbxAnswer.Text.Trim();
            }
        }

    }
}