
using System.Windows;


namespace COMIGHT
{
    /// <summary>
    /// InputDialog.xaml 的交互逻辑
    /// </summary>
    public partial class InputDialog : Window
    {

        public InputDialog(string question, string defaultAnswer = "")
        {
            InitializeComponent();
            txtblkQuestion.Text = question; //将问题值赋值给问题文本块
            txtbxAnswer.Text = defaultAnswer; //将默认答案值赋值给答案文本框 
        }

        private void btnDialogOk_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true; //对话框返回值设为true
        }

        private void btnDialogCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false; //对话框返回值设为false
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            txtbxAnswer.SelectAll(); //全选答案文本框文字
            txtbxAnswer.Focus(); //答案文本框获取焦点
        }

        public string Answer
        {
            get { return txtbxAnswer.Text.Trim(); } //移除答案文本框的文字的首尾空白字符，赋值给答案属性
        }

    }
}