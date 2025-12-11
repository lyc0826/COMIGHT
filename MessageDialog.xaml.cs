using System.Windows;

namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for MessageDialog.xaml
    /// </summary>
    public partial class MessageDialog : Window
    {
        public MessageDialog(string message)
        {
            InitializeComponent();

            txtbxMessage.Text = message; //将问题值赋值给问题文本块
        }

        private void btnDialogOk_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsModal())
            {
                this.DialogResult = true; //对话框返回值设为true
            }
            this.Close();
        }

        private void btnDialogCancel_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsModal())
            {
                this.DialogResult = false; //对话框返回值设为false
            }
            this.Close();
        }

    }
}
