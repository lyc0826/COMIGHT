using System.Windows.Controls;
using System.Windows.Input;

namespace COMIGHT
{
    public static class TextBoxCommands
    {
        // 定义静态命令 "Clear"
        public static readonly RoutedUICommand Clear = new RoutedUICommand("Clear", "Clear", typeof(TextBoxCommands));

        // 静态构造函数：在程序启动时，告诉所有 TextBox 如何处理这个命令
        static TextBoxCommands()
        {
            // 为 TextBox 类型注册命令绑定
            CommandManager.RegisterClassCommandBinding(typeof(TextBox), new CommandBinding(Clear, OnClearExecuted, OnClearCanExecute));
            CommandManager.RegisterClassInputBinding(typeof(TextBox), new InputBinding(Clear, new KeyGesture(Key.D, ModifierKeys.Control)));
        }

        // 执行逻辑：清空
        private static void OnClearExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                textBox.Clear();
            }
        }

        // 判断逻辑：只有当 TextBox 有内容且未被禁用时，菜单项才可用
        private static void OnClearCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                e.CanExecute = textBox.IsEnabled && !textBox.IsReadOnly;
            }
        }
    }
}

