//using System.Configuration;
//using System.Data;
//using System.Windows;

//namespace COMIGHT
//{
//    /// <summary>
//    /// Interaction logic for App.xaml
//    /// </summary>
//    public partial class App : Application
//    {

//    }

//}


using System.Windows;

// 定义应用程序命名空间
namespace COMIGHT
{
    // 定义 App 类，继承自 WPF 的 Application 类，使用 partial 关键字支持代码分离
    public partial class App : Application
    {
        // 声明静态私有字段，用于存储互斥量对象，初始化为 null
        private static Mutex? _mutex = null;
        // 声明常量字符串，作为互斥量的唯一标识符，需要在系统范围内唯一
        private const string MutexName = "COMIGHT";

        // 重写 Application 类的 OnStartup 方法，在应用程序启动时调用
        protected override void OnStartup(StartupEventArgs e)
        {
            // 声明布尔变量，用于接收互斥量是否为新创建的结果
            bool createdNew;
            // 创建命名互斥量对象
            // 参数1: true - 表示当前线程希望立即拥有互斥量
            // 参数2: MutexName - 互斥量的系统级唯一名称
            // 参数3: out createdNew - 输出参数，true表示创建了新互斥量，false表示互斥量已存在
            _mutex = new Mutex(true, MutexName, out createdNew);

            // 如果 createdNew 为 false，说明互斥量已经存在，即已有应用程序实例在运行
            if (!createdNew)
            {
                // 显示消息框提示用户应用程序已经在运行

                MessageBox.Show("Application already running.", "Warning",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                // 调用方法尝试激活已存在的应用程序窗口
                ActivateExistingWindow();

                // 关闭当前应用程序实例
                Current.Shutdown();
                // 退出当前方法，不执行后续的启动流程
                return;
            }

            // 如果是第一个实例，调用基类的 OnStartup 方法继续正常的应用程序启动流程
            base.OnStartup(e);
        }

        // 重写 Application 类的 OnExit 方法，在应用程序退出时调用
        protected override void OnExit(ExitEventArgs e)
        {
            // 使用空条件运算符，如果 _mutex 不为 null，则释放互斥量的所有权
            _mutex?.ReleaseMutex();
            // 使用空条件运算符，如果 _mutex 不为 null，则释放互斥量占用的系统资源
            _mutex?.Dispose();
            // 调用基类的 OnExit 方法，执行正常的应用程序退出流程
            base.OnExit(e);
        }

        // 定义私有方法，用于查找并激活已存在的应用程序窗口
        private void ActivateExistingWindow()
        {
            // 获取当前进程对象
            var current = System.Diagnostics.Process.GetCurrentProcess();
            // 根据进程名称获取所有同名进程的数组
            var processes = System.Diagnostics.Process.GetProcessesByName(current.ProcessName);

            // 遍历所有同名进程
            foreach (var process in processes)
            {
                // 检查两个条件：
                // 1. 进程ID不等于当前进程ID（排除自身）
                // 2. 进程有主窗口句柄（IntPtr.Zero 表示空句柄）
                if (process.Id != current.Id && process.MainWindowHandle != IntPtr.Zero)
                {
                    // 调用 Windows API 恢复窗口显示状态（如果窗口被最小化）
                    // 参数1: 窗口句柄
                    // 参数2: 显示命令（SW_RESTORE = 9，表示恢复窗口）
                    ShowWindow(process.MainWindowHandle, SW_RESTORE);
                    // 调用 Windows API 将指定窗口设置为前台窗口（激活窗口）
                    SetForegroundWindow(process.MainWindowHandle);
                    // 找到第一个符合条件的进程后退出循环
                    break;
                }
            }
        }

        // 使用 P/Invoke 导入 Windows API 函数 ShowWindow
        // 该函数用于控制窗口的显示状态
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // 使用 P/Invoke 导入 Windows API 函数 SetForegroundWindow
        // 该函数用于将指定窗口带到前台并激活它
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        // 定义常量，表示恢复窗口的命令值
        // SW_RESTORE = 9：激活并显示窗口，如果窗口被最小化或最大化，则恢复到原来的尺寸和位置
        private const int SW_RESTORE = 9;
    }
}
