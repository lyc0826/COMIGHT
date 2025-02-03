using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using static COMIGHT.Methods;


namespace COMIGHT
{
    internal class ExternalAppManager : INotifyPropertyChanged
    {
        private readonly string appPath; // 定义应用程序路径字段
        private CancellationTokenSource? cancellationTokenSource; // 定义取消令牌源字段，用于创建和管理取消令牌
        private string appName; // 定义应用程序名称字段
        private bool isAppRunning; // 定义“应用程序是否运行”字段

        public event PropertyChangedEventHandler? PropertyChanged; // 定义属性变更事件变量

        protected void OnPropertyChanged([CallerMemberName] string? name = null) // 定义属性变更事件处理方法
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name)); // 触发属性变更
        }

        private bool IsAppRunning // 定义“应用程序是否运行”私有属性
        {
            get => isAppRunning;
            set // 当前应用程序运行状态发生变化时，更新属性并触发属性变更
            {
                if (isAppRunning != value)
                {
                    isAppRunning = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(StatusText));  // 触发StatusText属性变更
                }
            }
        }

        public string StatusText => 
            $"{appName} is{(IsAppRunning ? "" : " not")} running"; // 设置状态文本属性的值：如果激活任务数量大于0，显示"Operation in Progress..."；否则，显示"Idle"


        public ExternalAppManager(string appPath) // 定义公共构造函数，接收应用程序路径作为参数
        {
            this.appPath = appPath; // 将参数传入的应用程序路径赋值给appPath字段

            appName = Path.GetFileNameWithoutExtension(appPath); // 将应用程序的文件主名赋值给appName字段
        }

        private bool CheckAppState() // 定义CheckAppState()方法，用于检查应用程序是否正在运行
        {
            return Process.GetProcessesByName(appName).Length > 0; // 检查应用程序进程数量是否大于0，如果大于0，则得到true；否则得到false；将结果赋值给函数返回值
        }

        private async Task MonitorApp(CancellationToken cancellationToken) // 定义MonitorApp异步方法，接收取消令牌作为参数
        {
            
            while (!cancellationToken.IsCancellationRequested) // 循环检查取消令牌的IsCancellationRequested属性，如果未请求取消，则继续循环
            {
                try
                {
                    
                    IsAppRunning = CheckAppState(); // 更新“应用程序是否运行”属性，等于CheckAppState()方法返回的结果
                    
                    if (!File.Exists(appPath) || CheckAppState()) // 如果应用程序不存在，或经IsAppRunning方法检查发现应用程序正在运行，则直接跳过进入下一个循环
                    {
                        continue;
                    }

                    ProcessStartInfo startInfo = new ProcessStartInfo // 创建ProcessStartInfo对象，用于配置进程启动信息，赋值给startInfo字段
                    {
                        FileName = appPath, // 设置要启动的应用程序的文件名
                        UseShellExecute = false, // 设置为false，表示不使用操作系统shell启动进程，允许更精细的控制
                        Verb = "runas", // 设置为"runas"，表示以管理员权限运行进程
                        CreateNoWindow = true, // 设置为true，表示不创建新窗口
                        WindowStyle = ProcessWindowStyle.Hidden //设置窗口样式为隐藏
                    }; 

                    Process.Start(startInfo); // 使用指定的启动信息启动进程
                    
                    await Task.Delay(5000, cancellationToken); // 异步等待 5 秒 (5000 毫秒)，并传入取消令牌，允许在等待期间响应取消请求
 
                }

                catch (TaskCanceledException) // 捕获TaskCanceledException异常，该异常在任务被取消时抛出
                {
                    break; // 如果任务被取消，则跳出循环
                }

                catch (Exception ex) // 捕获其他类型的异常
                {
                    Application.Current.Dispatcher.Invoke(() =>
                        ShowMessage($"Monitoring application failed: {ex.Message}") // 在UI线程上显示消息框，提示监控进程失败，并显示异常信息
                    );
                    await Task.Delay(10000, cancellationToken); // 暂停10秒，并传入取消令牌，允许在等待期间响应取消请求
                    break; // 退出循环
                }
            }
        }

         public void StartMonitoring() // 定义StartMonitoring方法，用于启动监控任务
        {
            cancellationTokenSource = new CancellationTokenSource(); // 创建一个取消令牌源，赋值给_cancellationTokenSource字段
            
            //StopApp(); // 调用StopApp方法，先停止应用程序，以便后期再以管理员权限重新启动
            
            Task.Run(() => MonitorApp(cancellationTokenSource.Token)); // 使用Task.Run启动一个新的后台任务，执行MonitorApp方法，并传递取消令牌
        }

        private void StopApp()
        {
            try
            {
                // 获取所有与指定应用程序名称匹配的进程
                Process[] processes = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(appPath));

                // 遍历每个进程
                foreach (Process process in processes)
                {
                    process.Kill(); // 强制终止进程

                    process.WaitForExit(5000); // 等待进程退出，最多等待 5 秒

                    // 检查进程是否已退出，如果没有退出，则抛出异常
                    if (!process.HasExited)
                    {
                        throw new Exception($"Stopping application failed: {process.ProcessName} (ID: {process.Id})");
                    }
                }
            }

            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() =>
                    ShowMessage($"Stopping application failed: {ex.Message}")
                );
            }
        }

        public void StopMonitoring() // 定义StopMonitoring方法，用于停止监控任务
        {
            cancellationTokenSource?.Cancel(); // 如果_cancellationTokenSource不为null，则调用其Cancel 方法，触发取消操作
            
            StopApp(); // 调用StopApp方法，停止应用程序
        }

        
    }
}

