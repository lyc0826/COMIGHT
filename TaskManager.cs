using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Threading;


namespace COMIGHT
{
    public class TaskManager : INotifyPropertyChanged
    {
        private readonly DispatcherTimer timer = new DispatcherTimer(); // 定义定时器变量
        private readonly List<Task> lstTasks = new List<Task>(); // 定义任务列表变量
        private int activeTasksCount; // 定义当前激活任务数量变量

        public event PropertyChangedEventHandler? PropertyChanged; // 定义属性变更事件变量

        protected void OnPropertyChanged([CallerMemberName] string? name = null) // 定义属性变更事件处理方法
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name)); // 触发属性变更
        }

        private int ActiveTasksCount // 定义激活任务数量属性
        {
            get => activeTasksCount;
            set // 当前激活的任务数量发生变化时，更新属性并触发属性变更
            {
                if (activeTasksCount != value)
                {
                    activeTasksCount = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(StatusText));  // 触发StatusText属性变更
                }
            }
        }

        public string StatusText => ActiveTasksCount > 0 ? "Operation in progress..." : "Idle"; // 设置状态文本属性的值：如果激活任务数量大于0，显示"Operation in Progress..."；否则，显示"Idle"

        public TaskManager()
        {
            // 设置定时器以每秒检查一次任务状态
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += Timer_Tick!;
            timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            UpdateActiveTasksCount(); // 更新当前激活的任务数量
        }

        private void UpdateActiveTasksCount()
        {
            ActiveTasksCount = lstTasks.Count(t => !t.IsCompleted); // 设置当前激活任务数量属性的值，等于任务列表的元素数量
        }

        public async Task RunTaskAsync(Func<Task> taskFunc) // 定义一个异步方法，接收一个异步任务函数作为参数
        {
            var task = taskFunc();
            lstTasks.Add(task); // 将任务添加到任务列表

            try
            {
                await task; // 等待任务完成
            }
            finally
            {
                lstTasks.Remove(task); // 从任务列表中移除任务
                UpdateActiveTasksCount(); // 更新当前激活的任务数量
            }
        }
    }
}