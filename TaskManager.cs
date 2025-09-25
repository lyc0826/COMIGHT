using NPOI.SS.Formula.Functions;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Threading;


namespace COMIGHT
{
    
    public class TaskManager : ObservableObject
    {
        private readonly DispatcherTimer timer = new DispatcherTimer();
        private readonly List<Task> lstTasks = new List<Task>();
        private int _activeTasksCount; // 定义活动任务数私有字段
                                       
        private int ActiveTasksCount // 定义 ActiveTasksCount 属性
        {
            get => _activeTasksCount;
            set => SetPropertyAndNotify(ref _activeTasksCount, value, [nameof(StatusText)]); // 使用 SetPropertyAndNotify, 更新ActiveTasksCount属性值，并触发 StatusText 属性的变更通知，同时也触发ActiveTasksCount属性本身的变更通知
        }

        // 获取StatusText的值：当 ActiveTasksCount 大于 0 时，返回 "Operation in progress..."，否则返回 "Idle"
        public string StatusText => ActiveTasksCount > 0 ? "Task Running in Backgroud..." : "No Task Running in Backgroud.";
        //显示当前时间

        public TaskManager()
        {
            timer.Interval = TimeSpan.FromSeconds(1); // 设置定时器间隔为 1 秒
            timer.Tick += Timer_Tick!; // 注册定时器事件
            timer.Start(); // 启动定时器
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            UpdateActiveTasksCount(); // 更新当前激活的任务数量
        }

        private void UpdateActiveTasksCount()
        {
            ActiveTasksCount = lstTasks.Count(t => !t.IsCompleted); // 获取当前激活的任务数量，等于任务列表的元素数量
        }

        public async Task RunTaskAsync(Func<Task> taskFunc)
        {
            var task = taskFunc(); // 创建任务
            lstTasks.Add(task); // 将任务添加到任务列表中
            UpdateActiveTasksCount(); // 任务开始时立即更新一次，提供即时反馈
            try
            {
                await task;
            }
            finally
            {
                // 任务结束后，无论成功还是失败，都从列表中移除并再次更新计数
                lstTasks.Remove(task);
                UpdateActiveTasksCount();
            }
        }
    }

}