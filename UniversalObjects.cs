using System.IO;

namespace COMIGHT
{
    public static class UniversalObjects
    {
        // 获取路径
        public static readonly string appPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Settings"); //获取程序所在文件夹路径
        //public static readonly string websitesJsonFilePath = Path.Combine(appPath, "Websites.json"); //获取网址Json文件路径全名
        public static readonly string settingsJsonFilePath = Path.Combine(appPath, "Settings.json"); //获取应用程序设置Json文件路径全名
        public static readonly string recordsJsonFilePath = Path.Combine(appPath, "Records.json"); //获取用户使用记录Json文件路径全名

        // 定义应用设置管理器、用户使用记录管理器对象，应用设置类、用户使用记录类对象，用于读取、保存应用设置和用户使用记录
        public static SettingsManager<AppSettings> settingsManager = new SettingsManager<AppSettings>(settingsJsonFilePath);
        public static SettingsManager<LatestRecords> recordsManager = new SettingsManager<LatestRecords>(recordsJsonFilePath);
        public static AppSettings appSettings = new AppSettings();
        public static LatestRecords latestRecords = new LatestRecords();

        public static TaskManager taskManager = new TaskManager(); //定义任务管理器对象变量，用于执行异步任务，并提供任务执行状态数据

    }
}
