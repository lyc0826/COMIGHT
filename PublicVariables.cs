using System.IO;
using System.Text.RegularExpressions;
using static COMIGHT.Properties.Settings;

namespace COMIGHT
{
    public static partial class PublicVariables
    {
        public static TaskManager taskManager = new TaskManager(); //定义任务管理器对象变量

        public static string appPath = AppDomain.CurrentDomain.BaseDirectory; //获取程序所在文件夹路径

        public static string websiteJsonFilePath = Path.Combine(appPath, "Websites.json"); //获取网址Json文件路径全名
        
        //public static string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //获取桌面文件夹路径
        public static string targetBaseFolderPath = Default.savingFolderPath; //获取目标基文件夹路径

        public static string manualUrl = @"https://github.com/lyc0826/COMIGHT/"; //定义用户手册网址

        public enum FileType { Excel, Word, WordAndExcel, Convertible, All } //定义文件类型枚举

        //定义中文句子正则表达式变量，匹配模式为：非“。；;”字符任意多个，“。；;”
        public static Regex regExCnSentence = new Regex(@"[^。；;]*[。；;]");

    }
}
