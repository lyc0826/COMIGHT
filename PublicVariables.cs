using COMIGHT.Properties;
using System.IO;
using System.Text.RegularExpressions;
using static COMIGHT.Properties.Settings;

namespace COMIGHT
{
    public static partial class PublicVariables
    {
        public static string dataBaseFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database.xlsx"); //获取数据库Excel工作簿文件路径全名

        //public static string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //获取桌面文件夹路径
        //public static string targetBaseFolderPath = Path.Combine(desktopPath, "COMIGHT Files"); //获取目标基文件夹路径
        public static string targetBaseFolderPath = Default.savingFolderPath; //获取目标基文件夹路径

        public static string manualUrl = @"https://github.com/lyc0826/COMIGHT/"; //定义用户手册网址

        public enum FileType { Excel, Word, Convertible, All } //定义文件类型枚举


        //定义中文小标题文字正则表达式变量，匹配模式为：从开头开始，非“。；;分页符换行符回车符”的字符2-40个；后方出现：“。”换行符回车符或结尾标记
        //public static Regex regExCnHeadingText = new Regex(@"^[^。；;\f\n\r]{2,40}(?=。|\n|\r|$)", RegexOptions.Multiline);

        //定义中文句子正则表达式变量，匹配模式为：非“。；;”字符任意多个，“。；;”
        public static Regex regExCnSentence = new Regex(@"[^。；;]*[。；;]");

    }
}
