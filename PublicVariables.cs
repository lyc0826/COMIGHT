using System.IO;

namespace COMIGHT
{
    // 定义公共变量类
    public partial class PublicVariables 
    {
        public static string appPath = AppDomain.CurrentDomain.BaseDirectory; //获取程序所在文件夹路径

        public static string websitesJsonFilePath = Path.Combine(appPath, "Websites.json"); //获取网址Json文件路径全名

        public static string settingsJsonFilePath = Path.Combine(appPath, "Settings.json"); //获取应用程序设置Json文件路径全名

        public static string recordsJsonFilePath = Path.Combine(appPath, "Records.json"); //获取用户使用记录Json文件路径全名

        //public static string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //获取桌面文件夹路径

        public enum FileType { Excel, Word, WordAndExcel, Convertible, All } //定义文件类型枚举

    }

    // 定义应用设置类
    public class AppSettings
    {
        public string SavingFolderPath { get; set; } = string.Empty;
        public string PandocPath { get; set; } = string.Empty;
        public string UserManualUrl { get; set; } = string.Empty;
        public string CnTitleFontName { get; set; } = string.Empty;
        public double CnTitleFontSize { get; set; }
        public string CnBodyFontName { get; set; } = string.Empty;
        public double CnBodyFontSize { get; set; }
        public string CnHeading0FontName { get; set; } = string.Empty;
        public double CnHeading0FontSize { get; set; }
        public string CnHeading1FontName { get; set; } = string.Empty;
        public double CnHeading1FontSize { get; set; }
        public string CnHeading2FontName { get; set; } = string.Empty;
        public double CnHeading2FontSize { get; set; }
        public string CnHeading3_4FontName { get; set; } = string.Empty;
        public double CnHeading3_4FontSize { get; set; }
        public double CnLineSpace { get; set; }
        public string EnTitleFontName { get; set; } = string.Empty;
        public double EnTitleFontSize { get; set; }
        public string EnBodyFontName { get; set; } = string.Empty;
        public double EnBodyFontSize { get; set; }
        public string EnHeading0FontName { get; set; } = string.Empty;
        public double EnHeading0FontSize { get; set; }
        public string EnHeading1FontName { get; set; } = string.Empty;
        public double EnHeading1FontSize { get; set; }
        public string EnHeading2FontName { get; set; } = string.Empty;
        public double EnHeading2FontSize { get; set; }
        public string EnHeading3_4FontName { get; set; } = string.Empty;
        public double EnHeading3_4FontSize { get; set; }
        public double EnLineSpace { get; set; }
        public string WorksheetFontName { get; set; } = string.Empty;
        public double WorksheetFontSize { get; set; }
        public string NameCardFontName { get; set; } = string.Empty;
    }

    //定义用户使用记录类
    public class LatestRecords
    {
        public string LatestFolderPath { get; set; } = string.Empty;
        public string LatestStockDataColumnNamesStr { get; set; } = string.Empty;
        public string LastestHeaderAndFooterRowCountStr { get; set; } = string.Empty;
        public string LatestKeyColumnLetter { get; set; } = string.Empty;
        public string LatestExcelWorksheetIndexesStr { get; set; } = string.Empty;
        public string LatestExcelWorksheetName { get; set; } = string.Empty;
        public string LatestOperatingRangeAddresses { get; set; } = string.Empty;
        public string LatestKeyDataColumnName { get; set; } = string.Empty;
        public int LatestSubpathDepth { get; set; }
        public string LatestNameCardFontName { get; set; } = string.Empty;
        public string LatestBatchProcessWorkbookOption { get; set; } = string.Empty;
        public string LatestSplitWorksheetOption { get; set; } = string.Empty;
        public string LatestUrl { get; set; } = string.Empty;
        public string LatestSubConverterBackEndUrl { get; set; } = string.Empty;
        public string LatestOriginalSubUrls { get; set; } = string.Empty;
        public string LatestExternalConfigUrl { get; set; } = string.Empty;
    }

}
