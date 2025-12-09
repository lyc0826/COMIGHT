using Newtonsoft.Json;
using System.IO;

namespace COMIGHT
{
    public class SettingsManager<T> where T : new() // 定义设置管理器泛型类，泛型参数T必须有一个无参数的公共构造函数，允许在类内部通过 new T()创建T类型的新实例。
    {
        private readonly string settingsFilePath; // 定义设置文件路径变量
        private T settings = new T(); // 定义设置对象变量

        public SettingsManager(string settingsFilePath)
        {
            this.settingsFilePath = settingsFilePath; // 从方法参数中获取设置JSON文件路径
            LoadSettings(); // 从设置JSON文件中加载设置
        }

        // 从设置JSON文件中加载设置
        private void LoadSettings()
        {
            if (File.Exists(settingsFilePath)) // 如果JSON文件存在
            {
                string json = File.ReadAllText(settingsFilePath);
                settings = JsonConvert.DeserializeObject<T>(json) ?? new T(); //读取Json文件内容并反序列化为设置对象（如果失败返回null，则得到默认初始化对象）
            }
            else // 否则
            {
                settings = new T(); // 定义新设置对象
            }

        }

        // 读取设置并返回给调用者
        public T GetSettings()
        {
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(settings)) ?? new T(); //将内部设置对象序列化后再反序列化，形成深拷贝（与原对象无引用关系），赋值给外部调用者（如果序列化失败，则得到默认初始化对象）
        }

        // 保存新设置
        public void SaveSettings(T newSettings)
        {
            //将当前设置对象和新设置对象序列化为JSON字符串
            string currentSettingsJson = JsonConvert.SerializeObject(settings, Formatting.None);
            string newSettingsJson = JsonConvert.SerializeObject(newSettings, Formatting.None);

            // 如果当前设置和新设置序列化后的JSON字符串不同
            if (currentSettingsJson != newSettingsJson)
            {
                settings = JsonConvert.DeserializeObject<T>(newSettingsJson) ?? new T(); // 将新设置对象的JSON字符串反序列化，形成深拷贝（与原对象无引用关系），赋值给内部设置对象（如果反序列化失败，则得到默认初始化对象）
                string json = JsonConvert.SerializeObject(settings, Formatting.Indented); // 将内部设置对象序列化为JSON字符串
                File.WriteAllText(settingsFilePath, json); // 将JSON字符串写入文件
            }
        }

    }

    // 定义应用设置类，继承自 ObservableObject
    public class AppSettings : ObservableObject
    {
        // 为每个属性创建私有后备字段
        private string _savingFolderPath = string.Empty;
        private string _userManualUrl = string.Empty;
        private string _cnTitleFontName = string.Empty;
        private double _cnTitleFontSize;
        private string _cnBodyFontName = string.Empty;
        private double _cnBodyFontSize;
        private string _cnHeading0FontName = string.Empty;
        private double _cnHeading0FontSize;
        private string _cnHeading1FontName = string.Empty;
        private double _cnHeading1FontSize;
        private string _cnHeading2FontName = string.Empty;
        private double _cnHeading2FontSize;
        private string _cnHeading3_4FontName = string.Empty;
        private double _cnHeading3_4FontSize;
        private double _cnLineSpace;
        private string _worksheetFontName = string.Empty;
        private double _worksheetFontSize;
        private string _nameCardFontName = string.Empty;
        private bool _keepEmojisInMarkdown = false;
        private EnumUserProfile _userProfile = EnumUserProfile.Profile1;


        // 定义所有属性，如果属性变化，使用 SetProperty 来更新字段并触发通知
        public string SavingFolderPath
        {
            get => _savingFolderPath;
            set => SetProperty(ref _savingFolderPath, value);
        }

        public string UserManualUrl
        {
            get => _userManualUrl;
            set => SetProperty(ref _userManualUrl, value);
        }

        public string CnTitleFontName
        {
            get => _cnTitleFontName;
            set => SetProperty(ref _cnTitleFontName, value);
        }

        public double CnTitleFontSize
        {
            get => _cnTitleFontSize;
            set => SetProperty(ref _cnTitleFontSize, value);
        }

        public string CnBodyFontName
        {
            get => _cnBodyFontName;
            set => SetProperty(ref _cnBodyFontName, value);
        }

        public double CnBodyFontSize
        {
            get => _cnBodyFontSize;
            set => SetProperty(ref _cnBodyFontSize, value);
        }

        public string CnHeading0FontName
        {
            get => _cnHeading0FontName;
            set => SetProperty(ref _cnHeading0FontName, value);
        }

        public double CnHeading0FontSize
        {
            get => _cnHeading0FontSize;
            set => SetProperty(ref _cnHeading0FontSize, value);
        }

        public string CnHeading1FontName
        {
            get => _cnHeading1FontName;
            set => SetProperty(ref _cnHeading1FontName, value);
        }

        public double CnHeading1FontSize
        {
            get => _cnHeading1FontSize;
            set => SetProperty(ref _cnHeading1FontSize, value);
        }

        public string CnHeading2FontName
        {
            get => _cnHeading2FontName;
            set => SetProperty(ref _cnHeading2FontName, value);
        }

        public double CnHeading2FontSize
        {
            get => _cnHeading2FontSize;
            set => SetProperty(ref _cnHeading2FontSize, value);
        }

        public string CnHeading3_4FontName
        {
            get => _cnHeading3_4FontName;
            set => SetProperty(ref _cnHeading3_4FontName, value);
        }

        public double CnHeading3_4FontSize
        {
            get => _cnHeading3_4FontSize;
            set => SetProperty(ref _cnHeading3_4FontSize, value);
        }

        public double CnLineSpace
        {
            get => _cnLineSpace;
            set => SetProperty(ref _cnLineSpace, value);
        }

        public string WorksheetFontName
        {
            get => _worksheetFontName;
            set => SetProperty(ref _worksheetFontName, value);
        }

        public double WorksheetFontSize
        {
            get => _worksheetFontSize;
            set => SetProperty(ref _worksheetFontSize, value);
        }

        public string NameCardFontName
        {
            get => _nameCardFontName;
            set => SetProperty(ref _nameCardFontName, value);
        }

        public bool KeepEmojisInMarkdown
        {
            get => _keepEmojisInMarkdown;
            set => SetProperty(ref _keepEmojisInMarkdown, value);
        }

        public EnumUserProfile UserProfile 
        {
            get => _userProfile;
            set => SetProperty(ref _userProfile, value);
        }
    }

    // 定义用户配置枚举
    public enum EnumUserProfile
    {
        Profile1,
        Profile2,
        Profile3
    }

    //定义用户使用记录类
    public class UserRecords
    {
        public string LatestFolderPath { get; set; } = string.Empty;
        public string LastestHeaderAndFooterRowCountStr { get; set; } = string.Empty;
        public string LatestKeyColumnLetter { get; set; } = string.Empty;
        public string LatestExcelWorksheetIndexesStr { get; set; } = string.Empty;
        public string LatestOperatingRangeAddresses { get; set; } = string.Empty;
        public int LatestSubpathDepth { get; set; }
        public string LatestBatchProcessWorkbooksOption { get; set; } = string.Empty;
        public string LatestBatchDisassembleWorkbooksOption { get; set; } = string.Empty;
        public string LatestUrl { get; set; } = string.Empty;
    }

    public static class Settings
    {
        // 获取路径
        public static readonly string appPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Settings"); //获取程序所在文件夹路径
        //public static readonly string websitesJsonFilePath = Path.Combine(appPath, "Websites.json"); //获取网址Json文件路径全名
        public static readonly string settingsJsonFilePath = Path.Combine(appPath, "Settings.json"); //获取应用程序设置Json文件路径全名
        public static readonly string recordsJsonFilePath = Path.Combine(appPath, "Records.json"); //获取用户使用记录Json文件路径全名

        // 定义应用设置管理器、用户使用记录管理器对象，应用设置类、用户使用记录类对象，用于读取、保存应用设置和用户使用记录
        public static SettingsManager<AppSettings> settingsManager = new SettingsManager<AppSettings>(settingsJsonFilePath);
        public static SettingsManager<UserRecords> recordsManager = new SettingsManager<UserRecords>(recordsJsonFilePath);
        public static AppSettings appSettings = new AppSettings();
        public static UserRecords userRecords = new UserRecords();
    }

}