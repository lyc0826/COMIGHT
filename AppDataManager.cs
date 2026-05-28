using Newtonsoft.Json;
using System.IO;

namespace COMIGHT
{
    public class AppDataAccessor<T> where T : new() // 定义应用程序数据存取器泛型类，泛型参数T必须有一个无参数的公共构造函数，允许在类内部通过 new T()创建T类型的新实例。
    {
        private readonly string _appDataFilePath; // 创建应用程序数据文件路径变量
        private T appData = new T(); // 创建应用程序数据对象变量

        public AppDataAccessor(string appDataFilePath)
        {
            this._appDataFilePath = appDataFilePath; // 从方法参数中获取应用程序数据JSON文件路径
            LoadData(); // 从应用程序数据JSON文件中加载数据
        }

        // 从应用程序数据JSON文件中加载数据
        private void LoadData()
        {
            if (File.Exists(_appDataFilePath)) // 如果JSON文件存在
            {
                string jsonStr = File.ReadAllText(_appDataFilePath);
                appData = JsonConvert.DeserializeObject<T>(jsonStr) ?? new T(); //读取Json文件内容并反序列化为应用程序数据对象（如果失败返回null，则得到默认初始化对象）
            }
            else // 否则
            {
                appData = new T(); // 创建新应用程序数据对象
            }

        }

        // 读取应用程序数据并返回给调用者
        public T GetData()
        {
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(appData)) ?? new T(); //将应用程序数据对象序列化后再反序列化，形成深拷贝（与原对象无引用关系），赋值给外部调用者（如果序列化失败，则得到默认初始化对象）
        }

        // 保存新应用程序数据
        public void SaveData(T newAppData)
        {
            //将当前应用程序数据对象和新对象序列化为JSON字符串
            string currentAppDataJsonStr = JsonConvert.SerializeObject(appData, Formatting.None);
            string newAppDataJsonStr = JsonConvert.SerializeObject(newAppData, Formatting.None);

            // 如果当前应用程序数据和新数据的JSON字符串不同
            if (currentAppDataJsonStr != newAppDataJsonStr)
            {
                appData = JsonConvert.DeserializeObject<T>(newAppDataJsonStr) ?? new T(); // 将新应用程序数据对象的JSON字符串反序列化，形成深拷贝（与原对象无引用关系），赋值给应用程序数据对象（如果反序列化失败，则得到默认初始化对象）
                string jsonStr = JsonConvert.SerializeObject(appData, Formatting.Indented); // 将应用程序数据对象序列化为JSON字符串
                File.WriteAllText(_appDataFilePath, jsonStr); // 将JSON字符串写入文件
            }
        }

    }

    // 定义文档版式选项枚举
    public enum EnumDocumentLayoutOption
    {
        Universal,
        Chinese_Official,
    }

    // 定义用户配置Profile枚举
    public enum EnumUserProfile
    {
        Profile1,
        Profile2,
        Profile3
    }

    // 定义应用设置类，继承自 ObservableObject
    public class AppSettings : ObservableObject
    {
        // 为每个属性创建私有后备字段
        private string _savingFolderPath = string.Empty;
        private string _userManualUrl = string.Empty;

        private EnumDocumentLayoutOption _documentLayoutOption = EnumDocumentLayoutOption.Universal;

        private string _udTitleFontName = string.Empty;
        private double _udTitleFontSize;
        private string _udBodyFontName = string.Empty;
        private double _udBodyFontSize;
        private string _udHeading0FontName = string.Empty;
        private double _udHeading0FontSize;
        private string _udHeading1FontName = string.Empty;
        private double _udHeading1FontSize;
        private string _udHeading2FontName = string.Empty;
        private double _udHeading2FontSize;
        private string _udHeading3_4FontName = string.Empty;
        private double _udHeading3_4FontSize;

        private string _codTitleFontName = string.Empty;
        private double _codTitleFontSize;
        private string _codBodyFontName = string.Empty;
        private double _codBodyFontSize;
        private string _codHeading0FontName = string.Empty;
        private double _codHeading0FontSize;
        private string _codHeading1FontName = string.Empty;
        private double _codHeading1FontSize;
        private string _codHeading2FontName = string.Empty;
        private double _codHeading2FontSize;
        private string _codHeading3_4FontName = string.Empty;
        private double _codHeading3_4FontSize;

        private double _lineSpace;

        private string _worksheetFontName = string.Empty;
        private double _worksheetFontSize;
        private string _placeCardFontName = string.Empty;
        private bool _keepEmojisInMarkupText = false;
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

        public EnumDocumentLayoutOption DocumentLayoutOption
        {
            get => _documentLayoutOption;
            set => SetProperty(ref _documentLayoutOption, value);
        }

        public string UDTitleFontName
        {
            get => _udTitleFontName;
            set => SetProperty(ref _udTitleFontName, value);
        }

        public double UDTitleFontSize
        {
            get => _udTitleFontSize;
            set => SetProperty(ref _udTitleFontSize, value);
        }

        public string UDBodyFontName
        {
            get => _udBodyFontName;
            set => SetProperty(ref _udBodyFontName, value);
        }

        public double UDBodyFontSize
        {
            get => _udBodyFontSize;
            set => SetProperty(ref _udBodyFontSize, value);
        }

        public string UDHeading0FontName
        {
            get => _udHeading0FontName;
            set => SetProperty(ref _udHeading0FontName, value);
        }

        public double UDHeading0FontSize {
            get => _udHeading0FontSize;
            set => SetProperty(ref _udHeading0FontSize, value);
        }

        public string UDHeading1FontName
            {
            get => _udHeading1FontName;
            set => SetProperty(ref _udHeading1FontName, value);
        }

        public double UDHeading1FontSize
        {
            get => _udHeading1FontSize;
            set => SetProperty(ref _udHeading1FontSize, value);
        }

        public string UDHeading2FontName
        {
            get => _udHeading2FontName;
            set => SetProperty(ref _udHeading2FontName, value);
        }

        public double UDHeading2FontSize
        {
            get => _udHeading2FontSize;
            set => SetProperty(ref _udHeading2FontSize, value);
        }

        public string UDHeading3_4FontName
        {
            get => _udHeading3_4FontName;
            set => SetProperty(ref _udHeading3_4FontName, value);
        }

        public double UDHeading3_4FontSize
        {
            get => _udHeading3_4FontSize;
            set => SetProperty(ref _udHeading3_4FontSize, value);
        }

        public string CODTitleFontName
        {
            get => _codTitleFontName;
            set => SetProperty(ref _codTitleFontName, value);
        }

        public double CODTitleFontSize
        {
            get => _codTitleFontSize;
            set => SetProperty(ref _codTitleFontSize, value);
        }

        public string CODBodyFontName
        {
            get => _codBodyFontName;
            set => SetProperty(ref _codBodyFontName, value);
        }

        public double CODBodyFontSize
        {
            get => _codBodyFontSize;
            set => SetProperty(ref _codBodyFontSize, value);
        }

        public string CODHeading0FontName
        {
            get => _codHeading0FontName;
            set => SetProperty(ref _codHeading0FontName, value);
        }

        public double CODHeading0FontSize
        {
            get => _codHeading0FontSize;
            set => SetProperty(ref _codHeading0FontSize, value);
        }

        public string CODHeading1FontName
        {
            get => _codHeading1FontName;
            set => SetProperty(ref _codHeading1FontName, value);
        }

        public double CODHeading1FontSize
        {
            get => _codHeading1FontSize;
            set => SetProperty(ref _codHeading1FontSize, value);
        }

        public string CODHeading2FontName
        {
            get => _codHeading2FontName;
            set => SetProperty(ref _codHeading2FontName, value);
        }

        public double CODHeading2FontSize
        {
            get => _codHeading2FontSize;
            set => SetProperty(ref _codHeading2FontSize, value);
        }

        public string CODHeading3_4FontName
        {
            get => _codHeading3_4FontName;
            set => SetProperty(ref _codHeading3_4FontName, value);
        }

        public double CODHeading3_4FontSize
        {
            get => _codHeading3_4FontSize;
            set => SetProperty(ref _codHeading3_4FontSize, value);
        }

        public double LineSpace
        {
            get => _lineSpace;
            set => SetProperty(ref _lineSpace, value);
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

        public string PlaceCardFontName
        {
            get => _placeCardFontName;
            set => SetProperty(ref _placeCardFontName, value);
        }

        public bool KeepEmojisInMarkupText
        {
            get => _keepEmojisInMarkupText;
            set => SetProperty(ref _keepEmojisInMarkupText, value);
        }

        public EnumUserProfile UserProfile 
        {
            get => _userProfile;
            set => SetProperty(ref _userProfile, value);
        }
    }

    

    // 定义标记文本类型枚举
    public enum EnumMarkupType { Markdown, HTML };


    // 定义用户使用记录类，继承自 ObservableObject
    public class UserRecords : ObservableObject
    {
        // 为每个属性创建私有后备字段，默认值与原类保持一致
        private string _folderPath = string.Empty;
        private string _headerAndFooterRowCountStr = string.Empty;
        private string _keyColumnLetter = string.Empty;
        private string _worksheetIndexesStr = string.Empty;
        private string _operatingRanges = string.Empty;
        private int _subpathDepth;
        private string _processWorkbooksOption = string.Empty;
        private string _disassembleWorkbookOption = string.Empty;
        private EnumMarkupType _markupType = EnumMarkupType.Markdown;
        private string _OpenedWebsite = string.Empty;

        // 定义所有属性，如果属性变化，使用 SetProperty 来更新字段并触发通知（与AppSettings格式统一）
        public string FolderPath
        {
            get => _folderPath;
            set => SetProperty(ref _folderPath, value);
        }

        public string HeaderAndFooterRowCountStr
        {
            get => _headerAndFooterRowCountStr;
            set => SetProperty(ref _headerAndFooterRowCountStr, value);
        }

        public string KeyColumnLetter
        {
            get => _keyColumnLetter;
            set => SetProperty(ref _keyColumnLetter, value);
        }

        public string WorksheetIndexesStr
        {
            get => _worksheetIndexesStr;
            set => SetProperty(ref _worksheetIndexesStr, value);
        }

        public string OperatingRanges
        {
            get => _operatingRanges;
            set => SetProperty(ref _operatingRanges, value);
        }

        public int SubpathDepth
        {
            get => _subpathDepth;
            set => SetProperty(ref _subpathDepth, value);
        }

        public string ProcessWorkbooksOption
        {
            get => _processWorkbooksOption;
            set => SetProperty(ref _processWorkbooksOption, value);
        }

        public string DisassembleWorkbookOption
        {
            get => _disassembleWorkbookOption;
            set => SetProperty(ref _disassembleWorkbookOption, value);
        }

        public EnumMarkupType MarkupType
        {
            get => _markupType;
            set => SetProperty(ref _markupType, value);
        }

        public string OpenedWebsite
        {
            get => _OpenedWebsite;
            set => SetProperty(ref _OpenedWebsite, value);
        }
    }

    // 定义网址标签类
    public class WebsiteTag
    {
        public string Label { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
    }

    // 定义网址数据类（网址标签类的集合）
    public class WebsiteData
    {
        public List<WebsiteTag> WebsiteTags { get; set; } = new List<WebsiteTag>();

        public WebsiteData() { } // 公共无参构造函数，满足 AppDataAccessor<T> 的 new() 约束
    }


    // 定义总应用数据类（顶层容器），继承自 ObservableObject，
    public class AppData : ObservableObject
    {

        // 创建私有后备字段，对应三个属性
        private AppSettings _appSettings = new AppSettings();
        private UserRecords _userRecords = new UserRecords();
        private WebsiteData _websiteData = new WebsiteData();
        
        // 定义三个属性
        public AppSettings AppSettings
        {
            get => _appSettings;
            set => SetProperty(ref _appSettings, value);
        }

        public UserRecords UserRecords
        {
            get => _userRecords;
            set => SetProperty(ref _userRecords, value);
        }

        public WebsiteData WebsiteData
        {
            get => _websiteData;
            set => SetProperty(ref _websiteData, value);
        }
    }

    // 定义应用程序数据管理器类
    public static class AppDataManager
    {
        
        public static readonly string appDataDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AppData"); //获取程序所在文件夹路径
        public static readonly string appDataFilePath = Path.Combine(appDataDir, "AppData.json"); //获取应用程序数据Json文件路径

        // 创建应用程序数据访问器对象，用于读取和保存Json文件
        public static AppDataAccessor<AppData> appDataAccessor = new AppDataAccessor<AppData>(appDataFilePath);
        // 创建应用程序数据对象，用于存储所有数据
        public static AppData appData = new AppData();

        static AppDataManager() // 静态构造函数，在类被首次引用时自动执行
        {
            appData = appDataAccessor.GetData(); // 从应用程序数据存取器中读取所有数据，赋值给应用程序数据对象变量
        }
        
        // 创建应用设置、用户使用记录、网址数据三个对象，分别指向 appData 对象的三个属性
        public static AppSettings appSettings => appData.AppSettings;
        public static UserRecords userRecords => appData.UserRecords;
        public static WebsiteData websiteData => appData.WebsiteData;
        
    }

}