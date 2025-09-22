namespace COMIGHT
{
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
    public class LatestRecords
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
}