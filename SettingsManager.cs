using Newtonsoft.Json;
using System.IO;


namespace COMIGHT
{

    public class SettingsManager<T> where T : new() // 定义一个泛型类，泛型参数T必须有一个无参数的公共构造函数，允许在类内部通过 new T()创建T类型的新实例。
    {
        private readonly string _settingsFilePath;
        private T _settings = default(T) ?? new T(); // 确保了即使 T 是一个引用类型，_settings 也不会是 null，而是至少包含一个默认初始化的对象

        public SettingsManager(string settingsFilePath)
        {
            _settingsFilePath = settingsFilePath;
            LoadSettings();
        }

        // 加载设置
        private void LoadSettings()
        {
            if (File.Exists(_settingsFilePath)) // 如果JSON文件存在
            {
                string json = File.ReadAllText(_settingsFilePath);
                _settings = JsonConvert.DeserializeObject<T>(json) ?? new T(); //读取文件内容并反序列化成对象：如果反序列化失败，则返回一个默认初始化的对象
            }
            else
            {
                _settings = new T(); // 否则，创建一个新对象
            }
        }

        // 获取设置
        public T GetSettings()
        {
            return _settings; // 将设置赋值给外部调用者
        }

        // 保存设置
        public void SaveSettings(T newSettings)
        {
            _settings = newSettings; // 将外部调用者传递的设置赋值给内部对象
            var json = JsonConvert.SerializeObject(_settings, Formatting.Indented); // 序列化对象为JSON字符串
            File.WriteAllText(_settingsFilePath, json); // 将JSON字符串写入文件
        }

    }
}

