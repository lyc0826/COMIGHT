using Newtonsoft.Json;
using System.IO;
using System.Windows;


namespace COMIGHT
{

    public class SettingsManager<T> where T : new() // 定义一个泛型类，泛型参数T必须有一个无参数的公共构造函数，允许在类内部通过 new T()创建T类型的新实例。
    {
        private readonly string _settingsFilePath;
        private T _settings  = new T(); // 定义设置对象变量

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
                _settings = JsonConvert.DeserializeObject<T>(json) ?? new T(); //读取Json文件内容并反序列化为设置对象（如果失败返回null，则得到默认初始化对象）
            }
            else // 否则
            {
                _settings = new T(); // 定义新设置对象
            }
            
        }

        // 获取设置
        public T GetSettings()
        {
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(_settings)) ?? new T(); //将内部设置对象序列化后再反序列化，形成深拷贝（与原对象无引用关系），赋值给外部调用者（如果序列化失败，则得到默认初始化对象）
        }

        // 保存设置
        public void SaveSettings(T newSettings)
        {
            //将当前设置对象和新设置对象序列化为JSON字符串
            string currentSettingsJson = JsonConvert.SerializeObject(_settings, Formatting.None);
            string newSettingsJson = JsonConvert.SerializeObject(newSettings, Formatting.None);

            // 如果当前设置和新设置序列化后的JSON字符串不同
            if (currentSettingsJson != newSettingsJson)
            {
                _settings = newSettings; // 将新设置对象赋值给内部设置对象
                string json = JsonConvert.SerializeObject(_settings, Formatting.Indented); // 序列化内部设置对象为JSON字符串
                File.WriteAllText(_settingsFilePath, json); // 将JSON字符串写入文件
            }
        }

    }
}

