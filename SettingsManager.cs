using Newtonsoft.Json;
using System.IO;


namespace COMIGHT
{

    public class SettingsManager<T> where T : new() // 定义一个泛型类，泛型参数T必须有一个无参数的公共构造函数，允许在类内部通过 new T()创建T类型的新实例。
    {
        private readonly string settingsFilePath; // 定义设置文件路径变量
        private T settings = new T(); // 定义设置对象变量

        public SettingsManager(string settingsFilePath)
        {
            this.settingsFilePath = settingsFilePath; // 从方法参数中获取设置JSON文件路径
            LoadSettings();
        }

        // 加载设置
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

        // 获取设置
        public T GetSettings()
        {
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(settings)) ?? new T(); //将内部设置对象序列化后再反序列化，形成深拷贝（与原对象无引用关系），赋值给外部调用者（如果序列化失败，则得到默认初始化对象）
        }

        // 保存设置
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
}

