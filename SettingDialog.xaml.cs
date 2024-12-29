using System.Data;
using System.Reflection;
using System.Windows;
using static COMIGHT.Properties.Settings;


namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for SettingDialog.xaml
    /// </summary>
    public partial class SettingDialog : Window
    {

        public SettingDialog()
        {
            InitializeComponent();

            LoadSettings();

            // 将设置DataTable绑定到设置DataGrid上
            dtgrdGeneralSettings.ItemsSource = generalSettingsTable!.DefaultView;
            dtgrdCnDocumentSettings.ItemsSource = cnSettingsTable!.DefaultView;
            dtgrdEnDocumentSettings.ItemsSource = enSettingsTable!.DefaultView;
        }

        private void btnDialogSave_Click(object sender, RoutedEventArgs e)
        {
            SaveSettings();
        }

        // 定义设置DataTable和DataSet
        private DataTable generalSettingsTable = new DataTable("generalSettingsTable");
        private DataTable cnSettingsTable = new DataTable("cnSettingsTable");
        private DataTable enSettingsTable = new DataTable("enSettingsTable");
        private DataSet dataSet = new DataSet();

        private class Setting // 定义设置类
        {
            public string DataTableName { get; } // 定义DataTable名称属性
            public string DataTableItem { get; } // 定义DataTable项属性
            public string PropertiesSettingItem { get; } // 定义设置属性

            public Setting(string dataTableName, string dataTableItem, string propertiesSettingItem)
            {
                DataTableName = dataTableName;
                DataTableItem = dataTableItem;
                PropertiesSettingItem = propertiesSettingItem;
            }
        }

        // 定义设置记录列表
        private List<Setting> lstSettings = new List<Setting>
            {
                new Setting ("generalSettingsTable", "Saving Folder Path", "savingFolderPath"),
                new Setting ("generalSettingsTable", "Pandoc Application Path", "pandocPath"),

                new Setting ("cnSettingsTable", "Chinese Title Font Name", "cnTitleFontName"),
                new Setting ("cnSettingsTable", "Chinese Title Font Size", "cnTitleFontSize"),
                new Setting ("cnSettingsTable", "Chinese Body Text Font Name", "cnBodyFontName"),
                new Setting ("cnSettingsTable", "Chinese Body Text Font Size", "cnBodyFontSize"),
                new Setting ("cnSettingsTable", "Chinese Heading Lv0 Font Name", "cnHeading0FontName"),
                new Setting ("cnSettingsTable", "Chinese Heading Lv0 Font Size", "cnHeading0FontSize"),
                new Setting ("cnSettingsTable", "Chinese Heading Lv1 Font Name", "cnHeading1FontName"),
                new Setting ("cnSettingsTable", "Chinese Heading Lv1 Font Size", "cnHeading1FontSize"),
                new Setting ("cnSettingsTable", "Chinese Heading Lv2 Font Name", "cnHeading2FontName"),
                new Setting ("cnSettingsTable", "Chinese Heading Lv2 Font Size", "cnHeading2FontSize"),
                new Setting ("cnSettingsTable", "Chinese Heading Lv3-4 Font Name", "cnHeading3_4FontName"),
                new Setting ("cnSettingsTable", "Chinese Heading Lv3-4 Font Size", "cnHeading3_4FontSize"),
                new Setting ("cnSettingsTable", "Chinese Line Space", "cnLineSpace"),

                new Setting ("enSettingsTable", "English Title Font Name", "enTitleFontName"),
                new Setting ("enSettingsTable", "English Title Font Size", "enTitleFontSize"),
                new Setting ("enSettingsTable", "English Body Text Font Name", "enBodyFontName"),
                new Setting ("enSettingsTable", "English Body Text Font Size", "enBodyFontSize"),
                new Setting ("enSettingsTable", "English Heading Lv0 Font Name", "enHeading0FontName"),
                new Setting ("enSettingsTable", "English Heading Lv0 Font Size", "enHeading0FontSize"),
                new Setting ("enSettingsTable", "English Heading Lv1 Font Name", "enHeading1FontName"),
                new Setting ("enSettingsTable", "English Heading Lv1 Font Size", "enHeading1FontSize"),
                new Setting ("enSettingsTable", "English Heading Lv2 Font Name", "enHeading2FontName"),
                new Setting ("enSettingsTable", "English Heading Lv2 Font Size", "enHeading2FontSize"),
                new Setting ("enSettingsTable", "English Heading Lv3-4 Font Name", "enHeading3_4FontName"),
                new Setting ("enSettingsTable", "English Heading Lv3-4 Font Size", "enHeading3_4FontSize"),
                new Setting ("enSettingsTable", "English Line Space", "enLineSpace")
            };

        private object GetDataTableValue(DataTable dataTable, string dataTableKeyField, string datatableValueField, string key, Type targetType)
        {
            if (dataTable.PrimaryKey == null ||
                (!dataTable.PrimaryKey.Any(pk => pk == dataTable.Columns[dataTableKeyField]))) // 如果主键不存在，或者没有包含指定数据列字段，则将指定数据列字段添加进主键
            {
                dataTable.PrimaryKey = new[] { dataTable.Columns[dataTableKeyField]! };
            }

            return Convert.ChangeType(dataTable.Rows.Find(key)?[datatableValueField] ?? "", targetType); // 将指定数据行指定数据列的值转换为指定类型，并赋值给函数返回值
        }

        private void LoadSettings()
        {
            try
            {
                // 定义设置DataTable数组，并添加到设置DataSet
                DataTable[] dataTables = new DataTable[] { generalSettingsTable, cnSettingsTable, enSettingsTable };
                dataSet.Tables.AddRange(dataTables); // 将设置DataTable数组添加到设置DataSet
                dataSet.AcceptChanges();

                // 遍历设置DataTable数组，为每个DataTable添加列
                foreach (DataTable dataTable in dataTables)
                {
                    dataTable.Columns.AddRange(new DataColumn[]
                    {
                        new DataColumn("Item", typeof(string)),
                        new DataColumn("Value", typeof(object)),
                    });
                }

                // 将Properties.Settings中设置项的值填入对应的设置DataTable中
                foreach (var setting in lstSettings) // 遍历设置记录列表
                {
                    PropertyInfo propertyInfo = Default.GetType().GetProperty(setting.PropertiesSettingItem)!; // 获取Properties.Settings中当前设置项的属性
                    object? value = propertyInfo.GetValue(Default); // 获取当前属性的值
                    if (value != null)
                    {
                        DataRow newDataRow = dataSet.Tables[setting.DataTableName]!.NewRow(); // 创建一个新数据行
                        newDataRow["Item"] = setting.DataTableItem; // 设置新数据行"Item"、"Value"数据列的值
                        newDataRow["Value"] = value;
                        dataSet.Tables[setting.DataTableName]!.Rows.Add(newDataRow); // 向当前设置DataTable添加新数据行
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void SaveSettings()
        {
            try
            {
                // 遍历设置记录列表，从设置DataTable中读取设置值并保存到Properties.Settings中
                foreach (var setting in lstSettings)
                {

                    PropertyInfo propertyInfo = Default.GetType().GetProperty(setting.PropertiesSettingItem)!; // 获取Properties.Settings中当前设置项的属性
                    if (propertyInfo != null && propertyInfo.CanWrite) // 如果属性不为null且可写入
                    {
                        object valueToSet = GetDataTableValue(dataSet.Tables[setting.DataTableName]!, "Item", "Value", setting.DataTableItem, propertyInfo.PropertyType); // 将数据行值转换为属性类型
                        propertyInfo.SetValue(Default, valueToSet); // 将值设置到属性中
                    }
                }

                Default.Save();

                MessageBox.Show("Settings saved successfully.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            this.Close();
        }

    }
}
