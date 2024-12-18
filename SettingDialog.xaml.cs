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

        // 定义设置DataTable和DataSet
        private DataTable generalSettingsTable = new DataTable("generalSettingsTable");
        private DataTable cnSettingsTable = new DataTable("cnSettingsTable");
        private DataTable enSettingsTable = new DataTable("enSettingsTable");
        private DataSet settingsDataSet = new DataSet();

        // 定义设置记录类型
        private record Setting(string TableName, string TableItem, string PropertiesSettingItem);

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
                new Setting ("enSettingsTable", "English Heading Lv1 Font Name", "enHeading1FontName"),
                new Setting ("enSettingsTable", "English Heading Lv1 Font Size", "enHeading1FontSize"),
                new Setting ("enSettingsTable", "English Heading Lv2 Font Name", "enHeading2FontName"),
                new Setting ("enSettingsTable", "English Heading Lv2 Font Size", "enHeading2FontSize"),
                new Setting ("enSettingsTable", "English Heading Lv3-4 Font Name", "enHeading3_4FontName"),
                new Setting ("enSettingsTable", "English Heading Lv3-4 Font Size", "enHeading3_4FontSize"),
                new Setting ("enSettingsTable", "English Line Space", "enLineSpace")
            };

        private void LoadSettings()
        {
            try
            {
                // 定义设置DataTable数组
                DataTable[] dataTables = new DataTable[] { generalSettingsTable, cnSettingsTable, enSettingsTable };
                settingsDataSet.Tables.AddRange(dataTables); // 将设置DataTable数组添加到设置DataSet
                settingsDataSet.AcceptChanges();
                // 遍历设置DataTable数组，为每个DataTable添加列，并设置主键便于快速查找
                foreach (DataTable dataTable in dataTables)
                {
                    //dataTable.Reset();
                    //DataColumn[] columnsToAdd = new[]
                    //{
                    //    new DataColumn("Item", typeof(string)),
                    //    new DataColumn("Value", typeof(object))
                    //};

                    //// 遍历要添加的列，并检查是否存在；如果不存在则添加
                    //foreach (DataColumn column in columnsToAdd)
                    //{
                    //    if (!dataTable.Columns.Contains(column.ColumnName))
                    //    {
                    //        dataTable.Columns.Add(column);
                    //    }
                    //}

                    //if (!dataTable.PrimaryKey.Any()) // 如果主键不存在，则添加主键
                    //{
                    //    dataTable.PrimaryKey = new[] { dataTable.Columns["Item"]! };
                    //}

                    dataTable.Columns.AddRange(new DataColumn[]
                    {
                        new DataColumn("Item", typeof(string)),
                        new DataColumn("Value", typeof(object)),
                    });
                    dataTable.PrimaryKey = new[] { dataTable.Columns["Item"]! };
                }

                // 将Properties.Settings中设置项的值填入对应的设置DataTable中
                foreach (var setting in lstSettings) // 遍历设置记录列表
                {
                    PropertyInfo propertyInfo = Default.GetType().GetProperty(setting.PropertiesSettingItem)!; // 获取Properties.Settings中当前设置项的属性
                    object? value = propertyInfo.GetValue(Default); // 获取对应属性的值
                    if (value != null) settingsDataSet.Tables[setting.TableName]!.Rows.Add(setting.TableItem, value); // 如果值不为null，则向当前设置DataTable中添加DataTable设置项和值
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void btnDialogSave_Click(object sender, RoutedEventArgs e)
        {
            SaveSettings();
        }

        private T GetSettingValue<T>(DataTable datatable, string dataTableKeyField, string datatableValueField)
        {
            return (T)Convert.ChangeType(datatable.Rows.Find(dataTableKeyField)?[datatableValueField] ?? "", typeof(T));
        }

        private void SaveSettings()
        {
            try
            {
                // 遍历设置记录列表，从设置DataTable中读取设置值并保存到Properties.Settings中
                foreach (var setting in lstSettings)
                {

                    DataRow? row = settingsDataSet.Tables[setting.TableName]!.Rows.Find(setting.TableItem); // 查找当前DataTable设置项对应的数据行
                    if (row != null && row["Value"] != DBNull.Value) // 如果数据行存在且值不为数据库空值
                    {
                        var propertyInfo = Default.GetType().GetProperty(setting.PropertiesSettingItem); // 获取Properties.Settings中当前设置项的属性
                        if (propertyInfo != null && propertyInfo.CanWrite) // 如果属性不为null且可写入
                        {
                            object valueToSet = Convert.ChangeType(row["Value"], propertyInfo.PropertyType); // 将数据行值转换为属性类型
                            propertyInfo.SetValue(Default, valueToSet); // 将值设置到属性中

                        }
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
