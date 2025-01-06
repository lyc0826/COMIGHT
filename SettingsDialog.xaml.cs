using System.Data;
using System.Reflection;
using System.Windows;
using static COMIGHT.MainWindow;


namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for SettingDialog.xaml
    /// </summary>
    public partial class SettingsDialog : Window
    {

        // 定义设置DataTable和DataSet
        private DataTable generalSettingsTable = new DataTable("generalSettingsTable");
        private DataTable cnSettingsTable = new DataTable("cnSettingsTable");
        private DataTable enSettingsTable = new DataTable("enSettingsTable");
        private DataSet dataSet = new DataSet();

        private class SettingsRelation // 定义设置对应关系类
        {
            public string DataTableName { get; } // 定义设置DataTable名称属性
            public string DataTableItem { get; } // 定义设置DataTable设置项属性
            public string AppSettingItem { get; } // 定义应用设置项属性

            public SettingsRelation(string dataTableName, string dataTableItem, string appSettingItem)
            {
                DataTableName = dataTableName;
                DataTableItem = dataTableItem;
                AppSettingItem = appSettingItem;
            }
        }


        private List<SettingsRelation> lstSettingsRelations = new List<SettingsRelation> // 定义设置对应关系列表
            {
                new SettingsRelation ("generalSettingsTable", "Saving Folder Path", "SavingFolderPath"),
                new SettingsRelation ("generalSettingsTable", "Pandoc App Path", "PandocPath"),
                new SettingsRelation ("generalSettingsTable", "SubConverter App Path", "SubConverterPath"),

                new SettingsRelation ("cnSettingsTable", "Chinese Title Font Name", "CnTitleFontName"),
                new SettingsRelation ("cnSettingsTable", "Chinese Title Font Size", "CnTitleFontSize"),
                new SettingsRelation ("cnSettingsTable", "Chinese Body Text Font Name", "CnBodyFontName"),
                new SettingsRelation ("cnSettingsTable", "Chinese Body Text Font Size", "CnBodyFontSize"),
                new SettingsRelation ("cnSettingsTable", "Chinese Heading Lv0 Font Name", "CnHeading0FontName"),
                new SettingsRelation ("cnSettingsTable", "Chinese Heading Lv0 Font Size", "CnHeading0FontSize"),
                new SettingsRelation ("cnSettingsTable", "Chinese Heading Lv1 Font Name", "CnHeading1FontName"),
                new SettingsRelation ("cnSettingsTable", "Chinese Heading Lv1 Font Size", "CnHeading1FontSize"),
                new SettingsRelation ("cnSettingsTable", "Chinese Heading Lv2 Font Name", "CnHeading2FontName"),
                new SettingsRelation ("cnSettingsTable", "Chinese Heading Lv2 Font Size", "CnHeading2FontSize"),
                new SettingsRelation ("cnSettingsTable", "Chinese Heading Lv3-4 Font Name", "CnHeading3_4FontName"),
                new SettingsRelation ("cnSettingsTable", "Chinese Heading Lv3-4 Font Size", "CnHeading3_4FontSize"),
                new SettingsRelation ("cnSettingsTable", "Chinese Line Space", "CnLineSpace"),

                new SettingsRelation ("enSettingsTable", "English Title Font Name", "EnTitleFontName"),
                new SettingsRelation ("enSettingsTable", "English Title Font Size", "EnTitleFontSize"),
                new SettingsRelation ("enSettingsTable", "English Body Text Font Name", "EnBodyFontName"),
                new SettingsRelation ("enSettingsTable", "English Body Text Font Size", "EnBodyFontSize"),
                new SettingsRelation ("enSettingsTable", "English Heading Lv0 Font Name", "EnHeading0FontName"),
                new SettingsRelation ("enSettingsTable", "English Heading Lv0 Font Size", "EnHeading0FontSize"),
                new SettingsRelation ("enSettingsTable", "English Heading Lv1 Font Name", "EnHeading1FontName"),
                new SettingsRelation ("enSettingsTable", "English Heading Lv1 Font Size", "EnHeading1FontSize"),
                new SettingsRelation ("enSettingsTable", "English Heading Lv2 Font Name", "EnHeading2FontName"),
                new SettingsRelation ("enSettingsTable", "English Heading Lv2 Font Size", "EnHeading2FontSize"),
                new SettingsRelation ("enSettingsTable", "English Heading Lv3-4 Font Name", "EnHeading3_4FontName"),
                new SettingsRelation ("enSettingsTable", "English Heading Lv3-4 Font Size", "EnHeading3_4FontSize"),
                new SettingsRelation ("enSettingsTable", "English Line Space", "EnLineSpace")
            };


        public SettingsDialog()
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
                appSettings = settingsManager.GetSettings();
                // 定义设置DataTable数组，并添加到设置DataSet
                DataTable[] dataTables = new DataTable[] { generalSettingsTable, cnSettingsTable, enSettingsTable };
                dataSet.Tables.AddRange(dataTables); // 将设置DataTable数组添加到设置DataSet
                dataSet.AcceptChanges();

                // 遍历设置DataTable，为每个DataTable添加列
                foreach (DataTable dataTable in dataTables)
                {
                    dataTable.Columns.AddRange(new DataColumn[]
                    {
                        new DataColumn("Item", typeof(string)),
                        new DataColumn("Value", typeof(object)),
                    });
                }

                // 将应用程序设置赋值到对应的设置DataTable中
                foreach (var settingRelation in lstSettingsRelations) // 遍历设置对应关系列表
                {
                    PropertyInfo propertyInfo = appSettings.GetType().GetProperty(settingRelation.AppSettingItem)!; // 获取应用程序设置中当前设置项属性
                    object? value = propertyInfo?.GetValue(appSettings); // 获取当前设置项属性的值
                    if (value != null)
                    {
                        DataRow newDataRow = dataSet.Tables[settingRelation.DataTableName]!.NewRow(); // 创建一个新数据行
                        // 将设置DataTable中与当前设置项属性相对应的项名和值分别赋值给新数据行的"Item"、"Value"数据列
                        newDataRow["Item"] = settingRelation.DataTableItem;
                        newDataRow["Value"] = value;
                        dataSet.Tables[settingRelation.DataTableName]!.Rows.Add(newDataRow); // 向当前设置DataTable添加新数据行
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
                // 将设置DataTable中的设置值保存到对应的应用程序设置中
                foreach (var settingRelation in lstSettingsRelations) // 遍历设置对应关系列表
                {
                    PropertyInfo propertyInfo = appSettings.GetType().GetProperty(settingRelation.AppSettingItem)!; // 获取应用程序设置中当前设置项属性
                    if (propertyInfo != null && propertyInfo.CanWrite) // 如果设置项属性不为null且可写入
                    {
                        object valueToSet = GetDataTableValue(dataSet.Tables[settingRelation.DataTableName]!, "Item", "Value", settingRelation.DataTableItem, propertyInfo.PropertyType); // 将设置DataTable中与当前设置项属性相对应的数据行的Value值转换为设置项属性的类型
                        propertyInfo.SetValue(appSettings, valueToSet); // 将值设置到设置项属性中
                    }
                }

                settingsManager.SaveSettings(appSettings); // 保存设置

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
