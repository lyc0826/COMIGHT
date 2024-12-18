using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Data;
using static COMIGHT.PublicVariables;
using static COMIGHT.Properties.Settings;
using System.Reflection;


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

        }

        // 定义设置DataTable
        private DataTable generalSettingsTable = new DataTable();
        private DataTable cnSettingsTable = new DataTable();
        private DataTable enSettingsTable = new DataTable();
        private List<(DataTable table, string item, object value)> lstSettings = new List<(DataTable table, string item, object value)>();


        private void LoadSettings()
        {
            List<(DataTable table, string item, object value)> lstSettings = new List<(DataTable table, string item, object value)>
                {
                    (generalSettingsTable, "Saving Folder Path", Default.savingFolderPath),
                    (generalSettingsTable, "Pandoc Application Path", Default.pandocPath),

                    (cnSettingsTable, "Chinese Title Font Name", Default.cnTitleFontName),
                    (cnSettingsTable, "Chinese Title Font Size", Default.cnTitleFontSize),
                    (cnSettingsTable, "Chinese Body Text Font Name", Default.cnBodyFontName),
                    (cnSettingsTable, "Chinese Body Text Font Size", Default.cnBodyFontSize),
                    (cnSettingsTable, "Chinese Heading Lv0 Font Name", Default.cnHeading0FontName),
                    (cnSettingsTable, "Chinese Heading Lv0 Font Size", Default.cnHeading0FontSize),
                    (cnSettingsTable, "Chinese Heading Lv1 Font Name", Default.cnHeading1FontName),
                    (cnSettingsTable, "Chinese Heading Lv1 Font Size", Default.cnHeading1FontSize),
                    (cnSettingsTable, "Chinese Heading Lv2 Font Name", Default.cnHeading2FontName),
                    (cnSettingsTable, "Chinese Heading Lv2 Font Size", Default.cnHeading2FontSize),
                    (cnSettingsTable, "Chinese Heading Lv3-4 Font Name", Default.cnHeading3_4FontName),
                    (cnSettingsTable, "Chinese Heading Lv3-4 Font Size", Default.cnHeading3_4FontSize),
                    (cnSettingsTable, "Chinese Line Space", Default.cnLineSpace),

                    (enSettingsTable, "English Title Font Name", Default.enTitleFontName),
                    (enSettingsTable, "English Title Font Size", Default.enTitleFontSize),
                    (enSettingsTable, "English Body Text Font Name", Default.enBodyFontName),
                    (enSettingsTable, "English Body Text Font Size", Default.enBodyFontSize),
                    (enSettingsTable, "English Heading Lv1 Font Name", Default.enHeading1FontName),
                    (enSettingsTable, "English Heading Lv1 Font Size", Default.enHeading1FontSize),
                    (enSettingsTable, "English Heading Lv2 Font Name", Default.enHeading2FontName),
                    (enSettingsTable, "English Heading Lv2 Font Size", Default.enHeading2FontSize),
                    (enSettingsTable, "English Heading Lv3-4 Font Name", Default.enHeading3_4FontName),
                    (enSettingsTable, "English Heading Lv3-4 Font Size", Default.enHeading3_4FontSize),
                    (enSettingsTable, "English Line Space", Default.enLineSpace)
                };

            try
            {
                // 定义设置DataTable数组
                DataTable[] dataTables = new DataTable[] { generalSettingsTable, cnSettingsTable, enSettingsTable };

                // 遍历设置DataTable数组，为每个数组添加列，并设置主键便于快速查找
                foreach (DataTable dataTable in dataTables)
                {

                    DataColumn[] columnsToAdd = new[]
                    {
                        new DataColumn("Item", typeof(string)),
                        new DataColumn("Value", typeof(object))
                    };

                    // 遍历要添加的列，并检查是否存在；如果不存在则添加
                    foreach (DataColumn column in columnsToAdd)
                    {
                        if (!dataTable.Columns.Contains(column.ColumnName))
                        {
                            dataTable.Columns.Add(column);
                        }
                    }

                    if (!dataTable.PrimaryKey.Any()) // 如果主键不存在，则添加主键
                    {
                        dataTable.PrimaryKey = new[] { dataTable.Columns["Item"]! };
                    }

                    //dataTable.Columns.AddRange(new DataColumn[]
                    //{
                    //    new DataColumn("Item", typeof(string)),
                    //    new DataColumn("Value", typeof(object)),
                    //});
                    // dataTable.PrimaryKey = new[] { dataTable.Columns["Item"]! };
                }

                // 定义设置列表，元素为值元组，包含设置DataTable名、设置项和值
                
                foreach ((DataTable table, string item, object value) setting in lstSettings) // 遍历设置列表
                {
                    setting.table.Rows.Add(setting.item, setting.value); // 向对应的设置DataTable中添加设置项和值
                }

                // 将设置DataTable绑定到设置DataGrid上
                dtgrdGeneralSettings.ItemsSource = generalSettingsTable!.DefaultView;
                dtgrdCnDocumentSettings.ItemsSource = cnSettingsTable!.DefaultView;
                dtgrdEnDocumentSettings.ItemsSource = enSettingsTable!.DefaultView;
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
            return (T) Convert.ChangeType(datatable.Rows.Find(dataTableKeyField)?[datatableValueField] ?? "", typeof(T));
        }

        private void SaveSettings()
        {
            try
            {
                // 保存设置

                //常规设置
                Default.savingFolderPath = GetSettingValue<string>(generalSettingsTable, "Saving Folder Path", "Value");
                Default.pandocPath = GetSettingValue<string>(generalSettingsTable, "Pandoc Application Path", "Value");

                // 中文文档设置
                Default.cnTitleFontName = GetSettingValue<string>(cnSettingsTable, "Chinese Title Font Name", "Value");
                Default.cnTitleFontSize = GetSettingValue<float>(cnSettingsTable, "Chinese Title Font Size", "Value");
                Default.cnBodyFontName = GetSettingValue<string>(cnSettingsTable, "Chinese Body Text Font Name", "Value");
                Default.cnBodyFontSize = GetSettingValue<float>(cnSettingsTable, "Chinese Body Text Font Size", "Value");
                Default.cnHeading1FontName = GetSettingValue<string>(cnSettingsTable, "Chinese Heading Lv1 Font Name", "Value");
                Default.cnHeading1FontSize = GetSettingValue<float>(cnSettingsTable, "Chinese Heading Lv1 Font Size", "Value");
                Default.cnHeading2FontName = GetSettingValue<string>(cnSettingsTable, "Chinese Heading Lv2 Font Name", "Value");
                Default.cnHeading2FontSize = GetSettingValue<float>(cnSettingsTable, "Chinese Heading Lv2 Font Size", "Value");
                Default.cnHeading3_4FontName = GetSettingValue<string>(cnSettingsTable, "Chinese Heading Lv3-4 Font Name", "Value");
                Default.cnHeading3_4FontSize = GetSettingValue<float>(cnSettingsTable, "Chinese Heading Lv3-4 Font Size", "Value");
                Default.cnLineSpace = GetSettingValue<float>(cnSettingsTable, "Chinese Line Space", "Value");

                // 英文文档设置
                Default.enTitleFontName = GetSettingValue<string>(enSettingsTable, "English Title Font Name", "Value");
                Default.enTitleFontSize = GetSettingValue<float>(enSettingsTable, "English Title Font Size", "Value");
                Default.enBodyFontName = GetSettingValue<string>(enSettingsTable, "English Body Text Font Name", "Value");
                Default.enBodyFontSize = GetSettingValue<float>(enSettingsTable, "English Body Text Font Size", "Value");
                Default.enHeading1FontName = GetSettingValue<string>(enSettingsTable, "English Heading Lv1 Font Name", "Value");
                Default.enHeading1FontSize = GetSettingValue<float>(enSettingsTable, "English Heading Lv1 Font Size", "Value");
                Default.enHeading2FontName = GetSettingValue<string>(enSettingsTable, "English Heading Lv2 Font Name", "Value");
                Default.enHeading2FontSize = GetSettingValue<float>(enSettingsTable, "English Heading Lv2 Font Size", "Value");
                Default.enHeading3_4FontName = GetSettingValue<string>(enSettingsTable, "English Heading Lv3-4 Font Name", "Value");
                Default.enHeading3_4FontSize = GetSettingValue<float>(enSettingsTable, "English Heading Lv3-4 Font Size", "Value");
                Default.enLineSpace = GetSettingValue<float>(enSettingsTable, "English Line Space", "Value");


                Default.Save(); // 保存默认设置

                MessageBox.Show("Settings saved successfully.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            this.Close(); // 关闭窗口
        }


        //private void SaveSettings()
        //{
        //    try
        //    {
        //        // 遍历设置列表，从DataTable中读取设置并保存到Properties.Settings
        //        foreach (var setting in lstSettings)
        //        {
        //            // 尝试找到对应的行
        //            DataRow? row = setting.table.Rows.Find(setting.item);
        //            if (row != null && row["Value"] != DBNull.Value)
        //            {
        //                // 使用反射动态设置属性值
        //                var propertyInfo = Default.GetType().GetProperty(nameof(setting.value).Replace("Default.",""), BindingFlags.Public | BindingFlags.Instance);
        //                if (propertyInfo != null && propertyInfo.CanWrite)
        //                {
        //                    // 确保类型匹配后再赋值
        //                    object valueToSet = Convert.ChangeType(row["Value"], propertyInfo.PropertyType);
        //                    propertyInfo.SetValue(Default, valueToSet);
        //                }
        //            }
        //        }

        //        // 保存设置到磁盘
        //        Default.Save();

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
        //    }

        //    this.Close();
        //}

    }
}
