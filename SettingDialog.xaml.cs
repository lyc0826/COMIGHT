using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Data;
using static COMIGHT.PublicVariables;
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
            dtgrdSettings.ItemsSource = settingTable!.DefaultView; // 将设置DataTable绑定到设置DataGrid上
        }

        private DataTable settingTable;

        private void LoadSettings()
        {
            try
            {
                // 创建设置DataTable, 并添加列
                settingTable = new DataTable();

                settingTable.Columns.AddRange(new DataColumn[]
                    {
                        new DataColumn("Item", typeof(string)),
                        new DataColumn("Value", typeof(object)),
                    });

                // 定义设置列表，元素为值元组，包含设置项和值
                List<(string item, object value)> lstSettings = new List<(string item, object value)>
                    {
                        ("Saving Folder Path", Default.savingFolderPath),
                        ("Pandoc Application Path", Default.pandocPath),
                        ("Chinese Title Font Name", Default.cnTitleFontName),
                        ("Chinese Title Font Size", Default.cnTitleFontSize),
                        ("Chinese Body Text Font Name", Default.cnBodyFontName),
                        ("Chinese Body Text Font Size", Default.cnBodyFontSize),
                        ("Chinese Heading Lv0 Font Name", Default.cnHeading0FontName),
                        ("Chinese Heading Lv0 Font Size", Default.cnHeading0FontSize),
                        ("Chinese Heading Lv1 Font Name", Default.cnHeading1FontName),
                        ("Chinese Heading Lv1 Font Size", Default.cnHeading1FontSize),
                        ("Chinese Heading Lv2 Font Name", Default.cnHeading2FontName),
                        ("Chinese Heading Lv2 Font Size", Default.cnHeading2FontSize),
                        ("Chinese Heading Lv3-4 Font Name", Default.cnHeading3_4FontName),
                        ("Chinese Heading Lv3-4 Font Size", Default.cnHeading3_4FontSize),
                        ("English Title Font Name", Default.enTitleFontName),
                        ("English Title Font Size", Default.enTitleFontSize),
                        ("English Body Text Font Name", Default.enBodyFontName),
                        ("English Body Text Font Size", Default.enBodyFontSize),
                        ("English Heading Font Name", Default.enHeadingFontName),
                        ("English Heading Font Size", Default.enHeadingFontSize),
                        ("Chinese Line Space", Default.cnLineSpace),
                        ("English Line Space", Default.enLineSpace)
                    };

                foreach ((string item, object value) setting in lstSettings) // 遍历设置列表
                {
                    settingTable.Rows.Add(setting.item, setting.value); // 向设置DataTable添加设置项和值
                }

                settingTable.PrimaryKey = new[] { settingTable.Columns["Item"]! }; // 设置 "Item" 列为主键以方便查找
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void btnDialogSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 保存设置

                Default.savingFolderPath = Convert.ToString(settingTable.Rows.Find("Saving Folder Path")?["Value"]);
                Default.pandocPath = Convert.ToString(settingTable.Rows.Find("Pandoc Application Path")?["Value"]);

                Default.cnTitleFontName = Convert.ToString(settingTable.Rows.Find("Chinese Title Font Name")?["Value"]);
                Default.cnTitleFontSize = Convert.ToInt32(settingTable.Rows.Find("Chinese Title Font Size")?["Value"]);
                Default.cnBodyFontName = Convert.ToString(settingTable.Rows.Find("Chinese Body Text Font Name")?["Value"]);
                Default.cnBodyFontSize = Convert.ToInt32(settingTable.Rows.Find("Chinese Body Text Font Size")?["Value"]);
                Default.cnHeading0FontName = Convert.ToString(settingTable.Rows.Find("Chinese Heading Lv0 Font Name")?["Value"]);
                Default.cnHeading0FontSize = Convert.ToInt32(settingTable.Rows.Find("Chinese Heading Lv0 Font Size")?["Value"]);
                Default.cnHeading1FontName = Convert.ToString(settingTable.Rows.Find("Chinese Heading Lv1 Font Name")?["Value"]);
                Default.cnHeading1FontSize = Convert.ToInt32(settingTable.Rows.Find("Chinese Heading Lv1 Font Size")?["Value"]);
                Default.cnHeading2FontName = Convert.ToString(settingTable.Rows.Find("Chinese Heading Lv2 Font Name")?["Value"]);
                Default.cnHeading2FontSize = Convert.ToInt32(settingTable.Rows.Find("Chinese Heading Lv2 Font Size")?["Value"]);
                Default.cnHeading3_4FontName = Convert.ToString(settingTable.Rows.Find("Chinese Heading Lv3-4 Font Name")?["Value"]);
                Default.cnHeading3_4FontSize = Convert.ToInt32(settingTable.Rows.Find("Chinese Heading Lv3-4 Font Size")?["Value"]);
                Default.enTitleFontName = Convert.ToString(settingTable.Rows.Find("English Title Font Name")?["Value"]);
                Default.enTitleFontSize = Convert.ToInt32(settingTable.Rows.Find("English Title Font Size")?["Value"]);
                Default.enBodyFontName = Convert.ToString(settingTable.Rows.Find("English Body Text Font Name")?["Value"]);
                Default.enBodyFontSize = Convert.ToInt32(settingTable.Rows.Find("English Body Text Font Size")?["Value"]);
                Default.enHeadingFontName = Convert.ToString(settingTable.Rows.Find("English Heading Font Name")?["Value"]);
                Default.enHeadingFontSize = Convert.ToInt32(settingTable.Rows.Find("English Heading Font Size")?["Value"]);
                Default.cnLineSpace = Convert.ToInt32(settingTable.Rows.Find("Chinese Line Space")?["Value"]);
                Default.enLineSpace = Convert.ToInt32(settingTable.Rows.Find("English Line Space")?["Value"]);

                Default.Save(); // 保存默认设置

                MessageBox.Show("Settings saved successfully.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            this.Close(); // 关闭窗口
        }

    }
}
