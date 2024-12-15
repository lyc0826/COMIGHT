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
                settingTable.Columns.Add("Item", typeof(string));
                settingTable.Columns.Add("Value", typeof(object));

                // 添加设置项和默认值到设置DataTable
                settingTable.Rows.Add("Saving Folder Path", Default.savingFolderPath);
                settingTable.Rows.Add("Chinese Title Font Name", Default.cnTitleFontName);
                settingTable.Rows.Add("Chinese Title Font Size", Default.cnTitleFontSize);
                settingTable.Rows.Add("Chinese Body Text Font Name", Default.cnBodyFontName);
                settingTable.Rows.Add("Chinese Body Text Font Size", Default.cnBodyFontSize);
                settingTable.Rows.Add("Chinese Heading Lv0 Font Name", Default.cnHeading0FontName);
                settingTable.Rows.Add("Chinese Heading Lv0 Font Size", Default.cnHeading0FontSize);
                settingTable.Rows.Add("Chinese Heading Lv1 Font Name", Default.cnHeading1FontName);
                settingTable.Rows.Add("Chinese Heading Lv1 Font Size", Default.cnHeading1FontSize);
                settingTable.Rows.Add("Chinese Heading Lv2 Font Name", Default.cnHeading2FontName);
                settingTable.Rows.Add("Chinese Heading Lv2 Font Size", Default.cnHeading2FontSize);
                settingTable.Rows.Add("Chinese Heading Lv3-4 Font Name", Default.cnHeading3_4FontName);
                settingTable.Rows.Add("Chinese Heading Lv3-4 Font Size", Default.cnHeading3_4FontSize);
                settingTable.Rows.Add("English Title Font Name", Default.enTitleFontName);
                settingTable.Rows.Add("English Title Font Size", Default.enTitleFontSize);
                settingTable.Rows.Add("English Body Text Font Name", Default.enBodyFontName);
                settingTable.Rows.Add("English Body Text Font Size", Default.enBodyFontSize);
                settingTable.Rows.Add("English Heading Font Name", Default.enHeadingFontName);
                settingTable.Rows.Add("English Heading Font Size", Default.enHeadingFontSize);
                settingTable.Rows.Add("Footer Font Name", Default.footerFontName);
                settingTable.Rows.Add("Footer Font Size", Default.footerFontSize);
                settingTable.Rows.Add("Chinese Line Space", Default.cnLineSpace);
                settingTable.Rows.Add("English Line Space", Default.enLineSpace);

                // 设置 "Item" 列为主键以方便查找
                settingTable.PrimaryKey = new[] { settingTable.Columns["Item"]! };
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
                DataRow? row = settingTable.Rows.Find("Saving Folder Path");
                if (row != null) Default.savingFolderPath = (string)row["Value"];

                row = settingTable.Rows.Find("Chinese Title Font Name");
                if (row != null) Default.cnTitleFontName = (string)row["Value"];
                row = settingTable.Rows.Find("Chinese Title Font Size");
                if (row != null) Default.cnTitleFontSize = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("Chinese Body Text Font Name");
                if (row != null) Default.cnBodyFontName = (string)row["Value"];
                row = settingTable.Rows.Find("Chinese Body Text Font Size");
                if (row != null) Default.cnBodyFontSize = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("Chinese Heading Lv0 Font Name");
                if (row != null) Default.cnHeading0FontName = (string)row["Value"];
                row = settingTable.Rows.Find("Chinese Heading Lv0 Font Size");
                if (row != null) Default.cnHeading0FontSize = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("Chinese Heading Lv1 Font Name");
                if (row != null) Default.cnHeading1FontName = (string)row["Value"];
                row = settingTable.Rows.Find("Chinese Heading Lv1 Font Size");
                if (row != null) Default.cnHeading1FontSize = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("Chinese Heading Lv2 Font Name");
                if (row != null) Default.cnHeading2FontName = (string)row["Value"];
                row = settingTable.Rows.Find("Chinese Heading Lv2 Font Size");
                if (row != null) Default.cnHeading2FontSize = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("Chinese Heading Lv3-4 Font Name");
                if (row != null) Default.cnHeading3_4FontName = (string)row["Value"];
                row = settingTable.Rows.Find("Chinese Heading Lv3-4 Font Size");
                if (row != null) Default.cnHeading3_4FontSize = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("English Title Font Name");
                if (row != null) Default.enTitleFontName = (string)row["Value"];
                row = settingTable.Rows.Find("English Title Font Size");
                if (row != null) Default.enTitleFontSize = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("English Body Text Font Name");
                if (row != null) Default.enBodyFontName = (string)row["Value"];
                row = settingTable.Rows.Find("English Body Text Font Size");
                if (row != null) Default.enBodyFontSize = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("English Heading Font Name");
                if (row != null) Default.enHeadingFontName = (string)row["Value"];
                row = settingTable.Rows.Find("English Heading Font Size");
                if (row != null) Default.enHeadingFontSize = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("Footer Font Name");
                if (row != null) Default.footerFontName = (string)row["Value"];
                row = settingTable.Rows.Find("Footer Font Size");
                if (row != null) Default.footerFontSize = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("Chinese Line Space");
                if (row != null) Default.cnLineSpace = Convert.ToInt32(row["Value"]);
                row = settingTable.Rows.Find("English Line Space");
                if (row != null) Default.enLineSpace = Convert.ToInt32(row["Value"]);

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
