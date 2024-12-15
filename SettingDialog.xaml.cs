using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using static COMIGHT.PublicVariables;
using static COMIGHT.Properties.Settings;


namespace COMIGHT
{
    /// <summary>
    /// Interaction logic for SettingDialog.xaml
    /// </summary>
    public partial class SettingDialog : Window
    {

        private ObservableCollection<SettingItem> settingItems;

        public SettingDialog()
        {
            InitializeComponent();

            LoadSettings();
            dtgrdSettings.ItemsSource = settingItems;
        }

        private void LoadSettings()
        {
            settingItems = new ObservableCollection<SettingItem>
            {
                // Load previous settings or default values here.
                // For example:
                new SettingItem { Item = "Saving Folder Path", Value = Default.savingFolderPath },

                new SettingItem { Item = "Chinese Title Font Name", Value = Default.cnTitleFontName },
                new SettingItem { Item = "Chinese Title Font Size", Value = Default.cnTitleFontSize },
                new SettingItem { Item = "Chinese Body Text Font Name", Value = Default.cnBodyFontName },
                new SettingItem { Item = "Chinese Body Text Font Size", Value = Default.cnBodyFontSize },
                new SettingItem { Item = "Chinese Heading Lv0 Font Name", Value = Default.cnHeading0FontName },
                new SettingItem { Item = "Chinese Heading Lv0 Font Size", Value = Default.cnHeading0FontSize },
                new SettingItem { Item = "Chinese Heading Lv1 Font Name", Value = Default.cnHeading1FontName },
                new SettingItem { Item = "Chinese Heading Lv1 Font Size", Value = Default.cnHeading1FontSize },
                new SettingItem { Item = "Chinese Heading Lv2 Font Name", Value = Default.cnHeading2FontName },
                new SettingItem { Item = "Chinese Heading Lv2 Font Size", Value = Default.cnHeading2FontSize },
                new SettingItem { Item = "Chinese Heading Lv3-4 Font Name", Value = Default.cnHeading3_4FontName },
                new SettingItem { Item = "Chinese Heading Lv3-4 Font Size", Value = Default.cnHeading3_4FontSize },
                new SettingItem { Item = "English Title Font Name", Value = Default.enTitleFontName },
                new SettingItem { Item = "English Title Font Size", Value = Default.enTitleFontSize },
                new SettingItem { Item = "English Body Text Font Name", Value = Default.enBodyFontName },
                new SettingItem { Item = "English Body Text Font Size", Value = Default.enBodyFontSize },
                new SettingItem { Item = "English Heading Font Name", Value = Default.enHeadingFontName },
                new SettingItem { Item = "English Heading Font Size", Value = Default.enHeadingFontSize },
                new SettingItem { Item = "Footer Font Name", Value = Default.footerFontName },
                new SettingItem { Item = "Footer Font Size", Value = Default.footerFontSize },
                new SettingItem { Item = "Chinese Line Space", Value = Default.cnLineSpace },
                new SettingItem { Item = "English Line Space", Value = Default.enLineSpace },

            };
        }

        private void btnDialogSave_Click(object sender, RoutedEventArgs e)
        {
            // 保存设置
            Default.savingFolderPath = (string)settingItems.FirstOrDefault(e => e.Item == "Saving Folder Path")!.Value;

            Default.cnTitleFontName = (string)settingItems.FirstOrDefault(e => e.Item == "Chinese Title Font Name")!.Value;
            Default.cnTitleFontSize = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "Chinese Title Font Size")!.Value);
            Default.cnBodyFontName = (string)settingItems.FirstOrDefault(e => e.Item == "Chinese Body Text Font Name")!.Value;
            Default.cnBodyFontSize = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "Chinese Body Text Font Size")!.Value);
            Default.cnHeading0FontName = (string)settingItems.FirstOrDefault(e => e.Item == "Chinese Heading Lv0 Font Name")!.Value;
            Default.cnHeading0FontSize = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "Chinese Heading Lv0 Font Size")!.Value);
            Default.cnHeading1FontName = (string)settingItems.FirstOrDefault(e => e.Item == "Chinese Heading Lv1 Font Name")!.Value;
            Default.cnHeading1FontSize = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "Chinese Heading Lv1 Font Size")!.Value);
            Default.cnHeading2FontName = (string)settingItems.FirstOrDefault(e => e.Item == "Chinese Heading Lv2 Font Name")!.Value;
            Default.cnHeading2FontSize = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "Chinese Heading Lv2 Font Size")!.Value);
            Default.cnHeading3_4FontName = (string)settingItems.FirstOrDefault(e => e.Item == "Chinese Heading Lv3-4 Font Name")!.Value;
            Default.cnHeading3_4FontSize = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "Chinese Heading Lv3-4 Font Size")!.Value);
            Default.enTitleFontName = (string)settingItems.FirstOrDefault(e => e.Item == "English Title Font Name")!.Value;
            Default.enTitleFontSize = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "English Title Font Size")!.Value);
            Default.enBodyFontName = (string)settingItems.FirstOrDefault(e => e.Item == "English Body Text Font Name")!.Value;
            Default.enBodyFontSize = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "English Body Text Font Size")!.Value);
            Default.enHeadingFontName = (string)settingItems.FirstOrDefault(e => e.Item == "English Heading Font Name")!.Value;
            Default.enHeadingFontSize = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "English Heading Font Size")!.Value);
            Default.footerFontName = (string)settingItems.FirstOrDefault(e => e.Item == "Footer Font Name")!.Value;
            Default.footerFontSize = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "Footer Font Size")!.Value);
            Default.cnLineSpace = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "Chinese Line Space")!.Value);
            Default.enLineSpace = Convert.ToInt32(settingItems.FirstOrDefault(e => e.Item == "English Line Space")!.Value);



            // Continue with the rest of your settings...

            Default.Save();

            MessageBox.Show("Settings saved successfully.", "Result", MessageBoxButton.OK, MessageBoxImage.Information);

            this.Close();
        }


        public class SettingItem : INotifyPropertyChanged
        {
            private string? _item;
            private object? _value;

            public string Item
            {
                get { return _item!; }
                set
                {
                    if (_item != value)
                    {
                        _item = value;
                        OnPropertyChanged(nameof(Item));
                    }
                }
            }

            public object Value
            {
                get { return _value!; }
                set
                {
                    if (this._value != value)
                    {
                        this._value = value;
                        OnPropertyChanged(nameof(Value));
                    }
                }
            }

            public event PropertyChangedEventHandler? PropertyChanged;

            protected void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
