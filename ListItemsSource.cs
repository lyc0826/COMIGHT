using System.Drawing.Text;

namespace COMIGHT
{
    public class ListItemsSource : ObservableObject
    {
        // 为每个属性创建私有后备字段
        private List<string> _fontList = new List<string>(); // 定义字体列表

        // 定义字体列表属性
        public List<string> FontList
        {
            get => _fontList;
            set => SetProperty(ref _fontList, value);
        }
    }

    
}
