using System.Drawing.Text;

namespace COMIGHT
{
    public class ListItemsSource : ObservableObject
    {
        // 为每个属性创建私有后备字段
        private List<string> _fontList = new List<string>(); // 定义字体列表
        private List<string> _markupTypeList = new List<string>(); // 定义字体家族列表

        // 定义字体列表属性
        public List<string> FontList
        {
            get => _fontList;
            set => SetProperty(ref _fontList, value);
        }
        
        // 定义标记文本类型列表属性
        public List<string> MarkupTypeList
        {
            get => _markupTypeList;
            set => SetProperty(ref _markupTypeList, value);
        }
    
    }

    
}
