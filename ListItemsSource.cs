namespace COMIGHT
{
    class ListItemsSource : ObservableObject
    {
        // 为每个属性创建私有后备字段
        private List<string> _fontList = new List<string>(); // 字体列表对象

        public List<string> FontList
        {
            get => _fontList;
            set => SetProperty(ref _fontList, value);
        }
    }
}
