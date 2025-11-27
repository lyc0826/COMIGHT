using System.Globalization;
using System.Windows.Data;

namespace COMIGHT
{
    // 定义枚举转换器，用于将枚举值转换为布尔值，继承IValueConverter接口
    public class EnumToBooleanConverter : IValueConverter
    {
        // 将枚举值转换为布尔值（用于确定RadioButton是否选中）
        // "value"：源枚举值（来自DataContext中的属性值）
        // "targetType"：目标类型（typeof(bool)）
        // "parameter"：转换参数（来自XAML中ConverterParameter指定的枚举值字符串）
        // "culture"当前区域性信息
        // 如果value与parameter匹配则返回true，否则返回false
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return false;

            string enumValue = value.ToString()!;
            string targetValue = parameter.ToString()!;

            return enumValue.Equals(targetValue, StringComparison.OrdinalIgnoreCase); // 枚举值字符串比较，返回比较结果布尔值
        }

        // 将布尔值转换回枚举值（当RadioButton被选中时更新源属性）
        // "value" 源布尔值（来自RadioButton的IsChecked属性）
        // "targetType"目标类型（枚举类型）
        // "parameter">转换参数（来自XAML中ConverterParameter指定的枚举值字符串）
        // "culture">当前区域性信息
        // 如果value为true则返回对应的枚举值，否则返回Binding.DoNothing表示不进行更新
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return Binding.DoNothing;

            bool isChecked = (bool)value;
            if (!isChecked)
            {
                return Binding.DoNothing;
            }
            return Enum.Parse(targetType, parameter.ToString()!); // 枚举值字符串转换成枚举值，返回
        }
    }
}