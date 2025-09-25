using System.Globalization;
using System.Windows.Data;

namespace COMIGHT
{
    // ����ö��ת���������ڽ�ö��ֵת��Ϊ����ֵ���̳�IValueConverter�ӿ�
    public class EnumToBooleanConverter : IValueConverter
    {
        // ��ö��ֵת��Ϊ����ֵ������ȷ��RadioButton�Ƿ�ѡ�У�
        // "value"��Դö��ֵ������DataContext�е�����ֵ��
        // "targetType"��Ŀ�����ͣ�typeof(bool)��
        // "parameter"��ת������������XAML��ConverterParameterָ����ö��ֵ�ַ�����
        // "culture"��ǰ��������Ϣ
        // ���value��parameterƥ���򷵻�true�����򷵻�false
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return false;

            string enumValue = value.ToString()!;
            string targetValue = parameter.ToString()!;

            return enumValue.Equals(targetValue, StringComparison.OrdinalIgnoreCase); // ö��ֵ�ַ����Ƚϣ����رȽϽ������ֵ
        }

        // ������ֵת����ö��ֵ����RadioButton��ѡ��ʱ����Դ���ԣ�
        // "value" Դ����ֵ������RadioButton��IsChecked���ԣ�
        // "targetType"Ŀ�����ͣ�ö�����ͣ�
        // "parameter">ת������������XAML��ConverterParameterָ����ö��ֵ�ַ�����
        // "culture">��ǰ��������Ϣ
        // ���valueΪtrue�򷵻ض�Ӧ��ö��ֵ�����򷵻�Binding.DoNothing��ʾ�����и���
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return Binding.DoNothing;

            bool isChecked = (bool)value;
            if (!isChecked)
                return Binding.DoNothing;

            return Enum.Parse(targetType, parameter.ToString()!); // ö��ֵ�ַ���ת����ö��ֵ������
        }
    }
}