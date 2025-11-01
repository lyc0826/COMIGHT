using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace COMIGHT
{
    // 一个可复用的基类，用于实现属性变更通知
    public class ObservableObject : INotifyPropertyChanged
    {
        // 定义INotifyPropertyChanged接口要求的事件
        public event PropertyChangedEventHandler? PropertyChanged;


        // 定义属性变更通知方法，使用 [CallerMemberName] 特性，调用时可以省略参数，编译器会自动填充调用方（即属性）的名称。
        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // 定义SetProperty方法，允许触发主属性的变更通知，field为属性的后备字段，value为要设置的新值，propertyName为属性名
        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null) // 使用 [CallerMemberName] 特性，调用时可以省略参数，编译器会自动填充调用方（即属性）的名称；如果在调用时显式地提供了这个参数的值，那么将覆盖 [CallerMemberName] 的默认行为。
        {
            // 如果新旧值相同，则不执行任何操作
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return false;
            }

            field = value; // 更新字段值
            OnPropertyChanged(propertyName); // 触发通知
            return true; // 返回 true 表示值已更改
        }

        // 定义SetPropertyAndNotify方法，允许触发多个依赖属性的变更通知，以及主属性的变更通知
        protected bool SetPropertyAndNotify<T>(ref T field, T value, string[] dependentPropertyNames, [CallerMemberName] string? propertyName = null)
        {
            // 调用 SetProperty 来处理主属性的更新和通知，如果主属性值已改变
            if (SetProperty(ref field, value, propertyName))
            {
                if (dependentPropertyNames != null)  // 如果依赖属性不为空
                {
                    foreach (var dependentPropertyName in dependentPropertyNames) // 遍历并通知所有依赖属性
                    {
                        OnPropertyChanged(dependentPropertyName);
                    }
                }
                return true; // 返回 true 表示主属性已更改
            }
            return false; // 返回 false 表示主属性未更改
        }

    }

}

