using System.Text.RegularExpressions;

namespace COMIGHT
{
    public static class Constants
    {
        // 定义表格标题正则表达式（需要兼顾常规字符串和Word中的文本；总长度限制在100个字符内；字符中不允许出现“。；;”）
        public static Regex regExTableTitle = new Regex(@"(?<=^|\n|\r)(?=.{1,100}(?:[\n\r]|$))[^。；;\f\n\r]*(?:表|单|录|册|回执|table|form|list|roll|roster)[^。；;\f\n\r]*(?:[\n\r]|$)", RegexOptions.Multiline | RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }
}
