using System.Globalization;
using System.Windows.Data;
using Wpf.Ui.Appearance;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// 列挙型をブール値に変換するコンバーター
    /// </summary>
    internal class EnumToBooleanConverter : IValueConverter
    {
        /// <summary>
        /// 値を変換します
        /// </summary>
        /// <param name="value">変換する値（ApplicationTheme列挙型）</param>
        /// <param name="targetType">変換先の型</param>
        /// <param name="parameter">変換パラメータ（列挙型の名前を表す文字列）</param>
        /// <param name="culture">カルチャ情報</param>
        /// <returns>値がパラメータで指定された列挙値と等しい場合はtrue、それ以外の場合はfalse</returns>
        /// <exception cref="ArgumentException">パラメータが文字列でない場合、または値がApplicationTheme列挙型でない場合にスローされます</exception>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (parameter is not String enumString)
            {
                throw new ArgumentException("ExceptionEnumToBooleanConverterParameterMustBeAnEnumName");
            }

            if (!Enum.IsDefined(typeof(ApplicationTheme), value))
            {
                throw new ArgumentException("ExceptionEnumToBooleanConverterValueMustBeAnEnum");
            }

            var enumValue = Enum.Parse(typeof(ApplicationTheme), enumString);

            return enumValue.Equals(value);
        }

        /// <summary>
        /// 値を逆変換します
        /// </summary>
        /// <param name="value">逆変換する値</param>
        /// <param name="targetType">変換先の型</param>
        /// <param name="parameter">変換パラメータ（列挙型の名前を表す文字列）</param>
        /// <param name="culture">カルチャ情報</param>
        /// <returns>パラメータで指定された列挙値</returns>
        /// <exception cref="ArgumentException">パラメータが文字列でない場合にスローされます</exception>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (parameter is not String enumString)
            {
                throw new ArgumentException("ExceptionEnumToBooleanConverterParameterMustBeAnEnumName");
            }

            return Enum.Parse(typeof(ApplicationTheme), enumString);
        }
    }
}
