using System.Globalization;
using System.Windows.Data;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// ブール値を分割ペインメニューのテキストに変換するコンバーター
    /// </summary>
    public class BooleanToSplitPaneMenuTextConverter : IValueConverter
    {
        /// <summary>
        /// 値を変換します
        /// </summary>
        /// <param name="value">変換する値（bool型のIsSplitPaneEnabled）</param>
        /// <param name="targetType">変換先の型</param>
        /// <param name="parameter">変換パラメータ</param>
        /// <param name="culture">カルチャ情報</param>
        /// <returns>分割ペインが有効な場合は"分割ペインを無効にする"、それ以外の場合は"分割ペインを有効にする"</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isEnabled)
            {
                return isEnabled ? "分割ペインを無効にする" : "分割ペインを有効にする";
            }
            return "分割ペインを有効にする";
        }

        /// <summary>
        /// 値を逆変換します（実装されていません）
        /// </summary>
        /// <param name="value">逆変換する値</param>
        /// <param name="targetType">変換先の型</param>
        /// <param name="parameter">変換パラメータ</param>
        /// <param name="culture">カルチャ情報</param>
        /// <returns>常に<see cref="NotImplementedException"/>をスローします</returns>
        /// <exception cref="NotImplementedException">常にスローされます</exception>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

