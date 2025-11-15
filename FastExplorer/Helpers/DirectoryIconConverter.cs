using System.Globalization;
using System.Windows.Data;
using Wpf.Ui.Controls;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// ディレクトリかどうかに基づいてアイコンを変換するコンバーター
    /// </summary>
    public class DirectoryIconConverter : IValueConverter
    {
        /// <summary>
        /// 値を変換します
        /// </summary>
        /// <param name="value">変換する値（bool型のisDirectory）</param>
        /// <param name="targetType">変換先の型</param>
        /// <param name="parameter">変換パラメータ</param>
        /// <param name="culture">カルチャ情報</param>
        /// <returns>ディレクトリの場合はFolder24、それ以外の場合はDocument24のシンボル</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isDirectory)
            {
                return isDirectory ? SymbolRegular.Folder24 : SymbolRegular.Document24;
            }
            return SymbolRegular.Document24;
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

