using System.Globalization;
using System.IO;
using System.Windows.Data;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// パスをブレッドクラム表示用の文字列に変換するコンバーター
    /// </summary>
    public class PathToBreadcrumbConverter : IValueConverter
    {
        /// <summary>
        /// 値を変換します
        /// </summary>
        /// <param name="value">変換する値（string型のパス）</param>
        /// <param name="targetType">変換先の型</param>
        /// <param name="parameter">変換パラメータ</param>
        /// <param name="culture">カルチャ情報</param>
        /// <returns>パスの最後のディレクトリ名、または空文字列</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string path && !string.IsNullOrEmpty(path))
            {
                try
                {
                    var dirInfo = new DirectoryInfo(path);
                    return dirInfo.Name;
                }
                catch
                {
                    // パスが無効な場合は、パスの最後の部分を返す
                    var parts = path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    return parts.Length > 0 ? parts[parts.Length - 1] : string.Empty;
                }
            }
            return string.Empty;
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

