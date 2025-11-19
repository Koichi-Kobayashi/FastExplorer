using System.Collections.Generic;
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
        /// <returns>パンくずリスト形式の文字列（例: "C:\ > Users > kobayashi > Documents"）、または空文字列</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string path && !string.IsNullOrEmpty(path))
            {
                try
                {
                    // パスを正規化
                    var normalizedPath = Path.GetFullPath(path);
                    
                    // ルートディレクトリを取得
                    var root = Path.GetPathRoot(normalizedPath);
                    if (string.IsNullOrEmpty(root))
                    {
                        // ルートがない場合は、パスを分割して処理
                        var parts = normalizedPath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
                        return string.Join(" > ", parts);
                    }
                    
                    // ルートを除いた部分を取得
                    var relativePath = normalizedPath.Substring(root.Length);
                    var parts2 = relativePath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
                    
                    // ルートと各セグメントを結合
                    var breadcrumbParts = new List<string> { root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) };
                    breadcrumbParts.AddRange(parts2);
                    
                    return string.Join(" > ", breadcrumbParts);
                }
                catch
                {
                    // パスが無効な場合は、パスを分割して処理
                    var parts = path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
                    return parts.Length > 0 ? string.Join(" > ", parts) : string.Empty;
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


