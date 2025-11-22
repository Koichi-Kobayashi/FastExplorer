using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Data;
using Cysharp.Text;

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
                        if (parts.Length == 0)
                            return string.Empty;
                        if (parts.Length == 1)
                            return parts[0];
                        
                        // 文字列結合を最適化（ZString.Joinを使用）
                        return ZString.Join(" > ", parts);
                    }
                    
                    // ルートを除いた部分を取得
                    var relativePath = normalizedPath.Substring(root.Length);
                    var parts2 = relativePath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
                    
                    // ルートと各セグメントを結合（StringBuilderを使用してメモリ割り当てを削減）
                    var rootTrimmed = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    if (parts2.Length == 0)
                        return rootTrimmed;
                    
                    // 文字列結合を最適化（ZString.Joinを使用）
                    using var sb2 = ZString.CreateStringBuilder();
                    sb2.Append(rootTrimmed);
                    if (parts2.Length > 0)
                    {
                        sb2.Append(" > ");
                        sb2.Append(parts2[0]);
                        for (int i = 1; i < parts2.Length; i++)
                        {
                            sb2.Append(" > ");
                            sb2.Append(parts2[i]);
                        }
                    }
                    return sb2.ToString();
                }
                catch
                {
                    // パスが無効な場合は、パスを分割して処理
                    var parts = path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length == 0)
                        return string.Empty;
                    if (parts.Length == 1)
                        return parts[0];
                    
                    // 文字列結合を最適化（ZString.Joinを使用）
                    return ZString.Join(" > ", parts);
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


