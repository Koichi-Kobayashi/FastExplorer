using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Data;
using FastExplorer.Models;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// パスをパンくずリスト用のアイテムコレクションに変換するコンバーター
    /// </summary>
    public class PathToBreadcrumbItemsConverter : IValueConverter
    {
        /// <summary>
        /// 値を変換します
        /// </summary>
        /// <param name="value">変換する値（string型のパス）</param>
        /// <param name="targetType">変換先の型</param>
        /// <param name="parameter">変換パラメータ</param>
        /// <param name="culture">カルチャ情報</param>
        /// <returns>パンくずリスト用のアイテムコレクション</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var items = new List<BreadcrumbItem>();
            
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
                        if (parts.Length > 0)
                        {
                            var currentPath = string.Empty;
                            foreach (var part in parts)
                            {
                                if (string.IsNullOrEmpty(currentPath))
                                {
                                    currentPath = part;
                                }
                                else
                                {
                                    currentPath = Path.Combine(currentPath, part);
                                }
                                items.Add(new BreadcrumbItem { Name = part, Path = currentPath });
                            }
                        }
                    }
                    else
                    {
                        // ルートを追加
                        var rootTrimmed = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                        items.Add(new BreadcrumbItem { Name = rootTrimmed, Path = root });
                        
                        // ルートを除いた部分を取得
                        var relativePath = normalizedPath.Substring(root.Length);
                        var parts = relativePath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
                        
                        var currentPath = root;
                        foreach (var part in parts)
                        {
                            // ルートディレクトリの場合は、Path.Combineが正しく動作しないため、手動で結合
                            if (currentPath.EndsWith(Path.DirectorySeparatorChar.ToString()) || currentPath.EndsWith(Path.AltDirectorySeparatorChar.ToString()))
                            {
                                currentPath = currentPath + part;
                            }
                            else
                            {
                                currentPath = Path.Combine(currentPath, part);
                            }
                            items.Add(new BreadcrumbItem { Name = part, Path = currentPath });
                        }
                    }
                }
                catch
                {
                    // パスが無効な場合は、パスを分割して処理
                    var parts = path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length > 0)
                    {
                        var currentPath = string.Empty;
                        foreach (var part in parts)
                        {
                            if (string.IsNullOrEmpty(currentPath))
                            {
                                currentPath = part;
                            }
                            else
                            {
                                currentPath = Path.Combine(currentPath, part);
                            }
                            items.Add(new BreadcrumbItem { Name = part, Path = currentPath });
                        }
                    }
                }
            }
            
            return items;
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

    /// <summary>
    /// パンくずリストのアイテムを表すクラス
    /// </summary>
    public class BreadcrumbItem
    {
        /// <summary>
        /// 表示名を取得または設定します
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// パスを取得または設定します
        /// </summary>
        public string Path { get; set; } = string.Empty;
    }
}

