using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Data;
using Cysharp.Text;
using FastExplorer.Models;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// パスをパンくずリスト用のアイテムコレクションに変換するコンバーター
    /// </summary>
    public class PathToBreadcrumbItemsConverter : IValueConverter
    {
        // 文字列定数（メモリ割り当てを削減）
        private static readonly char DirectorySeparator = Path.DirectorySeparatorChar;
        private static readonly char AltDirectorySeparator = Path.AltDirectorySeparatorChar;

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
            if (value is not string path || string.IsNullOrEmpty(path))
            {
                return new List<BreadcrumbItem>(0);
            }

            // リストの容量を事前に確保（平均的なパス深度を想定）
            var items = new List<BreadcrumbItem>(8);
            
            try
            {
                // パスを正規化
                var normalizedPath = Path.GetFullPath(path);
                
                // ルートディレクトリを取得
                var root = Path.GetPathRoot(normalizedPath);
                if (string.IsNullOrEmpty(root))
                {
                    // ルートがない場合は、パスを分割して処理
                    var parts = normalizedPath.Split(DirectorySeparator, AltDirectorySeparator, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length == 0)
                    {
                        return items;
                    }

                    // リストの容量を事前に確保
                    items.Capacity = parts.Length;
                    
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
                        // パスを正規化して、完全なパスを取得
                        if (!string.IsNullOrEmpty(currentPath))
                        {
                            try
                            {
                                currentPath = Path.GetFullPath(currentPath);
                            }
                            catch
                            {
                                // パスが無効な場合は、そのまま使用
                            }
                        }
                        items.Add(new BreadcrumbItem { Name = part, Path = currentPath });
                    }
                }
                else
                {
                    // ルートを追加
                    var rootTrimmed = root.TrimEnd(DirectorySeparator, AltDirectorySeparator);
                    items.Add(new BreadcrumbItem { Name = rootTrimmed, Path = root });
                    
                    // ルートを除いた部分を取得
                    var relativePath = normalizedPath.Substring(root.Length);
                    var parts = relativePath.Split(DirectorySeparator, AltDirectorySeparator, StringSplitOptions.RemoveEmptyEntries);
                    
                    if (parts.Length == 0)
                    {
                        return items;
                    }

                    // リストの容量を事前に確保
                    items.Capacity = items.Count + parts.Length;
                    
                    var currentPath = root;
                    
                    foreach (var part in parts)
                    {
                        // Path.Combineを使用してパスを正しく結合
                        // Path.Combineはルートディレクトリの場合でも正しく動作する
                        currentPath = Path.Combine(currentPath, part);
                        // パスを正規化して、完全なパスを取得
                        currentPath = Path.GetFullPath(currentPath);
                        items.Add(new BreadcrumbItem { Name = part, Path = currentPath });
                    }
                }
            }
            catch
            {
                // パスが無効な場合は、パスを分割して処理
                var parts = path.Split(DirectorySeparator, AltDirectorySeparator, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 0)
                {
                    return items;
                }

                // リストの容量を事前に確保
                items.Capacity = parts.Length;
                
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
                    // パスを正規化して、完全なパスを取得
                    if (!string.IsNullOrEmpty(currentPath))
                    {
                        try
                        {
                            currentPath = Path.GetFullPath(currentPath);
                        }
                        catch
                        {
                            // パスが無効な場合は、そのまま使用
                        }
                    }
                    items.Add(new BreadcrumbItem { Name = part, Path = currentPath });
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

