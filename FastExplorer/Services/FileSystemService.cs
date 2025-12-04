using System.IO;
using FastExplorer.Models;

namespace FastExplorer.Services
{
    /// <summary>
    /// ファイルシステム操作を提供するサービス
    /// </summary>
    public class FileSystemService
    {
        #region ファイル・ディレクトリ取得

        /// <summary>
        /// 指定されたパスのディレクトリとファイルの一覧を取得します
        /// </summary>
        /// <param name="path">取得するパス</param>
        /// <returns>ディレクトリを先に、その後ファイルを名前順でソートしたアイテムのコレクション</returns>
        public IEnumerable<FileSystemItem> GetItems(string path)
        {
            if (!Directory.Exists(path))
                return Enumerable.Empty<FileSystemItem>();

            // リストの容量を事前に推定（平均的なディレクトリサイズを想定）
            // 実際の数がわからないため、初期容量を256に設定（必要に応じて拡張される）
            // 32から256に増やすことで、メモリ再割り当てを削減（高速化）
            var items = new List<FileSystemItem>(256);

            try
            {
                var dirInfo = new DirectoryInfo(path);
                
                // EnumerateFileSystemInfosを使用してディレクトリとファイルを一度に取得
                // これにより、GetDirectories()とGetFiles()を別々に呼び出すよりも高速
                foreach (var info in dirInfo.EnumerateFileSystemInfos("*", new EnumerationOptions
                {
                    IgnoreInaccessible = true,
                    ReturnSpecialDirectories = false
                }))
                {
                    try
                    {
                        var isDirectory = (info.Attributes & FileAttributes.Directory) == FileAttributes.Directory;
                        long size = 0;
                        
                        // 型チェックを最適化（パターンマッチングを使用して高速化）
                        if (!isDirectory && info is FileInfo fileInfo)
                        {
                            size = fileInfo.Length;
                        }
                        
                        // プロパティアクセスをキャッシュ（高速化）
                        var name = info.Name;
                        var fullPath = info.FullName;
                        // 拡張子の取得を最適化（isDirectoryがtrueの場合はstring.Emptyを直接使用）
                        var extension = isDirectory ? string.Empty : info.Extension;
                        var lastModified = info.LastWriteTime;
                        var attributes = info.Attributes;
                        
                        items.Add(new FileSystemItem
                        {
                            Name = name,
                            FullPath = fullPath,
                            Extension = extension,
                            Size = size,
                            LastModified = lastModified,
                            IsDirectory = isDirectory,
                            Attributes = attributes
                        });
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // アクセス権限がない場合はスキップ
                        continue;
                    }
                    catch (FileNotFoundException)
                    {
                        // ファイルが削除された場合はスキップ
                        continue;
                    }
                }
            }
            catch (Exception)
            {
                return Enumerable.Empty<FileSystemItem>();
            }

            // ディレクトリを先に、その後ファイルを名前順でソート
            // ソートを最適化：リストが空の場合は早期リターン
            if (items.Count == 0)
                return Enumerable.Empty<FileSystemItem>();

            // ソートを最適化：Array.Sortを使用してメモリ割り当てを削減
            // 比較関数を最適化：ディレクトリの比較を高速化（boolの比較を直接行う）
            items.Sort((x, y) =>
            {
                // ディレクトリを先に（boolの比較を直接行うことで高速化）
                if (x.IsDirectory != y.IsDirectory)
                {
                    return y.IsDirectory ? 1 : -1;
                }
                
                // 名前でソート（大文字小文字を区別しない）
                // ReadOnlySpan<char>を使用してメモリ割り当てを削減
                var xNameSpan = x.Name.AsSpan();
                var yNameSpan = y.Name.AsSpan();
                return xNameSpan.CompareTo(yNameSpan, StringComparison.OrdinalIgnoreCase);
            });

            return items;
        }

        #endregion

        #region パス操作

        /// <summary>
        /// 指定されたパスの親ディレクトリのパスを取得します
        /// </summary>
        /// <param name="path">親パスを取得するパス</param>
        /// <returns>親ディレクトリのパス。取得できない場合は空文字列</returns>
        public string GetParentPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return string.Empty;

            try
            {
                var parent = Directory.GetParent(path);
                return parent?.FullName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion

        #region パス検証

        /// <summary>
        /// 指定されたパスが有効かどうかを確認します
        /// </summary>
        /// <param name="path">確認するパス</param>
        /// <returns>パスが有効な場合はtrue、それ以外の場合はfalse</returns>
        public bool IsValidPath(string path)
        {
            try
            {
                return Directory.Exists(path) || File.Exists(path);
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region ドライブ操作

        /// <summary>
        /// システム上のすべての論理ドライブを取得します
        /// </summary>
        /// <returns>ドライブ文字列の配列</returns>
        public string[] GetDrives()
        {
            return Directory.GetLogicalDrives();
        }

        #endregion

        #region ファイル操作

        /// <summary>
        /// ファイルまたはフォルダーを指定されたディレクトリに移動します
        /// </summary>
        /// <param name="sourcePath">移動元のパス</param>
        /// <param name="destinationDirectory">移動先のディレクトリパス</param>
        /// <returns>移動に成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool MoveItem(string sourcePath, string destinationDirectory)
        {
            if (string.IsNullOrEmpty(sourcePath) || string.IsNullOrEmpty(destinationDirectory))
                return false;

            try
            {
                // 移動先のディレクトリが存在しない場合は作成
                if (!Directory.Exists(destinationDirectory))
                {
                    Directory.CreateDirectory(destinationDirectory);
                }

                // 移動先のパスを構築
                var fileName = Path.GetFileName(sourcePath);
                var destinationPath = Path.Combine(destinationDirectory, fileName);

                // 同じパスの場合は何もしない
                if (string.Equals(sourcePath, destinationPath, StringComparison.OrdinalIgnoreCase))
                    return false;

                // 移動先に既に同じ名前のファイル/フォルダーが存在する場合はエラー
                if (File.Exists(destinationPath) || Directory.Exists(destinationPath))
                    return false;

                // ファイルまたはディレクトリを移動
                if (File.Exists(sourcePath))
                {
                    File.Move(sourcePath, destinationPath);
                }
                else if (Directory.Exists(sourcePath))
                {
                    Directory.Move(sourcePath, destinationPath);
                }
                else
                {
                    return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region 削除操作

        /// <summary>
        /// ファイルまたはフォルダーをゴミ箱に移動します
        /// </summary>
        /// <param name="path">削除するファイルまたはフォルダーのパス</param>
        /// <returns>削除に成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool DeleteToRecycleBin(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            try
            {
                // Microsoft.VisualBasic.FileIO.FileSystemを使用してゴミ箱に移動
                if (File.Exists(path))
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(
                        path,
                        Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                        Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                }
                else if (Directory.Exists(path))
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(
                        path,
                        Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                        Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                }
                else
                {
                    return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

    }
}

