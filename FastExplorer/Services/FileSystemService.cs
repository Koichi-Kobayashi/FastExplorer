using System.IO;
using FastExplorer.Models;

namespace FastExplorer.Services
{
    /// <summary>
    /// ファイルシステム操作を提供するサービス
    /// </summary>
    public class FileSystemService
    {
        /// <summary>
        /// 指定されたパスのディレクトリとファイルの一覧を取得します
        /// </summary>
        /// <param name="path">取得するパス</param>
        /// <returns>ディレクトリを先に、その後ファイルを名前順でソートしたアイテムのコレクション</returns>
        public IEnumerable<FileSystemItem> GetItems(string path)
        {
            if (!Directory.Exists(path))
                return Enumerable.Empty<FileSystemItem>();

            var items = new List<FileSystemItem>();

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
                        
                        if (!isDirectory && info is FileInfo fileInfo)
                        {
                            size = fileInfo.Length;
                        }
                        
                        items.Add(new FileSystemItem
                        {
                            Name = info.Name,
                            FullPath = info.FullName,
                            Extension = isDirectory ? string.Empty : info.Extension,
                            Size = size,
                            LastModified = info.LastWriteTime,
                            IsDirectory = isDirectory,
                            Attributes = info.Attributes
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
            return items.OrderByDescending(x => x.IsDirectory).ThenBy(x => x.Name, StringComparer.OrdinalIgnoreCase);
        }

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

        /// <summary>
        /// システム上のすべての論理ドライブを取得します
        /// </summary>
        /// <returns>ドライブ文字列の配列</returns>
        public string[] GetDrives()
        {
            return Directory.GetLogicalDrives();
        }

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
    }
}

