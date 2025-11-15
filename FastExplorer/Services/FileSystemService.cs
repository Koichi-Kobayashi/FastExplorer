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
                // ディレクトリを先に追加
                var directories = Directory.GetDirectories(path);
                foreach (var dir in directories)
                {
                    try
                    {
                        var dirInfo = new DirectoryInfo(dir);
                        items.Add(new FileSystemItem
                        {
                            Name = dirInfo.Name,
                            FullPath = dirInfo.FullName,
                            Extension = string.Empty,
                            Size = 0,
                            LastModified = dirInfo.LastWriteTime,
                            IsDirectory = true,
                            Attributes = dirInfo.Attributes
                        });
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // アクセス権限がない場合はスキップ
                        continue;
                    }
                }

                // ファイルを追加
                var files = Directory.GetFiles(path);
                foreach (var file in files)
                {
                    try
                    {
                        var fileInfo = new FileInfo(file);
                        items.Add(new FileSystemItem
                        {
                            Name = fileInfo.Name,
                            FullPath = fileInfo.FullName,
                            Extension = fileInfo.Extension,
                            Size = fileInfo.Length,
                            LastModified = fileInfo.LastWriteTime,
                            IsDirectory = false,
                            Attributes = fileInfo.Attributes
                        });
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // アクセス権限がない場合はスキップ
                        continue;
                    }
                }
            }
            catch (Exception)
            {
                return Enumerable.Empty<FileSystemItem>();
            }

            // ディレクトリを先に、その後ファイルを名前順でソート
            return items.OrderByDescending(x => x.IsDirectory).ThenBy(x => x.Name);
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
    }
}

