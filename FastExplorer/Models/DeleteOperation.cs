using System.IO;
using FastExplorer.Services;
using Microsoft.VisualBasic.FileIO;

namespace FastExplorer.Models
{
    /// <summary>
    /// ファイル/フォルダーの削除操作を表すクラス
    /// </summary>
    public class DeleteOperation : IUndoableOperation
    {
        private readonly string _path;
        private readonly bool _isDirectory;
        private string? _restoredPath;
        private readonly RecycleBinService? _recycleBinService;
        private readonly DateTime _deletedAt; // 削除時刻

        /// <summary>
        /// 操作の説明
        /// </summary>
        public string Description => $"削除: {Path.GetFileName(_path)}";

        /// <summary>
        /// <see cref="DeleteOperation"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="path">削除されたファイル/フォルダーのパス</param>
        /// <param name="isDirectory">ディレクトリかどうか</param>
        /// <param name="recycleBinService">ゴミ箱サービス（オプション）</param>
        public DeleteOperation(string path, bool isDirectory, RecycleBinService? recycleBinService = null)
        {
            _path = path;
            _isDirectory = isDirectory;
            _recycleBinService = recycleBinService ?? new RecycleBinService();
            _deletedAt = DateTime.Now; // 削除時刻を記録
        }

        /// <summary>
        /// 操作をUndoします（ゴミ箱から復元）
        /// </summary>
        /// <returns>Undoに成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool Undo()
        {
            try
            {
                // ゴミ箱から復元を試みる
                if (_recycleBinService != null)
                {
                    if (_recycleBinService.RestoreFromRecycleBin(_path, _deletedAt))
                    {
                        _restoredPath = _path;
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 操作をRedoします（再度削除）
        /// </summary>
        /// <returns>Redoに成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool Redo()
        {
            try
            {
                // 復元されたパスが存在することを確認
                if (string.IsNullOrEmpty(_restoredPath))
                    return false;

                if (!File.Exists(_restoredPath) && !Directory.Exists(_restoredPath))
                    return false;

                // 再度ゴミ箱に移動
                if (_isDirectory)
                {
                    FileSystem.DeleteDirectory(
                        _restoredPath,
                        UIOption.OnlyErrorDialogs,
                        RecycleOption.SendToRecycleBin);
                }
                else
                {
                    FileSystem.DeleteFile(
                        _restoredPath,
                        UIOption.OnlyErrorDialogs,
                        RecycleOption.SendToRecycleBin);
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

