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
            System.Diagnostics.Debug.WriteLine($"[DeleteOperation] Undo開始: {_path}, IsDirectory: {_isDirectory}, DeletedAt: {_deletedAt}");
            try
            {
                // ゴミ箱から復元を試みる
                if (_recycleBinService != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[DeleteOperation] RecycleBinService.RestoreFromRecycleBinを呼び出します: {_path}");
                    var restoreResult = _recycleBinService.RestoreFromRecycleBin(_path, _deletedAt);
                    System.Diagnostics.Debug.WriteLine($"[DeleteOperation] RestoreFromRecycleBinの結果: {restoreResult}");
                    if (restoreResult)
                    {
                        _restoredPath = _path;
                        System.Diagnostics.Debug.WriteLine($"[DeleteOperation] Undo成功: {_path}");
                        
                        // 復元後のファイル存在確認（少し待機してから確認）
                        System.Threading.Thread.Sleep(200);
                        bool fileExists = _isDirectory ? Directory.Exists(_path) : File.Exists(_path);
                        System.Diagnostics.Debug.WriteLine($"[DeleteOperation] 復元後のファイル存在確認: {fileExists}, パス: {_path}");
                        
                        return true;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[DeleteOperation] RestoreFromRecycleBinがfalseを返しました: {_path}");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[DeleteOperation] _recycleBinServiceがnullです: {_path}");
                }
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DeleteOperation] Undoで例外が発生しました: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[DeleteOperation] スタックトレース: {ex.StackTrace}");
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

