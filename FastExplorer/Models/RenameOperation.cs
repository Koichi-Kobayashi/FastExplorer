using System.IO;

namespace FastExplorer.Models
{
    /// <summary>
    /// ファイル/フォルダーのリネーム操作を表すクラス
    /// </summary>
    public class RenameOperation : IUndoableOperation
    {
        private readonly string _oldPath;
        private readonly string _newPath;
        private readonly bool _isDirectory;

        /// <summary>
        /// 操作の説明
        /// </summary>
        public string Description => $"名前を変更: {Path.GetFileName(_oldPath)} → {Path.GetFileName(_newPath)}";

        /// <summary>
        /// <see cref="RenameOperation"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="oldPath">変更前のパス</param>
        /// <param name="newPath">変更後のパス</param>
        /// <param name="isDirectory">ディレクトリかどうか</param>
        public RenameOperation(string oldPath, string newPath, bool isDirectory)
        {
            _oldPath = oldPath;
            _newPath = newPath;
            _isDirectory = isDirectory;
        }

        /// <summary>
        /// 操作をUndoします（元の名前に戻す）
        /// </summary>
        /// <returns>Undoに成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool Undo()
        {
            try
            {
                // 新しいパスが存在することを確認
                if (!File.Exists(_newPath) && !Directory.Exists(_newPath))
                    return false;

                // 古いパスが既に存在する場合は失敗
                if (File.Exists(_oldPath) || Directory.Exists(_oldPath))
                    return false;

                // リネームを実行
                if (_isDirectory)
                {
                    Directory.Move(_newPath, _oldPath);
                }
                else
                {
                    File.Move(_newPath, _oldPath);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 操作をRedoします（新しい名前に戻す）
        /// </summary>
        /// <returns>Redoに成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool Redo()
        {
            try
            {
                // 古いパスが存在することを確認
                if (!File.Exists(_oldPath) && !Directory.Exists(_oldPath))
                    return false;

                // 新しいパスが既に存在する場合は失敗
                if (File.Exists(_newPath) || Directory.Exists(_newPath))
                    return false;

                // リネームを実行
                if (_isDirectory)
                {
                    Directory.Move(_oldPath, _newPath);
                }
                else
                {
                    File.Move(_oldPath, _newPath);
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

