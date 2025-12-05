namespace FastExplorer.Models
{
    /// <summary>
    /// Undo可能な操作を表すインターフェース
    /// </summary>
    public interface IUndoableOperation
    {
        /// <summary>
        /// 操作の説明
        /// </summary>
        string Description { get; }

        /// <summary>
        /// 操作をUndoします
        /// </summary>
        /// <returns>Undoに成功した場合はtrue、それ以外の場合はfalse</returns>
        bool Undo();

        /// <summary>
        /// 操作をRedoします
        /// </summary>
        /// <returns>Redoに成功した場合はtrue、それ以外の場合はfalse</returns>
        bool Redo();
    }
}

