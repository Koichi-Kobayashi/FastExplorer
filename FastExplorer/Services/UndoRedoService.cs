using System.Collections.Generic;
using FastExplorer.Models;

namespace FastExplorer.Services
{
    /// <summary>
    /// Undo/Redo操作を管理するサービス
    /// </summary>
    public class UndoRedoService
    {
        private readonly Stack<IUndoableOperation> _undoStack = new();
        private readonly Stack<IUndoableOperation> _redoStack = new();
        private const int MaxHistorySize = 50; // 最大履歴数

        /// <summary>
        /// Undo可能な操作があるかどうか
        /// </summary>
        public bool CanUndo => _undoStack.Count > 0;

        /// <summary>
        /// Redo可能な操作があるかどうか
        /// </summary>
        public bool CanRedo => _redoStack.Count > 0;

        /// <summary>
        /// 操作を履歴に追加します
        /// </summary>
        /// <param name="operation">追加する操作</param>
        public void AddOperation(IUndoableOperation operation)
        {
            if (operation == null)
                return;

            _undoStack.Push(operation);

            // 履歴が最大数を超えた場合、古い操作を削除
            if (_undoStack.Count > MaxHistorySize)
            {
                var tempStack = new Stack<IUndoableOperation>();
                for (int i = 0; i < MaxHistorySize; i++)
                {
                    tempStack.Push(_undoStack.Pop());
                }
                _undoStack.Clear();
                while (tempStack.Count > 0)
                {
                    _undoStack.Push(tempStack.Pop());
                }
            }

            // 新しい操作が追加されたら、Redoスタックをクリア
            _redoStack.Clear();
        }

        /// <summary>
        /// 最後の操作をUndoします
        /// </summary>
        /// <returns>Undoに成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool Undo()
        {
            if (!CanUndo)
                return false;

            var operation = _undoStack.Pop();
            try
            {
                if (operation.Undo())
                {
                    _redoStack.Push(operation);
                    return true;
                }
                else
                {
                    // Undoに失敗した場合はスタックに戻す
                    _undoStack.Push(operation);
                    return false;
                }
            }
            catch
            {
                // 例外が発生した場合はスタックに戻さず、操作を破棄
                return false;
            }
        }

        /// <summary>
        /// 最後のUndo操作をRedoします
        /// </summary>
        /// <returns>Redoに成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool Redo()
        {
            if (!CanRedo)
                return false;

            var operation = _redoStack.Pop();
            try
            {
                if (operation.Redo())
                {
                    _undoStack.Push(operation);
                    return true;
                }
                else
                {
                    // Redoに失敗した場合はスタックに戻す
                    _redoStack.Push(operation);
                    return false;
                }
            }
            catch
            {
                // 例外が発生した場合はスタックに戻さず、操作を破棄
                return false;
            }
        }

        /// <summary>
        /// 履歴をクリアします
        /// </summary>
        public void Clear()
        {
            _undoStack.Clear();
            _redoStack.Clear();
        }
    }
}

