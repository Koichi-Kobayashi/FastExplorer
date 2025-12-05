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
            {
                System.Diagnostics.Debug.WriteLine("[UndoRedoService] AddOperation: operationがnullです");
                return;
            }

            System.Diagnostics.Debug.WriteLine($"[UndoRedoService] AddOperation: {operation.Description} を追加します。現在のスタックサイズ: {_undoStack.Count}");
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
            System.Diagnostics.Debug.WriteLine($"[UndoRedoService] AddOperation完了。スタックサイズ: {_undoStack.Count}");
        }

        /// <summary>
        /// 最後の操作をUndoします
        /// </summary>
        /// <returns>Undoに成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool Undo()
        {
            if (!CanUndo)
            {
                System.Diagnostics.Debug.WriteLine("[UndoRedoService] Undo: CanUndoがfalseです。スタックサイズ: {_undoStack.Count}");
                return false;
            }

            var operation = _undoStack.Pop();
            System.Diagnostics.Debug.WriteLine($"[UndoRedoService] Undo: {operation.Description} をUndoします。残りのスタックサイズ: {_undoStack.Count}");
            try
            {
                var undoResult = operation.Undo();
                System.Diagnostics.Debug.WriteLine($"[UndoRedoService] Undo: {operation.Description} のUndo結果: {undoResult}");
                if (undoResult)
                {
                    _redoStack.Push(operation);
                    System.Diagnostics.Debug.WriteLine($"[UndoRedoService] Undo成功。Redoスタックサイズ: {_redoStack.Count}");
                    return true;
                }
                else
                {
                    // Undoに失敗した場合はスタックに戻す
                    _undoStack.Push(operation);
                    System.Diagnostics.Debug.WriteLine($"[UndoRedoService] Undo失敗。スタックに戻しました。スタックサイズ: {_undoStack.Count}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                // 例外が発生した場合はスタックに戻さず、操作を破棄
                System.Diagnostics.Debug.WriteLine($"[UndoRedoService] Undoで例外が発生しました: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[UndoRedoService] スタックトレース: {ex.StackTrace}");
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

