using System;
using System.IO;
using System.Runtime.InteropServices;

namespace FastExplorer.Services
{
    /// <summary>
    /// ゴミ箱操作を提供するサービス
    /// </summary>
    public class RecycleBinService
    {
        // Shell32 COM定数
        private const int CSIDL_BITBUCKET = 0x000a; // ゴミ箱

        /// <summary>
        /// ゴミ箱からファイルを復元します
        /// </summary>
        /// <param name="originalPath">元のパス</param>
        /// <param name="deletedAt">削除時刻（オプション、最新のアイテムを検索するために使用）</param>
        /// <returns>復元に成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool RestoreFromRecycleBin(string originalPath, DateTime? deletedAt = null)
        {
            // COM方式を使用（より確実）
            return RestoreFromRecycleBinUsingCOM(originalPath, deletedAt);
        }

        /// <summary>
        /// COM方式でゴミ箱からファイルを復元します（フォールバック用）
        /// </summary>
        /// <param name="originalPath">元のパス</param>
        /// <param name="deletedAt">削除時刻（オプション、最新のアイテムを検索するために使用）</param>
        /// <returns>復元に成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool RestoreFromRecycleBinUsingCOM(string originalPath, DateTime? deletedAt = null)
        {
            if (string.IsNullOrEmpty(originalPath))
                return false;

            object? shell = null;
            object? recycleBin = null;
            object? items = null;

            try
            {
                // Shell32 COMを使用してゴミ箱から復元
                Type? shellType = Type.GetTypeFromProgID("Shell.Application");
                if (shellType == null)
                    return false;

                shell = Activator.CreateInstance(shellType);
                if (shell == null)
                    return false;

                // dynamicにキャストして使用
                dynamic dynamicShell = shell;

                // ゴミ箱フォルダーを取得
                try
                {
                    recycleBin = dynamicShell.NameSpace(CSIDL_BITBUCKET);
                }
                catch
                {
                    return false;
                }

                if (recycleBin == null)
                    return false;

                dynamic dynamicRecycleBin = recycleBin;

                // ゴミ箱内のアイテムを取得
                try
                {
                    items = dynamicRecycleBin.Items();
                }
                catch
                {
                    return false;
                }

                if (items == null)
                    return false;

                dynamic dynamicItems = items;

                // アイテムのコレクションを反復処理
                int count;
                try
                {
                    count = dynamicItems.Count;
                }
                catch
                {
                    return false;
                }

                string fileName = Path.GetFileName(originalPath);
                string? originalDirectory = Path.GetDirectoryName(originalPath);

                // まず最新の10件を検索（削除時刻が指定されている場合）
                int searchEndIndex = deletedAt.HasValue && count > 10 ? Math.Min(10, count) : count;
                
                for (int i = 0; i < searchEndIndex; i++)
                {
                    object? item = null;
                    try
                    {
                        try
                        {
                            item = dynamicItems.Item(i);
                        }
                        catch
                        {
                            continue;
                        }

                        if (item == null)
                            continue;

                        dynamic dynamicItem = item;

                        // アイテムの元のパスを取得（複数の方法を試す）
                        string? originalLocation = null;
                        try
                        {
                            // 方法1: System.Recycle.OriginalLocation
                            var extendedProperty = dynamicItem.ExtendedProperty("System.Recycle.OriginalLocation");
                            originalLocation = extendedProperty?.ToString();
                        }
                        catch
                        {
                            // ExtendedPropertyの取得に失敗した場合は、別の方法を試す
                        }

                        // 元のパスと一致するアイテムを探す
                        // ファイル名も確認（元のパスが取得できない場合に備える）
                        string? itemName = null;
                        try
                        {
                            itemName = dynamicItem.Name?.ToString();
                        }
                        catch
                        {
                            continue;
                        }

                        bool matches = false;
                        if (!string.IsNullOrEmpty(originalLocation) &&
                            string.Equals(originalLocation, originalPath, StringComparison.OrdinalIgnoreCase))
                        {
                            matches = true;
                        }
                        else if (!string.IsNullOrEmpty(itemName) &&
                                 string.Equals(itemName, fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            // OriginalLocationが取得できない場合は、ファイル名が一致していれば削除時刻に関係なく一致とみなす
                            // （OriginalLocationが取得できない場合、削除時刻の情報が不正確な可能性があるため）
                            if (string.IsNullOrEmpty(originalLocation))
                            {
                                matches = true;
                            }
                            else if (deletedAt.HasValue)
                            {
                                // OriginalLocationが取得できた場合は、削除時刻も確認
                                try
                                {
                                    // 削除日時を取得
                                    var dateDeletedProperty = dynamicItem.ExtendedProperty("System.Recycle.DateDeleted");
                                    if (dateDeletedProperty != null)
                                    {
                                        var dateDeleted = Convert.ToDateTime(dateDeletedProperty);
                                        // 削除時刻が近い場合（5分以内）は一致とみなす
                                        var timeDiff = Math.Abs((dateDeleted - deletedAt.Value).TotalMinutes);
                                        if (timeDiff <= 5)
                                        {
                                            matches = true;
                                        }
                                    }
                                    else
                                    {
                                        // 削除日時が取得できない場合は、ファイル名のみで一致とみなす
                                        matches = true;
                                    }
                                }
                                catch
                                {
                                    // 削除日時の取得に失敗した場合は、ファイル名のみで一致とみなす
                                    matches = true;
                                }
                            }
                            else
                            {
                                // 削除時刻が指定されていない場合は、ファイル名のみで一致とみなす
                                matches = true;
                            }
                        }

                        if (matches)
                        {
                            // 復元コマンドを実行
                            // まずVerbs()を使用して正確なverb名を取得してから実行
                            object? verbs = null;
                            try
                            {
                                verbs = dynamicItem.Verbs();
                            }
                            catch
                            {
                                continue;
                            }

                            if (verbs == null)
                                continue;

                            dynamic dynamicVerbs = verbs;

                            int verbCount;
                            try
                            {
                                verbCount = dynamicVerbs.Count;
                            }
                            catch
                            {
                                continue;
                            }

                            // 復元verbを探す
                            bool restoreVerbFound = false;
                            for (int j = 0; j < verbCount; j++)
                            {
                                object? verb = null;
                                try
                                {
                                    try
                                    {
                                        verb = dynamicVerbs.Item(j);
                                    }
                                    catch
                                    {
                                        continue;
                                    }

                                    if (verb == null)
                                        continue;

                                    dynamic dynamicVerb = verb;

                                    string? verbName = null;
                                    try
                                    {
                                        verbName = dynamicVerb.Name?.ToString();
                                    }
                                    catch
                                    {
                                        continue;
                                    }

                                    // "復元"または"Restore"コマンドを探す
                                    if (!string.IsNullOrEmpty(verbName))
                                    {
                                        // 日本語環境では"復元"または"元に戻す"、英語環境では"Restore"など
                                        bool isRestoreVerb = verbName.Contains("復元", StringComparison.OrdinalIgnoreCase) ||
                                                             verbName.Contains("元に戻す", StringComparison.OrdinalIgnoreCase) ||
                                                             verbName.Contains("戻す", StringComparison.OrdinalIgnoreCase) ||
                                                             verbName.Contains("Restore", StringComparison.OrdinalIgnoreCase) ||
                                                             verbName.Contains("&Restore", StringComparison.OrdinalIgnoreCase) ||
                                                             verbName.Contains("ESTORE", StringComparison.OrdinalIgnoreCase) ||
                                                             verbName.IndexOf("estore", StringComparison.OrdinalIgnoreCase) >= 0;

                                        if (isRestoreVerb)
                                        {
                                            restoreVerbFound = true;
                                            try
                                            {
                                                // DoIt()を呼び出して復元
                                                dynamicVerb.DoIt();
                                                // 少し待機して復元が完了するのを待つ
                                                System.Threading.Thread.Sleep(100);
                                                return true;
                                            }
                                            catch
                                            {
                                                // DoItの実行に失敗した場合は次のverbを試す
                                                continue;
                                            }
                                        }
                                    }
                                }
                                catch
                                {
                                    // verbの処理中にエラーが発生した場合はスキップ
                                    continue;
                                }
                            }

                            // Verbs()で見つからなかった場合、InvokeVerb("Restore")を試す
                            if (!restoreVerbFound)
                            {
                                try
                                {
                                    dynamicItem.InvokeVerb("Restore");
                                    System.Threading.Thread.Sleep(100);
                                    return true;
                                }
                                catch
                                {
                                    // InvokeVerbも失敗した場合はスキップ
                                    continue;
                                }
                            }
                        }
                    }
                    catch
                    {
                        // itemの処理中にエラーが発生した場合はスキップ
                        continue;
                    }
                }

                // 最新の10件で見つからなかった場合、残りのアイテムも検索
                if (deletedAt.HasValue && count > 10 && searchEndIndex < count)
                {
                    for (int i = searchEndIndex; i < count; i++)
                    {
                        object? item = null;
                        try
                        {
                            try
                            {
                                item = dynamicItems.Item(i);
                            }
                            catch
                            {
                                continue;
                            }

                            if (item == null)
                                continue;

                            dynamic dynamicItem = item;

                            // アイテムの元のパスを取得
                            string? originalLocation = null;
                            try
                            {
                                var extendedProperty = dynamicItem.ExtendedProperty("System.Recycle.OriginalLocation");
                                originalLocation = extendedProperty?.ToString();
                            }
                            catch
                            {
                                // ExtendedPropertyの取得に失敗した場合はスキップ
                                continue;
                            }

                            // 元のパスと一致するアイテムを探す
                            if (!string.IsNullOrEmpty(originalLocation) &&
                                string.Equals(originalLocation, originalPath, StringComparison.OrdinalIgnoreCase))
                            {
                                // 復元コマンドを実行
                                // まずVerbs()を使用して正確なverb名を取得してから実行
                                object? verbs = null;
                                try
                                {
                                    verbs = dynamicItem.Verbs();
                                }
                                catch
                                {
                                    continue;
                                }

                                if (verbs == null)
                                    continue;

                                dynamic dynamicVerbs = verbs;

                                int verbCount;
                                try
                                {
                                    verbCount = dynamicVerbs.Count;
                                }
                                catch
                                {
                                    continue;
                                }

                                // 復元verbを探す
                                for (int j = 0; j < verbCount; j++)
                                {
                                    object? verb = null;
                                    try
                                    {
                                        try
                                        {
                                            verb = dynamicVerbs.Item(j);
                                        }
                                        catch
                                        {
                                            continue;
                                        }

                                        if (verb == null)
                                            continue;

                                        dynamic dynamicVerb = verb;

                                        string? verbName = null;
                                        try
                                        {
                                            verbName = dynamicVerb.Name?.ToString();
                                        }
                                        catch
                                        {
                                            continue;
                                        }

                                        // "復元"または"Restore"コマンドを探す
                                        if (!string.IsNullOrEmpty(verbName))
                                        {
                                            bool isRestoreVerb = verbName.Contains("復元", StringComparison.OrdinalIgnoreCase) ||
                                                                 verbName.Contains("元に戻す", StringComparison.OrdinalIgnoreCase) ||
                                                                 verbName.Contains("戻す", StringComparison.OrdinalIgnoreCase) ||
                                                                 verbName.Contains("Restore", StringComparison.OrdinalIgnoreCase) ||
                                                                 verbName.Contains("&Restore", StringComparison.OrdinalIgnoreCase) ||
                                                                 verbName.Contains("ESTORE", StringComparison.OrdinalIgnoreCase) ||
                                                                 verbName.IndexOf("estore", StringComparison.OrdinalIgnoreCase) >= 0;

                                            if (isRestoreVerb)
                                            {
                                                try
                                                {
                                                    dynamicVerb.DoIt();
                                                    System.Threading.Thread.Sleep(100);
                                                    return true;
                                                }
                                                catch
                                                {
                                                    continue;
                                                }
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }

                                // Verbs()で見つからなかった場合、InvokeVerb("Restore")を試す
                                try
                                {
                                    dynamicItem.InvokeVerb("Restore");
                                    System.Threading.Thread.Sleep(100);
                                    return true;
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                        catch
                        {
                            // itemの処理中にエラーが発生した場合はスキップ
                            continue;
                        }
                    }
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                // COMオブジェクトを解放
                try
                {
                    if (items != null && Marshal.IsComObject(items))
                    {
                        Marshal.ReleaseComObject(items);
                    }
                }
                catch { }

                try
                {
                    if (recycleBin != null && Marshal.IsComObject(recycleBin))
                    {
                        Marshal.ReleaseComObject(recycleBin);
                    }
                }
                catch { }

                try
                {
                    if (shell != null && Marshal.IsComObject(shell))
                    {
                        Marshal.ReleaseComObject(shell);
                    }
                }
                catch { }
            }

            return false;
        }
    }
}

