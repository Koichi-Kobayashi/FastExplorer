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
            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] RestoreFromRecycleBin開始: {originalPath}, DeletedAt: {deletedAt}");
            // COM方式を使用（より確実）
            var result = RestoreFromRecycleBinUsingCOM(originalPath, deletedAt);
            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] RestoreFromRecycleBin結果: {result}");
            return result;
        }

        /// <summary>
        /// COM方式でゴミ箱からファイルを復元します（フォールバック用）
        /// </summary>
        /// <param name="originalPath">元のパス</param>
        /// <param name="deletedAt">削除時刻（オプション、最新のアイテムを検索するために使用）</param>
        /// <returns>復元に成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool RestoreFromRecycleBinUsingCOM(string originalPath, DateTime? deletedAt = null)
        {
            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] RestoreFromRecycleBinUsingCOM開始: {originalPath}, DeletedAt: {deletedAt}");
            if (string.IsNullOrEmpty(originalPath))
            {
                System.Diagnostics.Debug.WriteLine("[RecycleBinService] originalPathがnullまたは空です");
                return false;
            }

            object? shell = null;
            object? recycleBin = null;
            object? items = null;

            try
            {
                // Shell32 COMを使用してゴミ箱から復元
                System.Diagnostics.Debug.WriteLine("[RecycleBinService] Shell.ApplicationのProgIDを取得します");
                Type? shellType = Type.GetTypeFromProgID("Shell.Application");
                if (shellType == null)
                {
                    System.Diagnostics.Debug.WriteLine("[RecycleBinService] Shell.ApplicationのProgIDが見つかりませんでした");
                    return false;
                }

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
                    System.Diagnostics.Debug.WriteLine($"[RecycleBinService] ゴミ箱内のアイテム数: {count}");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[RecycleBinService] アイテム数の取得に失敗しました");
                    return false;
                }

                string fileName = Path.GetFileName(originalPath);
                string? originalDirectory = Path.GetDirectoryName(originalPath);
                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] 検索対象: ファイル名={fileName}, ディレクトリ={originalDirectory}");

                // まず最新の10件を検索（削除時刻が指定されている場合）
                int searchEndIndex = deletedAt.HasValue && count > 10 ? Math.Min(10, count) : count;
                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] 検索範囲: 0-{searchEndIndex} (全{count}件)");
                
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
                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] OriginalLocation: {originalLocation ?? "(null)"}");
                        }
                        catch (Exception ex)
                        {
                            // ExtendedPropertyの取得に失敗した場合は、別の方法を試す
                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] OriginalLocationの取得に失敗: {ex.Message}");
                        }

                        // 元のパスと一致するアイテムを探す
                        // ファイル名も確認（元のパスが取得できない場合に備える）
                        string? itemName = null;
                        try
                        {
                            itemName = dynamicItem.Name?.ToString();
                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] Name: {itemName ?? "(null)"}");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] Nameの取得に失敗: {ex.Message}");
                            continue;
                        }

                        bool matches = false;
                        if (!string.IsNullOrEmpty(originalLocation) &&
                            string.Equals(originalLocation, originalPath, StringComparison.OrdinalIgnoreCase))
                        {
                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] 元のパスが一致しました: {originalLocation}");
                            matches = true;
                        }
                        else if (!string.IsNullOrEmpty(itemName) &&
                                 string.Equals(itemName, fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] ファイル名が一致しました: {itemName}");
                            
                            // OriginalLocationが取得できない場合は、ファイル名が一致していれば削除時刻に関係なく一致とみなす
                            // （OriginalLocationが取得できない場合、削除時刻の情報が不正確な可能性があるため）
                            if (string.IsNullOrEmpty(originalLocation))
                            {
                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] OriginalLocationが取得できないため、ファイル名のみで一致とみなします");
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
                                        System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] DateDeleted: {dateDeleted}, 削除時刻: {deletedAt.Value}");
                                        // 削除時刻が近い場合（5分以内）は一致とみなす
                                        var timeDiff = Math.Abs((dateDeleted - deletedAt.Value).TotalMinutes);
                                        System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] 時間差: {timeDiff}分");
                                        if (timeDiff <= 5)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] 削除時刻が近いため一致とみなします");
                                            matches = true;
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] 削除時刻が遠いため一致しません");
                                        }
                                    }
                                    else
                                    {
                                        // 削除日時が取得できない場合は、ファイル名のみで一致とみなす
                                        System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] DateDeletedが取得できないため、ファイル名のみで一致とみなします");
                                        matches = true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // 削除日時の取得に失敗した場合は、ファイル名のみで一致とみなす
                                    System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] DateDeletedの取得に失敗したため、ファイル名のみで一致とみなします: {ex.Message}");
                                    matches = true;
                                }
                            }
                            else
                            {
                                // 削除時刻が指定されていない場合は、ファイル名のみで一致とみなす
                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] 削除時刻が指定されていないため、ファイル名のみで一致とみなします");
                                matches = true;
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] 一致しませんでした。originalLocation={originalLocation ?? "(null)"}, itemName={itemName ?? "(null)"}, fileName={fileName}");
                        }

                        if (matches)
                        {
                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] 一致するアイテムが見つかりました: ファイル名={itemName}, 元のパス={originalLocation}");
                            // 復元コマンドを実行
                            // まずVerbs()を使用して正確なverb名を取得してから実行
                            object? verbs = null;
                            try
                            {
                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] Verbs()を取得します");
                                verbs = dynamicItem.Verbs();
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] Verbs()の取得に失敗しました: {ex.Message}");
                                continue;
                            }

                            if (verbs == null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] verbsがnullです");
                                continue;
                            }

                            dynamic dynamicVerbs = verbs;

                            int verbCount;
                            try
                            {
                                verbCount = dynamicVerbs.Count;
                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] verb数: {verbCount}");
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] verb数の取得に失敗しました: {ex.Message}");
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
                                        System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] verb[{j}]: {verbName}");
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
                                            System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] 復元verbが見つかりました: {verbName}");
                                            restoreVerbFound = true;
                                            try
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] DoIt()を呼び出します: {verbName}");
                                                // DoIt()を呼び出して復元
                                                dynamicVerb.DoIt();
                                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] DoIt()の呼び出しが完了しました");
                                                // 少し待機して復元が完了するのを待つ
                                                System.Threading.Thread.Sleep(500);
                                                System.Diagnostics.Debug.WriteLine("[RecycleBinService] 復元が成功しました");
                                                return true;
                                            }
                                            catch (Exception ex)
                                            {
                                                // DoItの実行に失敗した場合は次のverbを試す
                                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] DoIt()の実行に失敗しました: {ex.Message}");
                                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] スタックトレース: {ex.StackTrace}");
                                                continue;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] verb[{j}]の名前がnullまたは空です");
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
                                System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] 復元verbが見つからなかったため、InvokeVerb(\"Restore\")を試します");
                                try
                                {
                                    dynamicItem.InvokeVerb("Restore");
                                    System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] InvokeVerb(\"Restore\")の呼び出しが完了しました");
                                    System.Threading.Thread.Sleep(500);
                                    System.Diagnostics.Debug.WriteLine("[RecycleBinService] InvokeVerbで復元が成功しました");
                                    return true;
                                }
                                catch (Exception ex)
                                {
                                    // InvokeVerbも失敗した場合はスキップ
                                    System.Diagnostics.Debug.WriteLine($"[RecycleBinService] アイテム[{i}] InvokeVerb(\"Restore\")の実行に失敗しました: {ex.Message}");
                                    System.Diagnostics.Debug.WriteLine($"[RecycleBinService] スタックトレース: {ex.StackTrace}");
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

