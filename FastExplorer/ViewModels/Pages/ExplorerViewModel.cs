using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Buffers;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Cysharp.Text;
using FastExplorer.Models;
using FastExplorer.Services;
using Wpf.Ui.Abstractions.Controls;

namespace FastExplorer.ViewModels.Pages
{
    /// <summary>
    /// エクスプローラーページのViewModel
    /// </summary>
    public partial class ExplorerViewModel : ObservableObject, INavigationAware
    {
        private readonly FileSystemService _fileSystemService;
        private readonly FavoriteService? _favoriteService;
        private bool _isInitialized = false;
        private readonly Stack<string> _backHistory = new();
        private readonly Stack<string> _forwardHistory = new();
        private bool _isNavigating = false;
        private readonly List<FileSystemItem> _recentFiles = new();
        private CancellationTokenSource? _navigationCancellationTokenSource;
        
        // 非同期処理の最適化（Dispatcherのキャッシュ）
        private System.Windows.Threading.Dispatcher? _cachedDispatcher;
        
        // 文字列定数（メモリ割り当てを削減）
        private const string ShellPrefix = "shell:";

        /// <summary>
        /// 現在表示しているファイルシステムアイテムのコレクション
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<FileSystemItem> _items = new();

        /// <summary>
        /// 現在表示しているパス
        /// </summary>
        [ObservableProperty]
        private string _currentPath = string.Empty;

        /// <summary>
        /// 選択されているアイテム
        /// </summary>
        [ObservableProperty]
        private FileSystemItem? _selectedItem;

        /// <summary>
        /// 読み込み中かどうかを示す値
        /// </summary>
        [ObservableProperty]
        private bool _isLoading;

        /// <summary>
        /// ピン留めされたフォルダーのコレクション
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<FavoriteItem> _pinnedFolders = new();

        /// <summary>
        /// ドライブ情報のコレクション
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<DriveInfoModel> _drives = new();

        /// <summary>
        /// 最近使用したファイルのコレクション
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<FileSystemItem> _recentFilesList = new();

        /// <summary>
        /// ホームページを表示するかどうか
        /// </summary>
        [ObservableProperty]
        private bool _isHomePage = true;

        /// <summary>
        /// <see cref="ExplorerViewModel"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="fileSystemService">ファイルシステムサービス</param>
        /// <param name="favoriteService">お気に入りサービス</param>
        public ExplorerViewModel(FileSystemService fileSystemService, FavoriteService? favoriteService = null)
        {
            _fileSystemService = fileSystemService;
            _favoriteService = favoriteService;
        }

        /// <summary>
        /// ページにナビゲートされたときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedToAsync()
        {
            // タブのViewModelとして使用される場合は、OnNavigatedToAsyncは呼ばれない
            // そのため、初期化はCreateNewTabで行う
            if (!_isInitialized && Items.Count == 0)
            {
                InitializeViewModel();
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// ページから離れるときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedFromAsync() => Task.CompletedTask;

        /// <summary>
        /// ViewModelを初期化します
        /// </summary>
        private void InitializeViewModel()
        {
            // デフォルトでホームページを表示
            if (!_isInitialized)
            {
                NavigateToHome();
                _isInitialized = true;
            }
        }

        /// <summary>
        /// 指定されたパスにナビゲートします
        /// </summary>
        /// <param name="path">ナビゲートするパス。nullまたは空の場合はホームページを表示</param>
        [RelayCommand]
        private void NavigateToPath(string? path)
        {
            // IsNullOrWhiteSpaceでIsNullOrEmptyもカバーされるため、1回のチェックで十分
            if (string.IsNullOrWhiteSpace(path))
            {
                NavigateToHome();
                return;
            }

            var trimmedPath = path.Trim();

            // shell:で始まる特殊なパス（ごみ箱、ネットワークなど）の処理
            if (trimmedPath.StartsWith(ShellPrefix, StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = trimmedPath,
                        UseShellExecute = true
                    });
                }
                catch (Exception)
                {
                    // エラーハンドリング
                }
                return;
            }

            NavigateToDirectory(trimmedPath);
        }

        /// <summary>
        /// 親ディレクトリにナビゲートします
        /// </summary>
        [RelayCommand]
        private void NavigateToParent()
        {
            // 戻る履歴がある場合は、履歴から戻る
            if (_backHistory.Count > 0)
            {
                // 現在のパスを進む履歴に追加
                if (!string.IsNullOrEmpty(CurrentPath))
                {
                    _forwardHistory.Push(CurrentPath);
                }

                var previousPath = _backHistory.Pop();
                if (string.IsNullOrEmpty(previousPath))
                {
                    NavigateToHome(addToHistory: false);
                }
                else
                {
                    NavigateToDirectory(previousPath, addToHistory: false);
                }
                return;
            }

            // 履歴がない場合は、親ディレクトリに移動
            if (string.IsNullOrEmpty(CurrentPath))
            {
                NavigateToHome();
                return;
            }

            // 現在のパスを進む履歴に追加
            _forwardHistory.Push(CurrentPath);

            var parentPath = _fileSystemService.GetParentPath(CurrentPath);
            if (!string.IsNullOrEmpty(parentPath))
            {
                NavigateToDirectory(parentPath, addToHistory: false);
            }
            else
            {
                NavigateToDrives(addToHistory: false);
            }
        }

        /// <summary>
        /// 進む履歴の次のパスにナビゲートします
        /// </summary>
        [RelayCommand]
        private void NavigateForward()
        {
            if (_forwardHistory.Count == 0)
                return;

            // 現在のパスを戻る履歴に追加
            if (!string.IsNullOrEmpty(CurrentPath))
            {
                _backHistory.Push(CurrentPath);
            }

            var nextPath = _forwardHistory.Pop();
            if (string.IsNullOrEmpty(nextPath))
            {
                NavigateToHome(addToHistory: false);
            }
            else
            {
                NavigateToDirectory(nextPath, addToHistory: false);
            }
        }

        /// <summary>
        /// 指定されたアイテムにナビゲートします（ディレクトリの場合は開く、ファイルの場合は実行）
        /// </summary>
        /// <param name="item">ナビゲートするアイテム</param>
        [RelayCommand]
        private void NavigateToItem(FileSystemItem? item)
        {
            if (item == null)
                return;

            if (item.IsDirectory)
            {
                NavigateToDirectory(item.FullPath);
            }
            else
            {
                // ファイルの場合は開く
                try
                {
                    // 最近使用したファイルに追加
                    AddToRecentFiles(item.FullPath);

                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = item.FullPath,
                        UseShellExecute = true
                    });
                }
                catch (Exception)
                {
                    // エラーハンドリング
                }
            }
        }

        /// <summary>
        /// 現在のディレクトリを更新します
        /// </summary>
        [RelayCommand]
        private void Refresh()
        {
            if (string.IsNullOrEmpty(CurrentPath))
            {
                NavigateToHome();
            }
            else
            {
                NavigateToDirectory(CurrentPath);
            }
        }

        /// <summary>
        /// 指定されたディレクトリにナビゲートします
        /// </summary>
        /// <param name="path">ナビゲートするディレクトリのパス</param>
        /// <param name="addToHistory">履歴に追加するかどうか</param>
        public void NavigateToDirectory(string path, bool addToHistory = true)
        {
            if (!_fileSystemService.IsValidPath(path) || !Directory.Exists(path))
                return;

            if (_isNavigating)
                return;

            // 前のナビゲーションをキャンセル
            _navigationCancellationTokenSource?.Cancel();
            _navigationCancellationTokenSource?.Dispose();
            _navigationCancellationTokenSource = new CancellationTokenSource();
            var cancellationToken = _navigationCancellationTokenSource.Token;

            _isNavigating = true;

            // 履歴に追加する場合、現在のパスを戻る履歴に追加し、進む履歴をクリア
            if (addToHistory && !string.IsNullOrEmpty(CurrentPath) && CurrentPath != path)
            {
                _backHistory.Push(CurrentPath);
                _forwardHistory.Clear();
            }

            // プロパティ変更を最適化（値が実際に変更された場合のみ通知）
            if (CurrentPath != path)
            {
                CurrentPath = path;
            }
            if (IsHomePage)
            {
                IsHomePage = false;
            }
            Items.Clear();
            IsLoading = true;

            // 非同期でファイル一覧を読み込み（UIスレッドをブロックしない）
            _ = Task.Run(async () =>
            {
                try
                {
                    // ToList()を削減：IEnumerableを直接使用してメモリ割り当てを削減
                    var fileItems = _fileSystemService.GetItems(path);
                    
                    // キャンセルされた場合は処理を中断
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    
                    // アイテム数をカウント（ToList()を避けるため、一度だけ反復）
                    // リストの容量を事前に確保（平均的なディレクトリサイズを想定）
                    var itemList = new List<FileSystemItem>(256);
                    foreach (var item in fileItems)
                    {
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        itemList.Add(item);
                    }
                    var itemCount = itemList.Count;
                    
                    // キャンセルされた場合は処理を中断
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    // UIスレッドでバッチ更新（200件ずつ追加してUIの応答性を保ちつつ、更新回数を削減）
                    const int batchSize = 200;
                    // Dispatcherをキャッシュ（パフォーマンス向上）
                    if (_cachedDispatcher == null)
                    {
                        _cachedDispatcher = System.Windows.Application.Current?.Dispatcher;
                    }
                    var dispatcher = _cachedDispatcher;
                    if (dispatcher == null)
                        return;

                    // 大量のアイテムがある場合は、一度に追加してパフォーマンスを向上
                    if (itemCount <= batchSize)
                    {
                        // 少量の場合は一度に追加
                        await dispatcher.InvokeAsync(() =>
                        {
                            if (!cancellationToken.IsCancellationRequested)
                            {
                                // コレクションの容量を事前に確保してメモリ再割り当てを削減
                                // ObservableCollectionは内部でListを使用しているため、容量を事前に確保できないが、
                                // 個別のAdd()呼び出しを最適化（ループ展開で小さいリストを高速化）
                                if (itemCount <= 4)
                                {
                                    // 小さいリストの場合はループ展開（条件分岐を削減）
                                    if (itemCount > 0) Items.Add(itemList[0]);
                                    if (itemCount > 1) Items.Add(itemList[1]);
                                    if (itemCount > 2) Items.Add(itemList[2]);
                                    if (itemCount > 3) Items.Add(itemList[3]);
                                }
                                else
                                {
                                    // 大きいリストの場合は通常のループ
                                    foreach (var item in itemList)
                                    {
                                        Items.Add(item);
                                    }
                                }
                            }
                        });
                    }
                    else
                    {
                        // 大量の場合はバッチ更新（Skip/Takeを削減して直接インデックスアクセス）
                        // 中間リストを作成せず、直接インデックスアクセスでメモリ割り当てを削減
                        // itemList.Countを一度だけ取得してキャッシュ（パフォーマンス向上）
                        var totalCount = itemList.Count;
                        for (int i = 0; i < totalCount; i += batchSize)
                        {
                            if (cancellationToken.IsCancellationRequested)
                                return;

                            var endIndex = Math.Min(i + batchSize, totalCount);
                            // バッチの開始インデックスと終了インデックスをキャプチャ
                            var startIdx = i;
                            var endIdx = endIndex;
                            
                            // バッチサイズを事前計算（ループ内での計算を削減）
                            var batchItemCount = endIdx - startIdx;
                            
                            await dispatcher.InvokeAsync(() =>
                            {
                                if (!cancellationToken.IsCancellationRequested)
                                {
                                    // 直接インデックスアクセスでメモリ割り当てを削減
                                    // ループ展開の最適化（小さいバッチの場合は展開）
                                    if (batchItemCount <= 4)
                                    {
                                        // 小さいバッチの場合は展開（条件分岐を削減）
                                        if (startIdx < itemList.Count) Items.Add(itemList[startIdx]);
                                        if (startIdx + 1 < itemList.Count && startIdx + 1 < endIdx) Items.Add(itemList[startIdx + 1]);
                                        if (startIdx + 2 < itemList.Count && startIdx + 2 < endIdx) Items.Add(itemList[startIdx + 2]);
                                        if (startIdx + 3 < itemList.Count && startIdx + 3 < endIdx) Items.Add(itemList[startIdx + 3]);
                                    }
                                    else
                                    {
                                        // 大きいバッチの場合は通常のループ
                                        for (int j = startIdx; j < endIdx; j++)
                                        {
                                            Items.Add(itemList[j]);
                                        }
                                    }
                                }
                            }, System.Windows.Threading.DispatcherPriority.Background);
                        }
                    }
                }
                catch (Exception)
                {
                    // エラーハンドリング
                }
                finally
                {
                    if (!cancellationToken.IsCancellationRequested)
                    {
                        // Dispatcherをキャッシュから使用
                        var dispatcher = _cachedDispatcher ?? System.Windows.Application.Current?.Dispatcher;
                        dispatcher?.Invoke(() =>
                        {
                            IsLoading = false;
                            _isNavigating = false;
                        }, System.Windows.Threading.DispatcherPriority.Normal);
                    }
                }
            }, cancellationToken);
        }

        /// <summary>
        /// ホームページを表示します
        /// </summary>
        /// <param name="addToHistory">履歴に追加するかどうか</param>
        public void NavigateToHome(bool addToHistory = true)
        {
            if (_isNavigating)
                return;

            _isNavigating = true;

            // 履歴に追加する場合、現在のパスを戻る履歴に追加し、進む履歴をクリア
            // ホームボタンを押したときもブラウザーバックで戻れるように、現在のパスが空でない場合のみ履歴に追加
            if (addToHistory && !string.IsNullOrEmpty(CurrentPath))
            {
                _backHistory.Push(CurrentPath);
                _forwardHistory.Clear();
            }

            // プロパティ変更を最適化（値が実際に変更された場合のみ通知）
            if (CurrentPath != string.Empty)
            {
                CurrentPath = string.Empty;
            }
            if (!IsHomePage)
            {
                IsHomePage = true;
            }
            Items.Clear();
            IsLoading = true;

            // ピン留めフォルダーを読み込み（同期的、軽量）
            LoadPinnedFolders();

            // 最近使用したファイルを読み込み（同期的、軽量）
            LoadRecentFiles();

            // ドライブ情報を非同期で読み込み（UIスレッドをブロックしない）
            _ = Task.Run(async () =>
            {
                try
                {
                    await LoadDrivesAsync();
                }
                catch (Exception)
                {
                    // エラーハンドリング
                }
                finally
                {
                    // Dispatcherをキャッシュから使用
                    var dispatcher = _cachedDispatcher ?? System.Windows.Application.Current?.Dispatcher;
                    dispatcher?.Invoke(() =>
                    {
                        IsLoading = false;
                        _isNavigating = false;
                    }, System.Windows.Threading.DispatcherPriority.Normal);
                }
            });
        }

        /// <summary>
        /// ピン留めフォルダーを読み込みます
        /// </summary>
        private void LoadPinnedFolders()
        {
            PinnedFolders.Clear();
            if (_favoriteService == null)
                return;

            var favorites = _favoriteService.GetFavorites();
            // 標準的なWindowsフォルダ名を定数化（メモリ割り当てを削減）
            const string Desktop = "デスクトップ";
            const string Downloads = "ダウンロード";
            const string Documents = "ドキュメント";
            const string Pictures = "ピクチャ";
            const string Music = "ミュージック";
            const string Videos = "ビデオ";
            const string RecycleBin = "ごみ箱";
            
            foreach (var favorite in favorites)
            {
                // 標準的なWindowsフォルダのみをピン留めとして表示
                // ReadOnlySpan<char>を使用してメモリ割り当てを削減（高速化）
                var name = favorite.Name;
                var nameSpan = name.AsSpan();
                if (nameSpan.SequenceEqual(Desktop.AsSpan()) ||
                    nameSpan.SequenceEqual(Downloads.AsSpan()) ||
                    nameSpan.SequenceEqual(Documents.AsSpan()) ||
                    nameSpan.SequenceEqual(Pictures.AsSpan()) ||
                    nameSpan.SequenceEqual(Music.AsSpan()) ||
                    nameSpan.SequenceEqual(Videos.AsSpan()) ||
                    nameSpan.SequenceEqual(RecycleBin.AsSpan()))
                {
                    PinnedFolders.Add(favorite);
                }
            }
        }

        /// <summary>
        /// ドライブ情報を非同期で読み込みます（UIスレッドをブロックしない）
        /// </summary>
        private async Task LoadDrivesAsync()
        {
            // Dispatcherをキャッシュから使用
            var dispatcher = _cachedDispatcher ?? System.Windows.Application.Current?.Dispatcher;
            if (dispatcher == null)
                return;

            // UIスレッドでクリア
            await dispatcher.InvokeAsync(() => Drives.Clear());

            try
            {
                var drives = _fileSystemService.GetDrives();
                // リストの容量を事前に確保（通常のドライブ数は10個以下）
                var driveModels = new List<DriveInfoModel>(drives.Length);

                // バックグラウンドスレッドでドライブ情報を取得
                foreach (var drive in drives)
                {
                    try
                    {
                        // ドライブ情報を非同期で取得（ネットワークドライブの応答待ちでUIスレッドをブロックしない）
                        var driveModel = await Task.Run(() =>
                        {
                            try
                            {
                                var driveInfo = new DriveInfo(drive);
                                // IsReadyプロパティの取得は同期的だが、ネットワークドライブの場合は時間がかかる可能性がある
                                // そのため、Task.Run内で実行してUIスレッドをブロックしない
                                if (driveInfo.IsReady)
                                {
                                    // 文字列操作を最適化（ZString.Concatを使用してボクシングを回避）
                                    var driveLetter = drive.TrimEnd('\\');
                                    var driveName = string.IsNullOrEmpty(driveInfo.VolumeLabel)
                                        ? ZString.Concat("ローカルディスク (", driveLetter, ")")
                                        : ZString.Concat(driveInfo.VolumeLabel, " (", driveLetter, ")");
                                    
                                    return new DriveInfoModel
                                    {
                                        Name = driveName,
                                        Path = driveInfo.RootDirectory.FullName,
                                        VolumeLabel = driveInfo.VolumeLabel,
                                        TotalSize = driveInfo.TotalSize,
                                        FreeSpace = driveInfo.AvailableFreeSpace
                                    };
                                }
                                return null;
                            }
                            catch
                            {
                                return null;
                            }
                        });

                        if (driveModel != null)
                        {
                            driveModels.Add(driveModel);
                        }
                    }
                    catch
                    {
                        // ドライブにアクセスできない場合はスキップ
                        continue;
                    }
                }

                // UIスレッドで一括追加
                await dispatcher.InvokeAsync(() =>
                {
                    foreach (var model in driveModels)
                    {
                        Drives.Add(model);
                    }
                });
            }
            catch (Exception)
            {
                // エラーハンドリング
            }
        }

        /// <summary>
        /// 最近使用したファイルを読み込みます
        /// </summary>
        private void LoadRecentFiles()
        {
            RecentFilesList.Clear();
            // 簡易実装：最近アクセスしたファイルを保持
            // 実際の実装では、Windowsのジャンプリストやファイルアクセス履歴を使用
            // LINQのTake()を直接ループに置き換え（メモリ割り当てを削減）
            var count = Math.Min(_recentFiles.Count, 10);
            for (int i = 0; i < count; i++)
            {
                var file = _recentFiles[i];
                if (File.Exists(file.FullPath) || Directory.Exists(file.FullPath))
                {
                    RecentFilesList.Add(file);
                }
            }
        }

        /// <summary>
        /// ファイルを最近使用したファイルリストに追加します
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        public void AddToRecentFiles(string filePath)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                var item = new FileSystemItem
                {
                    Name = fileInfo.Name,
                    FullPath = fileInfo.FullName,
                    Extension = fileInfo.Extension,
                    Size = fileInfo.Length,
                    LastModified = fileInfo.LastWriteTime,
                    IsDirectory = false,
                    Attributes = fileInfo.Attributes
                };

                // 既に存在する場合は削除
                // RemoveAllのラムダ式を直接ループに置き換え（メモリ割り当てを削減）
                // 文字列比較を最適化（ReadOnlySpanを使用してメモリ割り当てを削減）
                var filePathSpan = filePath.AsSpan();
                for (int i = _recentFiles.Count - 1; i >= 0; i--)
                {
                    var fullPathSpan = _recentFiles[i].FullPath.AsSpan();
                    if (fullPathSpan.Length == filePathSpan.Length && 
                        fullPathSpan.CompareTo(filePathSpan, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        _recentFiles.RemoveAt(i);
                    }
                }
                // 先頭に追加
                _recentFiles.Insert(0, item);
                // 最大20件まで保持
                // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                var recentFilesCount = _recentFiles.Count;
                if (recentFilesCount > 20)
                {
                    _recentFiles.RemoveRange(20, recentFilesCount - 20);
                }

                // ホームページ表示中の場合は更新
                if (IsHomePage)
                {
                    LoadRecentFiles();
                }
            }
            catch
            {
                // エラーハンドリング
            }
        }

        /// <summary>
        /// ドライブ一覧を表示します（後方互換性のため残す）
        /// </summary>
        /// <param name="addToHistory">履歴に追加するかどうか</param>
        public void NavigateToDrives(bool addToHistory = true)
        {
            NavigateToHome(addToHistory);
        }
    }
}

