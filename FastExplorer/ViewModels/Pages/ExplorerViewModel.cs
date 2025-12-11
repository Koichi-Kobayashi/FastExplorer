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
        #region フィールド

        private readonly FileSystemService _fileSystemService;
        private readonly FavoriteService? _favoriteService;
        private bool _isInitialized = false;
        private readonly Stack<string> _backHistory = new();
        private readonly Stack<string> _forwardHistory = new();
        private bool _isNavigating = false;
        private readonly List<FileSystemItem> _recentFiles = new();
        private readonly List<FavoriteItem> _recentFolders = new();
        private CancellationTokenSource? _navigationCancellationTokenSource;
        
        // 非同期処理の最適化（Dispatcherのキャッシュ）
        private System.Windows.Threading.Dispatcher? _cachedDispatcher;
        
        // 文字列定数（メモリ割り当てを削減）
        private const string ShellPrefix = "shell:";
        
        // FileSystemWatcher関連
        private FileSystemWatcher? _fileSystemWatcher;
        private System.Windows.Threading.DispatcherTimer? _refreshTimer;
        private bool _refreshPending = false;

        #endregion

        #region プロパティ

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
        /// CurrentPathが変更されたときに呼び出されます
        /// </summary>
        partial void OnCurrentPathChanged(string value)
        {
            // FileSystemWatcherを更新
            UpdateFileSystemWatcher(value);
        }

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
        /// 最近使用したフォルダーのコレクション
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<FavoriteItem> _recentFoldersList = new();

        /// <summary>
        /// ホームページを表示するかどうか
        /// </summary>
        [ObservableProperty]
        private bool _isHomePage = true;

        /// <summary>
        /// ステータスバーに表示するテキスト
        /// </summary>
        [ObservableProperty]
        private string _statusBarText = "準備完了";

        /// <summary>
        /// 現在のソート列（"Name", "Size", "Extension", "LastModified"）
        /// </summary>
        private string _sortColumn = "Name";

        /// <summary>
        /// ソート方向（true: 昇順、false: 降順）
        /// </summary>
        private bool _sortAscending = true;

        /// <summary>
        /// ListViewのフォントサイズ（デフォルト: 14）
        /// </summary>
        [ObservableProperty]
        private double _listViewFontSize = 14.0;

        /// <summary>
        /// ListViewの行の高さ（デフォルト: 32）
        /// </summary>
        [ObservableProperty]
        private double _listViewItemHeight = 32.0;

        #endregion

        #region コンストラクタ

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

        #endregion

        #region ナビゲーション

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
        public Task OnNavigatedFromAsync()
        {
            // FileSystemWatcherをクリーンアップ
            CleanupFileSystemWatcher();
            return Task.CompletedTask;
        }

        #endregion

        #region 初期化

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

        #endregion

        #region パスナビゲーション

        /// <summary>
        /// 指定されたパスにナビゲートします
        /// </summary>
        /// <param name="path">ナビゲートするパス。nullまたは空の場合はホームページを表示</param>
        [RelayCommand]
        private void NavigateToPath(string? path)
        {
            System.Diagnostics.Debug.WriteLine($"[ExplorerViewModel.NavigateToPath] 呼び出されました: path='{path}'");
            
            // IsNullOrWhiteSpaceでIsNullOrEmptyもカバーされるため、1回のチェックで十分
            if (string.IsNullOrWhiteSpace(path))
            {
                System.Diagnostics.Debug.WriteLine($"[ExplorerViewModel.NavigateToPath] パスが空のため、NavigateToHome()を呼び出します");
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

        #endregion

        #region 履歴ナビゲーション

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

        #endregion

        #region アイテムナビゲーション

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

        #endregion

        #region ディレクトリ操作

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

            // 最近使用したフォルダーに追加（ホームページ以外の場合）
            if (!string.IsNullOrEmpty(path) && path != CurrentPath)
            {
                AddToRecentFolders(path);
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

                    // アイテムをソート（バックグラウンドスレッドで実行）
                    SortItemList(itemList);

                    // UIスレッドでバッチ更新（500件ずつ追加してUIの応答性を保ちつつ、更新回数を削減）
                    // バッチサイズを200から500に増やすことで、Dispatcher呼び出し回数を削減（高速化）
                    const int batchSize = 500;
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
                                    // プロパティアクセスをキャッシュして高速化
                                    var items = Items;
                                    foreach (var item in itemList)
                                    {
                                        items.Add(item);
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
                                    // endIdxは既にMath.Minで計算されているため、endIdx <= itemList.Countが保証されている
                                    // そのため、itemList.Countのチェックは不要（条件分岐を削減して高速化）
                                    if (batchItemCount <= 4)
                                    {
                                        // 小さいバッチの場合は展開（条件分岐を削減）
                                        if (startIdx < endIdx) Items.Add(itemList[startIdx]);
                                        if (startIdx + 1 < endIdx) Items.Add(itemList[startIdx + 1]);
                                        if (startIdx + 2 < endIdx) Items.Add(itemList[startIdx + 2]);
                                        if (startIdx + 3 < endIdx) Items.Add(itemList[startIdx + 3]);
                                    }
                                    else
                                    {
                                        // 大きいバッチの場合は通常のループ
                                        // ループ変数をキャッシュしてプロパティアクセスを削減
                                        var items = Items;
                                        for (int j = startIdx; j < endIdx; j++)
                                        {
                                            items.Add(itemList[j]);
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

        #endregion

        #region ホームページ

        /// <summary>
        /// ホームページを表示します
        /// </summary>
        /// <param name="addToHistory">履歴に追加するかどうか</param>
        public void NavigateToHome(bool addToHistory = true)
        {
            System.Diagnostics.Debug.WriteLine($"[ExplorerViewModel.NavigateToHome] 呼び出されました: addToHistory={addToHistory}, _isNavigating={_isNavigating}, CurrentPath='{CurrentPath}', IsHomePage={IsHomePage}");
            
            if (_isNavigating)
            {
                System.Diagnostics.Debug.WriteLine($"[ExplorerViewModel.NavigateToHome] _isNavigatingがtrueのため、早期リターンします");
                return;
            }

            _isNavigating = true;
            System.Diagnostics.Debug.WriteLine($"[ExplorerViewModel.NavigateToHome] _isNavigatingをtrueに設定しました");

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
            // IsHomePageを常に設定して、PropertyChangedイベントを確実に発火させる
            // （既にtrueの場合でも、UIの更新を確実にするため）
            IsHomePage = true;
            // PropertyChangedイベントを明示的に発火させて、UIの更新を確実にする
            OnPropertyChanged(nameof(IsHomePage));
            System.Diagnostics.Debug.WriteLine($"[ExplorerViewModel.NavigateToHome] IsHomePageをtrueに設定し、PropertyChangedイベントを発火しました");
            Items.Clear();
            IsLoading = true;
            
            // FileSystemWatcherを停止（ホームページでは監視しない）
            CleanupFileSystemWatcher();

            // ピン留めフォルダーを読み込み（同期的、軽量）
            LoadPinnedFolders();

            // 最近使用したファイルを読み込み（同期的、軽量）
            LoadRecentFiles();

            // 最近使用したフォルダーを読み込み（同期的、軽量）
            LoadRecentFolders();

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
                        System.Diagnostics.Debug.WriteLine($"[ExplorerViewModel.NavigateToHome] 非同期処理完了: IsLoading=false, _isNavigating=false, IsHomePage={IsHomePage}");
                    }, System.Windows.Threading.DispatcherPriority.Normal);
                }
            });
        }

        #endregion

        #region ホームページデータ読み込み

        /// <summary>
        /// ピン留めフォルダーを読み込みます
        /// </summary>
        private void LoadPinnedFolders()
        {
            var pinnedFolders = PinnedFolders;
            pinnedFolders.Clear();
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
            
            // ReadOnlySpanを事前に作成（ループ内での作成を削減）
            var desktopSpan = Desktop.AsSpan();
            var downloadsSpan = Downloads.AsSpan();
            var documentsSpan = Documents.AsSpan();
            var picturesSpan = Pictures.AsSpan();
            var musicSpan = Music.AsSpan();
            var videosSpan = Videos.AsSpan();
            var recycleBinSpan = RecycleBin.AsSpan();
            
            foreach (var favorite in favorites)
            {
                // 標準的なWindowsフォルダのみをピン留めとして表示
                // ReadOnlySpan<char>を使用してメモリ割り当てを削減（高速化）
                var name = favorite.Name;
                var nameSpan = name.AsSpan();
                if (nameSpan.SequenceEqual(desktopSpan) ||
                    nameSpan.SequenceEqual(downloadsSpan) ||
                    nameSpan.SequenceEqual(documentsSpan) ||
                    nameSpan.SequenceEqual(picturesSpan) ||
                    nameSpan.SequenceEqual(musicSpan) ||
                    nameSpan.SequenceEqual(videosSpan) ||
                    nameSpan.SequenceEqual(recycleBinSpan))
                {
                    pinnedFolders.Add(favorite);
                }
            }
        }

        /// <summary>
        /// 最近使用したフォルダーを読み込みます
        /// </summary>
        private void LoadRecentFolders()
        {
            var recentFoldersList = RecentFoldersList;
            recentFoldersList.Clear();
            // 最大20件まで保持、表示は最大10件
            var count = Math.Min(_recentFolders.Count, 10);
            // リストの容量を事前に確保
            var validFolders = new List<FavoriteItem>(count);
            for (int i = 0; i < count; i++)
            {
                var folder = _recentFolders[i];
                var path = folder.Path;
                // フォルダーが存在する場合のみ追加
                if (!string.IsNullOrEmpty(path))
                {
                    // shell:で始まるパスは常に有効
                    if (path.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
                    {
                        validFolders.Add(folder);
                    }
                    else if (Directory.Exists(path))
                    {
                        validFolders.Add(folder);
                    }
                }
            }
            // 一括追加（メモリ割り当てを削減）
            foreach (var folder in validFolders)
            {
                recentFoldersList.Add(folder);
            }
        }

        /// <summary>
        /// フォルダーを最近使用したフォルダーリストに追加します
        /// </summary>
        /// <param name="folderPath">フォルダーパス</param>
        private void AddToRecentFolders(string folderPath)
        {
            try
            {
                // ホームページの場合は追加しない
                if (string.IsNullOrEmpty(folderPath))
                    return;

                var folderName = Path.GetFileName(folderPath);
                if (string.IsNullOrEmpty(folderName))
                {
                    // ルートディレクトリの場合は、ドライブ名を使用
                    var root = Path.GetPathRoot(folderPath);
                    if (!string.IsNullOrEmpty(root))
                    {
                        folderName = root.TrimEnd('\\');
                    }
                    else
                    {
                        folderName = folderPath;
                    }
                }

                var favorite = new FavoriteItem
                {
                    Name = folderName,
                    Path = folderPath
                };

                // 既に存在する場合は削除
                var folderPathSpan = folderPath.AsSpan();
                var folderPathLength = folderPathSpan.Length;
                var recentFoldersCount = _recentFolders.Count;
                for (int i = recentFoldersCount - 1; i >= 0; i--)
                {
                    var existingPath = _recentFolders[i].Path;
                    var existingPathSpan = existingPath.AsSpan();
                    // 長さチェックを先に行う（高速化）
                    if (existingPathSpan.Length == folderPathLength &&
                        existingPathSpan.CompareTo(folderPathSpan, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        _recentFolders.RemoveAt(i);
                        break; // 一度見つかったら終了
                    }
                }
                // 先頭に追加
                _recentFolders.Insert(0, favorite);
                // 最大20件まで保持
                var currentRecentFoldersCount = _recentFolders.Count;
                if (currentRecentFoldersCount > 20)
                {
                    _recentFolders.RemoveRange(20, currentRecentFoldersCount - 20);
                }

                // ホームページ表示中の場合は更新
                if (IsHomePage)
                {
                    LoadRecentFolders();
                }
            }
            catch
            {
                // エラーハンドリング
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
                                    // プロパティアクセスをキャッシュ（高速化）
                                    var volumeLabel = driveInfo.VolumeLabel;
                                    var rootDirectory = driveInfo.RootDirectory;
                                    var totalSize = driveInfo.TotalSize;
                                    var freeSpace = driveInfo.AvailableFreeSpace;
                                    
                                    // 文字列操作を最適化（ZString.Concatを使用してボクシングを回避）
                                    var driveLetter = drive.TrimEnd('\\');
                                    var driveName = string.IsNullOrEmpty(volumeLabel)
                                        ? ZString.Concat("ローカルディスク (", driveLetter, ")")
                                        : ZString.Concat(volumeLabel, " (", driveLetter, ")");
                                    
                                    return new DriveInfoModel
                                    {
                                        Name = driveName,
                                        Path = rootDirectory.FullName,
                                        VolumeLabel = volumeLabel,
                                        TotalSize = totalSize,
                                        FreeSpace = freeSpace
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
                // プロパティアクセスをキャッシュ（高速化）
                await dispatcher.InvokeAsync(() =>
                {
                    var drives = Drives;
                    foreach (var model in driveModels)
                    {
                        drives.Add(model);
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
            var recentFilesList = RecentFilesList;
            recentFilesList.Clear();
            // 簡易実装：最近アクセスしたファイルを保持
            // 実際の実装では、Windowsのジャンプリストやファイルアクセス履歴を使用
            // LINQのTake()を直接ループに置き換え（メモリ割り当てを削減）
            var recentFilesCount = _recentFiles.Count;
            var count = Math.Min(recentFilesCount, 10);
            // リストの容量を事前に確保（最大10件）
            var validFiles = new List<FileSystemItem>(count);
            for (int i = 0; i < count; i++)
            {
                var file = _recentFiles[i];
                var fullPath = file.FullPath;
                // ファイルシステムアクセスを最適化（IsDirectoryで分岐）
                if (file.IsDirectory)
                {
                    if (Directory.Exists(fullPath))
                    {
                        validFiles.Add(file);
                    }
                }
                else
                {
                    if (File.Exists(fullPath))
                    {
                        validFiles.Add(file);
                    }
                }
            }
            // 一括追加（メモリ割り当てを削減）
            foreach (var file in validFiles)
            {
                recentFilesList.Add(file);
            }
        }

        #endregion

        #region 最近使用したファイル

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
                var filePathLength = filePathSpan.Length;
                // 最適化：一度見つかったら削除してbreak（通常は1件しか存在しないため）
                var recentFilesCount = _recentFiles.Count;
                for (int i = recentFilesCount - 1; i >= 0; i--)
                {
                    var fullPath = _recentFiles[i].FullPath;
                    var fullPathSpan = fullPath.AsSpan();
                    // 長さチェックを先に行う（高速化）
                    if (fullPathSpan.Length == filePathLength && 
                        fullPathSpan.CompareTo(filePathSpan, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        _recentFiles.RemoveAt(i);
                        break; // 一度見つかったら終了（通常は1件しか存在しないため）
                    }
                }
                // 先頭に追加
                _recentFiles.Insert(0, item);
                // 最大20件まで保持
                // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                // 注意：recentFilesCountは既に上で定義されているため、新しい変数名を使用
                var currentRecentFilesCount = _recentFiles.Count;
                if (currentRecentFilesCount > 20)
                {
                    _recentFiles.RemoveRange(20, currentRecentFilesCount - 20);
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

        #endregion

        #region ドライブ操作

        /// <summary>
        /// ドライブ一覧を表示します（後方互換性のため残す）
        /// </summary>
        /// <param name="addToHistory">履歴に追加するかどうか</param>
        public void NavigateToDrives(bool addToHistory = true)
        {
            NavigateToHome(addToHistory);
        }

        #endregion

        #region ソート

        /// <summary>
        /// 指定された列でソートします
        /// </summary>
        /// <param name="columnName">ソートする列名（"Name", "Size", "Type"/"Extension", "DateModified"/"LastModified", "DateCreated", "Ascending", "Descending"）</param>
        public void SortByColumn(string columnName)
        {
            // ソート方向の変更を処理
            if (columnName == "Ascending")
            {
                _sortAscending = true;
                SortItems();
                return;
            }
            else if (columnName == "Descending")
            {
                _sortAscending = false;
                SortItems();
                return;
            }

            // 列名のマッピング（UIから渡される名前を内部の列名に変換）
            string mappedColumnName = columnName switch
            {
                "Type" => "Extension",
                "DateModified" => "LastModified",
                "DateCreated" => "LastModified", // DateCreatedプロパティがないため、LastModifiedを使用
                _ => columnName
            };

            // 同じ列をクリックした場合はソート方向を反転
            if (_sortColumn == mappedColumnName)
            {
                _sortAscending = !_sortAscending;
            }
            else
            {
                _sortColumn = mappedColumnName;
                _sortAscending = true;
            }

            // 現在のアイテムをソート
            SortItems();
        }

        /// <summary>
        /// アイテムを現在のソート設定に基づいてソートします
        /// </summary>
        private void SortItems()
        {
            var itemsCount = Items.Count;
            if (itemsCount == 0)
                return;

            // 一時リストにコピーしてソート（容量を事前に確保）
            var sortedItems = new List<FileSystemItem>(itemsCount);
            sortedItems.AddRange(Items);
            SortItemList(sortedItems);

            // UIスレッドで更新（一括更新でパフォーマンス向上）
            var dispatcher = _cachedDispatcher ?? System.Windows.Application.Current?.Dispatcher;
            dispatcher?.Invoke(() =>
            {
                // Clear + 個別にAdd（ObservableCollectionにはAddRangeがないため）
                var items = Items;
                items.Clear();
                
                // プロパティアクセスをキャッシュして高速化
                foreach (var item in sortedItems)
                {
                    items.Add(item);
                }
            });
        }

        /// <summary>
        /// アイテムリストを現在のソート設定に基づいてソートします（インライン、最適化済み）
        /// </summary>
        /// <param name="itemList">ソートするアイテムリスト</param>
        private void SortItemList(List<FileSystemItem> itemList)
        {
            var count = itemList.Count;
            if (count == 0)
                return;

            // ソート列名を定数として使用（メモリ割り当てを削減）
            const string NameColumn = "Name";
            const string SizeColumn = "Size";
            const string ExtensionColumn = "Extension";
            const string LastModifiedColumn = "LastModified";

            // ソート実行（比較関数を最適化）
            itemList.Sort((x, y) =>
            {
                // ディレクトリを常に先に表示（早期リターンでパフォーマンス向上）
                var xIsDir = x.IsDirectory;
                var yIsDir = y.IsDirectory;
                if (xIsDir != yIsDir)
                {
                    return xIsDir ? -1 : 1;
                }

                // 列に応じてソート（switch式を使用してパフォーマンス向上）
                // Extension列でソートする場合、拡張子が空のディレクトリは適切に処理される
                int result = _sortColumn switch
                {
                    NameColumn => string.Compare(x.Name, y.Name, StringComparison.OrdinalIgnoreCase),
                    SizeColumn => x.Size.CompareTo(y.Size),
                    ExtensionColumn => string.Compare(x.Extension, y.Extension, StringComparison.OrdinalIgnoreCase),
                    LastModifiedColumn => x.LastModified.CompareTo(y.LastModified),
                    _ => string.Compare(x.Name, y.Name, StringComparison.OrdinalIgnoreCase)
                };

                // ソート方向を適用（三項演算子で分岐を削減）
                return _sortAscending ? result : -result;
            });
        }

        #endregion

        #region FileSystemWatcher

        /// <summary>
        /// FileSystemWatcherを更新します
        /// </summary>
        /// <param name="path">監視するパス</param>
        private void UpdateFileSystemWatcher(string? path)
        {
            // 既存のFileSystemWatcherを停止して破棄
            if (_fileSystemWatcher != null)
            {
                _fileSystemWatcher.EnableRaisingEvents = false;
                _fileSystemWatcher.Dispose();
                _fileSystemWatcher = null;
            }

            // パスが空、またはホームページの場合は監視しない
            if (string.IsNullOrEmpty(path) || IsHomePage)
                return;

            // パスが存在しない場合は監視しない
            if (!Directory.Exists(path))
                return;

            try
            {
                // 新しいFileSystemWatcherを作成
                _fileSystemWatcher = new FileSystemWatcher(path)
                {
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite,
                    IncludeSubdirectories = false,
                    EnableRaisingEvents = true
                };

                // イベントハンドラーを登録
                _fileSystemWatcher.Created += OnFileSystemChanged;
                _fileSystemWatcher.Deleted += OnFileSystemChanged;
                _fileSystemWatcher.Renamed += OnFileSystemRenamed;
                _fileSystemWatcher.Changed += OnFileSystemChanged;
            }
            catch
            {
                // エラーが発生した場合は監視を無効化
                _fileSystemWatcher?.Dispose();
                _fileSystemWatcher = null;
            }
        }

        /// <summary>
        /// ファイルシステムの変更イベントハンドラー
        /// </summary>
        private void OnFileSystemChanged(object sender, FileSystemEventArgs e)
        {
            // スロットリング：一定時間内の複数のイベントを1つにまとめる
            ScheduleRefresh();
        }

        /// <summary>
        /// ファイルシステムのリネームイベントハンドラー
        /// </summary>
        private void OnFileSystemRenamed(object sender, RenamedEventArgs e)
        {
            // スロットリング：一定時間内の複数のイベントを1つにまとめる
            ScheduleRefresh();
        }

        /// <summary>
        /// リフレッシュをスケジュールします（スロットリング付き）
        /// </summary>
        private void ScheduleRefresh()
        {
            // 既にリフレッシュが保留中の場合は何もしない
            if (_refreshPending)
                return;

            _refreshPending = true;

            // タイマーが存在しない場合は作成
            if (_refreshTimer == null)
            {
                _refreshTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(500) // 500ms間隔で更新
                };
                _refreshTimer.Tick += (s, e) =>
                {
                    _refreshTimer.Stop();
                    _refreshPending = false;
                    
                    // UIスレッドでリフレッシュを実行
                    var dispatcher = _cachedDispatcher ?? System.Windows.Application.Current?.Dispatcher;
                    dispatcher?.BeginInvoke(() =>
                    {
                        // 現在のパスがまだ有効な場合のみリフレッシュ
                        if (!string.IsNullOrEmpty(CurrentPath) && Directory.Exists(CurrentPath))
                        {
                            Refresh();
                        }
                    }, System.Windows.Threading.DispatcherPriority.Background);
                };
            }

            // タイマーをリセットして再開始
            _refreshTimer.Stop();
            _refreshTimer.Start();
        }

        /// <summary>
        /// FileSystemWatcherをクリーンアップします
        /// </summary>
        private void CleanupFileSystemWatcher()
        {
            if (_fileSystemWatcher != null)
            {
                _fileSystemWatcher.EnableRaisingEvents = false;
                _fileSystemWatcher.Dispose();
                _fileSystemWatcher = null;
            }

            if (_refreshTimer != null)
            {
                _refreshTimer.Stop();
                _refreshTimer = null;
            }

            _refreshPending = false;
        }

        #endregion

    }
}

