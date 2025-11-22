using System.Collections.ObjectModel;
using System.Collections.Generic;
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
            if (string.IsNullOrEmpty(path))
            {
                NavigateToHome();
                return;
            }

            // パスが空文字列の場合はホームページに戻る
            if (string.IsNullOrWhiteSpace(path))
            {
                NavigateToHome();
                return;
            }

            var trimmedPath = path.Trim();

            // shell:で始まる特殊なパス（ごみ箱、ネットワークなど）の処理
            if (trimmedPath.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
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
                                // コレクション変更通知を抑制して高速化（ただし、WPFのバインディングには通知が必要）
                                foreach (var item in itemList)
                                {
                                    Items.Add(item);
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
                            
                            await dispatcher.InvokeAsync(() =>
                            {
                                if (!cancellationToken.IsCancellationRequested)
                                {
                                    // 直接インデックスアクセスでメモリ割り当てを削減
                                    for (int j = startIdx; j < endIdx; j++)
                                    {
                                        Items.Add(itemList[j]);
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
            foreach (var favorite in favorites)
            {
                // 標準的なWindowsフォルダのみをピン留めとして表示
                var name = favorite.Name;
                if (name == "デスクトップ" || name == "ダウンロード" || name == "ドキュメント" ||
                    name == "ピクチャ" || name == "ミュージック" || name == "ビデオ" || name == "ごみ箱")
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
            foreach (var file in _recentFiles.Take(10))
            {
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
                _recentFiles.RemoveAll(f => f.FullPath.Equals(filePath, StringComparison.OrdinalIgnoreCase));
                // 先頭に追加
                _recentFiles.Insert(0, item);
                // 最大20件まで保持
                if (_recentFiles.Count > 20)
                {
                    _recentFiles.RemoveRange(20, _recentFiles.Count - 20);
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

