using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.IO;
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

            _isNavigating = true;

            // 履歴に追加する場合、現在のパスを戻る履歴に追加し、進む履歴をクリア
            if (addToHistory && !string.IsNullOrEmpty(CurrentPath) && CurrentPath != path)
            {
                _backHistory.Push(CurrentPath);
                _forwardHistory.Clear();
            }

            IsLoading = true;
            CurrentPath = path;
            Items.Clear();
            IsHomePage = false;

            try
            {
                var fileItems = _fileSystemService.GetItems(path);
                foreach (var item in fileItems)
                {
                    Items.Add(item);
                }
            }
            catch (Exception)
            {
                // エラーハンドリング
            }
            finally
            {
                IsLoading = false;
                _isNavigating = false;
            }
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
            if (addToHistory && !string.IsNullOrEmpty(CurrentPath))
            {
                _backHistory.Push(CurrentPath);
                _forwardHistory.Clear();
            }

            IsLoading = true;
            CurrentPath = string.Empty;
            Items.Clear();
            IsHomePage = true;

            try
            {
                // ピン留めフォルダーを読み込み
                LoadPinnedFolders();

                // ドライブ情報を読み込み
                LoadDrives();

                // 最近使用したファイルを読み込み
                LoadRecentFiles();
            }
            catch (Exception)
            {
                // エラーハンドリング
            }
            finally
            {
                IsLoading = false;
                _isNavigating = false;
            }
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
        /// ドライブ情報を読み込みます
        /// </summary>
        private void LoadDrives()
        {
            Drives.Clear();
            try
            {
                var drives = _fileSystemService.GetDrives();
                foreach (var drive in drives)
                {
                    try
                    {
                        var driveInfo = new DriveInfo(drive);
                        if (driveInfo.IsReady)
                        {
                            var driveName = string.IsNullOrEmpty(driveInfo.VolumeLabel)
                                ? $"ローカルディスク ({drive.TrimEnd('\\')})"
                                : $"{driveInfo.VolumeLabel} ({drive.TrimEnd('\\')})";
                            
                            Drives.Add(new DriveInfoModel
                            {
                                Name = driveName,
                                Path = driveInfo.RootDirectory.FullName,
                                VolumeLabel = driveInfo.VolumeLabel,
                                TotalSize = driveInfo.TotalSize,
                                FreeSpace = driveInfo.AvailableFreeSpace
                            });
                        }
                    }
                    catch
                    {
                        // ドライブにアクセスできない場合はスキップ
                        continue;
                    }
                }
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

