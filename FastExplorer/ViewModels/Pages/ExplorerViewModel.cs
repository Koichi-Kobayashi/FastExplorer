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
        private bool _isInitialized = false;
        private readonly Stack<string> _backHistory = new();
        private readonly Stack<string> _forwardHistory = new();
        private bool _isNavigating = false;

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
        /// <see cref="ExplorerViewModel"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="fileSystemService">ファイルシステムサービス</param>
        public ExplorerViewModel(FileSystemService fileSystemService)
        {
            _fileSystemService = fileSystemService;
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
            // デフォルトでマイコンピュータ（ドライブ一覧）を表示
            if (!_isInitialized)
            {
                NavigateToDrives();
                _isInitialized = true;
            }
        }

        /// <summary>
        /// 指定されたパスにナビゲートします
        /// </summary>
        /// <param name="path">ナビゲートするパス。nullまたは空の場合はドライブ一覧を表示</param>
        [RelayCommand]
        private void NavigateToPath(string? path)
        {
            if (string.IsNullOrEmpty(path))
            {
                NavigateToDrives();
                return;
            }

            // パスが空文字列の場合はドライブ一覧に戻る
            if (string.IsNullOrWhiteSpace(path))
            {
                NavigateToDrives();
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
                    NavigateToDrives(addToHistory: false);
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
                NavigateToDrives();
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
                NavigateToDrives(addToHistory: false);
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
                NavigateToDrives();
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
        /// ドライブ一覧を表示します
        /// </summary>
        /// <param name="addToHistory">履歴に追加するかどうか</param>
        public void NavigateToDrives(bool addToHistory = true)
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
                            Items.Add(new FileSystemItem
                            {
                                Name = $"{driveInfo.Name} ({driveInfo.VolumeLabel})",
                                FullPath = driveInfo.RootDirectory.FullName,
                                Extension = string.Empty,
                                Size = driveInfo.TotalSize,
                                LastModified = driveInfo.RootDirectory.LastWriteTime,
                                IsDirectory = true,
                                Attributes = FileAttributes.Directory
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
            finally
            {
                IsLoading = false;
                _isNavigating = false;
            }
        }
    }
}

