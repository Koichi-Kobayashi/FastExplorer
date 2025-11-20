using System.Collections.ObjectModel;
using System.IO;
using FastExplorer.Models;
using FastExplorer.Services;
using Wpf.Ui.Abstractions.Controls;

namespace FastExplorer.ViewModels.Pages
{
    /// <summary>
    /// エクスプローラーページのタブ管理を行うViewModel
    /// </summary>
    public partial class ExplorerPageViewModel : ObservableObject, INavigationAware
    {
        private readonly FileSystemService _fileSystemService;
        private readonly FavoriteService? _favoriteService;
        private readonly WindowSettingsService? _windowSettingsService;
        
        // MainWindowViewModelをキャッシュ（パフォーマンス向上）
        private ViewModels.Windows.MainWindowViewModel? _cachedMainWindowViewModel;

        /// <summary>
        /// エクスプローラータブのコレクション
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<ExplorerTab> _tabs = new();

        /// <summary>
        /// 現在選択されているタブ
        /// </summary>
        [ObservableProperty]
        private ExplorerTab? _selectedTab;

        /// <summary>
        /// <see cref="ExplorerPageViewModel"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="fileSystemService">ファイルシステムサービス</param>
        /// <param name="favoriteService">お気に入りサービス</param>
        /// <param name="windowSettingsService">ウィンドウ設定サービス</param>
        public ExplorerPageViewModel(FileSystemService fileSystemService, FavoriteService? favoriteService = null, WindowSettingsService? windowSettingsService = null)
        {
            _fileSystemService = fileSystemService;
            _favoriteService = favoriteService;
            _windowSettingsService = windowSettingsService;
        }

        /// <summary>
        /// ページにナビゲートされたときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedToAsync()
        {
            if (Tabs.Count == 0)
            {
                // 保存されたタブ情報を復元
                RestoreTabs();
                
                // タブが復元されなかった場合は新しいタブを作成
                if (Tabs.Count == 0)
                {
                    CreateNewTab();
                }
            }
            return Task.CompletedTask;
        }

        /// <summary>
        /// ページから離れるときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedFromAsync() => Task.CompletedTask;

        /// <summary>
        /// 新しいタブを作成します
        /// </summary>
        [RelayCommand]
        private void CreateNewTab()
        {
            var viewModel = new ExplorerViewModel(_fileSystemService, _favoriteService);
            var tab = new ExplorerTab
            {
                Title = "ホーム",
                CurrentPath = string.Empty,
                ViewModel = viewModel
            };

            // CurrentPathが変更されたときにTitleとCurrentPathを更新
            viewModel.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == "CurrentPath" || e.PropertyName == nameof(ExplorerViewModel.CurrentPath))
                {
                    tab.CurrentPath = viewModel.CurrentPath;
                    UpdateTabTitle(tab);
                    UpdateStatusBar();
                }
                else if (e.PropertyName == nameof(ExplorerViewModel.Items))
                {
                    UpdateStatusBar();
                }
            };

            // タブを追加する前に初期化を完了させる
            // これにより、タブが追加された直後にフォルダーをダブルクリックしても問題が発生しない
            tab.ViewModel.NavigateToHome();
            
            Tabs.Add(tab);
            SelectedTab = tab;
            
            UpdateTabTitle(tab);
        }

        /// <summary>
        /// 指定されたタブを閉じます
        /// </summary>
        /// <param name="tab">閉じるタブ。nullの場合、またはタブが1つしかない場合は何もしません</param>
        [RelayCommand]
        private void CloseTab(ExplorerTab? tab)
        {
            if (tab == null || Tabs.Count <= 1)
                return;

            var index = Tabs.IndexOf(tab);
            Tabs.Remove(tab);

            if (SelectedTab == tab)
            {
                if (index > 0)
                    SelectedTab = Tabs[index - 1];
                else if (Tabs.Count > 0)
                    SelectedTab = Tabs[0];
            }
        }

        /// <summary>
        /// タブのタイトルを現在のパスに基づいて更新します
        /// </summary>
        /// <param name="tab">タイトルを更新するタブ</param>
        private void UpdateTabTitle(ExplorerTab tab)
        {
            if (string.IsNullOrEmpty(tab.ViewModel.CurrentPath))
            {
                tab.Title = "ホーム";
            }
            else
            {
                tab.Title = tab.ViewModel.CurrentPath;
            }
        }

        /// <summary>
        /// 現在のパスをお気に入りに追加します
        /// </summary>
        [RelayCommand]
        private void AddCurrentPathToFavorites()
        {
            if (SelectedTab == null || string.IsNullOrEmpty(SelectedTab.ViewModel.CurrentPath))
                return;

            var path = SelectedTab.ViewModel.CurrentPath;
            var name = Path.GetFileName(path) ?? path;
            
            _favoriteService?.AddFavorite(name, path);
            
            // MainWindowViewModelを更新（キャッシュを使用）
            if (_cachedMainWindowViewModel == null)
            {
                _cachedMainWindowViewModel = App.Services.GetService(typeof(ViewModels.Windows.MainWindowViewModel)) as ViewModels.Windows.MainWindowViewModel;
            }
            _cachedMainWindowViewModel?.LoadFavorites();
        }

        /// <summary>
        /// SelectedTabが変更されたときに呼び出されます
        /// </summary>
        partial void OnSelectedTabChanged(ExplorerTab? value)
        {
            if (value != null)
            {
                UpdateTabTitle(value);
            }
            UpdateStatusBar();
        }

        /// <summary>
        /// ステータスバーのテキストを更新します
        /// </summary>
        private void UpdateStatusBar()
        {
            // MainWindowViewModelをキャッシュから取得（なければ取得してキャッシュ）
            if (_cachedMainWindowViewModel == null)
            {
                _cachedMainWindowViewModel = App.Services.GetService(typeof(ViewModels.Windows.MainWindowViewModel)) as ViewModels.Windows.MainWindowViewModel;
            }
            
            if (_cachedMainWindowViewModel == null)
                return;

            if (SelectedTab != null)
            {
                // タブのタイトルも更新
                UpdateTabTitle(SelectedTab);
            }

            if (SelectedTab?.ViewModel != null)
            {
                var path = SelectedTab.ViewModel.CurrentPath;
                var itemCount = SelectedTab.ViewModel.Items.Count;
                
                if (string.IsNullOrEmpty(path))
                {
                    _cachedMainWindowViewModel.StatusBarText = $"パス: ホーム {itemCount}個の項目";
                }
                else
                {
                    _cachedMainWindowViewModel.StatusBarText = $"パス: {path} {itemCount}個の項目";
                }
            }
            else
            {
                _cachedMainWindowViewModel.StatusBarText = "準備完了";
            }
        }

        /// <summary>
        /// 保存されたタブ情報を復元します
        /// </summary>
        private void RestoreTabs()
        {
            try
            {
                // WindowSettingsServiceを取得（コンストラクタで取得できなかった場合はApp.Servicesから取得）
                var windowSettingsService = _windowSettingsService ?? 
                    App.Services.GetService(typeof(WindowSettingsService)) as WindowSettingsService;
                
                if (windowSettingsService == null)
                    return;

                var settings = windowSettingsService.GetSettings();
                var tabPaths = settings.TabPaths;

                if (tabPaths == null || tabPaths.Count == 0)
                    return;

                // 保存されたパスに基づいてタブを復元
                foreach (var path in tabPaths)
                {
                    var viewModel = new ExplorerViewModel(_fileSystemService, _favoriteService);
                    var tab = new ExplorerTab
                    {
                        Title = string.IsNullOrEmpty(path) ? "ホーム" : path,
                        CurrentPath = path ?? string.Empty,
                        ViewModel = viewModel
                    };

                    // CurrentPathが変更されたときにTitleとCurrentPathを更新
                    viewModel.PropertyChanged += (s, e) =>
                    {
                        if (e.PropertyName == "CurrentPath" || e.PropertyName == nameof(ExplorerViewModel.CurrentPath))
                        {
                            tab.CurrentPath = viewModel.CurrentPath;
                            UpdateTabTitle(tab);
                            UpdateStatusBar();
                        }
                        else if (e.PropertyName == nameof(ExplorerViewModel.Items))
                        {
                            UpdateStatusBar();
                        }
                    };

                    // パスが空の場合はホームに、そうでない場合は指定されたパスに移動
                    if (string.IsNullOrEmpty(path))
                    {
                        tab.ViewModel.NavigateToHome();
                    }
                    else
                    {
                        // パスが存在するか確認してから移動
                        if (System.IO.Directory.Exists(path))
                        {
                            tab.ViewModel.NavigateToPathCommand.Execute(path);
                        }
                        else
                        {
                            // パスが存在しない場合はホームに移動
                            tab.ViewModel.NavigateToHome();
                        }
                    }

                    Tabs.Add(tab);
                    UpdateTabTitle(tab);
                }

                // 最初のタブを選択
                if (Tabs.Count > 0)
                {
                    SelectedTab = Tabs[0];
                }
            }
            catch
            {
                // エラーハンドリング：復元に失敗した場合は新しいタブを作成
            }
        }
    }
}

