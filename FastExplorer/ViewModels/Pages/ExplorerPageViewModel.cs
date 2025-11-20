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
        /// 左ペインのタブのコレクション
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<ExplorerTab> _leftPaneTabs = new();

        /// <summary>
        /// 右ペインのタブのコレクション
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<ExplorerTab> _rightPaneTabs = new();

        /// <summary>
        /// 現在選択されているタブ
        /// </summary>
        [ObservableProperty]
        private ExplorerTab? _selectedTab;

        /// <summary>
        /// 左ペインで現在選択されているタブ
        /// </summary>
        [ObservableProperty]
        private ExplorerTab? _selectedLeftPaneTab;

        /// <summary>
        /// 右ペインで現在選択されているタブ
        /// </summary>
        [ObservableProperty]
        private ExplorerTab? _selectedRightPaneTab;

        /// <summary>
        /// 分割ペインが有効かどうか
        /// </summary>
        [ObservableProperty]
        private bool _isSplitPaneEnabled;

        /// <summary>
        /// IsSplitPaneEnabledが変更されたときに呼び出されます
        /// </summary>
        partial void OnIsSplitPaneEnabledChanged(bool value)
        {
            // ToggleSplitPaneメソッド内で直接更新するため、ここでは何もしない
            // （無限ループを防ぐため）
        }

        /// <summary>
        /// 分割ペインを切り替えます
        /// </summary>
        [RelayCommand]
        private void ToggleSplitPane()
        {
            var windowSettingsService = _windowSettingsService ?? 
                App.Services.GetService(typeof(WindowSettingsService)) as WindowSettingsService;
            
            if (windowSettingsService == null)
                return;

            var settings = windowSettingsService.GetSettings();
            IsSplitPaneEnabled = !IsSplitPaneEnabled;
            settings.IsSplitPaneEnabled = IsSplitPaneEnabled;
            windowSettingsService.SaveSettings(settings);
            
            // SettingsViewModelを更新
            var settingsViewModel = App.Services.GetService(typeof(SettingsViewModel)) as SettingsViewModel;
            if (settingsViewModel != null)
            {
                settingsViewModel.IsSplitPaneEnabled = IsSplitPaneEnabled;
            }

            // 分割ペインを切り替えた場合、現在のタブを適切なペインに移動
            if (IsSplitPaneEnabled)
            {
                // 通常モードから分割モードへ
                if (Tabs.Count > 0)
                {
                    // 現在のタブを左ペインに移動
                    var selectedTab = SelectedTab;
                    foreach (var tab in Tabs)
                    {
                        LeftPaneTabs.Add(tab);
                    }
                    Tabs.Clear();
                    // 選択されていたタブを左ペインの選択タブに設定
                    if (selectedTab != null)
                    {
                        SelectedLeftPaneTab = selectedTab;
                    }
                    else if (LeftPaneTabs.Count > 0)
                    {
                        // 選択タブがなかった場合は、最初のタブを選択
                        SelectedLeftPaneTab = LeftPaneTabs[0];
                    }
                    SelectedTab = null;
                }
                else
                {
                    // タブが存在しない場合は、新しいタブを作成
                    if (LeftPaneTabs.Count == 0)
                    {
                        CreateNewLeftPaneTab();
                    }
                    else if (SelectedLeftPaneTab == null && LeftPaneTabs.Count > 0)
                    {
                        // タブは存在するが選択されていない場合は、最初のタブを選択
                        SelectedLeftPaneTab = LeftPaneTabs[0];
                    }
                }
                // 右ペインに新しいタブを作成（存在しない場合）
                if (RightPaneTabs.Count == 0)
                {
                    CreateNewRightPaneTab();
                }
                else if (SelectedRightPaneTab == null && RightPaneTabs.Count > 0)
                {
                    // タブは存在するが選択されていない場合は、最初のタブを選択
                    SelectedRightPaneTab = RightPaneTabs[0];
                }
            }
            else
            {
                // 分割モードから通常モードへ
                // 左ペインのタブを通常のタブに移動
                var selectedLeftTab = SelectedLeftPaneTab;
                foreach (var tab in LeftPaneTabs)
                {
                    Tabs.Add(tab);
                }
                LeftPaneTabs.Clear();
                
                // 右ペインのタブも通常のタブに移動（最初のタブのみ保持する場合は、ここで処理）
                // 現在は右ペインのタブは保持しない
                RightPaneTabs.Clear();
                SelectedRightPaneTab = null;
                
                // 選択タブを設定
                if (selectedLeftTab != null)
                {
                    SelectedTab = selectedLeftTab;
                }
                else if (Tabs.Count > 0)
                {
                    // 選択タブがなかった場合は、最初のタブを選択
                    SelectedTab = Tabs[0];
                }
                SelectedLeftPaneTab = null;
                
                // タブが存在しない場合は、新しいタブを作成
                if (Tabs.Count == 0)
                {
                    CreateNewTab();
                }
            }
        }

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
            // 分割ペインの設定を読み込む
            LoadSplitPaneSettings();

            if (IsSplitPaneEnabled)
            {
                if (LeftPaneTabs.Count == 0 && RightPaneTabs.Count == 0)
                {
                    // 保存されたタブ情報を復元
                    RestoreTabs();
                    
                    // タブが復元されなかった場合は新しいタブを作成
                    if (LeftPaneTabs.Count == 0)
                    {
                        CreateNewLeftPaneTab();
                    }
                    if (RightPaneTabs.Count == 0)
                    {
                        CreateNewRightPaneTab();
                    }
                }
                // 選択タブが設定されていない場合は、最初のタブを選択
                if (SelectedLeftPaneTab == null && LeftPaneTabs.Count > 0)
                {
                    SelectedLeftPaneTab = LeftPaneTabs[0];
                }
                if (SelectedRightPaneTab == null && RightPaneTabs.Count > 0)
                {
                    SelectedRightPaneTab = RightPaneTabs[0];
                }
            }
            else
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
            var tab = CreateTabInternal();
            Tabs.Add(tab);
            SelectedTab = tab;
            UpdateTabTitle(tab);
        }

        /// <summary>
        /// 左ペインに新しいタブを作成します
        /// </summary>
        [RelayCommand]
        private void CreateNewLeftPaneTab()
        {
            var tab = CreateTabInternal();
            LeftPaneTabs.Add(tab);
            SelectedLeftPaneTab = tab;
            UpdateTabTitle(tab);
        }

        /// <summary>
        /// 右ペインに新しいタブを作成します
        /// </summary>
        [RelayCommand]
        private void CreateNewRightPaneTab()
        {
            var tab = CreateTabInternal();
            RightPaneTabs.Add(tab);
            SelectedRightPaneTab = tab;
            UpdateTabTitle(tab);
        }

        /// <summary>
        /// タブを作成する内部メソッド
        /// </summary>
        private ExplorerTab CreateTabInternal()
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

            return tab;
        }

        /// <summary>
        /// 指定されたタブを閉じます
        /// </summary>
        /// <param name="tab">閉じるタブ。nullの場合、またはタブが1つしかない場合は何もしません</param>
        [RelayCommand]
        private void CloseTab(ExplorerTab? tab)
        {
            if (tab == null)
                return;

            if (IsSplitPaneEnabled)
            {
                // 左ペインのタブかチェック
                if (LeftPaneTabs.Contains(tab))
                {
                    if (LeftPaneTabs.Count <= 1)
                        return;
                    var index = LeftPaneTabs.IndexOf(tab);
                    LeftPaneTabs.Remove(tab);
                    if (SelectedLeftPaneTab == tab)
                    {
                        if (index > 0)
                            SelectedLeftPaneTab = LeftPaneTabs[index - 1];
                        else if (LeftPaneTabs.Count > 0)
                            SelectedLeftPaneTab = LeftPaneTabs[0];
                    }
                }
                // 右ペインのタブかチェック
                else if (RightPaneTabs.Contains(tab))
                {
                    if (RightPaneTabs.Count <= 1)
                        return;
                    var index = RightPaneTabs.IndexOf(tab);
                    RightPaneTabs.Remove(tab);
                    if (SelectedRightPaneTab == tab)
                    {
                        if (index > 0)
                            SelectedRightPaneTab = RightPaneTabs[index - 1];
                        else if (RightPaneTabs.Count > 0)
                            SelectedRightPaneTab = RightPaneTabs[0];
                    }
                }
            }
            else
            {
                if (Tabs.Count <= 1)
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
        /// SelectedLeftPaneTabが変更されたときに呼び出されます
        /// </summary>
        partial void OnSelectedLeftPaneTabChanged(ExplorerTab? value)
        {
            if (value != null)
            {
                UpdateTabTitle(value);
            }
            UpdateStatusBar();
        }

        /// <summary>
        /// SelectedRightPaneTabが変更されたときに呼び出されます
        /// </summary>
        partial void OnSelectedRightPaneTabChanged(ExplorerTab? value)
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

            ExplorerTab? activeTab = null;
            if (IsSplitPaneEnabled)
            {
                // 分割ペインの場合は、フォーカスがある方のタブを表示
                // 簡易実装として、右ペインを優先
                activeTab = SelectedRightPaneTab ?? SelectedLeftPaneTab;
            }
            else
            {
                activeTab = SelectedTab;
            }

            if (activeTab != null)
            {
                // タブのタイトルも更新
                UpdateTabTitle(activeTab);
            }

            if (activeTab?.ViewModel != null)
            {
                var path = activeTab.ViewModel.CurrentPath;
                var itemCount = activeTab.ViewModel.Items.Count;
                
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
        /// 分割ペインの設定を読み込みます
        /// </summary>
        private void LoadSplitPaneSettings()
        {
            try
            {
                var windowSettingsService = _windowSettingsService ?? 
                    App.Services.GetService(typeof(WindowSettingsService)) as WindowSettingsService;
                
                if (windowSettingsService == null)
                    return;

                var settings = windowSettingsService.GetSettings();
                IsSplitPaneEnabled = settings.IsSplitPaneEnabled;
            }
            catch
            {
                IsSplitPaneEnabled = false;
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

                if (IsSplitPaneEnabled)
                {
                    // 分割ペインの場合は左右のペインのタブを復元
                    RestorePaneTabs(settings.LeftPaneTabPaths, LeftPaneTabs, (tab) => SelectedLeftPaneTab = tab);
                    RestorePaneTabs(settings.RightPaneTabPaths, RightPaneTabs, (tab) => SelectedRightPaneTab = tab);
                }
                else
                {
                    // 通常モードの場合は従来通り
                    var tabPaths = settings.TabPaths;
                    if (tabPaths == null || tabPaths.Count == 0)
                        return;

                    RestorePaneTabs(tabPaths, Tabs, (tab) => SelectedTab = tab);
                }
            }
            catch
            {
                // エラーハンドリング：復元に失敗した場合は新しいタブを作成
            }
        }

        /// <summary>
        /// 指定されたペインのタブを復元します
        /// </summary>
        private void RestorePaneTabs(List<string> tabPaths, ObservableCollection<ExplorerTab> tabs, Action<ExplorerTab> setSelectedTab)
        {
            if (tabPaths == null || tabPaths.Count == 0)
                return;

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

                tabs.Add(tab);
                UpdateTabTitle(tab);
            }

            // 最初のタブを選択
            if (tabs.Count > 0)
            {
                setSelectedTab(tabs[0]);
            }
        }
    }
}

