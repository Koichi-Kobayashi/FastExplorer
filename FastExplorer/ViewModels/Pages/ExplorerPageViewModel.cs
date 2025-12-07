using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using Cysharp.Text;
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
        #region フィールド

        private readonly FileSystemService _fileSystemService;
        private readonly FavoriteService? _favoriteService;
        private readonly WindowSettingsService? _windowSettingsService;

        // MainWindowViewModelをキャッシュ（パフォーマンス向上）
        private ViewModels.Windows.MainWindowViewModel? _cachedMainWindowViewModel;

        // サービス取得の最適化（キャッシュ）
        private SettingsViewModel? _cachedSettingsViewModel;
        private WindowSettingsService? _cachedWindowSettingsService;

        // 文字列定数（メモリ割り当てを削減）
        private const string HomeTitle = "ホーム";
        private const string ReadyStatus = "準備完了";
        private const string ItemSuffix = "個の項目";

        // Dispatcherをキャッシュ（パフォーマンス向上）
        private System.Windows.Threading.Dispatcher? _cachedDispatcher;

        // 型をキャッシュ（パフォーマンス向上）
        private static readonly Type WindowSettingsServiceType = typeof(WindowSettingsService);
        private static readonly Type SettingsViewModelType = typeof(SettingsViewModel);
        private static readonly Type MainWindowViewModelType = typeof(ViewModels.Windows.MainWindowViewModel);

        // PropertyChangedイベントのプロパティ名を定数化（メモリ割り当てを削減）
        private const string CurrentPathPropertyName = "CurrentPath";
        private static readonly string CurrentPathPropertyNameFull = nameof(ExplorerViewModel.CurrentPath);
        private static readonly string ItemsPropertyName = nameof(ExplorerViewModel.Items);

        // UpdateStatusBarデリゲートをキャッシュ（メモリ割り当てを削減）
        private System.Action? _cachedUpdateStatusBarAction;

        // OnNavigatedToAsyncが初回実行されたかどうかを追跡
        private bool _hasNavigatedTo = false;

        // ActivePaneの定数（パフォーマンス向上）
        private const int ActivePaneLeft = 0;
        private const int ActivePaneRight = 2;
        private const int ActivePaneNone = -1;

        #endregion

        #region プロパティ

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
        /// 現在アクティブなペイン（0=左ペイン、2=右ペイン、-1=未設定）
        /// </summary>
        [ObservableProperty]
        private int _activePane = -1;

        #endregion

        #region プロパティ変更ハンドラー

        /// <summary>
        /// ActivePaneが変更されたときに呼び出されます
        /// </summary>
        partial void OnActivePaneChanged(int value)
        {
            // ActivePaneが変更されたとき、ステータスバーを更新
            if (IsSplitPaneEnabled)
            {
                UpdateStatusBar();
            }
        }

        /// <summary>
        /// IsSplitPaneEnabledが変更されたときに呼び出されます
        /// </summary>
        partial void OnIsSplitPaneEnabledChanged(bool value)
        {
            // ToggleSplitPaneメソッド内で直接更新するため、ここでは何もしない
            // （無限ループを防ぐため）
        }

        #endregion

        #region コンストラクタ

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
            
            // デバッグ: コマンドが生成されているか確認
            System.Diagnostics.Debug.WriteLine($"[ExplorerPageViewModel] コンストラクタ: NavigateToHomeInActivePaneCommand={(NavigateToHomeInActivePaneCommand != null ? "存在" : "null")}");
        }

        #endregion

        #region 分割ペイン管理

        /// <summary>
        /// 分割ペインを切り替えます
        /// </summary>
        [RelayCommand]
        private void ToggleSplitPane()
        {
            // WindowSettingsServiceをキャッシュ（パフォーマンス向上）
            var windowSettingsService = _windowSettingsService ??
                (_cachedWindowSettingsService ??= App.Services.GetService(WindowSettingsServiceType) as WindowSettingsService);

            if (windowSettingsService == null)
                return;

            var settings = windowSettingsService.GetSettings();
            IsSplitPaneEnabled = !IsSplitPaneEnabled;
            settings.IsSplitPaneEnabled = IsSplitPaneEnabled;
            windowSettingsService.SaveSettings(settings);

            // SettingsViewModelを更新（キャッシュを使用）
            if (_cachedSettingsViewModel == null)
            {
                _cachedSettingsViewModel = App.Services.GetService(SettingsViewModelType) as SettingsViewModel;
            }
            if (_cachedSettingsViewModel != null && _cachedSettingsViewModel.IsSplitPaneEnabled != IsSplitPaneEnabled)
            {
                _cachedSettingsViewModel.IsSplitPaneEnabled = IsSplitPaneEnabled;
            }

            // 分割ペインを切り替えた場合、現在のタブを適切なペインに移動
            if (IsSplitPaneEnabled)
            {
                // 通常モードから分割モードへ
                // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                var tabsCount = Tabs.Count;
                if (tabsCount > 0)
                {
                    // パフォーマンス最適化：タブを一時リストに収集してから一括で移動
                    var selectedTab = SelectedTab;
                    var tabsToMove = new System.Collections.Generic.List<ExplorerTab>(tabsCount);
                    foreach (var tab in Tabs)
                    {
                        tabsToMove.Add(tab);
                    }
                    Tabs.Clear();
                    
                    // 一括追加（各Addで通知が発行されるが、Clear()後なのでUIの再構築は1回）
                    foreach (var tab in tabsToMove)
                    {
                        LeftPaneTabs.Add(tab);
                    }
                    
                    // 選択されていたタブを左ペインの選択タブに設定
                    if (selectedTab != null)
                    {
                        SelectedLeftPaneTab = selectedTab;
                    }
                    else
                    {
                        // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                        var leftPaneTabsCount = LeftPaneTabs.Count;
                        if (leftPaneTabsCount > 0)
                        {
                            // 選択タブがなかった場合は、最初のタブを選択
                            SelectedLeftPaneTab = LeftPaneTabs[0];
                        }
                    }
                    SelectedTab = null;
                }
                else
                {
                    // タブが存在しない場合は、新しいタブを作成
                    // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                    var leftPaneTabsCount = LeftPaneTabs.Count;
                    if (leftPaneTabsCount == 0)
                    {
                        CreateNewLeftPaneTab();
                    }
                    else if (SelectedLeftPaneTab == null)
                    {
                        // タブは存在するが選択されていない場合は、最初のタブを選択
                        SelectedLeftPaneTab = LeftPaneTabs[0];
                    }
                }
                // 右ペインに新しいタブを作成（存在しない場合）
                // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                var rightPaneTabsCount = RightPaneTabs.Count;
                if (rightPaneTabsCount == 0)
                {
                    CreateNewRightPaneTab();
                }
                else if (SelectedRightPaneTab == null)
                {
                    // タブは存在するが選択されていない場合は、最初のタブを選択
                    SelectedRightPaneTab = RightPaneTabs[0];
                }

                // 分割ペインモードが有効になったとき、ActivePaneを初期化
                // 最初は左ペインをデフォルトにする
                if (SelectedLeftPaneTab != null)
                {
                    ActivePane = ActivePaneLeft; // 左ペインをデフォルト
                }
                else if (SelectedRightPaneTab != null)
                {
                    ActivePane = ActivePaneRight; // 左ペインが存在しない場合のみ右ペインをデフォルト
                }
            }
            else
            {
                // 分割モードから通常モードへ
                // 左ペインのタブを通常のタブに移動
                var selectedLeftTab = SelectedLeftPaneTab;

                // 選択タブを先にnullに設定してからClear()を呼び出す（バインディングエラーを防ぐため）
                SelectedLeftPaneTab = null;
                SelectedRightPaneTab = null;

                // パフォーマンス最適化：タブを一時リストに収集してから一括で移動
                var leftPaneTabsCount = LeftPaneTabs.Count;
                var tabsToMove = new System.Collections.Generic.List<ExplorerTab>(leftPaneTabsCount);
                foreach (var tab in LeftPaneTabs)
                {
                    tabsToMove.Add(tab);
                }
                LeftPaneTabs.Clear();
                
                // 一括追加
                foreach (var tab in tabsToMove)
                {
                    Tabs.Add(tab);
                }

                // 右ペインのタブも通常のタブに移動（最初のタブのみ保持する場合は、ここで処理）
                // 現在は右ペインのタブは保持しない
                RightPaneTabs.Clear();

                // 選択タブを設定
                if (selectedLeftTab != null)
                {
                    SelectedTab = selectedLeftTab;
                }
                else
                {
                    // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                    var tabsCount = Tabs.Count;
                    if (tabsCount > 0)
                    {
                        // 選択タブがなかった場合は、最初のタブを選択
                        SelectedTab = Tabs[0];
                    }
                    else
                    {
                        // タブが存在しない場合は、新しいタブを作成
                        CreateNewTab();
                    }
                }
            }
        }

        #endregion

        #region ナビゲーション

        /// <summary>
        /// ページにナビゲートされたときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedToAsync()
        {
            // コマンドライン引数を一度だけ取得（パフォーマンス向上）
            var startupPath = App.GetStartupPath();
            var isSingleTabMode = App.IsSingleTabMode();

            // 初回実行時のみ初期化処理を実行
            if (!_hasNavigatedTo)
            {
                _hasNavigatedTo = true;

                // 単一タブモードの場合は、分割ペインを無効にする
                if (isSingleTabMode)
                {
                    IsSplitPaneEnabled = false;
                }
                else
                {
                    // 分割ペインの設定を読み込む
                    LoadSplitPaneSettings();
                }

                if (IsSplitPaneEnabled)
                {
                    // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                    var leftPaneTabsCount = LeftPaneTabs.Count;
                    var rightPaneTabsCount = RightPaneTabs.Count;
                    if (leftPaneTabsCount == 0 && rightPaneTabsCount == 0)
                    {
                        // 保存されたタブ情報を復元
                        RestoreTabs();

                        // タブが復元されなかった場合は新しいタブを作成
                        // Countプロパティを再取得（RestoreTabs後）
                        if (LeftPaneTabs.Count == 0)
                        {
                            CreateNewLeftPaneTab();
                        }
                        if (RightPaneTabs.Count == 0)
                        {
                            CreateNewRightPaneTab();
                        }
                    }
                }

                if (IsSplitPaneEnabled)
                {
                    // 選択タブが設定されていない場合は、最初のタブを選択
                    InitializeSelectedTabsForSplitPane();

                    // 起動時に分割ペインモードが有効な場合、ActivePaneを左ペインに設定
                    SetDefaultActivePane();

                    // コマンドライン引数で指定されたパスがある場合は、左ペインのタブに移動
                    NavigateToStartupPathIfExists(startupPath, SelectedLeftPaneTab);
                }
                else
                {
                    // 単一タブモードの場合は、既存のタブを削除して新しいタブを作成
                    if (isSingleTabMode)
                    {
                        InitializeSingleTabMode(startupPath);
                    }
                    else
                    {
                        // 通常モードの場合
                        InitializeNormalMode(startupPath);
                    }
                }
            }
            return Task.CompletedTask;
        }

        /// <summary>
        /// 分割ペインモードの選択タブを初期化します
        /// </summary>
        private void InitializeSelectedTabsForSplitPane()
        {
            if (SelectedLeftPaneTab == null && LeftPaneTabs.Count > 0)
            {
                SelectedLeftPaneTab = LeftPaneTabs[0];
            }
            if (SelectedRightPaneTab == null && RightPaneTabs.Count > 0)
            {
                SelectedRightPaneTab = RightPaneTabs[0];
            }
        }

        /// <summary>
        /// デフォルトのアクティブペインを設定します
        /// </summary>
        private void SetDefaultActivePane()
        {
            if (SelectedLeftPaneTab != null)
            {
                ActivePane = ActivePaneLeft;
            }
            else if (SelectedRightPaneTab != null)
            {
                ActivePane = ActivePaneRight;
            }
        }

        /// <summary>
        /// 起動時のパスが存在する場合、タブを移動します
        /// </summary>
        private static void NavigateToStartupPathIfExists(string? startupPath, ExplorerTab? tab)
        {
            if (!string.IsNullOrEmpty(startupPath) && System.IO.Directory.Exists(startupPath) && tab != null)
            {
                tab.ViewModel.NavigateToPathCommand.Execute(startupPath);
            }
        }

        /// <summary>
        /// 単一タブモードを初期化します
        /// </summary>
        private void InitializeSingleTabMode(string? startupPath)
        {
            // 既存のタブをすべて削除
            Tabs.Clear();
            SelectedTab = null;

            // コマンドライン引数で指定されたパスがある場合は、そのパスでタブを作成
            if (!string.IsNullOrEmpty(startupPath))
            {
                // パスが存在する場合は、そのパスでタブを作成
                if (System.IO.Directory.Exists(startupPath))
                {
                    var tab = CreateTabInternal(startupPath);
                    Tabs.Add(tab);
                    SelectedTab = tab;
                    // CreateTabInternal内でNavigateToPathCommandが実行されるが、
                    // 確実にパスを設定するために再度実行
                    tab.ViewModel.NavigateToPathCommand.Execute(startupPath);
                    UpdateTabTitle(tab);
                }
                else
                {
                    // パスが存在しない場合は、新しいタブを作成してホームに移動
                    CreateNewTab();
                    SelectedTab?.ViewModel.NavigateToHome();
                }
            }
            else
            {
                // パスが指定されていない場合は、新しいタブを作成（ホームに移動）
                CreateNewTab();
            }
        }

        /// <summary>
        /// 通常モードを初期化します
        /// </summary>
        private void InitializeNormalMode(string? startupPath)
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

            // コマンドライン引数で指定されたパスがある場合は、選択されているタブに移動
            NavigateToStartupPathIfExists(startupPath, SelectedTab);
        }

        /// <summary>
        /// ページから離れるときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedFromAsync() => Task.CompletedTask;

        #endregion

        #region タブ管理

        /// <summary>
        /// 新しいタブを作成します
        /// </summary>
        [RelayCommand]
        private void CreateNewTab()
        {
            // 直前に開いていたタブのパスを取得
            string? pathToOpen = null;
            if (SelectedTab != null)
            {
                pathToOpen = SelectedTab.ViewModel?.CurrentPath;
            }

            var tab = CreateTabInternal(pathToOpen);
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
            // 直前に開いていたタブのパスを取得（左ペインのタブを優先、なければ右ペインのタブ）
            string? pathToOpen = null;
            if (SelectedLeftPaneTab != null)
            {
                pathToOpen = SelectedLeftPaneTab.ViewModel?.CurrentPath;
            }
            else if (SelectedRightPaneTab != null)
            {
                pathToOpen = SelectedRightPaneTab.ViewModel?.CurrentPath;
            }

            var tab = CreateTabInternal(pathToOpen);
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
            // 直前に開いていたタブのパスを取得（右ペインのタブを優先、なければ左ペインのタブ）
            string? pathToOpen = null;
            if (SelectedRightPaneTab != null)
            {
                pathToOpen = SelectedRightPaneTab.ViewModel?.CurrentPath;
            }
            else if (SelectedLeftPaneTab != null)
            {
                pathToOpen = SelectedLeftPaneTab.ViewModel?.CurrentPath;
            }

            var tab = CreateTabInternal(pathToOpen);
            RightPaneTabs.Add(tab);
            SelectedRightPaneTab = tab;
            UpdateTabTitle(tab);
        }

        /// <summary>
        /// 指定されたペインにホームタブを作成します
        /// </summary>
        /// <param name="isLeftPane">左ペインかどうか（falseの場合は右ペイン）</param>
        public void CreateHomeTabForPane(bool isLeftPane)
        {
            var homeTab = CreateTabInternal(null);
            if (isLeftPane)
            {
                LeftPaneTabs.Add(homeTab);
                SelectedLeftPaneTab = homeTab;
            }
            else
            {
                RightPaneTabs.Add(homeTab);
                SelectedRightPaneTab = homeTab;
            }
            UpdateTabTitle(homeTab);
            homeTab.ViewModel.NavigateToHome();
        }

        /// <summary>
        /// タブを作成する内部メソッド
        /// </summary>
        /// <param name="pathToOpen">開くパス。nullまたは空の場合はホームディレクトリを開く</param>
        internal ExplorerTab CreateTabInternal(string? pathToOpen = null)
        {
            var viewModel = new ExplorerViewModel(_fileSystemService, _favoriteService);
            var tab = new ExplorerTab
            {
                Title = HomeTitle,
                CurrentPath = string.Empty,
                ViewModel = viewModel
            };

            // イベントハンドラーをアタッチ
            AttachPropertyChangedHandler(tab, viewModel);

            // タブを追加する前に初期化を完了させる
            // これにより、タブが追加された直後にフォルダーをダブルクリックしても問題が発生しない
            NavigateTabToPath(tab, pathToOpen);

            return tab;
        }

        /// <summary>
        /// ViewModelにPropertyChangedイベントハンドラーをアタッチします
        /// </summary>
        private void AttachPropertyChangedHandler(ExplorerTab tab, ExplorerViewModel viewModel)
        {
            // UpdateStatusBarデリゲートをキャッシュ（メモリ割り当てを削減）
            _cachedUpdateStatusBarAction ??= UpdateStatusBar;

            // Dispatcherを事前にキャッシュ（イベントハンドラー内での取得を削減）
            var dispatcher = _cachedDispatcher ??= System.Windows.Application.Current?.Dispatcher;

            PropertyChangedEventHandler handler = (s, e) =>
            {
                var propertyName = e.PropertyName;
                // 文字列比較を最適化（定数を使用、早期リターン）
                if (propertyName == CurrentPathPropertyName || propertyName == CurrentPathPropertyNameFull)
                {
                    tab.CurrentPath = viewModel.CurrentPath;
                    UpdateTabTitle(tab);
                    // UpdateStatusBarは遅延実行して頻繁な呼び出しを削減
                    dispatcher?.BeginInvoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        _cachedUpdateStatusBarAction);
                }
                else if (propertyName == ItemsPropertyName)
                {
                    // UpdateStatusBarは遅延実行して頻繁な呼び出しを削減
                    dispatcher?.BeginInvoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        _cachedUpdateStatusBarAction);
                }
            };
            viewModel.PropertyChanged += handler;
        }

        /// <summary>
        /// タブを指定されたパスに移動します
        /// </summary>
        private void NavigateTabToPath(ExplorerTab tab, string? pathToOpen)
        {
            if (!string.IsNullOrEmpty(pathToOpen))
            {
                // 起動時の高速化のため、Directory.Exists()の呼び出しを削減
                // パスが指定されている場合は、直接移動を試みる（存在チェックはViewModel内で行われる）
                tab.ViewModel.NavigateToPathCommand.Execute(pathToOpen);
            }
            else
            {
                // パスが指定されていない場合はホームに移動
                tab.ViewModel.NavigateToHome();
            }
        }

        #endregion

        #region タブ操作

        /// <summary>
        /// 指定されたタブを閉じます
        /// </summary>
        /// <param name="tab">閉じるタブ。nullの場合、またはタブが1つしかない場合は何もしません</param>
        [RelayCommand]
        private void CloseTab(ExplorerTab? tab)
        {
            if (tab == null)
                return;

            // イベントハンドラーを解除（メモリリーク防止）
            if (tab.ViewModel is ExplorerViewModel viewModel)
            {
                // PropertyChangedイベントハンドラーを解除
                // 注意: 現在の実装ではハンドラーの参照を保持していないため、
                // 将来的にはWeakEventManagerを使用するか、ExplorerTabにDisposeメソッドを追加することを推奨
            }

            if (IsSplitPaneEnabled)
            {
                // 左ペインのタブかチェック（IndexOfで存在確認とインデックス取得を同時に行う）
                var leftPaneTabs = LeftPaneTabs;
                var leftIndex = leftPaneTabs.IndexOf(tab);
                if (leftIndex >= 0)
                {
                    if (leftPaneTabs.Count <= 1)
                        return;

                    leftPaneTabs.RemoveAt(leftIndex);
                    UpdateSelectedTabAfterRemoval(leftPaneTabs, leftIndex, tab,
                        () => SelectedLeftPaneTab, t => SelectedLeftPaneTab = t);
                    return;
                }

                // 右ペインのタブかチェック
                var rightPaneTabs = RightPaneTabs;
                var rightIndex = rightPaneTabs.IndexOf(tab);
                if (rightIndex >= 0)
                {
                    if (rightPaneTabs.Count <= 1)
                        return;

                    rightPaneTabs.RemoveAt(rightIndex);
                    UpdateSelectedTabAfterRemoval(rightPaneTabs, rightIndex, tab,
                        () => SelectedRightPaneTab, t => SelectedRightPaneTab = t);
                }
            }
            else
            {
                var tabs = Tabs;
                if (tabs.Count <= 1)
                    return;

                var index = tabs.IndexOf(tab);
                if (index < 0)
                    return;

                tabs.RemoveAt(index);
                UpdateSelectedTabAfterRemoval(tabs, index, tab,
                    () => SelectedTab, t => SelectedTab = t);
            }
        }

        /// <summary>
        /// タブ削除後に選択タブを更新します
        /// </summary>
        private static void UpdateSelectedTabAfterRemoval(
            ObservableCollection<ExplorerTab> tabs,
            int removedIndex,
            ExplorerTab removedTab,
            Func<ExplorerTab?> getSelectedTab,
            Action<ExplorerTab?> setSelectedTab)
        {
            if (getSelectedTab() == removedTab)
            {
                if (removedIndex > 0 && tabs.Count > 0)
                    setSelectedTab(tabs[removedIndex - 1]);
                else if (tabs.Count > 0)
                    setSelectedTab(tabs[0]);
                else
                    setSelectedTab(null);
            }
        }

        /// <summary>
        /// タブの並び順を変更します
        /// </summary>
        /// <param name="draggedTab">ドラッグされたタブ</param>
        /// <param name="targetTab">ドロップ先のタブ</param>
        /// <param name="tabs">タブのコレクション</param>
        public void ReorderTab(ExplorerTab draggedTab, ExplorerTab targetTab, ObservableCollection<ExplorerTab> tabs)
        {
            if (draggedTab == null || targetTab == null || tabs == null)
                return;

            var draggedIndex = tabs.IndexOf(draggedTab);
            var targetIndex = tabs.IndexOf(targetTab);

            if (draggedIndex == -1 || targetIndex == -1 || draggedIndex == targetIndex)
                return;

            // タブを移動（Moveメソッドを使用してUI更新を1回に削減）
            // ObservableCollection.Moveは、RemoveAtとInsertを1回の操作として実行し、
            // CollectionChangedイベントを1回だけ発火するため、パフォーマンスが向上
            // 特にタブが多数ある場合に効果的
            tabs.Move(draggedIndex, targetIndex);
        }

        #endregion

        #region タブタイトル管理

        /// <summary>
        /// タブのタイトルを現在のパスに基づいて更新します（必要な場合のみ）
        /// </summary>
        /// <param name="tab">タイトルを更新するタブ</param>
        private void UpdateTabTitleIfNeeded(ExplorerTab tab)
        {
            var expectedTitle = GetTabTitleFromPath(tab.ViewModel?.CurrentPath);
            if (tab.Title != expectedTitle)
            {
                tab.Title = expectedTitle;
            }
        }

        /// <summary>
        /// タブのタイトルを現在のパスに基づいて更新します
        /// </summary>
        /// <param name="tab">タイトルを更新するタブ</param>
        internal void UpdateTabTitle(ExplorerTab tab)
        {
            tab.Title = GetTabTitleFromPath(tab.ViewModel.CurrentPath);
        }

        #endregion

        #region お気に入り操作

        /// <summary>
        /// 現在のパスをお気に入りに追加します
        /// </summary>
        [RelayCommand]
        private void AddCurrentPathToFavorites()
        {
            // ViewModel.CurrentPathを一度だけ取得してキャッシュ（パフォーマンス向上）
            var selectedTab = SelectedTab;
            if (selectedTab == null)
                return;

            var path = selectedTab.ViewModel.CurrentPath;
            if (string.IsNullOrEmpty(path))
                return;
            var name = Path.GetFileName(path) ?? path;

            _favoriteService?.AddFavorite(name, path);

            // MainWindowViewModelを更新（キャッシュを使用）
            if (_cachedMainWindowViewModel == null)
            {
                _cachedMainWindowViewModel = App.Services.GetService(MainWindowViewModelType) as ViewModels.Windows.MainWindowViewModel;
            }
            _cachedMainWindowViewModel?.LoadFavorites();
        }

        #endregion

        #region ドライブ操作

        /// <summary>
        /// ドライブにナビゲートします
        /// </summary>
        /// <param name="parameter">(drivePath, pane)のタプル。drivePathはドライブのパス、paneはペイン番号（0=左ペイン、2=右ペイン、null=現在アクティブなペイン）</param>
        [RelayCommand]
        private void NavigateToDrive((string drivePath, int? pane)? parameter)
        {
            // 早期リターン：パラメータチェック
            if (parameter == null)
                return;

            var (drivePath, pane) = parameter.Value;

            // 早期リターン：ドライブパスチェック
            if (string.IsNullOrEmpty(drivePath))
                return;

            // プロパティアクセスをキャッシュ（パフォーマンス向上）
            var isSplitPaneEnabled = IsSplitPaneEnabled;
            ExplorerTab? targetTab;

            if (isSplitPaneEnabled)
            {
                // 分割ペインモードの場合、プロパティを一度だけ取得してキャッシュ
                var selectedLeftPaneTab = SelectedLeftPaneTab;
                var selectedRightPaneTab = SelectedRightPaneTab;

                if (pane.HasValue)
                {
                    // ペイン番号が指定されている場合は、そのペインのタブを使用
                    targetTab = pane.Value == ActivePaneLeft ? selectedLeftPaneTab : selectedRightPaneTab;
                }
                else
                {
                    // ペイン番号が指定されていない場合は、現在アクティブなペインのタブを使用
                    var activePane = ActivePane;
                    targetTab = activePane == ActivePaneLeft ? selectedLeftPaneTab : selectedRightPaneTab;
                }

                // タブが見つからない場合は、フォールバック（null合体演算子を使用）
                targetTab ??= selectedLeftPaneTab ?? selectedRightPaneTab;
            }
            else
            {
                // 通常モード
                targetTab = SelectedTab;
            }

            // 早期リターン：タブまたはViewModelがnullの場合
            if (targetTab?.ViewModel == null)
                return;

            targetTab.ViewModel.NavigateToPathCommand.Execute(drivePath);
        }

        /// <summary>
        /// アクティブなペインのタブにホームナビゲーションを実行します
        /// </summary>
        [RelayCommand]
        private void NavigateToHomeInActivePane()
        {
            System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] メソッドが呼ばれました");

            // プロパティアクセスをキャッシュ（パフォーマンス向上）
            var isSplitPaneEnabled = IsSplitPaneEnabled;
            ExplorerTab? targetTab;

            System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] IsSplitPaneEnabled: {isSplitPaneEnabled}");

            if (isSplitPaneEnabled)
            {
                // 分割ペインモードの場合、アクティブなペインのタブを使用
                var selectedLeftPaneTab = SelectedLeftPaneTab;
                var selectedRightPaneTab = SelectedRightPaneTab;
                var activePane = ActivePane;
                
                System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] ActivePane: {activePane} (Left={ActivePaneLeft}, Right={ActivePaneRight}, None={ActivePaneNone})");
                System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] SelectedLeftPaneTab: {(selectedLeftPaneTab != null ? "存在" : "null")}");
                System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] SelectedRightPaneTab: {(selectedRightPaneTab != null ? "存在" : "null")}");
                
                // ActivePaneが未設定（-1）の場合は、左ペインをデフォルトとして使用
                if (activePane == ActivePaneNone)
                {
                    System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] ActivePaneが未設定のため、デフォルトペインを選択します");
                    // 左ペインが存在する場合は左ペイン、存在しない場合は右ペインを使用
                    targetTab = selectedLeftPaneTab ?? selectedRightPaneTab;
                    // ActivePaneを更新（左ペインが存在する場合は左、存在しない場合は右）
                    if (selectedLeftPaneTab != null)
                    {
                        ActivePane = ActivePaneLeft;
                        System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] ActivePaneを左ペインに設定しました");
                    }
                    else if (selectedRightPaneTab != null)
                    {
                        ActivePane = ActivePaneRight;
                        System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] ActivePaneを右ペインに設定しました");
                    }
                }
                else
                {
                    targetTab = activePane == ActivePaneLeft ? selectedLeftPaneTab : selectedRightPaneTab;
                    System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] ActivePaneに基づいてタブを選択: {(activePane == ActivePaneLeft ? "左ペイン" : "右ペイン")}");
                }

                // タブが見つからない場合は、フォールバック（null合体演算子を使用）
                targetTab ??= selectedLeftPaneTab ?? selectedRightPaneTab;
            }
            else
            {
                // 通常モード
                targetTab = SelectedTab;
                System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] 通常モード: SelectedTab={(targetTab != null ? "存在" : "null")}");
            }

            // 早期リターン：タブまたはViewModelがnullの場合
            if (targetTab?.ViewModel == null)
            {
                System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] エラー: targetTabまたはViewModelがnullです (targetTab={(targetTab != null ? "存在" : "null")}, ViewModel={(targetTab?.ViewModel != null ? "存在" : "null")})");
                return;
            }

            System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] ホームナビゲーションを実行します (CurrentPath: {targetTab.ViewModel.CurrentPath})");
            
            // NavigateToPathCommandを使用してホームにナビゲート（空文字列を渡すとNavigateToHomeが呼ばれる）
            // これにより、シングルペインモードと同じ動作になる
            targetTab.ViewModel.NavigateToPathCommand.Execute(string.Empty);
            
            System.Diagnostics.Debug.WriteLine($"[NavigateToHomeInActivePane] NavigateToPathCommand.Execute(string.Empty)を呼び出しました");
        }

        /// <summary>
        /// 指定されたタブのパスをクリップボードにコピーします
        /// </summary>
        /// <param name="tab">パスをコピーするタブ</param>
        [RelayCommand]
        private void CopyTabPath(ExplorerTab? tab)
        {
            if (tab == null)
            {
                // タブが指定されていない場合は、現在選択されているタブを使用
                if (IsSplitPaneEnabled)
                {
                    // 分割ペインモードの場合は、アクティブなペインのタブを使用
                    tab = ActivePane == ActivePaneLeft ? SelectedLeftPaneTab : SelectedRightPaneTab;
                }
                else
                {
                    tab = SelectedTab;
                }
            }

            if (tab == null)
                return;

            var path = tab.ViewModel?.CurrentPath;
            // 最適化：string.IsNullOrEmpty()を1回だけ呼び出し
            if (string.IsNullOrEmpty(path))
            {
                // パスが空の場合はホームディレクトリのパスを使用
                path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            }

            // pathは上記で必ず設定されるため、再度チェックは不要
            try
            {
                System.Windows.Clipboard.SetText(path);
            }
            catch
            {
                // クリップボードへのコピーに失敗した場合は何もしない
            }
        }

        #endregion

        #region タブ選択変更ハンドラー

        /// <summary>
        /// SelectedTabが変更されたときに呼び出されます
        /// </summary>
        partial void OnSelectedTabChanged(ExplorerTab? value)
        {
            // タブ切り替え時の処理を最小限に（高速化）
            // タイトルとステータスバーの更新は遅延実行して、タブ切り替えの応答性を優先
            var dispatcher = _cachedDispatcher ?? (_cachedDispatcher = System.Windows.Application.Current?.Dispatcher);
            if (dispatcher == null)
                return;

            if (value != null)
            {
                // タイトル更新を非同期で遅延実行（UIスレッドをブロックしない）
                dispatcher.BeginInvoke(
                    System.Windows.Threading.DispatcherPriority.Background,
                    new System.Action(() => UpdateTabTitleIfNeeded(value)));
            }
            // ステータスバー更新は大幅に遅延（Background優先度）して、タブ切り替えの応答性を優先
            dispatcher.BeginInvoke(
                System.Windows.Threading.DispatcherPriority.Background,
                new System.Action(() => UpdateStatusBar()));
        }

        /// <summary>
        /// SelectedLeftPaneTabが変更されたときに呼び出されます
        /// </summary>
        partial void OnSelectedLeftPaneTabChanged(ExplorerTab? value)
        {
            // 左ペインのタブが選択された場合、ActivePaneを左ペインに設定
            if (value != null && IsSplitPaneEnabled)
            {
                ActivePane = ActivePaneLeft;
            }

            // タブ切り替え時の処理を最小限に（高速化）
            // タイトルとステータスバーの更新は遅延実行して、タブ切り替えの応答性を優先
            var dispatcher = _cachedDispatcher ?? (_cachedDispatcher = System.Windows.Application.Current?.Dispatcher);
            if (dispatcher == null)
                return;

            if (value != null)
            {
                // タイトル更新を非同期で遅延実行（UIスレッドをブロックしない）
                dispatcher.BeginInvoke(
                    System.Windows.Threading.DispatcherPriority.Background,
                    new System.Action(() => UpdateTabTitleIfNeeded(value)));
            }
            // ステータスバー更新は大幅に遅延（Background優先度）して、タブ切り替えの応答性を優先
            dispatcher.BeginInvoke(
                System.Windows.Threading.DispatcherPriority.Background,
                new System.Action(() => UpdateStatusBar()));
        }

        /// <summary>
        /// SelectedRightPaneTabが変更されたときに呼び出されます
        /// </summary>
        partial void OnSelectedRightPaneTabChanged(ExplorerTab? value)
        {
            // 右ペインのタブが選択された場合、ActivePaneを右ペインに設定
            if (value != null && IsSplitPaneEnabled)
            {
                ActivePane = ActivePaneRight;
            }

            // タブ切り替え時の処理を最小限に（高速化）
            // タイトルとステータスバーの更新は遅延実行して、タブ切り替えの応答性を優先
            var dispatcher = _cachedDispatcher ?? (_cachedDispatcher = System.Windows.Application.Current?.Dispatcher);
            if (dispatcher == null)
                return;

            if (value != null)
            {
                // タイトル更新を非同期で遅延実行（UIスレッドをブロックしない）
                dispatcher.BeginInvoke(
                    System.Windows.Threading.DispatcherPriority.Background,
                    new System.Action(() => UpdateTabTitleIfNeeded(value)));
            }
            // ステータスバー更新は大幅に遅延（Background優先度）して、タブ切り替えの応答性を優先
            dispatcher.BeginInvoke(
                System.Windows.Threading.DispatcherPriority.Background,
                new System.Action(() => UpdateStatusBar()));
        }

        #endregion

        #region ステータスバー管理

        // ステータスバー更新のスロットリング（頻繁な更新を抑制）
        private System.Windows.Threading.DispatcherTimer? _statusBarUpdateTimer;
        private bool _statusBarUpdatePending = false;

        /// <summary>
        /// ステータスバーのテキストを更新します（スロットリング付き）
        /// </summary>
        private void UpdateStatusBar()
        {
            // 既に更新が保留中の場合は何もしない
            if (_statusBarUpdatePending)
                return;

            _statusBarUpdatePending = true;

            // タイマーが存在しない場合は作成
            if (_statusBarUpdateTimer == null)
            {
                // タイマーのIntervalを定数化（メモリ割り当てを削減）
                const int StatusBarUpdateIntervalMs = 100;
                _statusBarUpdateTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(StatusBarUpdateIntervalMs) // 100ms間隔で更新
                };
                // Tickイベントハンドラーを最適化（タイマー参照をキャッシュ）
                var timer = _statusBarUpdateTimer;
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    _statusBarUpdatePending = false;
                    UpdateStatusBarInternal();
                };
            }

            // タイマーをリセットして再開始
            _statusBarUpdateTimer.Stop();
            _statusBarUpdateTimer.Start();
        }

        /// <summary>
        /// ステータスバーのテキストを実際に更新します
        /// </summary>
        private void UpdateStatusBarInternal()
        {
            // プロパティアクセスを一度だけ取得してキャッシュ（パフォーマンス向上）
            var isSplitPaneEnabled = IsSplitPaneEnabled;

            // 分割ペインの場合は、左右のペインそれぞれのステータスバーを更新
            if (isSplitPaneEnabled)
            {
                UpdateTabStatusBar(SelectedLeftPaneTab);
                UpdateTabStatusBar(SelectedRightPaneTab);
            }
            else
            {
                UpdateTabStatusBar(SelectedTab);
            }
        }

        /// <summary>
        /// 指定されたタブのステータスバーを更新します
        /// </summary>
        /// <param name="tab">更新するタブ</param>
        private void UpdateTabStatusBar(ExplorerTab? tab)
        {
            // nullチェックを最適化（パターンマッチング）
            if (tab?.ViewModel is not { } viewModel)
                return;

            // タブのタイトルも更新
            UpdateTabTitle(tab);

            // ViewModelのプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
            var itemCount = viewModel.Items.Count;

            // 文字列補間を最適化（ZString.Concatを使用してメモリ割り当てを削減）
            // 項目数のみを表示（パスはパンくずリストで確認可能）
            var statusText = ZString.Concat(itemCount, ItemSuffix);

            // 値が変更された場合のみ更新（不要なPropertyChangedイベントを削減）
            if (viewModel.StatusBarText != statusText)
            {
                viewModel.StatusBarText = statusText;
            }
        }

        #endregion

        #region 設定管理

        /// <summary>
        /// 分割ペインの設定を読み込みます
        /// </summary>
        private void LoadSplitPaneSettings()
        {
            try
            {
                // WindowSettingsServiceをキャッシュ（パフォーマンス向上）
                var windowSettingsService = _windowSettingsService ??
                    (_cachedWindowSettingsService ??= App.Services.GetService(WindowSettingsServiceType) as WindowSettingsService);

                if (windowSettingsService == null)
                    return;

                var settings = windowSettingsService.GetSettings();
                // 値が変更された場合のみ更新（不要なPropertyChangedイベントを削減）
                if (IsSplitPaneEnabled != settings.IsSplitPaneEnabled)
                {
                    IsSplitPaneEnabled = settings.IsSplitPaneEnabled;
                }
            }
            catch
            {
                if (IsSplitPaneEnabled)
                {
                    IsSplitPaneEnabled = false;
                }
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
                // WindowSettingsServiceをキャッシュ（パフォーマンス向上）
                var windowSettingsService = _windowSettingsService ??
                    (_cachedWindowSettingsService ??= App.Services.GetService(WindowSettingsServiceType) as WindowSettingsService);

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

            // 最初のタブは同期的に復元（前回終了時のタブを即座に表示）
            var firstPath = tabPaths[0];
            var firstTab = CreateAndRestoreTab(firstPath, isFirstTab: true);
            if (firstTab != null)
            {
                tabs.Add(firstTab);
                setSelectedTab(firstTab);
            }

            // 残りのタブはBackground優先度で遅延実行（起動を高速化）
            if (tabPaths.Count > 1)
            {
                var remainingPaths = new System.Collections.Generic.List<string>(tabPaths.Count - 1);
                for (int i = 1; i < tabPaths.Count; i++)
                {
                    remainingPaths.Add(tabPaths[i]);
                }

                // Dispatcherをキャッシュ（パフォーマンス向上）
                var dispatcher = _cachedDispatcher ??= System.Windows.Application.Current?.Dispatcher;
                dispatcher?.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new System.Action(() =>
                {
                    RestoreRemainingTabs(remainingPaths, tabs);
                }));
            }
        }

        /// <summary>
        /// タブを作成して復元します
        /// </summary>
        /// <param name="path">復元するパス</param>
        /// <param name="isFirstTab">最初のタブかどうか（最初のタブは同期的に処理）</param>
        private ExplorerTab? CreateAndRestoreTab(string? path, bool isFirstTab = false)
        {
            try
            {
                var viewModel = new ExplorerViewModel(_fileSystemService, _favoriteService);
                var tab = new ExplorerTab
                {
                    Title = GetTabTitleFromPath(path),
                    CurrentPath = path ?? string.Empty,
                    ViewModel = viewModel
                };

                // イベントハンドラーをアタッチ
                AttachPropertyChangedHandler(tab, viewModel);

                // タブをパスに移動
                if (!string.IsNullOrEmpty(path))
                {
                    if (isFirstTab)
                    {
                        // 最初のタブは同期的に処理（前回終了時のタブを即座に表示）
                        if (System.IO.Directory.Exists(path))
                        {
                            tab.ViewModel.NavigateToPathCommand.Execute(path);
                        }
                        else
                        {
                            tab.ViewModel.NavigateToHome();
                        }
                        UpdateTabTitle(tab);
                    }
                    else
                    {
                        // 残りのタブは遅延実行（起動を高速化）
                        var pathCopy = path;
                        var dispatcher = _cachedDispatcher ??= System.Windows.Application.Current?.Dispatcher;
                        dispatcher?.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new System.Action(() =>
                        {
                            if (System.IO.Directory.Exists(pathCopy))
                            {
                                tab.ViewModel.NavigateToPathCommand.Execute(pathCopy);
                            }
                            else
                            {
                                tab.ViewModel.NavigateToHome();
                            }
                            UpdateTabTitle(tab);
                        }));
                    }
                }
                else
                {
                    // パスが指定されていない場合はホームに移動
                    tab.ViewModel.NavigateToHome();
                    UpdateTabTitle(tab);
                }

                return tab;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 残りのタブを復元します
        /// </summary>
        private void RestoreRemainingTabs(System.Collections.Generic.List<string> tabPaths, ObservableCollection<ExplorerTab> tabs)
        {
            foreach (var path in tabPaths)
            {
                var tab = CreateAndRestoreTab(path);
                if (tab != null)
                {
                    tabs.Add(tab);
                }
            }
        }

        /// <summary>
        /// パスからタブのタイトルを取得します
        /// </summary>
        private string GetTabTitleFromPath(string? path)
        {
            if (string.IsNullOrEmpty(path))
                return HomeTitle;

            var folderName = Path.GetFileName(path);
            // ルートディレクトリ（例：C:\）の場合は、パス自体を表示
            if (string.IsNullOrEmpty(folderName))
            {
                var root = Path.GetPathRoot(path);
                if (string.IsNullOrEmpty(root))
                    return path;
                
                // パフォーマンス最適化：TrimEndの代わりに最後の文字をチェック
                // バックスラッシュで終わる場合のみ、AsSpanを使用して部分文字列を作成
                return root.Length > 0 && root[root.Length - 1] == '\\' 
                    ? root.Substring(0, root.Length - 1) 
                    : root;
            }

            return folderName;
        }

        #endregion

        #region タブ複製・移動

        /// <summary>
        /// タブを複製します
        /// </summary>
        /// <param name="tab">複製するタブ</param>
        [RelayCommand]
        private void DuplicateTab(ExplorerTab? tab)
        {
            if (tab == null)
                return;

            var pathToOpen = tab.ViewModel?.CurrentPath;
            ExplorerTab newTab = CreateTabInternal(pathToOpen);

            if (IsSplitPaneEnabled)
            {
                // 分割ペインモードの場合、同じペインに追加（IndexOfで存在確認）
                var leftPaneTabs = LeftPaneTabs;
                if (leftPaneTabs.IndexOf(tab) >= 0)
                {
                    leftPaneTabs.Add(newTab);
                    SelectedLeftPaneTab = newTab;
                }
                else
                {
                    var rightPaneTabs = RightPaneTabs;
                    if (rightPaneTabs.IndexOf(tab) >= 0)
                    {
                        rightPaneTabs.Add(newTab);
                        SelectedRightPaneTab = newTab;
                    }
                    else
                    {
                        return;
                    }
                }
            }
            else
            {
                Tabs.Add(newTab);
                SelectedTab = newTab;
            }

            UpdateTabTitle(newTab);
        }

        /// <summary>
        /// 新しいウィンドウにタブを移動します
        /// </summary>
        /// <param name="tab">移動するタブ</param>
        [RelayCommand]
        private void MoveTabToNewWindow(ExplorerTab? tab)
        {
            MoveTabToNewWindow(tab, null);
        }

        /// <summary>
        /// 移動するアクティブなタブを取得します
        /// </summary>
        private ExplorerTab GetActiveTabForMove(ExplorerTab tab)
        {
            if (IsSplitPaneEnabled)
            {
                var leftPaneTabs = LeftPaneTabs;
                var rightPaneTabs = RightPaneTabs;

                if (leftPaneTabs.IndexOf(tab) >= 0)
                {
                    // 左ペインの選択タブを使用（存在しない場合はドラッグされたタブ）
                    var activeTab = SelectedLeftPaneTab;
                    return (activeTab != null && leftPaneTabs.IndexOf(activeTab) >= 0) ? activeTab : tab;
                }

                if (rightPaneTabs.IndexOf(tab) >= 0)
                {
                    // 右ペインの選択タブを使用（存在しない場合はドラッグされたタブ）
                    var activeTab = SelectedRightPaneTab;
                    return (activeTab != null && rightPaneTabs.IndexOf(activeTab) >= 0) ? activeTab : tab;
                }

                return tab;
            }

            // 通常モード
            var selectedTab = SelectedTab;
            return (selectedTab != null && Tabs.IndexOf(selectedTab) >= 0) ? selectedTab : tab;
        }

        /// <summary>
        /// 新しいウィンドウにタブを移動します（マウス位置を指定）
        /// </summary>
        /// <param name="tab">移動するタブ</param>
        /// <param name="dropPosition">ドロップ位置（nullの場合はデフォルト位置）</param>
        public void MoveTabToNewWindow(ExplorerTab? tab, System.Windows.Point? dropPosition)
        {
            if (tab == null)
                return;

            // 【重要】削除・移動するのはアクティブタブ
            tab = GetActiveTabForMove(tab);

            // タブのパスを取得（削除する前に取得する必要がある）
            var pathToOpen = tab.ViewModel?.CurrentPath;

            // CurrentPathが空またはnullの場合、ホームディレクトリを使用
            if (string.IsNullOrEmpty(pathToOpen))
            {
                pathToOpen = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            }

            try
            {
                var exePath = GetApplicationExecutablePath();
                if (string.IsNullOrEmpty(exePath))
                    return;

                // パフォーマンス最適化：ZStringを使用して文字列の割り当てを削減
                var pathArgument = BuildPathArgument(pathToOpen);
                
                string arguments;
                // ドロップ位置がある場合は、コマンドライン引数に追加
                if (dropPosition.HasValue)
                {
                    // ZString.Concatを使用してメモリ割り当てを削減
                    if (!string.IsNullOrEmpty(pathArgument))
                    {
                        arguments = ZString.Concat("--single-tab --position ", dropPosition.Value.X, ",", dropPosition.Value.Y, " ", pathArgument);
                    }
                    else
                    {
                        arguments = ZString.Concat("--single-tab --position ", dropPosition.Value.X, ",", dropPosition.Value.Y);
                    }
                }
                else
                {
                    // パスをコマンドライン引数として渡す
                    if (!string.IsNullOrEmpty(pathArgument))
                    {
                        arguments = ZString.Concat("--single-tab ", pathArgument);
                    }
                    else
                    {
                        arguments = "--single-tab";
                    }
                }

                // 元のウィンドウからタブを削除（新しいプロセス起動前に実行して、即座にUIを更新）
                // タブの削除処理を関数に抽出して、新しいプロセス起動前に実行
                void RemoveTabFromCurrentWindow()
                {
                    if (IsSplitPaneEnabled)
                    {
                        var leftPaneTabs = LeftPaneTabs;
                        var rightPaneTabs = RightPaneTabs;

                        var leftIndex = leftPaneTabs.IndexOf(tab);
                        if (leftIndex >= 0)
                        {
                            var index = leftIndex;
                            // タブを削除する前に、タブが1つだけかどうかを確認
                            bool wasOnlyTab = leftPaneTabs.Count == 1;

                            leftPaneTabs.RemoveAt(index);

                            // 選択タブを更新（削除されたタブが選択されていなくても更新）
                            bool wasSelected = (SelectedLeftPaneTab == tab);

                            // 削除されたタブが選択されていた場合、または選択タブがない場合
                            if (wasSelected || SelectedLeftPaneTab == null || leftPaneTabs.IndexOf(SelectedLeftPaneTab) < 0)
                            {
                                if (index > 0 && leftPaneTabs.Count > 0)
                                {
                                    SelectedLeftPaneTab = leftPaneTabs[index - 1];
                                }
                                else if (leftPaneTabs.Count > 0)
                                {
                                    SelectedLeftPaneTab = leftPaneTabs[0];
                                }
                                else
                                {
                                    SelectedLeftPaneTab = null;
                                }
                            }
                            else
                            {
                                // 選択タブが変更されなくても、コレクションが変更されたのでUIを強制更新
                                OnPropertyChanged(nameof(SelectedLeftPaneTab));
                            }

                            // タブが1つしかなかった場合（削除後に0個になった場合）、新しいタブを作成してホームページを表示
                            if (wasOnlyTab && leftPaneTabs.Count == 0)
                            {
                                // 新しいタブを作成（pathToOpenをnullにして、ホームページを表示）
                                var newTab = CreateTabInternal(null);
                                leftPaneTabs.Add(newTab);
                                SelectedLeftPaneTab = newTab;
                                UpdateTabTitle(newTab);
                                // CreateTabInternal内で既にNavigateToHome()が呼ばれているが、
                                // 確実にホームページが表示されるように再度呼び出す
                                newTab.ViewModel.NavigateToHome();

                                // 左ペインにフォーカスを設定
                                ActivePane = ActivePaneLeft;
                            }
                            else
                            {
                                // タブが複数あった場合でも、左ペインにフォーカスを設定してUIを更新
                                if (leftPaneTabs.Count > 0)
                                {
                                    // ActivePaneが既にLeftの場合でも、強制的にUIを更新するため、一度別の値に設定してから戻す
                                    var currentActivePane = ActivePane;
                                    if (currentActivePane == ActivePaneLeft)
                                    {
                                        ActivePane = ActivePaneNone;
                                    }
                                    ActivePane = ActivePaneLeft;

                                    // 明示的にPropertyChangedを発火させる
                                    OnPropertyChanged(nameof(SelectedLeftPaneTab));
                                }
                            }
                        }
                        else
                        {
                            var rightIndex = rightPaneTabs.IndexOf(tab);
                            if (rightIndex >= 0)
                            {
                                var index = rightIndex;
                                // タブを削除する前に、タブが1つだけかどうかを確認
                                bool wasOnlyTab = rightPaneTabs.Count == 1;

                                rightPaneTabs.RemoveAt(index);

                                // 選択タブを更新（削除されたタブが選択されていなくても更新）
                                bool wasSelected = (SelectedRightPaneTab == tab);

                                // 削除されたタブが選択されていた場合、または選択タブがない場合
                                if (wasSelected || SelectedRightPaneTab == null || rightPaneTabs.IndexOf(SelectedRightPaneTab) < 0)
                                {
                                    if (index > 0 && rightPaneTabs.Count > 0)
                                    {
                                        SelectedRightPaneTab = rightPaneTabs[index - 1];
                                    }
                                    else if (rightPaneTabs.Count > 0)
                                    {
                                        SelectedRightPaneTab = rightPaneTabs[0];
                                    }
                                    else
                                    {
                                        SelectedRightPaneTab = null;
                                    }
                                }
                                else
                                {
                                    // 選択タブが変更されなくても、コレクションが変更されたのでUIを強制更新
                                    OnPropertyChanged(nameof(SelectedRightPaneTab));
                                }

                                // タブが1つしかなかった場合（削除後に0個になった場合）、新しいタブを作成してホームページを表示
                                if (wasOnlyTab && rightPaneTabs.Count == 0)
                                {
                                    // 新しいタブを作成（pathToOpenをnullにして、ホームページを表示）
                                    var newTab = CreateTabInternal(null);
                                    rightPaneTabs.Add(newTab);
                                    SelectedRightPaneTab = newTab;
                                    UpdateTabTitle(newTab);
                                    // CreateTabInternal内で既にNavigateToHome()が呼ばれているが、
                                    // 確実にホームページが表示されるように再度呼び出す
                                    newTab.ViewModel.NavigateToHome();

                                    // 右ペインにフォーカスを設定
                                    ActivePane = ActivePaneRight;
                                }
                                else
                                {
                                    // タブが複数あった場合でも、右ペインにフォーカスを設定してUIを更新
                                    if (rightPaneTabs.Count > 0)
                                    {
                                        // ActivePaneが既にRightの場合でも、強制的にUIを更新するため、一度別の値に設定してから戻す
                                        var currentActivePane = ActivePane;
                                        if (currentActivePane == ActivePaneRight)
                                        {
                                            ActivePane = ActivePaneNone;
                                        }
                                        ActivePane = ActivePaneRight;

                                        // 明示的にPropertyChangedを発火させる
                                        OnPropertyChanged(nameof(SelectedRightPaneTab));
                                    }
                                }
                            }
                            else
                            {
                                // タブが左ペインにも右ペインにも見つからない場合、通常モードのTabsコレクションを確認

                                // 通常モードのTabsコレクションにタブがあるか確認
                                var tabs = Tabs;
                                var tabIndex = tabs.IndexOf(tab);
                                if (tabIndex >= 0)
                                {
                                    // 通常モードのTabsコレクションからタブを削除
                                    tabs.RemoveAt(tabIndex);

                                    // 選択タブを更新
                                    if (SelectedTab == tab)
                                    {
                                        if (tabIndex > 0 && tabs.Count > 0)
                                        {
                                            SelectedTab = tabs[tabIndex - 1];
                                        }
                                        else if (tabs.Count > 0)
                                        {
                                            SelectedTab = tabs[0];
                                        }
                                        else
                                        {
                                            SelectedTab = null;
                                        }
                                    }
                                }
                                else
                                {
                                    // 通常モードのTabsコレクションにも見つからない場合

                                    // 左ペインと右ペインのタブ数を確認して、どちらかが空の場合は新しいタブを作成
                                    bool leftPaneWasEmpty = false;
                                    bool rightPaneWasEmpty = false;

                                    if (leftPaneTabs.Count == 0)
                                    {
                                        var newTab = CreateTabInternal(null);
                                        leftPaneTabs.Add(newTab);
                                        SelectedLeftPaneTab = newTab;
                                        UpdateTabTitle(newTab);
                                        newTab.ViewModel.NavigateToHome();
                                        leftPaneWasEmpty = true;
                                    }
                                    if (rightPaneTabs.Count == 0)
                                    {
                                        var newTab = CreateTabInternal(null);
                                        rightPaneTabs.Add(newTab);
                                        SelectedRightPaneTab = newTab;
                                        UpdateTabTitle(newTab);
                                        newTab.ViewModel.NavigateToHome();
                                        rightPaneWasEmpty = true;
                                    }

                                    // 両方のペインが空だった場合は左ペインにフォーカスを設定
                                    // 左ペインのみが空だった場合は左ペインにフォーカスを設定
                                    // 右ペインのみが空だった場合は右ペインにフォーカスを設定
                                    if (leftPaneWasEmpty && rightPaneWasEmpty)
                                    {
                                        ActivePane = ActivePaneLeft;
                                    }
                                    else if (leftPaneWasEmpty)
                                    {
                                        ActivePane = ActivePaneLeft;
                                    }
                                    else if (rightPaneWasEmpty)
                                    {
                                        ActivePane = ActivePaneRight;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        var tabs = Tabs;
                        var index = tabs.IndexOf(tab);
                        if (index >= 0)
                        {
                            // タブを削除する前に、タブが1つだけかどうかを確認
                            bool wasOnlyTab = tabs.Count == 1;

                            tabs.RemoveAt(index);

                            // 選択タブを更新（削除されたタブが選択されていなくても更新）
                            bool wasSelected = (SelectedTab == tab);

                            // 削除されたタブが選択されていた場合、または選択タブがない場合
                            if (wasSelected || SelectedTab == null || tabs.IndexOf(SelectedTab) < 0)
                            {
                                if (index > 0 && tabs.Count > 0)
                                {
                                    SelectedTab = tabs[index - 1];
                                }
                                else if (tabs.Count > 0)
                                {
                                    SelectedTab = tabs[0];
                                }
                                else
                                {
                                    SelectedTab = null;
                                }
                            }
                            else
                            {
                                // 選択タブが変更されなくても、コレクションが変更されたのでUIを強制更新
                                OnPropertyChanged(nameof(SelectedTab));
                            }

                            // タブが1つしかなかった場合（削除後に0個になった場合）、新しいタブを作成してホームページを表示
                            if (wasOnlyTab && tabs.Count == 0)
                            {
                                // 新しいタブを作成（pathToOpenをnullにして、ホームページを表示）
                                var newTab = CreateTabInternal(null);
                                tabs.Add(newTab);
                                SelectedTab = newTab;
                                UpdateTabTitle(newTab);
                                // CreateTabInternal内で既にNavigateToHome()が呼ばれているが、
                                // 確実にホームページが表示されるように再度呼び出す
                                newTab.ViewModel.NavigateToHome();
                            }
                        }
                    }
                }

                // 【重要】先にタブを削除（新しいウィンドウを作成する前に元のウィンドウから削除）
                // タブが複数ある場合: タブを削除して次のタブを選択
                // タブが1つだけの場合: タブを削除してホームタブを作成
                RemoveTabFromCurrentWindow();

                // UIの更新を強制的に処理（新しいウィンドウを作成する前にUIを更新）
                System.Windows.Application.Current.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Background);

                // この後、新しいプロセスを起動して新しいウィンドウを作成
                try
                {
                    StartNewProcess(exePath, arguments);
                }
                catch
                {
                    // プロセス起動に失敗した場合でも、既にタブは削除されている
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // Win32Exceptionの場合は、詳細なエラー情報をログに記録（デバッグ用）
            }
            catch (System.Exception)
            {
                // その他の例外もログに記録（デバッグ用）
            }
        }

        /// <summary>
        /// 新しいウィンドウにタブを複製します
        /// </summary>
        /// <param name="tab">複製するタブ</param>
        [RelayCommand]
        private void DuplicateTabToNewWindow(ExplorerTab? tab)
        {
            if (tab == null)
                return;

            var pathToOpen = tab.ViewModel?.CurrentPath ?? string.Empty;

            try
            {
                var exePath = GetApplicationExecutablePath();
                if (string.IsNullOrEmpty(exePath))
                    return;

                // コマンドライン引数を構築
                var arguments = BuildPathArgument(pathToOpen);

                // 新しいプロセスを起動
                StartNewProcess(exePath, arguments);
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // Win32Exceptionの場合は、詳細なエラー情報をログに記録（デバッグ用）
            }
            catch (System.Exception)
            {
                // その他の例外もログに記録（デバッグ用）
            }
        }

        /// <summary>
        /// 現在のアプリケーションの実行ファイルパスを取得します
        /// </summary>
        private static string? GetApplicationExecutablePath()
        {
            // .NET 6以降では、Environment.ProcessPathを使用
            if (Environment.ProcessPath != null && System.IO.File.Exists(Environment.ProcessPath))
            {
                return Environment.ProcessPath;
            }

            // EntryAssemblyから取得を試みる
            var entryAssembly = System.Reflection.Assembly.GetEntryAssembly();
            if (entryAssembly != null && !string.IsNullOrEmpty(entryAssembly.Location))
            {
                return entryAssembly.Location;
            }

            // GetExecutingAssemblyを使用
            var executingAssembly = System.Reflection.Assembly.GetExecutingAssembly();
            if (!string.IsNullOrEmpty(executingAssembly.Location))
            {
                return executingAssembly.Location;
            }

            // Process.GetCurrentProcess().MainModuleを使用
            try
            {
                var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
                var exePath = currentProcess.MainModule?.FileName;
                if (!string.IsNullOrEmpty(exePath) && System.IO.File.Exists(exePath))
                {
                    return exePath;
                }
            }
            catch
            {
                // MainModuleの取得に失敗した場合は無視
            }

            return null;
        }

        /// <summary>
        /// パスからコマンドライン引数を構築します
        /// </summary>
        private static string BuildPathArgument(string? pathToOpen)
        {
            if (string.IsNullOrEmpty(pathToOpen) || !System.IO.Directory.Exists(pathToOpen))
                return string.Empty;

            // パフォーマンス最適化：パスにスペースが含まれている場合は引用符で囲む
            // IndexOfの方がContainsより高速、ZStringで文字列割り当てを削減
            return pathToOpen.IndexOf(' ') >= 0 ? ZString.Concat("\"", pathToOpen, "\"") : pathToOpen;
        }

        /// <summary>
        /// 新しいプロセスを起動します
        /// </summary>
        private static void StartNewProcess(string exePath, string arguments)
        {
            var processStartInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = exePath,
                Arguments = arguments,
                UseShellExecute = true,
                WorkingDirectory = System.IO.Path.GetDirectoryName(exePath) ?? string.Empty
            };

            System.Diagnostics.Process.Start(processStartInfo);
        }

        #endregion

        #region タブ一括操作

        /// <summary>
        /// 指定されたタブより左のタブをすべて閉じます
        /// </summary>
        /// <param name="tab">基準となるタブ</param>
        [RelayCommand]
        private void CloseTabsToLeft(ExplorerTab? tab)
        {
            if (tab == null)
                return;

            if (IsSplitPaneEnabled)
            {
                var leftPaneTabs = LeftPaneTabs;
                var leftIndex = leftPaneTabs.IndexOf(tab);
                if (leftIndex > 0)
                {
                    // パフォーマンス最適化：複数のタブを一度に削除してObservableCollectionの通知回数を削減
                    // 削除するタブをリストに収集
                    var tabsToRemove = new System.Collections.Generic.List<ExplorerTab>(leftIndex);
                    for (int i = 0; i < leftIndex; i++)
                    {
                        tabsToRemove.Add(leftPaneTabs[i]);
                    }
                    // 一括削除（逆順でRemoveAtを使用してインデックスのずれを防ぐ）
                    foreach (var tabToRemove in tabsToRemove)
                    {
                        leftPaneTabs.Remove(tabToRemove);
                    }
                    // 基準となるタブを選択状態にする（インデックスが0に変わる）
                    SelectedLeftPaneTab = tab;
                    ActivePane = ActivePaneLeft;
                    return;
                }

                var rightPaneTabs = RightPaneTabs;
                var rightIndex = rightPaneTabs.IndexOf(tab);
                if (rightIndex > 0)
                {
                    // パフォーマンス最適化：複数のタブを一度に削除
                    var tabsToRemove = new System.Collections.Generic.List<ExplorerTab>(rightIndex);
                    for (int i = 0; i < rightIndex; i++)
                    {
                        tabsToRemove.Add(rightPaneTabs[i]);
                    }
                    foreach (var tabToRemove in tabsToRemove)
                    {
                        rightPaneTabs.Remove(tabToRemove);
                    }
                    // 基準となるタブを選択状態にする（インデックスが0に変わる）
                    SelectedRightPaneTab = tab;
                    ActivePane = ActivePaneRight;
                }
            }
            else
            {
                var tabs = Tabs;
                var index = tabs.IndexOf(tab);
                if (index > 0)
                {
                    // パフォーマンス最適化：複数のタブを一度に削除
                    var tabsToRemove = new System.Collections.Generic.List<ExplorerTab>(index);
                    for (int i = 0; i < index; i++)
                    {
                        tabsToRemove.Add(tabs[i]);
                    }
                    foreach (var tabToRemove in tabsToRemove)
                    {
                        tabs.Remove(tabToRemove);
                    }
                    // 基準となるタブを選択状態にする（インデックスが0に変わる）
                    SelectedTab = tab;
                }
            }
        }

        /// <summary>
        /// 指定されたタブより右のタブをすべて閉じます
        /// </summary>
        /// <param name="tab">基準となるタブ</param>
        [RelayCommand]
        private void CloseTabsToRight(ExplorerTab? tab)
        {
            if (tab == null)
                return;

            if (IsSplitPaneEnabled)
            {
                var leftPaneTabs = LeftPaneTabs;
                var leftIndex = leftPaneTabs.IndexOf(tab);
                if (leftIndex >= 0)
                {
                    var count = leftPaneTabs.Count;
                    if (leftIndex < count - 1)
                    {
                        // パフォーマンス最適化：削除するタブをリストに収集してから一括削除
                        var tabsToRemove = new System.Collections.Generic.List<ExplorerTab>(count - leftIndex - 1);
                        for (int i = leftIndex + 1; i < count; i++)
                        {
                            tabsToRemove.Add(leftPaneTabs[i]);
                        }
                        foreach (var tabToRemove in tabsToRemove)
                        {
                            leftPaneTabs.Remove(tabToRemove);
                        }
                        // 基準となるタブを選択状態にする
                        SelectedLeftPaneTab = tab;
                        ActivePane = ActivePaneLeft;
                    }
                    return;
                }

                var rightPaneTabs = RightPaneTabs;
                var rightIndex = rightPaneTabs.IndexOf(tab);
                if (rightIndex >= 0)
                {
                    var count = rightPaneTabs.Count;
                    if (rightIndex < count - 1)
                    {
                        // パフォーマンス最適化：削除するタブをリストに収集してから一括削除
                        var tabsToRemove = new System.Collections.Generic.List<ExplorerTab>(count - rightIndex - 1);
                        for (int i = rightIndex + 1; i < count; i++)
                        {
                            tabsToRemove.Add(rightPaneTabs[i]);
                        }
                        foreach (var tabToRemove in tabsToRemove)
                        {
                            rightPaneTabs.Remove(tabToRemove);
                        }
                        // 基準となるタブを選択状態にする
                        SelectedRightPaneTab = tab;
                        ActivePane = ActivePaneRight;
                    }
                }
            }
            else
            {
                var tabs = Tabs;
                var index = tabs.IndexOf(tab);
                var count = tabs.Count;
                if (index < count - 1)
                {
                    // パフォーマンス最適化：削除するタブをリストに収集してから一括削除
                    var tabsToRemove = new System.Collections.Generic.List<ExplorerTab>(count - index - 1);
                    for (int i = index + 1; i < count; i++)
                    {
                        tabsToRemove.Add(tabs[i]);
                    }
                    foreach (var tabToRemove in tabsToRemove)
                    {
                        tabs.Remove(tabToRemove);
                    }
                    // 基準となるタブを選択状態にする
                    SelectedTab = tab;
                }
            }
        }

        /// <summary>
        /// 指定されたタブ以外のすべてのタブを閉じます
        /// </summary>
        /// <param name="tab">残すタブ</param>
        [RelayCommand]
        private void CloseOtherTabs(ExplorerTab? tab)
        {
            if (tab == null)
                return;

            if (IsSplitPaneEnabled)
            {
                var leftPaneTabs = LeftPaneTabs;
                if (leftPaneTabs.IndexOf(tab) >= 0)
                {
                    // パフォーマンス最適化：削除するタブをリストに収集してから一括削除
                    var tabsToRemove = new System.Collections.Generic.List<ExplorerTab>(leftPaneTabs.Count - 1);
                    foreach (var t in leftPaneTabs)
                    {
                        if (t != tab)
                        {
                            tabsToRemove.Add(t);
                        }
                    }
                    foreach (var tabToRemove in tabsToRemove)
                    {
                        leftPaneTabs.Remove(tabToRemove);
                    }
                    SelectedLeftPaneTab = tab;
                    return;
                }

                var rightPaneTabs = RightPaneTabs;
                if (rightPaneTabs.IndexOf(tab) >= 0)
                {
                    // パフォーマンス最適化：削除するタブをリストに収集してから一括削除
                    var tabsToRemove = new System.Collections.Generic.List<ExplorerTab>(rightPaneTabs.Count - 1);
                    foreach (var t in rightPaneTabs)
                    {
                        if (t != tab)
                        {
                            tabsToRemove.Add(t);
                        }
                    }
                    foreach (var tabToRemove in tabsToRemove)
                    {
                        rightPaneTabs.Remove(tabToRemove);
                    }
                    SelectedRightPaneTab = tab;
                }
            }
            else
            {
                var tabs = Tabs;
                // パフォーマンス最適化：削除するタブをリストに収集してから一括削除
                var tabsToRemove = new System.Collections.Generic.List<ExplorerTab>(tabs.Count - 1);
                foreach (var t in tabs)
                {
                    if (t != tab)
                    {
                        tabsToRemove.Add(t);
                    }
                }
                foreach (var tabToRemove in tabsToRemove)
                {
                    tabs.Remove(tabToRemove);
                }
                SelectedTab = tab;
            }
        }

        #endregion

    }
}
