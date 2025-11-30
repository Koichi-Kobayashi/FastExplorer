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
                
                foreach (var tab in LeftPaneTabs)
                {
                    Tabs.Add(tab);
                }
                LeftPaneTabs.Clear();
                
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
            // 初回実行時のみ初期化処理を実行
            if (!_hasNavigatedTo)
            {
                _hasNavigatedTo = true;
                
                // コマンドライン引数で指定されたパスを取得
                var startupPath = App.GetStartupPath();
                var isSingleTabMode = App.IsSingleTabMode();
                
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
            }
            
            // 初回実行時のみ実行する処理を続行
            if (_hasNavigatedTo)
            {
                // コマンドライン引数で指定されたパスを取得
                var startupPath = App.GetStartupPath();
                var isSingleTabMode = App.IsSingleTabMode();

                if (IsSplitPaneEnabled)
                {
                    // 選択タブが設定されていない場合は、最初のタブを選択
                    // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                    var leftPaneTabsCount2 = LeftPaneTabs.Count;
                    if (SelectedLeftPaneTab == null && leftPaneTabsCount2 > 0)
                    {
                        SelectedLeftPaneTab = LeftPaneTabs[0];
                    }
                    var rightPaneTabsCount2 = RightPaneTabs.Count;
                    if (SelectedRightPaneTab == null && rightPaneTabsCount2 > 0)
                    {
                        SelectedRightPaneTab = RightPaneTabs[0];
                    }
                    
                    // 起動時に分割ペインモードが有効な場合、ActivePaneを左ペインに設定
                    if (SelectedLeftPaneTab != null)
                    {
                        ActivePane = ActivePaneLeft; // 左ペインをデフォルト
                    }
                    else if (SelectedRightPaneTab != null)
                    {
                        ActivePane = ActivePaneRight; // 左ペインが存在しない場合のみ右ペインをデフォルト
                    }
                    
                    // コマンドライン引数で指定されたパスがある場合は、左ペインのタブに移動
                    if (!string.IsNullOrEmpty(startupPath) && System.IO.Directory.Exists(startupPath) && SelectedLeftPaneTab != null)
                    {
                        SelectedLeftPaneTab.ViewModel.NavigateToPathCommand.Execute(startupPath);
                    }
                }
                else
                {
                    // 単一タブモードの場合は、既存のタブを削除して新しいタブを作成
                    if (isSingleTabMode)
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
                                if (SelectedTab != null)
                                {
                                    SelectedTab.ViewModel.NavigateToHome();
                                }
                            }
                        }
                        else
                        {
                            // パスが指定されていない場合は、新しいタブを作成（ホームに移動）
                            CreateNewTab();
                        }
                    }
                    else
                    {
                        // 通常モードの場合
                        // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                        var tabsCount = Tabs.Count;
                        if (tabsCount == 0)
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
                        if (!string.IsNullOrEmpty(startupPath) && System.IO.Directory.Exists(startupPath) && SelectedTab != null)
                        {
                            SelectedTab.ViewModel.NavigateToPathCommand.Execute(startupPath);
                        }
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
        /// タブを作成する内部メソッド
        /// </summary>
        /// <param name="pathToOpen">開くパス。nullまたは空の場合はホームディレクトリを開く</param>
        private ExplorerTab CreateTabInternal(string? pathToOpen = null)
        {
            var viewModel = new ExplorerViewModel(_fileSystemService, _favoriteService);
            var tab = new ExplorerTab
            {
                Title = HomeTitle,
                CurrentPath = string.Empty,
                ViewModel = viewModel
            };

            // CurrentPathが変更されたときにTitleとCurrentPathを更新
            // イベントハンドラーを弱い参照で管理（メモリリーク防止）
            // UpdateStatusBarデリゲートをキャッシュ（メモリ割り当てを削減）
            if (_cachedUpdateStatusBarAction == null)
            {
                _cachedUpdateStatusBarAction = UpdateStatusBar;
            }
            
            // Dispatcherを事前にキャッシュ（イベントハンドラー内での取得を削減）
            var dispatcher = _cachedDispatcher ?? (_cachedDispatcher = System.Windows.Application.Current?.Dispatcher);
            
            PropertyChangedEventHandler? handler = null;
            handler = (s, e) =>
            {
                var propertyName = e.PropertyName;
                // 文字列比較を最適化（定数を使用、早期リターン）
                if (propertyName == CurrentPathPropertyName || propertyName == CurrentPathPropertyNameFull)
                {
                    tab.CurrentPath = viewModel.CurrentPath;
                    UpdateTabTitle(tab);
                    // UpdateStatusBarは遅延実行して頻繁な呼び出しを削減
                    // キャッシュされたデリゲートとDispatcherを使用（メモリ割り当てを削減）
                    dispatcher?.BeginInvoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        _cachedUpdateStatusBarAction);
                }
                else if (propertyName == ItemsPropertyName)
                {
                    // UpdateStatusBarは遅延実行して頻繁な呼び出しを削減
                    // キャッシュされたデリゲートとDispatcherを使用（メモリ割り当てを削減）
                    dispatcher?.BeginInvoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        _cachedUpdateStatusBarAction);
                }
            };
            viewModel.PropertyChanged += handler;
            
            // タブが削除されたときにイベントハンドラーを解除するための参照を保持
            // （ExplorerTabにDisposeメソッドを追加するか、タブ削除時に明示的に解除）

            // タブを追加する前に初期化を完了させる
            // これにより、タブが追加された直後にフォルダーをダブルクリックしても問題が発生しない
            if (!string.IsNullOrEmpty(pathToOpen) && System.IO.Directory.Exists(pathToOpen))
            {
                // パスが指定されていて存在する場合は、そのパスに移動
                tab.ViewModel.NavigateToPathCommand.Execute(pathToOpen);
            }
            else
            {
                // パスが指定されていない、または存在しない場合はホームに移動
                tab.ViewModel.NavigateToHome();
            }
            
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

            // イベントハンドラーを解除（メモリリーク防止）
            if (tab.ViewModel is ExplorerViewModel viewModel)
            {
                // PropertyChangedイベントハンドラーを解除
                // 注意: 現在の実装ではハンドラーの参照を保持していないため、
                // 将来的にはWeakEventManagerを使用するか、ExplorerTabにDisposeメソッドを追加することを推奨
            }

            if (IsSplitPaneEnabled)
            {
                // 左ペインのタブかチェック
                // Contains()とIndexOf()を1回の呼び出しに統合（パフォーマンス向上）
                var leftPaneTabs = LeftPaneTabs;
                var leftIndex = leftPaneTabs.IndexOf(tab);
                if (leftIndex >= 0)
                {
                    // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                    var leftPaneTabsCount = leftPaneTabs.Count;
                    if (leftPaneTabsCount <= 1)
                        return;
                    leftPaneTabs.RemoveAt(leftIndex);
                    if (SelectedLeftPaneTab == tab)
                    {
                        if (leftIndex > 0)
                            SelectedLeftPaneTab = leftPaneTabs[leftIndex - 1];
                        else if (leftPaneTabs.Count > 0)
                            SelectedLeftPaneTab = leftPaneTabs[0];
                    }
                }
                // 右ペインのタブかチェック
                else
                {
                    var rightPaneTabs = RightPaneTabs;
                    var rightIndex = rightPaneTabs.IndexOf(tab);
                    if (rightIndex >= 0)
                    {
                        // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                        var rightPaneTabsCount = rightPaneTabs.Count;
                        if (rightPaneTabsCount <= 1)
                            return;
                        rightPaneTabs.RemoveAt(rightIndex);
                        if (SelectedRightPaneTab == tab)
                        {
                            if (rightIndex > 0)
                                SelectedRightPaneTab = rightPaneTabs[rightIndex - 1];
                            else if (rightPaneTabs.Count > 0)
                                SelectedRightPaneTab = rightPaneTabs[0];
                        }
                    }
                }
            }
            else
            {
                var tabs = Tabs;
                // Countプロパティを一度だけ取得してキャッシュ（パフォーマンス向上）
                var tabsCount = tabs.Count;
                if (tabsCount <= 1)
                    return;

                var index = tabs.IndexOf(tab);
                tabs.Remove(tab);

                if (SelectedTab == tab)
                {
                    if (index > 0)
                        SelectedTab = tabs[index - 1];
                    else if (tabs.Count > 0)
                        SelectedTab = tabs[0];
                }
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

        /// <summary>
        /// タブのタイトルを現在のパスに基づいて更新します（必要な場合のみ）
        /// </summary>
        /// <param name="tab">タイトルを更新するタブ</param>
        private void UpdateTabTitleIfNeeded(ExplorerTab tab)
        {
            // 現在のタイトルとパスをキャッシュ（高速化）
            var currentTitle = tab.Title;
            var currentPath = tab.ViewModel?.CurrentPath;
            
            // 早期リターン: パスがnullまたは空で、タイトルが既にHomeTitleの場合は何もしない
            if (string.IsNullOrEmpty(currentPath))
            {
                if (currentTitle == HomeTitle)
                    return;
                tab.Title = HomeTitle;
                return;
            }
            
            // フォルダー名を取得（高速化：Spanを使用してメモリ割り当てを削減）
            var folderName = Path.GetFileName(currentPath);
            if (!string.IsNullOrEmpty(folderName))
            {
                // フォルダー名が取得できた場合
                if (currentTitle == folderName)
                    return;
                tab.Title = folderName;
                return;
            }
            
            // ルートディレクトリの場合
            var root = Path.GetPathRoot(currentPath);
            if (string.IsNullOrEmpty(root))
            {
                if (currentTitle == currentPath)
                    return;
                tab.Title = currentPath;
                return;
            }
            
            // ルートパスの末尾のバックスラッシュを削除
            var rootLength = root.Length;
            string expectedTitle;
            if (rootLength > 0 && root[rootLength - 1] == '\\')
            {
                expectedTitle = root.Substring(0, rootLength - 1);
            }
            else
            {
                expectedTitle = root;
            }
            
            if (currentTitle == expectedTitle)
                return;
            tab.Title = expectedTitle;
        }

        /// <summary>
        /// タブのタイトルを現在のパスに基づいて更新します
        /// </summary>
        /// <param name="tab">タイトルを更新するタブ</param>
        private void UpdateTabTitle(ExplorerTab tab)
        {
            // ViewModel.CurrentPathを一度だけ取得してキャッシュ（パフォーマンス向上）
            var currentPath = tab.ViewModel.CurrentPath;
            if (string.IsNullOrEmpty(currentPath))
            {
                tab.Title = HomeTitle;
            }
            else
            {
                // フォルダー名のみを表示（パス全体ではなく）
                var folderName = Path.GetFileName(currentPath);
                // ルートディレクトリ（例：C:\）の場合は、パス自体を表示
                if (string.IsNullOrEmpty(folderName))
                {
                    var root = Path.GetPathRoot(currentPath);
                    // 文字列操作を最適化（TrimEnd()の結果を直接使用）
                    if (string.IsNullOrEmpty(root))
                    {
                        tab.Title = currentPath;
                    }
                    else
                    {
                        // TrimEnd()を最適化（末尾のバックスラッシュのみをチェック）
                        var rootLength = root.Length;
                        if (rootLength > 0 && root[rootLength - 1] == '\\')
                        {
                            tab.Title = root.Substring(0, rootLength - 1);
                        }
                        else
                        {
                            tab.Title = root;
                        }
                    }
                }
                else
                {
                    tab.Title = folderName;
                }
            }
        }

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

                foreach (var path in tabPaths)
                {
                    var viewModel = new ExplorerViewModel(_fileSystemService, _favoriteService);
                    // フォルダー名のみを表示（パス全体ではなく）
                    string tabTitle;
                    if (string.IsNullOrEmpty(path))
                    {
                        tabTitle = HomeTitle;
                    }
                    else
                    {
                        var folderName = Path.GetFileName(path);
                        // ルートディレクトリ（例：C:\）の場合は、パス自体を表示
                        if (string.IsNullOrEmpty(folderName))
                        {
                            var root = Path.GetPathRoot(path);
                            tabTitle = string.IsNullOrEmpty(root) ? path : root.TrimEnd('\\');
                        }
                        else
                        {
                            tabTitle = folderName;
                        }
                    }
                    var tab = new ExplorerTab
                    {
                        Title = tabTitle,
                        CurrentPath = path ?? string.Empty,
                        ViewModel = viewModel
                    };

                    // CurrentPathが変更されたときにTitleとCurrentPathを更新
                    // イベントハンドラーを弱い参照で管理（メモリリーク防止）
                    // UpdateStatusBarデリゲートをキャッシュ（メモリ割り当てを削減）
                    if (_cachedUpdateStatusBarAction == null)
                    {
                        _cachedUpdateStatusBarAction = UpdateStatusBar;
                    }
                    
                    // Dispatcherを事前にキャッシュ（イベントハンドラー内での取得を削減）
                    var dispatcher3 = _cachedDispatcher ?? (_cachedDispatcher = System.Windows.Application.Current?.Dispatcher);
                    
                    PropertyChangedEventHandler? handler = null;
                    handler = (s, e) =>
                    {
                        var propertyName = e.PropertyName;
                        // 文字列比較を最適化（定数を使用、早期リターン）
                        if (propertyName == CurrentPathPropertyName || propertyName == CurrentPathPropertyNameFull)
                        {
                            tab.CurrentPath = viewModel.CurrentPath;
                            UpdateTabTitle(tab);
                            // UpdateStatusBarは遅延実行して頻繁な呼び出しを削減
                            // キャッシュされたデリゲートとDispatcherを使用（メモリ割り当てを削減）
                            dispatcher3?.BeginInvoke(
                                System.Windows.Threading.DispatcherPriority.Background,
                                _cachedUpdateStatusBarAction);
                        }
                        else if (propertyName == ItemsPropertyName)
                        {
                            // UpdateStatusBarは遅延実行して頻繁な呼び出しを削減
                            // キャッシュされたデリゲートとDispatcherを使用（メモリ割り当てを削減）
                            dispatcher3?.BeginInvoke(
                                System.Windows.Threading.DispatcherPriority.Background,
                                _cachedUpdateStatusBarAction);
                        }
                    };
                    viewModel.PropertyChanged += handler;

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
            ExplorerTab newTab;

            if (IsSplitPaneEnabled)
            {
                // 分割ペインモードの場合、同じペインに追加
                var leftPaneTabs = LeftPaneTabs;
                var rightPaneTabs = RightPaneTabs;
                
                if (leftPaneTabs.Contains(tab))
                {
                    newTab = CreateTabInternal(pathToOpen);
                    leftPaneTabs.Add(newTab);
                    SelectedLeftPaneTab = newTab;
                }
                else if (rightPaneTabs.Contains(tab))
                {
                    newTab = CreateTabInternal(pathToOpen);
                    rightPaneTabs.Add(newTab);
                    SelectedRightPaneTab = newTab;
                }
                else
                {
                    return;
                }
            }
            else
            {
                newTab = CreateTabInternal(pathToOpen);
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
        /// 新しいウィンドウにタブを移動します（マウス位置を指定）
        /// </summary>
        /// <param name="tab">移動するタブ</param>
        /// <param name="dropPosition">ドロップ位置（nullの場合はデフォルト位置）</param>
        public void MoveTabToNewWindow(ExplorerTab? tab, System.Windows.Point? dropPosition)
        {
            if (tab == null)
                return;

            // 【重要】削除・移動するのはアクティブタブ
            // ドラッグされたタブではなく、現在選択されているタブを使用
            System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] ========================================");
            System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] アクティブタブの取得を開始");
            System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] IsSplitPaneEnabled={IsSplitPaneEnabled}, ActivePane={ActivePane}");
            System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] ドラッグされたタブ: {tab?.Title}");
            
            ExplorerTab? activeTab = null;
            if (IsSplitPaneEnabled)
            {
                // 分割ペインモードの場合、ドラッグされたタブがどちらのコレクションにあるか確認
                var leftPaneTabs = LeftPaneTabs;
                var rightPaneTabs = RightPaneTabs;
                
                if (leftPaneTabs.Contains(tab))
                {
                    // ドラッグされたタブが左ペインにある場合、左ペインの選択タブを取得
                    activeTab = SelectedLeftPaneTab;
                    System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 分割ペインモード（ドラッグされたタブは左ペイン）: 選択タブ={activeTab?.Title}");
                    
                    // 選択タブがない、またはコレクションにない場合は、ドラッグされたタブを使用
                    if (activeTab == null || !leftPaneTabs.Contains(activeTab))
                    {
                        activeTab = tab;
                        System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 左ペインの選択タブがないため、ドラッグされたタブを使用={activeTab?.Title}");
                    }
                }
                else if (rightPaneTabs.Contains(tab))
                {
                    // ドラッグされたタブが右ペインにある場合、右ペインの選択タブを取得
                    activeTab = SelectedRightPaneTab;
                    System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 分割ペインモード（ドラッグされたタブは右ペイン）: 選択タブ={activeTab?.Title}");
                    
                    // 選択タブがない、またはコレクションにない場合は、ドラッグされたタブを使用
                    if (activeTab == null || !rightPaneTabs.Contains(activeTab))
                    {
                        activeTab = tab;
                        System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 右ペインの選択タブがないため、ドラッグされたタブを使用={activeTab?.Title}");
                    }
                }
                else
                {
                    // ドラッグされたタブがどちらのペインにもない場合は、ドラッグされたタブを使用
                    activeTab = tab;
                    System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 分割ペインモード（ドラッグされたタブはどちらのペインにもない）: ドラッグされたタブを使用={activeTab?.Title}");
                }
            }
            else
            {
                // 通常モードの場合、SelectedTabを使用
                activeTab = SelectedTab;
                System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 通常モード: 選択タブ={activeTab?.Title}");
                
                // 選択タブがない、またはコレクションにない場合は、ドラッグされたタブを使用
                if (activeTab == null || !Tabs.Contains(activeTab))
                {
                    activeTab = tab;
                    System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 通常モードの選択タブがないため、ドラッグされたタブを使用={activeTab?.Title}");
                }
            }
            
            // アクティブタブが取得できない場合は、ドラッグされたタブを使用
            if (activeTab == null)
            {
                activeTab = tab;
                System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] アクティブタブがnullのため、ドラッグされたタブを使用={activeTab?.Title}");
            }
            
            // この後、activeTabを使用して処理を進める
            tab = activeTab;
            System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 【確定】使用するタブ: {tab.Title}, Path={tab.ViewModel?.CurrentPath}");
            System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] ========================================");

            // タブのパスを取得（削除する前に取得する必要がある）
            var pathToOpen = tab.ViewModel?.CurrentPath;
            
            // CurrentPathが空またはnullの場合、ホームディレクトリを使用
            if (string.IsNullOrEmpty(pathToOpen))
            {
                pathToOpen = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            }

            try
            {
                // 現在のアプリケーションのEXEパスを取得
                string? exePath = null;
                
                // .NET 6以降では、Environment.ProcessPathを使用
                if (Environment.ProcessPath != null && System.IO.File.Exists(Environment.ProcessPath))
                {
                    exePath = Environment.ProcessPath;
                }
                
                // 取得できない場合は、EntryAssemblyから取得を試みる
                if (string.IsNullOrEmpty(exePath))
                {
                    var entryAssembly = System.Reflection.Assembly.GetEntryAssembly();
                    if (entryAssembly != null)
                    {
                        exePath = entryAssembly.Location;
                    }
                }
                
                // 取得できない場合は、GetExecutingAssemblyを使用
                if (string.IsNullOrEmpty(exePath))
                {
                    var executingAssembly = System.Reflection.Assembly.GetExecutingAssembly();
                    exePath = executingAssembly.Location;
                }
                
                // それでも取得できない場合は、Process.GetCurrentProcess().MainModuleを使用
                if (string.IsNullOrEmpty(exePath))
                {
                    try
                    {
                        var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
                        exePath = currentProcess.MainModule?.FileName;
                    }
                    catch
                    {
                        // MainModuleの取得に失敗した場合は無視
                    }
                }

                // EXEパスが取得できなかった場合、またはファイルが存在しない場合は終了
                if (string.IsNullOrEmpty(exePath) || !System.IO.File.Exists(exePath))
                    return;

                // コマンドライン引数を構築（--single-tabフラグを追加）
                var arguments = "--single-tab";
                
                // ドロップ位置がある場合は、コマンドライン引数に追加
                if (dropPosition.HasValue)
                {
                    arguments += $" --position {dropPosition.Value.X},{dropPosition.Value.Y}";
                }
                
                // パスをコマンドライン引数として渡す（pathToOpenは既に設定されている）
                if (!string.IsNullOrEmpty(pathToOpen))
                {
                    // パスをコマンドライン引数として渡す（スペースがある場合は引用符で囲む）
                    if (pathToOpen.Contains(" "))
                    {
                        arguments += $" \"{pathToOpen}\"";
                    }
                    else
                    {
                        arguments += $" {pathToOpen}";
                    }
                }

                // 元のウィンドウからタブを削除（新しいプロセス起動前に実行して、即座にUIを更新）
                // タブの削除処理を関数に抽出して、新しいプロセス起動前に実行
                void RemoveTabFromCurrentWindow()
                {
                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 【ステップ1開始】元のウィンドウからタブを削除します: IsSplitPaneEnabled={IsSplitPaneEnabled}");
                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 削除するタブ: Title={tab?.Title}, Path={tab?.ViewModel?.CurrentPath}");
                    
                    if (IsSplitPaneEnabled)
                    {
                        var leftPaneTabs = LeftPaneTabs;
                        var rightPaneTabs = RightPaneTabs;
                        
                        System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 分割ペインモード: leftPaneTabs.Count={leftPaneTabs.Count}, rightPaneTabs.Count={rightPaneTabs.Count}");
                        System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] leftPaneTabs.Contains(tab)={leftPaneTabs.Contains(tab)}, rightPaneTabs.Contains(tab)={rightPaneTabs.Contains(tab)}");
                        
                        if (leftPaneTabs.Contains(tab))
                        {
                            var index = leftPaneTabs.IndexOf(tab);
                            // タブを削除する前に、タブが1つだけかどうかを確認
                            bool wasOnlyTab = leftPaneTabs.Count == 1;
                            
                            System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインからタブを削除: index={index}, wasOnlyTab={wasOnlyTab}, leftPaneTabs.Count={leftPaneTabs.Count}");
                            
                            leftPaneTabs.RemoveAt(index);
                            
                            // 選択タブを更新（削除されたタブが選択されていなくても更新）
                            bool wasSelected = (SelectedLeftPaneTab == tab);
                            System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 削除されたタブは選択されていた: {wasSelected}");
                            
                            // 削除されたタブが選択されていた場合、または選択タブがない場合
                            if (wasSelected || SelectedLeftPaneTab == null || !leftPaneTabs.Contains(SelectedLeftPaneTab))
                            {
                                if (index > 0 && leftPaneTabs.Count > 0)
                                {
                                    SelectedLeftPaneTab = leftPaneTabs[index - 1];
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインの選択タブを前のタブに変更: index={index - 1}, Title={SelectedLeftPaneTab?.Title}");
                                }
                                else if (leftPaneTabs.Count > 0)
                                {
                                    SelectedLeftPaneTab = leftPaneTabs[0];
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインの選択タブを最初のタブに変更: index=0, Title={SelectedLeftPaneTab?.Title}");
                                }
                                else
                                {
                                    SelectedLeftPaneTab = null;
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインの選択タブをnullに設定");
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 選択タブは変更しない: Title={SelectedLeftPaneTab?.Title}");
                                // 選択タブが変更されなくても、コレクションが変更されたのでUIを強制更新
                                OnPropertyChanged(nameof(SelectedLeftPaneTab));
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインのUIを強制更新");
                            }
                            
                            // タブが1つしかなかった場合（削除後に0個になった場合）、新しいタブを作成してホームページを表示
                            if (wasOnlyTab && leftPaneTabs.Count == 0)
                            {
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインに新しいホームタブを作成");
                                
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
                                
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインに新しいホームタブを作成完了、ActivePane={ActivePane}");
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
                                    
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインにフォーカスを設定してUIを更新、ActivePane={ActivePane}（強制更新）");
                                }
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインに新しいタブを作成しない: wasOnlyTab={wasOnlyTab}, leftPaneTabs.Count={leftPaneTabs.Count}");
                            }
                        }
                        else if (rightPaneTabs.Contains(tab))
                        {
                            var index = rightPaneTabs.IndexOf(tab);
                            // タブを削除する前に、タブが1つだけかどうかを確認
                            bool wasOnlyTab = rightPaneTabs.Count == 1;
                            
                            System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインからタブを削除: index={index}, wasOnlyTab={wasOnlyTab}, rightPaneTabs.Count={rightPaneTabs.Count}");
                            
                            rightPaneTabs.RemoveAt(index);
                            
                            // 選択タブを更新（削除されたタブが選択されていなくても更新）
                            bool wasSelected = (SelectedRightPaneTab == tab);
                            System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 削除されたタブは選択されていた: {wasSelected}");
                            
                            // 削除されたタブが選択されていた場合、または選択タブがない場合
                            if (wasSelected || SelectedRightPaneTab == null || !rightPaneTabs.Contains(SelectedRightPaneTab))
                            {
                                if (index > 0 && rightPaneTabs.Count > 0)
                                {
                                    SelectedRightPaneTab = rightPaneTabs[index - 1];
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインの選択タブを前のタブに変更: index={index - 1}, Title={SelectedRightPaneTab?.Title}");
                                }
                                else if (rightPaneTabs.Count > 0)
                                {
                                    SelectedRightPaneTab = rightPaneTabs[0];
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインの選択タブを最初のタブに変更: index=0, Title={SelectedRightPaneTab?.Title}");
                                }
                                else
                                {
                                    SelectedRightPaneTab = null;
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインの選択タブをnullに設定");
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 選択タブは変更しない: Title={SelectedRightPaneTab?.Title}");
                                // 選択タブが変更されなくても、コレクションが変更されたのでUIを強制更新
                                OnPropertyChanged(nameof(SelectedRightPaneTab));
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインのUIを強制更新");
                            }
                            
                            // タブが1つしかなかった場合（削除後に0個になった場合）、新しいタブを作成してホームページを表示
                            if (wasOnlyTab && rightPaneTabs.Count == 0)
                            {
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインに新しいホームタブを作成");
                                
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
                                
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインに新しいホームタブを作成完了、ActivePane={ActivePane}");
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
                                    
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインにフォーカスを設定してUIを更新、ActivePane={ActivePane}（強制更新）");
                                }
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインに新しいタブを作成しない: wasOnlyTab={wasOnlyTab}, rightPaneTabs.Count={rightPaneTabs.Count}");
                            }
                        }
                        else
                        {
                            // タブが左ペインにも右ペインにも見つからない場合、通常モードのTabsコレクションを確認
                            System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 分割ペインモードでタブが見つかりません: leftPaneTabs.Count={leftPaneTabs.Count}, rightPaneTabs.Count={rightPaneTabs.Count}");
                            
                            // 通常モードのTabsコレクションにタブがあるか確認
                            var tabs = Tabs;
                            var tabIndex = tabs.IndexOf(tab);
                            if (tabIndex >= 0)
                            {
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードのTabsコレクションにタブが見つかりました: index={tabIndex}, tabs.Count={tabs.Count}");
                                
                                // 通常モードのTabsコレクションからタブを削除
                                tabs.RemoveAt(tabIndex);
                                
                                // 選択タブを更新
                                if (SelectedTab == tab)
                                {
                                    if (tabIndex > 0 && tabs.Count > 0)
                                    {
                                        SelectedTab = tabs[tabIndex - 1];
                                        System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードの選択タブを前のタブに変更: index={tabIndex - 1}");
                                    }
                                    else if (tabs.Count > 0)
                                    {
                                        SelectedTab = tabs[0];
                                        System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードの選択タブを最初のタブに変更: index=0");
                                    }
                                    else
                                    {
                                        SelectedTab = null;
                                        System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードの選択タブをnullに設定");
                                    }
                                }
                                
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードのTabsコレクションからタブを削除しました: 残りのタブ数={tabs.Count}");
                            }
                            else
                            {
                                // 通常モードのTabsコレクションにも見つからない場合
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードのTabsコレクションにもタブが見つかりません: tabs.Count={tabs.Count}");
                                
                                // 左ペインと右ペインのタブ数を確認して、どちらかが空の場合は新しいタブを作成
                                bool leftPaneWasEmpty = false;
                                bool rightPaneWasEmpty = false;
                                
                                if (leftPaneTabs.Count == 0)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインが空のため、新しいホームタブを作成");
                                    var newTab = CreateTabInternal(null);
                                    leftPaneTabs.Add(newTab);
                                    SelectedLeftPaneTab = newTab;
                                    UpdateTabTitle(newTab);
                                    newTab.ViewModel.NavigateToHome();
                                    leftPaneWasEmpty = true;
                                }
                                if (rightPaneTabs.Count == 0)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインが空のため、新しいホームタブを作成");
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
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 両方のペインが空だったため、左ペインにフォーカスを設定、ActivePane={ActivePane}");
                                }
                                else if (leftPaneWasEmpty)
                                {
                                    ActivePane = ActivePaneLeft;
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 左ペインが空だったため、左ペインにフォーカスを設定、ActivePane={ActivePane}");
                                }
                                else if (rightPaneWasEmpty)
                                {
                                    ActivePane = ActivePaneRight;
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 右ペインが空だったため、右ペインにフォーカスを設定、ActivePane={ActivePane}");
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
                            
                            System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードからタブを削除: index={index}, wasOnlyTab={wasOnlyTab}, tabs.Count={tabs.Count}");
                            
                            tabs.RemoveAt(index);
                            
                            // 選択タブを更新（削除されたタブが選択されていなくても更新）
                            bool wasSelected = (SelectedTab == tab);
                            System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 削除されたタブは選択されていた: {wasSelected}");
                            
                            // 削除されたタブが選択されていた場合、または選択タブがない場合
                            if (wasSelected || SelectedTab == null || !tabs.Contains(SelectedTab))
                            {
                                if (index > 0 && tabs.Count > 0)
                                {
                                    SelectedTab = tabs[index - 1];
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードの選択タブを前のタブに変更: index={index - 1}, Title={SelectedTab?.Title}");
                                }
                                else if (tabs.Count > 0)
                                {
                                    SelectedTab = tabs[0];
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードの選択タブを最初のタブに変更: index=0, Title={SelectedTab?.Title}");
                                }
                                else
                                {
                                    SelectedTab = null;
                                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードの選択タブをnullに設定");
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 選択タブは変更しない: Title={SelectedTab?.Title}");
                                // 選択タブが変更されなくても、コレクションが変更されたのでUIを強制更新
                                OnPropertyChanged(nameof(SelectedTab));
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードのUIを強制更新");
                            }
                            
                            // タブが1つしかなかった場合（削除後に0個になった場合）、新しいタブを作成してホームページを表示
                            if (wasOnlyTab && tabs.Count == 0)
                            {
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードに新しいホームタブを作成");
                                
                                // 新しいタブを作成（pathToOpenをnullにして、ホームページを表示）
                                var newTab = CreateTabInternal(null);
                                tabs.Add(newTab);
                                SelectedTab = newTab;
                                UpdateTabTitle(newTab);
                                // CreateTabInternal内で既にNavigateToHome()が呼ばれているが、
                                // 確実にホームページが表示されるように再度呼び出す
                                newTab.ViewModel.NavigateToHome();
                                
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードに新しいホームタブを作成完了");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードに新しいタブを作成しない: wasOnlyTab={wasOnlyTab}, tabs.Count={tabs.Count}");
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 通常モードでタブが見つかりません");
                        }
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"[RemoveTabFromCurrentWindow] 【ステップ1完了】元のウィンドウからタブを削除しました");
                }
                
                // 【重要】先にタブを削除（新しいウィンドウを作成する前に元のウィンドウから削除）
                // タブが複数ある場合: タブを削除して次のタブを選択
                // タブが1つだけの場合: タブを削除してホームタブを作成
                System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] ========================================");
                System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 処理開始: タブをデスクトップにドロップ");
                System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] ========================================");
                RemoveTabFromCurrentWindow();
                System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] ----------------------------------------");
                
                // UIの更新を強制的に処理（新しいウィンドウを作成する前にUIを更新）
                System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] UIの更新を待機します...");
                System.Windows.Application.Current.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Background);
                System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] UIの更新が完了しました");
                
                // この後、新しいプロセスを起動して新しいウィンドウを作成
                var processStartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = arguments,
                    UseShellExecute = true,
                    WorkingDirectory = System.IO.Path.GetDirectoryName(exePath) ?? string.Empty
                };

                System.Diagnostics.Process? newProcess = null;
                try
                {
                    System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 【ステップ2】新しいウィンドウを作成します: exePath={exePath}, arguments={arguments}");
                    newProcess = System.Diagnostics.Process.Start(processStartInfo);
                    System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 【ステップ2完了】新しいウィンドウを作成しました: newProcess={newProcess?.Id}");
                }
                catch (System.Exception ex)
                {
                    // プロセス起動に失敗した場合でも、既にタブは削除されている
                    System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 【エラー】新しいウィンドウの作成に失敗: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] ※タブは既に元のウィンドウから削除されています");
                    newProcess = null;
                }
                
                if (newProcess == null)
                {
                    // プロセスの起動に失敗した場合でも、タブは既に削除されている
                    System.Diagnostics.Debug.WriteLine("[MoveTabToNewWindow] 【エラー】新しいウィンドウの作成に失敗しました。タブは既に元のウィンドウから削除されています。");
                    return;
                }
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                // Win32Exceptionの場合は、詳細なエラー情報をログに記録（デバッグ用）
                System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 外側のcatch: Win32Exception: {ex.Message}");
            }
            catch (System.Exception ex)
            {
                // その他の例外もログに記録（デバッグ用）
                System.Diagnostics.Debug.WriteLine($"[MoveTabToNewWindow] 外側のcatch: Exception: {ex.Message}");
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
                // 現在のアプリケーションのEXEパスを取得
                string? exePath = null;
                
                // .NET 6以降では、Environment.ProcessPathを使用
                if (Environment.ProcessPath != null && System.IO.File.Exists(Environment.ProcessPath))
                {
                    exePath = Environment.ProcessPath;
                }
                
                // 取得できない場合は、EntryAssemblyから取得を試みる
                if (string.IsNullOrEmpty(exePath))
                {
                    var entryAssembly = System.Reflection.Assembly.GetEntryAssembly();
                    if (entryAssembly != null)
                    {
                        exePath = entryAssembly.Location;
                    }
                }
                
                // 取得できない場合は、GetExecutingAssemblyを使用
                if (string.IsNullOrEmpty(exePath))
                {
                    var executingAssembly = System.Reflection.Assembly.GetExecutingAssembly();
                    exePath = executingAssembly.Location;
                }
                
                // それでも取得できない場合は、Process.GetCurrentProcess().MainModuleを使用
                if (string.IsNullOrEmpty(exePath))
                {
                    try
                    {
                        var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
                        exePath = currentProcess.MainModule?.FileName;
                    }
                    catch
                    {
                        // MainModuleの取得に失敗した場合は無視
                    }
                }

                // EXEパスが取得できなかった場合、またはファイルが存在しない場合は終了
                if (string.IsNullOrEmpty(exePath) || !System.IO.File.Exists(exePath))
                    return;

                // コマンドライン引数を構築
                var arguments = string.Empty;
                if (!string.IsNullOrEmpty(pathToOpen) && System.IO.Directory.Exists(pathToOpen))
                {
                    // パスをコマンドライン引数として渡す（既に引用符で囲まれている場合はそのまま使用）
                    // パスにスペースが含まれている場合は引用符で囲む
                    if (pathToOpen.Contains(" "))
                    {
                        arguments = $"\"{pathToOpen}\"";
                    }
                    else
                    {
                        arguments = pathToOpen;
                    }
                }

                // 新しいプロセスを起動
                var processStartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = arguments,
                    UseShellExecute = true,
                    WorkingDirectory = System.IO.Path.GetDirectoryName(exePath) ?? string.Empty
                };

                var newProcess = System.Diagnostics.Process.Start(processStartInfo);
                if (newProcess == null)
                {
                    // プロセスの起動に失敗した場合
                    return;
                }
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                // Win32Exceptionの場合は、詳細なエラー情報をログに記録（デバッグ用）
                System.Diagnostics.Debug.WriteLine($"新しいプロセスの起動に失敗しました: {ex.Message}");
            }
            catch (System.Exception ex)
            {
                // その他の例外もログに記録（デバッグ用）
                System.Diagnostics.Debug.WriteLine($"新しいプロセスの起動中にエラーが発生しました: {ex.Message}");
            }
        }

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
                var rightPaneTabs = RightPaneTabs;
                
                if (leftPaneTabs.Contains(tab))
                {
                    var index = leftPaneTabs.IndexOf(tab);
                    if (index > 0)
                    {
                        // 左側のタブをすべて削除
                        for (int i = index - 1; i >= 0; i--)
                        {
                            leftPaneTabs.RemoveAt(i);
                        }
                        // 基準となるタブを選択状態にする（インデックスが0に変わる）
                        SelectedLeftPaneTab = tab;
                        ActivePane = 0;
                    }
                }
                else if (rightPaneTabs.Contains(tab))
                {
                    var index = rightPaneTabs.IndexOf(tab);
                    if (index > 0)
                    {
                        // 左側のタブをすべて削除
                        for (int i = index - 1; i >= 0; i--)
                        {
                            rightPaneTabs.RemoveAt(i);
                        }
                        // 基準となるタブを選択状態にする（インデックスが0に変わる）
                        SelectedRightPaneTab = tab;
                        ActivePane = 2;
                    }
                }
            }
            else
            {
                var tabs = Tabs;
                var index = tabs.IndexOf(tab);
                if (index > 0)
                {
                    // 左側のタブをすべて削除
                    for (int i = index - 1; i >= 0; i--)
                    {
                        tabs.RemoveAt(i);
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
                var rightPaneTabs = RightPaneTabs;
                
                if (leftPaneTabs.Contains(tab))
                {
                    var index = leftPaneTabs.IndexOf(tab);
                    var count = leftPaneTabs.Count;
                    if (index < count - 1)
                    {
                        // 右側のタブをすべて削除
                        for (int i = count - 1; i > index; i--)
                        {
                            leftPaneTabs.RemoveAt(i);
                        }
                        // 基準となるタブを選択状態にする
                        SelectedLeftPaneTab = tab;
                        ActivePane = 0;
                    }
                }
                else if (rightPaneTabs.Contains(tab))
                {
                    var index = rightPaneTabs.IndexOf(tab);
                    var count = rightPaneTabs.Count;
                    if (index < count - 1)
                    {
                        // 右側のタブをすべて削除
                        for (int i = count - 1; i > index; i--)
                        {
                            rightPaneTabs.RemoveAt(i);
                        }
                        // 基準となるタブを選択状態にする
                        SelectedRightPaneTab = tab;
                        ActivePane = 2;
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
                    // 右側のタブをすべて削除
                    for (int i = count - 1; i > index; i--)
                    {
                        tabs.RemoveAt(i);
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
                var rightPaneTabs = RightPaneTabs;
                
                if (leftPaneTabs.Contains(tab))
                {
                    // 左ペインの他のタブをすべて削除
                    for (int i = leftPaneTabs.Count - 1; i >= 0; i--)
                    {
                        if (leftPaneTabs[i] != tab)
                        {
                            leftPaneTabs.RemoveAt(i);
                        }
                    }
                    SelectedLeftPaneTab = tab;
                }
                else if (rightPaneTabs.Contains(tab))
                {
                    // 右ペインの他のタブをすべて削除
                    for (int i = rightPaneTabs.Count - 1; i >= 0; i--)
                    {
                        if (rightPaneTabs[i] != tab)
                        {
                            rightPaneTabs.RemoveAt(i);
                        }
                    }
                    SelectedRightPaneTab = tab;
                }
            }
            else
            {
                var tabs = Tabs;
                // 他のタブをすべて削除
                for (int i = tabs.Count - 1; i >= 0; i--)
                {
                    if (tabs[i] != tab)
                    {
                        tabs.RemoveAt(i);
                    }
                }
                SelectedTab = tab;
            }
        }
    }
}

