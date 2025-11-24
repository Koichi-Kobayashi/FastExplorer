using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Linq;
using System.Windows.Interop;
using FastExplorer.Services;
using FastExplorer.ViewModels.Windows;
using FastExplorer.Helpers;
using FastExplorer.ShellContextMenu;
using Wpf.Ui;
using Wpf.Ui.Abstractions;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

namespace FastExplorer.Views.Windows
{
    /// <summary>
    /// メインウィンドウを表すクラス
    /// </summary>
    public partial class MainWindow : INavigationWindow
    {
        private readonly WindowSettingsService _windowSettingsService;
        private readonly INavigationService _navigationService;
        private HwndSource? _hwndSource;
        private HwndSourceHook? _wndProcHook; // メッセージフックの参照を保持
        
        // リフレクション結果をキャッシュ（パフォーマンス向上）
        private static PropertyInfo? _cachedInvokedItemContainerProperty;
        private static Type? _cachedArgsType;
        
        // ViewModelをキャッシュ（パフォーマンス向上）
        private ViewModels.Pages.ExplorerPageViewModel? _cachedExplorerPageViewModel;
        
        // 文字列定数（パフォーマンス向上）
        private const string HomeTag = "HOME";
        
        // 画面サイズをキャッシュ（パフォーマンス向上）
        private static double? _cachedScreenWidth;
        private static double? _cachedScreenHeight;
        
        // 型をキャッシュ（パフォーマンス向上）
        private static readonly Type ExplorerPageType = typeof(Views.Pages.ExplorerPage);
        private static readonly Type ExplorerPageViewModelType = typeof(ViewModels.Pages.ExplorerPageViewModel);

        /// <summary>
        /// メインウィンドウのViewModelを取得します
        /// </summary>
        public MainWindowViewModel ViewModel { get; }

        /// <summary>
        /// <see cref="MainWindow"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">メインウィンドウのViewModel</param>
        /// <param name="navigationViewPageProvider">ナビゲーションビューページプロバイダー</param>
        /// <param name="navigationService">ナビゲーションサービス</param>
        /// <param name="windowSettingsService">ウィンドウ設定サービス</param>
        public MainWindow(
            MainWindowViewModel viewModel,
            INavigationViewPageProvider navigationViewPageProvider,
            INavigationService navigationService,
            WindowSettingsService windowSettingsService
        )
        {
            ViewModel = viewModel;
            DataContext = this;
            _windowSettingsService = windowSettingsService;
            _navigationService = navigationService;

            ShowInTaskbar = true; // タスクバーに表示

            InitializeComponent();
            
            // すべての初期化処理を1つのLoadedイベントハンドラーに統合（起動時の高速化）
            void InitializeHandler(object? s, RoutedEventArgs e)
            {
                Loaded -= InitializeHandler; // 一度だけ実行されるように解除
                
                // RootNavigationの初期化（1回だけ取得して再利用）
                var nav = RootNavigation;
                if (nav != null)
                {
                    SetPageService(navigationViewPageProvider);
                    navigationService.SetNavigationControl(nav);
                }
                
                // ViewModelにNavigationServiceを設定
                viewModel.SetNavigationService(navigationService);
                
                // テーマカラーを適用（ウィンドウ表示前に適用することでチラつきを防ぐ）
                var settings = _windowSettingsService.GetSettings();
                var themeColorCode = settings.ThemeColorCode;
                // 高速化：nullチェックとLengthチェックを1回に統合
                if (themeColorCode != null && themeColorCode.Length != 0)
                {
                    // 色計算を1回だけ実行（高速なカスタム変換を使用）
                    var mainColor = Helpers.FastColorConverter.ParseHexColor(themeColorCode);
                    var secondaryColor = Helpers.FastColorConverter.ParseHexColor(settings.ThemeSecondaryColorCode ?? "#FCFCFC");
                    var mainBrush = new System.Windows.Media.SolidColorBrush(mainColor);
                    var secondaryBrush = new System.Windows.Media.SolidColorBrush(secondaryColor);
                    // 輝度計算を最適化（定数を事前計算）
                    var luminance = (0.299 * mainColor.R + 0.587 * mainColor.G + 0.114 * mainColor.B) * 0.00392156862745098; // 1/255を事前計算
                    var statusBarTextColor = luminance > 0.5 ? System.Windows.Media.Colors.Black : System.Windows.Media.Colors.White;
                    var statusBarTextBrush = new System.Windows.Media.SolidColorBrush(statusBarTextColor);
                    
                    // リソースを更新（計算済みの色を渡して重複計算を回避）
                    App.ApplyThemeColorFromSettings(settings, (mainColor, secondaryColor));
                    
                    // すべての色を同じタイミングで設定（リソース更新直後、計算済みの色を使用）
                    // プロパティ設定を最適化（nullチェックを削除して高速化）
                    Background = mainBrush;
                    // FluentWindowのBackgroundを設定（型チェックを削除して高速化）
                    var fluentWindow = this as Wpf.Ui.Controls.FluentWindow;
                    if (fluentWindow != null)
                    {
                        fluentWindow.Background = mainBrush;
                    }
                    // navは既に取得済みなので、nullチェックを削除（navがnullの場合は既にreturnしている）
                    if (nav != null)
                    {
                        nav.Background = secondaryBrush;
                    }

                    // タブとListViewの選択中の色を更新するため、スタイルを無効化してDynamicResourceの再評価を強制
                    // ContentRenderedイベントで実行（確実にExplorerPageが読み込まれた後）
                    // 起動を高速化するため、さらに遅延実行（デリゲートのメモリアロケーションを削減）
                    void ContentRenderedHandler(object? s, EventArgs e)
                    {
                        ContentRendered -= ContentRenderedHandler;
                        // 遅延実行して起動を高速化（静的メソッド参照を使用してメモリアロケーションを削減）
                        var window = this;
                        _ = Dispatcher.BeginInvoke(DispatcherPriority.Background, new System.Action(() =>
                        {
                            InvalidateTabAndListViewStyles(window);
                        }));
                    }
                    ContentRendered += ContentRenderedHandler;
                }
                
                // テーマカラー適用後にウィンドウを表示（チラつきを防ぐ）
                // UpdateLayout()を削除して起動を高速化（レイアウトは自動的に更新される）
                if (Visibility == Visibility.Hidden)
                {
                    Visibility = Visibility.Visible;
                    // ShowInTaskbarは既にコンストラクタで設定済み
                }
                
                // SystemThemeWatcherを遅延実行（起動を最速化）
                // ウィンドウ位置とサイズの復元はShowWindow()で実行されるため、ここでは実行しない
                // デリゲートのメモリアロケーションを削減
                var window = this;
                _ = Dispatcher.BeginInvoke(DispatcherPriority.Background, new System.Action(() =>
                {
                    SystemThemeWatcher.Watch(window);
                }));
            }
            Loaded += InitializeHandler;
        }


        #region INavigationWindow methods

        /// <summary>
        /// ナビゲーションコントロールを取得します
        /// </summary>
        /// <returns>ナビゲーションビュー</returns>
        public INavigationView GetNavigation() => RootNavigation;

        /// <summary>
        /// 指定されたページタイプにナビゲートします
        /// </summary>
        /// <param name="pageType">ナビゲートするページのタイプ</param>
        /// <returns>ナビゲートに成功した場合はtrue、それ以外の場合はfalse</returns>
        public bool Navigate(Type pageType)
        {
            // ナビゲーションハンドラーを定義（重複を避けるため）
            void NavigateHandler(object? s, RoutedEventArgs e)
            {
                Loaded -= NavigateHandler; // 一度だけ実行されるように解除
                var nav = RootNavigation;
                if (nav != null && IsLoaded)
                {
                    try
                    {
                        nav.Navigate(pageType);
                    }
                    catch
                    {
                        // ナビゲーションに失敗した場合は無視
                    }
                }
            }
            
            var rootNav = RootNavigation;
            var isLoaded = IsLoaded;
            if (rootNav == null || !isLoaded)
            {
                // RootNavigationが初期化されていない、またはウィンドウが読み込まれていない場合は、Loadedイベントでナビゲート
                Loaded += NavigateHandler;
                return false;
            }
            
            try
            {
                return rootNav.Navigate(pageType);
            }
            catch
            {
                // ナビゲーションに失敗した場合は、Loadedイベントで再試行
                Loaded += NavigateHandler;
                return false;
            }
        }

        /// <summary>
        /// ページサービスを設定します
        /// </summary>
        /// <param name="navigationViewPageProvider">ナビゲーションビューページプロバイダー</param>
        public void SetPageService(INavigationViewPageProvider navigationViewPageProvider)
        {
            var rootNav = RootNavigation;
            if (rootNav != null)
            {
                rootNav.SetPageProviderService(navigationViewPageProvider);
            }
            else
            {
                // RootNavigationが初期化されていない場合は、Loadedイベントで設定
                // 一度だけ実行されるように、既に登録されているかチェック（簡易的な実装）
                void SetPageServiceHandler(object? s, RoutedEventArgs e)
                {
                    Loaded -= SetPageServiceHandler; // 一度だけ実行されるように解除
                    var nav = RootNavigation;
                    if (nav != null)
                    {
                        nav.SetPageProviderService(navigationViewPageProvider);
                    }
                }
                Loaded += SetPageServiceHandler;
            }
        }

        /// <summary>
        /// ウィンドウを表示します
        /// </summary>
        public void ShowWindow()
        {
            // テーマ適用と同じタイミングでウィンドウ位置とサイズを復元
            var settings = _windowSettingsService.GetSettings();
            var isMaximized = settings.State == WindowState.Maximized;
            
            // ウィンドウ位置とサイズを復元
            RestoreWindowSettings();
            
            // テーマカラーは既にApp.xaml.csのOnStartupで適用されているため、
            // ここでは再適用しない（ちらつきを防ぐため）
            
            // ウィンドウを表示（LoadedイベントでVisibilityがVisibleに設定されるまで非表示のまま）
            // これにより、テーマカラーが適用されてから表示されるためチラつきを防ぐ
            Show();
            WindowState = isMaximized ? WindowState.Maximized : WindowState.Normal;
        }

        /// <summary>
        /// ウィンドウを閉じます
        /// </summary>
        public void CloseWindow() => Close();

        #endregion INavigationWindow methods

        /// <summary>
        /// ウィンドウが閉じられようとしているときに呼び出されます
        /// </summary>
        /// <param name="e">キャンセルイベント引数</param>
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            // ウィンドウサイズと位置、テーマを一度のGetSettings呼び出しで保存（最適化）
            try
            {
                var settings = _windowSettingsService.GetSettings();
                
                // ウィンドウサイズと位置を保存（プロパティアクセスを最適化）
                var width = Width;
                var height = Height;
                var left = Left;
                var top = Top;
                var state = WindowState;
                
                settings.Width = width;
                settings.Height = height;
                settings.Left = left;
                settings.Top = top;
                settings.State = state;
                
                // 現在のテーマを保存（switch式で高速化）
                var currentTheme = ApplicationThemeManager.GetAppTheme();
                settings.Theme = currentTheme switch
                {
                    ApplicationTheme.Light => "Light",
                    ApplicationTheme.Dark => "Dark",
                    _ => "System"
                };
                
                // タブ情報を保存（キャッシュされたViewModelを使用して高速化）
                try
                {
                    // キャッシュされたViewModelを使用（なければ取得してキャッシュ）
                    if (_cachedExplorerPageViewModel == null)
                    {
                        _cachedExplorerPageViewModel = App.Services.GetService(ExplorerPageViewModelType) as ViewModels.Pages.ExplorerPageViewModel;
                    }
                    
                    if (_cachedExplorerPageViewModel != null)
                    {
                        if (_cachedExplorerPageViewModel.IsSplitPaneEnabled)
                        {
                            // 分割ペインモードの場合、左右のペインのタブ情報をそれぞれ保存
                            var leftPaneTabs = _cachedExplorerPageViewModel.LeftPaneTabs;
                            var leftPaneTabPaths = new System.Collections.Generic.List<string>(leftPaneTabs.Count);
                            foreach (var tab in leftPaneTabs)
                            {
                                leftPaneTabPaths.Add(tab.CurrentPath ?? string.Empty);
                            }
                            settings.LeftPaneTabPaths = leftPaneTabPaths;

                            var rightPaneTabs = _cachedExplorerPageViewModel.RightPaneTabs;
                            var rightPaneTabPaths = new System.Collections.Generic.List<string>(rightPaneTabs.Count);
                            foreach (var tab in rightPaneTabs)
                            {
                                rightPaneTabPaths.Add(tab.CurrentPath ?? string.Empty);
                            }
                            settings.RightPaneTabPaths = rightPaneTabPaths;
                        }
                        else
                        {
                            // 通常モードの場合、従来通り
                            var tabs = _cachedExplorerPageViewModel.Tabs;
                            var tabPaths = new System.Collections.Generic.List<string>(tabs.Count);
                            foreach (var tab in tabs)
                            {
                                // CurrentPathが空の場合はホームなので、空文字列として保存
                                tabPaths.Add(tab.CurrentPath ?? string.Empty);
                            }
                            settings.TabPaths = tabPaths;
                        }
                    }
                }
                catch
                {
                    // タブ情報の取得に失敗した場合は無視
                }
                
                _windowSettingsService.SaveSettings(settings);
            }
            catch
            {
                // エラーハンドリング：保存に失敗してもウィンドウの閉じる処理は続行
            }
            
            base.OnClosing(e);
        }

        /// <summary>
        /// ウィンドウが閉じられたときに呼び出されます
        /// </summary>
        /// <param name="e">イベント引数</param>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            // Make sure that closing this window will begin the process of closing the application.
            // Application.Currentは常に存在するため、nullチェックを削除して高速化
            Application.Current.Shutdown();
        }

        /// <summary>
        /// 保存されたウィンドウ設定を復元します
        /// </summary>
        private void RestoreWindowSettings()
        {
            var settings = _windowSettingsService.GetSettings();
            
            // 最小サイズを設定（保存された値が小さすぎる場合に備えて）
            const double minWidth = 800;
            const double minHeight = 600;
            MinWidth = minWidth;
            MinHeight = minHeight;

            // ウィンドウサイズを復元
            var width = settings.Width;
            var height = settings.Height;
            if (width >= minWidth && height >= minHeight)
            {
                Width = width;
                Height = height;
            }

            // ウィンドウ位置を復元（有効な値の場合のみ）
            var left = settings.Left;
            var top = settings.Top;
            
            // 早期リターンで不要な処理をスキップ（高速化）
            if (double.IsNaN(left) || double.IsNaN(top))
            {
                WindowStartupLocation = WindowStartupLocation.CenterScreen;
                return;
            }
            
            // 画面の範囲内にあることを確認（キャッシュを使用）
            if (!_cachedScreenWidth.HasValue || !_cachedScreenHeight.HasValue)
            {
                _cachedScreenWidth = SystemParameters.PrimaryScreenWidth;
                _cachedScreenHeight = SystemParameters.PrimaryScreenHeight;
            }
            
            var screenWidth = _cachedScreenWidth.Value;
            var screenHeight = _cachedScreenHeight.Value;
            
            if (left >= 0 && left < screenWidth && top >= 0 && top < screenHeight)
            {
                Left = left;
                Top = top;
                WindowStartupLocation = WindowStartupLocation.Manual;
            }
            else
            {
                WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }

            // ウィンドウ状態を復元（起動時は復元しない）
            // ShowWindow()でNormalに設定されるため、ここでは復元しない
        }


        /// <summary>
        /// ナビゲーションコントロールを取得します（明示的なインターフェース実装）
        /// </summary>
        /// <returns>ナビゲーションビュー</returns>
        /// <exception cref="NotImplementedException">常にスローされます</exception>
        INavigationView INavigationWindow.GetNavigation()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// サービスプロバイダーを設定します
        /// </summary>
        /// <param name="serviceProvider">サービスプロバイダー</param>
        /// <exception cref="NotImplementedException">常にスローされます</exception>
        public void SetServiceProvider(IServiceProvider serviceProvider)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// ビジュアルツリー内の指定された型の子要素を検索します
        /// </summary>
        /// <typeparam name="T">検索する型</typeparam>
        /// <param name="parent">親要素</param>
        /// <returns>見つかった要素、見つからない場合はnull</returns>
        private static T? FindVisualChild<T>(System.Windows.DependencyObject parent) where T : System.Windows.DependencyObject
        {
            if (parent == null)
                return null;

            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T t)
                {
                    return t;
                }

                var childOfChild = FindVisualChild<T>(child);
                if (childOfChild != null)
                {
                    return childOfChild;
                }
            }

            return null;
        }

        /// <summary>
        /// タブとListViewのスタイルを無効化してDynamicResourceの再評価を強制します
        /// </summary>
        /// <param name="window">対象のウィンドウ</param>
        public static void InvalidateTabAndListViewStyles(System.Windows.Window window)
        {
            var explorerPage = FindVisualChild<Views.Pages.ExplorerPage>(window);
            if (explorerPage == null)
                return;

            // TabControlとListViewを一度の走査で見つける（高速化）
            System.Windows.Controls.TabControl? tabControl = null;
            System.Windows.Controls.ListView? listView = null;
            
            FindTabControlAndListView(explorerPage, ref tabControl, ref listView);

            // TabControl内のすべてのTabItemのスタイルを無効化
            if (tabControl != null)
            {
                // プロパティをキャッシュして高速化
                var templateProperty = System.Windows.Controls.Control.TemplateProperty;
                tabControl.InvalidateProperty(templateProperty);
                InvalidateTabItems(tabControl);
            }

            // ListView内のすべてのListViewItemのスタイルを無効化
            if (listView != null)
            {
                // プロパティをキャッシュして高速化
                var itemContainerStyleProperty = System.Windows.Controls.ItemsControl.ItemContainerStyleProperty;
                listView.InvalidateProperty(itemContainerStyleProperty);
                InvalidateListViewItems(listView);
            }
        }

        /// <summary>
        /// ビジュアルツリーを走査してTabControlとListViewを見つけます
        /// </summary>
        private static void FindTabControlAndListView(
            System.Windows.DependencyObject parent,
            ref System.Windows.Controls.TabControl? tabControl,
            ref System.Windows.Controls.ListView? listView)
        {
            if (parent == null)
                return;

            // 両方見つかった場合は早期終了（ループの最初でチェックして高速化）
            if (tabControl != null && listView != null)
                return;

            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
            {
                // 両方見つかった場合は早期終了
                if (tabControl != null && listView != null)
                    return;

                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                
                // 型チェックを最適化（if-elseの方がswitch式よりわずかに高速）
                if (tabControl == null && child is System.Windows.Controls.TabControl tc)
                {
                    tabControl = tc;
                }
                else if (listView == null && child is System.Windows.Controls.ListView lv)
                {
                    listView = lv;
                }

                // 再帰的に検索
                FindTabControlAndListView(child, ref tabControl, ref listView);
            }
        }

        /// <summary>
        /// TabControl内のすべてのTabItemのスタイルを無効化します
        /// </summary>
        private static void InvalidateTabItems(System.Windows.Controls.TabControl tabControl)
        {
            // プロパティをキャッシュして高速化（静的フィールドアクセスを削減）
            var styleProperty = System.Windows.FrameworkElement.StyleProperty;
            var backgroundProperty = System.Windows.Controls.TabItem.BackgroundProperty;
            
            // VisualTreeHelperをキャッシュ（メソッド呼び出しを削減）
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(tabControl);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(tabControl, i);
                if (child is System.Windows.Controls.TabItem tabItem)
                {
                    tabItem.InvalidateProperty(styleProperty);
                    tabItem.InvalidateProperty(backgroundProperty);
                }
                
                // FindVisualChildの代わりに直接走査（高速化）
                var childChildrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(child);
                for (int j = 0; j < childChildrenCount; j++)
                {
                    var grandChild = System.Windows.Media.VisualTreeHelper.GetChild(child, j);
                    if (grandChild is System.Windows.Controls.Primitives.TabPanel tabPanel)
                    {
                        var panelChildrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(tabPanel);
                        for (int k = 0; k < panelChildrenCount; k++)
                        {
                            var tabItemChild = System.Windows.Media.VisualTreeHelper.GetChild(tabPanel, k);
                            if (tabItemChild is System.Windows.Controls.TabItem tabItem2)
                            {
                                tabItem2.InvalidateProperty(styleProperty);
                                tabItem2.InvalidateProperty(backgroundProperty);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// ListView内のすべてのListViewItemのスタイルを無効化します
        /// </summary>
        private static void InvalidateListViewItems(System.Windows.Controls.ListView listView)
        {
            // プロパティをキャッシュして高速化
            var styleProperty = System.Windows.FrameworkElement.StyleProperty;
            var backgroundProperty = System.Windows.Controls.Control.BackgroundProperty;
            
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(listView);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(listView, i);
                if (child is System.Windows.Controls.ListViewItem listViewItem)
                {
                    listViewItem.InvalidateProperty(styleProperty);
                    listViewItem.InvalidateProperty(backgroundProperty);
                }
            }
        }

        /// <summary>
        /// ナビゲーションビューのアイテムが選択されたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="args">ナビゲーションイベント引数</param>
        private void RootNavigation_ItemInvoked(object sender, object args)
        {
            // NavigationViewItemがクリックされた場合の処理
            // リフレクションを使用してInvokedItemContainerプロパティにアクセス（キャッシュを使用）
            var argsType = args.GetType();
            
            // キャッシュされた型と一致しない場合は、プロパティを再取得（ReferenceEqualsで高速化）
            if (!ReferenceEquals(_cachedArgsType, argsType))
            {
                _cachedArgsType = argsType;
                _cachedInvokedItemContainerProperty = argsType.GetProperty("InvokedItemContainer");
            }
            
            if (_cachedInvokedItemContainerProperty == null)
                return;
            
            var invokedItem = _cachedInvokedItemContainerProperty.GetValue(args) as NavigationViewItem;
            if (invokedItem?.Tag is not string tag)
                return;
            
            // ViewModelをキャッシュから取得（なければ取得してキャッシュ）
            if (_cachedExplorerPageViewModel == null)
            {
                _cachedExplorerPageViewModel = App.Services.GetService(ExplorerPageViewModelType) as ViewModels.Pages.ExplorerPageViewModel;
            }
            
            var selectedTab = _cachedExplorerPageViewModel?.SelectedTab;
            
            // エクスプローラーページにナビゲート（共通処理、nullチェック削減で高速化）
            _navigationService?.Navigate(ExplorerPageType);
            
            // ページが読み込まれるのを待ってから処理を実行
            // DispatcherPriority.Normalを使用することで、レイアウト完了を待たずに高速に実行される
            if (selectedTab?.ViewModel == null)
                return;
            
            var viewModel = selectedTab.ViewModel;
            // ReadOnlySpan<char>を使用してメモリ割り当てを削減（高速化）
            var isHome = tag.AsSpan().SequenceEqual(HomeTag.AsSpan());
            
            // ホームアイテムとお気に入りアイテムの処理を統合（メモリ割り当て削減）
            // DispatcherPriority.Normalに変更して高速化（Loadedはレイアウト完了まで待つため遅い）
            // 1つのBeginInvokeに統合してオーバーヘッドを削減
            // 条件分岐を簡略化してキャプチャする変数を削減（高速化）
            if (isHome)
            {
                _ = Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    // ホームボタンを押したときも履歴に追加して、ブラウザーバックで戻れるようにする
                    // NavigateToHome内でCurrentPathが空でない場合のみ履歴に追加される
                    viewModel.NavigateToHome(addToHistory: true);
                }), DispatcherPriority.Normal);
            }
            else
            {
                var path = tag; // クロージャで使用するため変数に保存
                _ = Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    viewModel.NavigateToPathCommand.Execute(path);
                }), DispatcherPriority.Normal);
            }
            // TargetPageTypeが設定されている場合（Settingsなど）は自動的にナビゲートされる
        }

        /// <summary>
        /// ウィンドウでキーが押されたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">キーイベント引数</param>
        /// <summary>
        /// ソース初期化時にメッセージフックを追加
        /// </summary>
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            _hwndSource = (HwndSource)PresentationSource.FromVisual(this);
            if (_hwndSource != null)
            {
                _wndProcHook = WndProc;
                _hwndSource.AddHook(_wndProcHook);
            }
        }

        /// <summary>
        /// メッセージフックを一時的に無効化します（WindowChromeの処理を妨げないようにするため）
        /// </summary>
        public void DisableMessageHook()
        {
            if (_hwndSource != null && _wndProcHook != null)
            {
                _hwndSource.RemoveHook(_wndProcHook);
            }
        }

        /// <summary>
        /// メッセージフックを再有効化します
        /// </summary>
        public void EnableMessageHook()
        {
            if (_hwndSource != null && _wndProcHook != null)
            {
                _hwndSource.AddHook(_wndProcHook);
            }
        }

        /// <summary>
        /// ウィンドウメッセージを処理（IContextMenu3のメッセージフック用）
        /// </summary>
        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            // メニュー関連のメッセージIDを定義（高速チェック用）
            const int WM_INITMENUPOPUP = 0x0117;
            const int WM_DRAWITEM = 0x002B;
            const int WM_MEASUREITEM = 0x002C;
            const int WM_MENUCHAR = 0x0120;
            
            // 重要: WindowChromeが非クライアント領域のメッセージを処理するため、
            // これらのメッセージはメッセージフックで処理せず、完全にスキップする
            // メッセージフックが呼ばれるだけで、WindowChromeの処理に影響を与える可能性があるため、
            // これらのメッセージは即座にスキップする
            // 特に、WM_NCHITTESTはWindowChromeが頻繁に呼び出すメッセージのため、
            // メッセージフックで処理するとWindowChromeの処理が妨げられる可能性がある
            // メッセージIDを直接チェックして、高速にスキップする
            // メニュー表示中でも、非クライアント領域のメッセージはWindowChromeに完全に委譲する
            switch (msg)
            {
                // 非クライアント領域のメッセージを即座にスキップ
                // これらのメッセージはWindowChromeが処理する必要があるため、
                // メッセージフックで何も処理せず、即座に返す
                case 0x0084: // WM_NCHITTEST - WindowChromeが頻繁に呼び出す（マウス移動時に毎回呼ばれる）
                case 0x00A1: // WM_NCLBUTTONDOWN
                case 0x00A2: // WM_NCLBUTTONUP
                case 0x00A4: // WM_NCRBUTTONDOWN
                case 0x00A5: // WM_NCRBUTTONUP
                case 0x00A0: // WM_NCMOUSEMOVE
                case 0x02A2: // WM_NCMOUSELEAVE
                case 0x0086: // WM_NCACTIVATE
                case 0x0112: // WM_SYSCOMMAND
                case 0x0010: // WM_CLOSE
                case 0x0012: // WM_QUIT
                case 0x0006: // WM_ACTIVATE
                case 0x001C: // WM_ACTIVATEAPP
                    // 重要: handledをfalseにして、WindowChromeに処理を委譲する
                    // メッセージフックで何も処理せず、即座に返すことで、WindowChromeの処理を妨げない
                    // メニュー表示中でも、これらのメッセージはWindowChromeに完全に委譲する
                    handled = false;
                    return IntPtr.Zero;
            }

            // メニュー関連のメッセージかどうかをチェック
            // メニュー関連のメッセージのみを処理し、それ以外は即座にスキップする
            if (msg != WM_INITMENUPOPUP && 
                msg != WM_DRAWITEM && 
                msg != WM_MEASUREITEM && 
                msg != WM_MENUCHAR)
            {
                // メニュー関連のメッセージでない場合は、メッセージフックで処理しない
                // WindowChromeやその他のWPFの処理に委譲する
                handled = false;
                return IntPtr.Zero;
            }

            // メニュー関連のメッセージの場合のみ、メニューが表示中かチェック
            if (!FastExplorer.ShellContextMenu.ShellContextMenuService.IsMenuShowing)
            {
                // メニューが表示されていない場合は、メッセージを通常処理に任せる
                handled = false;
                return IntPtr.Zero;
            }

            // メニューが表示中で、メニュー関連のメッセージの場合のみProcessWindowMessageを呼び出す
            bool wasHandled = false;
            IntPtr result = FastExplorer.ShellContextMenu.ShellContextMenuService.ProcessWindowMessage(hwnd, msg, wParam, lParam, ref wasHandled);
            
            // ProcessWindowMessageがメッセージを処理した場合のみhandledをtrueにする
            if (wasHandled)
            {
                handled = true;
                return result;
            }

            // 処理されなかった場合は、通常処理に任せる
            handled = false;
            return IntPtr.Zero;
        }

        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            // バックスペースキーが押された場合、戻るボタンと同じ動作をする
            if (e.Key != Key.Back)
                return;

            // テキストボックスやエディットコントロールにフォーカスがある場合は処理しない（switch式で高速化）
            var source = e.OriginalSource;
            if (source is System.Windows.Controls.TextBox or 
                System.Windows.Controls.TextBlock or 
                System.Windows.Controls.RichTextBox)
            {
                return;
            }

            // ViewModelをキャッシュから取得（なければ取得してキャッシュ）
            if (_cachedExplorerPageViewModel == null)
            {
                _cachedExplorerPageViewModel = App.Services.GetService(ExplorerPageViewModelType) as ViewModels.Pages.ExplorerPageViewModel;
            }

            var selectedTab = _cachedExplorerPageViewModel?.SelectedTab;
            if (selectedTab?.ViewModel != null)
            {
                selectedTab.ViewModel.NavigateToParentCommand.Execute(null);
                e.Handled = true;
            }
        }
    }
}
