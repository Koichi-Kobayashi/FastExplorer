using System;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using FastExplorer.Services;
using FastExplorer.ViewModels.Windows;
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

            ShowInTaskbar = false; // 初期状態ではタスクバーに表示しない

            InitializeComponent();
            
            // すべての初期化処理を1つのLoadedイベントハンドラーに統合（起動時の高速化）
            void InitializeHandler(object? s, RoutedEventArgs e)
            {
                Loaded -= InitializeHandler; // 一度だけ実行されるように解除
                
                // RootNavigationの初期化
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
                if (themeColorCode != null && themeColorCode.Length > 0)
                {
                    // リソースを更新（App.ApplyThemeColorFromSettingsを呼び出す）
                    App.ApplyThemeColorFromSettings(settings);
                    
                    // ウィンドウの背景色を直接設定（DynamicResourceが反映されるまでの間）
                    var mainColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(themeColorCode);
                    var mainBrush = new System.Windows.Media.SolidColorBrush(mainColor);
                    var secondaryColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(settings.ThemeSecondaryColorCode ?? "#FCFCFC");
                    var secondaryBrush = new System.Windows.Media.SolidColorBrush(secondaryColor);
                    var luminance = (0.299 * mainColor.R + 0.587 * mainColor.G + 0.114 * mainColor.B) / 255.0;
                    var statusBarTextColor = luminance > 0.5 ? System.Windows.Media.Colors.Black : System.Windows.Media.Colors.White;
                    var statusBarTextBrush = new System.Windows.Media.SolidColorBrush(statusBarTextColor);
                    
                    Background = mainBrush;
                    if (this is Wpf.Ui.Controls.FluentWindow fluentWindow)
                    {
                        fluentWindow.Background = mainBrush;
                    }
                    if (nav != null)
                    {
                        nav.Background = secondaryBrush;
                    }
                    if (StatusBar != null)
                    {
                        StatusBar.Background = mainBrush;
                    }
                    if (StatusBarText != null)
                    {
                        StatusBarText.Foreground = statusBarTextBrush;
                    }

                    // タブとListViewの選択中の色を更新するため、スタイルを無効化してDynamicResourceの再評価を強制
                    // ContentRenderedイベントで実行（確実にExplorerPageが読み込まれた後）
                    void ContentRenderedHandler(object? s, EventArgs e)
                    {
                        ContentRendered -= ContentRenderedHandler;
                        InvalidateTabAndListViewStyles(this);
                    }
                    ContentRendered += ContentRenderedHandler;
                }
                
                // レイアウトを更新してからウィンドウを表示（チラつきを防ぐ）
                UpdateLayout();
                
                // テーマカラー適用後にウィンドウを表示（チラつきを防ぐ）
                if (Visibility == Visibility.Hidden)
                {
                    Visibility = Visibility.Visible;
                    ShowInTaskbar = true;
                }
                
                // SystemThemeWatcherを遅延実行（起動を最速化）
                // ウィンドウ位置とサイズの復元はShowWindow()で実行されるため、ここでは実行しない
                _ = Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    SystemThemeWatcher.Watch(this);
                }), DispatcherPriority.Background);
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
                
                // 現在のテーマを保存
                var currentTheme = ApplicationThemeManager.GetAppTheme();
                if (currentTheme == ApplicationTheme.Light)
                    settings.Theme = "Light";
                else if (currentTheme == ApplicationTheme.Dark)
                    settings.Theme = "Dark";
                else
                    settings.Theme = "System";
                
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
                            var leftPaneTabPaths = new System.Collections.Generic.List<string>();
                            foreach (var tab in _cachedExplorerPageViewModel.LeftPaneTabs)
                            {
                                leftPaneTabPaths.Add(tab.CurrentPath ?? string.Empty);
                            }
                            settings.LeftPaneTabPaths = leftPaneTabPaths;

                            var rightPaneTabPaths = new System.Collections.Generic.List<string>();
                            foreach (var tab in _cachedExplorerPageViewModel.RightPaneTabs)
                            {
                                rightPaneTabPaths.Add(tab.CurrentPath ?? string.Empty);
                            }
                            settings.RightPaneTabPaths = rightPaneTabPaths;
                        }
                        else
                        {
                            // 通常モードの場合、従来通り
                            var tabPaths = new System.Collections.Generic.List<string>();
                            foreach (var tab in _cachedExplorerPageViewModel.Tabs)
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
            if (explorerPage != null)
            {
                // TabControl内のすべてのTabItemのスタイルを無効化
                var tabControl = FindVisualChild<System.Windows.Controls.TabControl>(explorerPage);
                if (tabControl != null)
                {
                    tabControl.InvalidateProperty(System.Windows.Controls.Control.TemplateProperty);
                    
                    for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(tabControl); i++)
                    {
                        var child = System.Windows.Media.VisualTreeHelper.GetChild(tabControl, i);
                        if (child is System.Windows.Controls.TabItem tabItem)
                        {
                            tabItem.InvalidateProperty(System.Windows.FrameworkElement.StyleProperty);
                            tabItem.InvalidateProperty(System.Windows.Controls.TabItem.BackgroundProperty);
                        }
                        var tabPanel = FindVisualChild<System.Windows.Controls.Primitives.TabPanel>(child);
                        if (tabPanel != null)
                        {
                            for (int j = 0; j < System.Windows.Media.VisualTreeHelper.GetChildrenCount(tabPanel); j++)
                            {
                                var tabItemChild = System.Windows.Media.VisualTreeHelper.GetChild(tabPanel, j);
                                if (tabItemChild is System.Windows.Controls.TabItem tabItem2)
                                {
                                    tabItem2.InvalidateProperty(System.Windows.FrameworkElement.StyleProperty);
                                    tabItem2.InvalidateProperty(System.Windows.Controls.TabItem.BackgroundProperty);
                                }
                            }
                        }
                    }
                }

                // ListView内のすべてのListViewItemのスタイルを無効化
                var listView = FindVisualChild<System.Windows.Controls.ListView>(explorerPage);
                if (listView != null)
                {
                    listView.InvalidateProperty(System.Windows.Controls.ItemsControl.ItemContainerStyleProperty);
                    
                    for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(listView); i++)
                    {
                        var child = System.Windows.Media.VisualTreeHelper.GetChild(listView, i);
                        if (child is System.Windows.Controls.ListViewItem listViewItem)
                        {
                            listViewItem.InvalidateProperty(System.Windows.FrameworkElement.StyleProperty);
                            listViewItem.InvalidateProperty(System.Windows.Controls.Control.BackgroundProperty);
                        }
                    }
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
            if (selectedTab == null)
                return;
            
            var viewModel = selectedTab.ViewModel;
            var isHome = string.Equals(tag, HomeTag, StringComparison.Ordinal);
            
            // ホームアイテムとお気に入りアイテムの処理を統合（メモリ割り当て削減）
            // DispatcherPriority.Normalに変更して高速化（Loadedはレイアウト完了まで待つため遅い）
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
        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            // バックスペースキーが押された場合、戻るボタンと同じ動作をする
            if (e.Key == Key.Back)
            {
                // テキストボックスやエディットコントロールにフォーカスがある場合は処理しない
                if (e.OriginalSource is System.Windows.Controls.TextBox || 
                    e.OriginalSource is System.Windows.Controls.TextBlock ||
                    e.OriginalSource is System.Windows.Controls.RichTextBox)
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
}
