using System;
using System.Reflection;
using System.Windows;
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
                    // テーマカラー選択時と同じ処理
                    var mainColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(themeColorCode);
                    var mainBrush = new System.Windows.Media.SolidColorBrush(mainColor);
                    var secondaryColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(settings.ThemeSecondaryColorCode ?? "#FCFCFC");
                    var secondaryBrush = new System.Windows.Media.SolidColorBrush(secondaryColor);
                    
                    // ステータスバーのテキスト色を計算
                    var luminance = (0.299 * mainColor.R + 0.587 * mainColor.G + 0.114 * mainColor.B) / 255.0;
                    var statusBarTextColor = luminance > 0.5 ? System.Windows.Media.Colors.Black : System.Windows.Media.Colors.White;
                    var statusBarTextBrush = new System.Windows.Media.SolidColorBrush(statusBarTextColor);
                    
                    // ウィンドウの背景色を直接設定
                    Background = mainBrush;

                    // FluentWindowの場合は、Backgroundプロパティも更新
                    if (this is Wpf.Ui.Controls.FluentWindow fluentWindow)
                    {
                        fluentWindow.Background = mainBrush;
                    }

                    // NavigationViewの背景色も更新
                    if (nav != null)
                    {
                        nav.Background = secondaryBrush;
                    }

                    // ステータスバーの背景色とテキスト色を直接更新
                    if (StatusBar != null)
                    {
                        StatusBar.Background = mainBrush;
                    }
                    
                    if (StatusBarText != null)
                    {
                        StatusBarText.Foreground = statusBarTextBrush;
                    }
                }
                
                // レイアウトを更新してからウィンドウを表示（チラつきを防ぐ）
                UpdateLayout();
                
                // UpdateLayout()の後に再度色を設定（色が消えないようにする）
                if (themeColorCode != null && themeColorCode.Length > 0)
                {
                    var mainColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(themeColorCode);
                    var mainBrush = new System.Windows.Media.SolidColorBrush(mainColor);
                    var secondaryColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(settings.ThemeSecondaryColorCode ?? "#FCFCFC");
                    var secondaryBrush = new System.Windows.Media.SolidColorBrush(secondaryColor);
                    
                    // ウィンドウの背景色を再度設定（UpdateLayout()でリセットされる可能性があるため）
                    Background = mainBrush;

                    // FluentWindowの場合は、Backgroundプロパティも更新
                    if (this is Wpf.Ui.Controls.FluentWindow fluentWindow)
                    {
                        fluentWindow.Background = mainBrush;
                    }

                    // NavigationViewの背景色も再度設定
                    if (nav != null)
                    {
                        nav.Background = secondaryBrush;
                    }

                    // ステータスバーの背景色とテキスト色も再度設定
                    if (StatusBar != null)
                    {
                        StatusBar.Background = mainBrush;
                    }
                    
                    if (StatusBarText != null)
                    {
                        var luminance = (0.299 * mainColor.R + 0.587 * mainColor.G + 0.114 * mainColor.B) / 255.0;
                        var statusBarTextColor = luminance > 0.5 ? System.Windows.Media.Colors.Black : System.Windows.Media.Colors.White;
                        StatusBarText.Foreground = new System.Windows.Media.SolidColorBrush(statusBarTextColor);
                    }
                }
                
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
            // DispatcherPriority.Loadedを使用することで、レイアウトが完了してから実行される
            if (selectedTab == null)
                return;
            
            var dispatcher = Dispatcher;
            var viewModel = selectedTab.ViewModel;
            var isHome = string.Equals(tag, HomeTag, StringComparison.Ordinal);
            
            // ホームアイテムとお気に入りアイテムの処理を統合（メモリ割り当て削減）
            if (isHome)
            {
                _ = dispatcher.BeginInvoke(new System.Action(() =>
                {
                    viewModel.NavigateToHome();
                }), DispatcherPriority.Loaded);
            }
            else
            {
                var path = tag; // クロージャで使用するため変数に保存
                _ = dispatcher.BeginInvoke(new System.Action(() =>
                {
                    viewModel.NavigateToPathCommand.Execute(path);
                }), DispatcherPriority.Loaded);
            }
            // TargetPageTypeが設定されている場合（Settingsなど）は自動的にナビゲートされる
        }
    }
}
