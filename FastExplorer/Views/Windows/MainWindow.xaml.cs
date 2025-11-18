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

            // 最初は非表示で作成（テーマ適用後に表示するため）
            Visibility = Visibility.Hidden;
            ShowInTaskbar = false; // 初期状態ではタスクバーに表示しない

            InitializeComponent();
            
            // RootNavigationが初期化されるのを待つ（起動時の高速化のため、Loadedイベントで遅延）
            void InitializeNavigationHandler(object? s, RoutedEventArgs e)
            {
                Loaded -= InitializeNavigationHandler; // 一度だけ実行されるように解除
                if (RootNavigation != null)
                {
                    SetPageService(navigationViewPageProvider);
                    navigationService.SetNavigationControl(RootNavigation);
                }
            }
            Loaded += InitializeNavigationHandler;
            
            // ViewModelにNavigationServiceを設定（起動時の高速化のため、Loadedイベントで遅延）
            Loaded += (s, e) =>
            {
                viewModel.SetNavigationService(navigationService);
                RestoreWindowSettings();
                // SystemThemeWatcherも遅延読み込み（起動時の高速化）
                SystemThemeWatcher.Watch(this);
            };
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
                if (RootNavigation != null && IsLoaded)
                {
                    try
                    {
                        RootNavigation.Navigate(pageType);
                    }
                    catch
                    {
                        // ナビゲーションに失敗した場合は無視
                    }
                }
            }
            
            if (RootNavigation == null || !IsLoaded)
            {
                // RootNavigationが初期化されていない、またはウィンドウが読み込まれていない場合は、Loadedイベントでナビゲート
                Loaded += NavigateHandler;
                return false;
            }
            
            try
            {
                return RootNavigation.Navigate(pageType);
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
            if (RootNavigation != null)
            {
                RootNavigation.SetPageProviderService(navigationViewPageProvider);
            }
            else
            {
                // RootNavigationが初期化されていない場合は、Loadedイベントで設定
                // 一度だけ実行されるように、既に登録されているかチェック（簡易的な実装）
                void SetPageServiceHandler(object? s, RoutedEventArgs e)
                {
                    Loaded -= SetPageServiceHandler; // 一度だけ実行されるように解除
                    if (RootNavigation != null)
                    {
                        RootNavigation.SetPageProviderService(navigationViewPageProvider);
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
            // テーマは既に起動時に適用されているため、ここでは適用しない（起動時の高速化）
            // 保存されたテーマカラーを適用
            var settings = _windowSettingsService.GetSettings();
            var themeColorCode = settings.ThemeColorCode;
            var hasThemeColor = themeColorCode != null && themeColorCode.Length > 0;
            var isMaximized = settings.State == WindowState.Maximized;
            
            if (hasThemeColor)
            {
                App.ApplyThemeColorFromSettings(settings);
            }
            
            // ウィンドウ位置を中央に設定（非表示の場合は位置が設定されていない可能性があるため）
            if (WindowStartupLocation != WindowStartupLocation.Manual)
            {
                WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
            
            // UIスレッドでウィンドウを表示（テーマ適用後に表示）
            // DispatcherPriority.Loadedを使用することで、レイアウトが完了してから実行される
            // メモリ割り当てを削減するため、クロージャで変数をキャプチャ
            _ = Dispatcher.BeginInvoke(new System.Action(() =>
            {
                Visibility = Visibility.Visible;
                Show();
                
                // タスクバーに表示するように戻す
                ShowInTaskbar = true;
                
                // 通常表示に設定（最大化が保存されていた場合は後で復元）
                WindowState = WindowState.Normal;
                
                // 保存されたウィンドウ状態を復元（最大化の場合）
                if (isMaximized)
                {
                    WindowState = WindowState.Maximized;
                }

                // ウィンドウが表示された後にテーマカラーを再適用（確実に反映させるため）
                // Render優先度で実行することで、レンダリング後に確実に適用される
                // ただし、起動時の高速化のため、テーマカラーが設定されている場合のみ実行
                if (hasThemeColor)
                {
                    Dispatcher.BeginInvoke(new System.Action(() =>
                    {
                        App.ApplyThemeColorFromSettings(settings);
                    }), System.Windows.Threading.DispatcherPriority.Render);
                }
            }), System.Windows.Threading.DispatcherPriority.Loaded);
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
            // ウィンドウサイズと位置を保存
            SaveWindowSettings();
            
            // 現在のテーマを保存
            SaveCurrentTheme();
            
            base.OnClosing(e);
        }

        /// <summary>
        /// 現在のテーマを保存します
        /// </summary>
        private void SaveCurrentTheme()
        {
            try
            {
                var currentTheme = ApplicationThemeManager.GetAppTheme();
                var settings = _windowSettingsService.GetSettings();
                // switch式を最適化（文字列リテラルを直接使用）
                settings.Theme = currentTheme == ApplicationTheme.Light ? "Light" :
                                currentTheme == ApplicationTheme.Dark ? "Dark" : "System";
                _windowSettingsService.SaveSettings(settings);
            }
            catch
            {
                // エラーハンドリング：保存に失敗してもウィンドウの閉じる処理は続行
            }
        }

        /// <summary>
        /// ウィンドウが閉じられたときに呼び出されます
        /// </summary>
        /// <param name="e">イベント引数</param>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            // Make sure that closing this window will begin the process of closing the application.
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
            if (settings.Width >= minWidth && settings.Height >= minHeight)
            {
                Width = settings.Width;
                Height = settings.Height;
            }

            // ウィンドウ位置を復元（有効な値の場合のみ、かつウィンドウが表示されている場合のみ）
            var isVisible = Visibility == Visibility.Visible;
            var left = settings.Left;
            var top = settings.Top;
            
            if (!double.IsNaN(left) && !double.IsNaN(top) && isVisible)
            {
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
                    return;
                }
            }
            
            if (!isVisible)
            {
                // 非表示の場合は中央に配置（表示時に適用される）
                WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }

            // ウィンドウ状態を復元（起動時は復元しない）
            // ShowWindow()でNormalに設定されるため、ここでは復元しない
        }

        /// <summary>
        /// ウィンドウ設定を保存します
        /// </summary>
        private void SaveWindowSettings()
        {
            // 既存の設定を取得して、ウィンドウのサイズと位置のみを更新
            var settings = _windowSettingsService.GetSettings();
            settings.Width = Width;
            settings.Height = Height;
            settings.Left = Left;
            settings.Top = Top;
            settings.State = WindowState;

            _windowSettingsService.SaveSettings(settings);
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
            
            // ホームアイテムの場合（文字列比較を最適化）
            if (string.Equals(tag, HomeTag, StringComparison.Ordinal))
            {
                // エクスプローラーページにナビゲート
                _navigationService?.Navigate(ExplorerPageType);
                
                // ページが読み込まれるのを待ってからホームページを表示
                // DispatcherPriority.Loadedを使用することで、レイアウトが完了してから実行される
                if (selectedTab != null)
                {
                    _ = Dispatcher.BeginInvoke(new System.Action(() =>
                    {
                        selectedTab.ViewModel.NavigateToHome();
                    }), DispatcherPriority.Loaded);
                }
            }
            // お気に入りアイテムの場合（Tagにパスが設定されている）
            else if (selectedTab != null)
            {
                // エクスプローラーページにナビゲート
                _navigationService?.Navigate(ExplorerPageType);
                
                // ページが読み込まれるのを待ってからパスを設定
                // DispatcherPriority.Loadedを使用することで、レイアウトが完了してから実行される
                var path = tag; // クロージャで使用するため変数に保存
                _ = Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    selectedTab.ViewModel.NavigateToPathCommand.Execute(path);
                }), DispatcherPriority.Loaded);
            }
            // TargetPageTypeが設定されている場合（Settingsなど）は自動的にナビゲートされる
        }
    }
}
