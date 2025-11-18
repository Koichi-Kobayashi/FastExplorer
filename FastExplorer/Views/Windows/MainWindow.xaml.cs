using System;
using System.Threading;
using System.Threading.Tasks;
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

            SystemThemeWatcher.Watch(this);

            // 最初は非表示で作成（テーマ適用後に表示するため）
            Visibility = Visibility.Hidden;
            ShowInTaskbar = false; // 初期状態ではタスクバーに表示しない

            InitializeComponent();
            
            // RootNavigationが初期化されるのを待つ
            if (RootNavigation != null)
            {
                SetPageService(navigationViewPageProvider);
                navigationService.SetNavigationControl(RootNavigation);
            }
            else
            {
                // RootNavigationが初期化されていない場合は、Loadedイベントで設定
                Loaded += (s, e) =>
                {
                    if (RootNavigation != null)
                    {
                        SetPageService(navigationViewPageProvider);
                        navigationService.SetNavigationControl(RootNavigation);
                    }
                };
            }
            
            // ViewModelにNavigationServiceを設定
            viewModel.SetNavigationService(navigationService);

            // 保存されたウィンドウ設定を復元（位置とサイズのみ、状態は復元しない）
            RestoreWindowSettings();
            
            // Loadedイベントでテーマを確認してから表示
            Loaded += MainWindow_Loaded;
        }

        /// <summary>
        /// ウィンドウが読み込まれたときに呼び出されます
        /// </summary>
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // テーマが適用されていることを確認
            App.UpdateThemeResourcesInternal();
            
            // 表示されていることを確認（ShowWindowで既にVisibleに設定されているはず）
            if (Visibility != Visibility.Visible)
            {
                Visibility = Visibility.Visible;
            }
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
            if (RootNavigation == null || !IsLoaded)
            {
                // RootNavigationが初期化されていない、またはウィンドウが読み込まれていない場合は、Loadedイベントでナビゲート
                Loaded += (s, e) =>
                {
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
                };
                return false;
            }
            
            try
            {
                return RootNavigation.Navigate(pageType);
            }
            catch
            {
                // ナビゲーションに失敗した場合は、Loadedイベントで再試行
                Loaded += (s, e) =>
                {
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
                };
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
                Loaded += (s, e) =>
                {
                    if (RootNavigation != null)
                    {
                        RootNavigation.SetPageProviderService(navigationViewPageProvider);
                    }
                };
            }
        }

        /// <summary>
        /// ウィンドウを表示します
        /// </summary>
        public async void ShowWindow()
        {
            // テーマを確認して適用（表示前に確実に適用）
            App.UpdateThemeResourcesInternal();
            
            // 保存されたテーマカラーを適用
            var settings = _windowSettingsService.GetSettings();
            if (!string.IsNullOrEmpty(settings.ThemeColorCode))
            {
                App.ApplyThemeColorFromSettings(settings);
            }
            
            // ウィンドウ位置を中央に設定（非表示の場合は位置が設定されていない可能性があるため）
            if (WindowStartupLocation != WindowStartupLocation.Manual)
            {
                WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
            
            // テーマが完全に適用されるまで待機
            await Task.Delay(100);
            
            // UIスレッドでウィンドウを表示（テーマ適用後に表示）
            _ = Dispatcher.BeginInvoke(new System.Action(() =>
            {
                Visibility = Visibility.Visible;
                Show();
                
                // タスクバーに表示するように戻す
                ShowInTaskbar = true;
                
                // 通常表示に設定（最大化が保存されていた場合は後で復元）
                WindowState = WindowState.Normal;
                
                // 保存されたウィンドウ状態を復元（最大化の場合）
                if (settings.State == WindowState.Maximized)
                {
                    WindowState = WindowState.Maximized;
                }

                // ウィンドウが表示された後にテーマカラーを再適用（確実に反映させるため）
                if (!string.IsNullOrEmpty(settings.ThemeColorCode))
                {
                    Dispatcher.BeginInvoke(new System.Action(() =>
                    {
                        App.ApplyThemeColorFromSettings(settings);
                    }), System.Windows.Threading.DispatcherPriority.Loaded);
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
                settings.Theme = currentTheme switch
                {
                    ApplicationTheme.Light => "Light",
                    ApplicationTheme.Dark => "Dark",
                    _ => "System" // ApplicationTheme.Unknownの場合は"System"として保存
                };
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
            MinWidth = 800;
            MinHeight = 600;

            // ウィンドウサイズを復元
            if (settings.Width >= MinWidth && settings.Height >= MinHeight)
            {
                Width = settings.Width;
                Height = settings.Height;
            }

            // ウィンドウ位置を復元（有効な値の場合のみ、かつウィンドウが表示されている場合のみ）
            if (!double.IsNaN(settings.Left) && !double.IsNaN(settings.Top) && Visibility == Visibility.Visible)
            {
                // 画面の範囲内にあることを確認
                var screenWidth = SystemParameters.PrimaryScreenWidth;
                var screenHeight = SystemParameters.PrimaryScreenHeight;
                
                if (settings.Left >= 0 && settings.Left < screenWidth &&
                    settings.Top >= 0 && settings.Top < screenHeight)
                {
                    Left = settings.Left;
                    Top = settings.Top;
                    WindowStartupLocation = WindowStartupLocation.Manual;
                }
            }
            else if (Visibility == Visibility.Hidden)
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
            // リフレクションを使用してInvokedItemContainerプロパティにアクセス
            var argsType = args.GetType();
            var invokedItemContainerProperty = argsType.GetProperty("InvokedItemContainer");
            
            if (invokedItemContainerProperty != null)
            {
                var invokedItem = invokedItemContainerProperty.GetValue(args) as NavigationViewItem;
                if (invokedItem != null)
                {
                    // ホームアイテムの場合
                    if (invokedItem.Tag is string tag && tag == "HOME")
                    {
                        // エクスプローラーページにナビゲート
                        _navigationService?.Navigate(typeof(Views.Pages.ExplorerPage));
                        
                        // 少し遅延してからホームページを表示（ページが読み込まれるのを待つ）
                        Task.Delay(100).ContinueWith(_ =>
                        {
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                var explorerPageViewModel = App.Services.GetService(typeof(ViewModels.Pages.ExplorerPageViewModel)) as ViewModels.Pages.ExplorerPageViewModel;
                                if (explorerPageViewModel != null && explorerPageViewModel.SelectedTab != null)
                                {
                                    // ホームページにナビゲート
                                    explorerPageViewModel.SelectedTab.ViewModel.NavigateToHome();
                                }
                            });
                        });
                    }
                    // お気に入りアイテムの場合（Tagにパスが設定されている）
                    else if (invokedItem.Tag is string path && path != "HOME")
                    {
                        var explorerPageViewModel = App.Services.GetService(typeof(ViewModels.Pages.ExplorerPageViewModel)) as ViewModels.Pages.ExplorerPageViewModel;
                        if (explorerPageViewModel != null && explorerPageViewModel.SelectedTab != null)
                        {
                            // エクスプローラーページにナビゲート
                            _navigationService?.Navigate(typeof(Views.Pages.ExplorerPage));
                            
                            // 少し遅延してからパスを設定（ページが読み込まれるのを待つ）
                            Task.Delay(100).ContinueWith(_ =>
                            {
                                Application.Current.Dispatcher.Invoke(() =>
                                {
                                    if (explorerPageViewModel.SelectedTab != null)
                                    {
                                        explorerPageViewModel.SelectedTab.ViewModel.NavigateToPathCommand.Execute(path);
                                    }
                                });
                            });
                        }
                    }
                    // TargetPageTypeが設定されている場合（Settingsなど）は自動的にナビゲートされる
                }
            }
        }
    }
}
