using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Threading;
using FastExplorer.Services;
using FastExplorer.ViewModels.Pages;
using FastExplorer.ViewModels.Windows;
using FastExplorer.Views.Pages;
using FastExplorer.Views.Windows;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Wpf.Ui;
using Wpf.Ui.Appearance;
using Wpf.Ui.DependencyInjection;

namespace FastExplorer
{
    /// <summary>
    /// App.xamlの相互作用ロジック
    /// </summary>
    public partial class App
    {
        private ResourceDictionary? _darkThemeResources;
        private Views.Windows.SplashWindow? _splashWindow;

        // The.NET Generic Host provides dependency injection, configuration, logging, and other services.
        // https://docs.microsoft.com/dotnet/core/extensions/generic-host
        // https://docs.microsoft.com/dotnet/core/extensions/dependency-injection
        // https://docs.microsoft.com/dotnet/core/extensions/configuration
        // https://docs.microsoft.com/dotnet/core/extensions/logging
        private static readonly IHost _host = Host
            .CreateDefaultBuilder()
            .ConfigureAppConfiguration(c => 
            { 
                var basePath = Path.GetDirectoryName(AppContext.BaseDirectory);
                if (!string.IsNullOrEmpty(basePath))
                {
                    c.SetBasePath(basePath);
                }
            })
            .ConfigureServices((context, services) =>
            {
                services.AddNavigationViewPageProvider();

                services.AddHostedService<ApplicationHostService>();

                // Theme manipulation
                services.AddSingleton<IThemeService, ThemeService>();

                // TaskBar manipulation
                services.AddSingleton<ITaskBarService, TaskBarService>();

                // Service containing navigation, same as INavigationWindow... but without window
                services.AddSingleton<INavigationService, NavigationService>();

                // Main window with navigation
                services.AddSingleton<INavigationWindow, MainWindow>();
                services.AddSingleton<MainWindowViewModel>();

                // FileSystemService
                services.AddSingleton<Services.FileSystemService>();

                // FavoriteService
                services.AddSingleton<Services.FavoriteService>();

                // WindowSettingsService
                services.AddSingleton<Services.WindowSettingsService>();

                services.AddSingleton<DashboardPage>();
                services.AddSingleton<DashboardViewModel>();
                services.AddSingleton<DataPage>();
                services.AddSingleton<DataViewModel>();
                services.AddSingleton<Views.Pages.ExplorerPage>();
                services.AddSingleton<ViewModels.Pages.ExplorerPageViewModel>();
                services.AddSingleton<SettingsPage>();
                services.AddSingleton<SettingsViewModel>();
            }).Build();

        /// <summary>
        /// サービスプロバイダーを取得します
        /// </summary>
        public static IServiceProvider Services
        {
            get { return _host.Services; }
        }

        /// <summary>
        /// アプリケーションが読み込まれるときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">スタートアップイベント引数</param>
        private async void OnStartup(object sender, StartupEventArgs e)
        {
            // ダークモード用のリソースディクショナリーを読み込む
            _darkThemeResources = new ResourceDictionary
            {
                Source = new Uri("pack://application:,,,/Resources/DarkThemeResources.xaml", UriKind.Absolute)
            };

            // テーマを適用してからウィンドウを表示
            LoadAndApplyThemeOnStartup();
            UpdateThemeResources();

            // スプラッシュウィンドウを作成（非表示状態）
            _splashWindow = new Views.Windows.SplashWindow();
            
            // リソースが完全に解決されるのを待ってから表示
            _ = Dispatcher.BeginInvoke(new System.Action(() =>
            {
                // 再度テーマを確認
                UpdateThemeResources();
                
                // スプラッシュウィンドウの色を適用
                if (_splashWindow != null && !_splashWindow.IsClosed())
                {
                    _splashWindow.ApplyThemeColors();
                }
                
                // UIスレッドでスプラッシュウィンドウを表示
                _ = Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    if (_splashWindow != null && !_splashWindow.IsClosed())
                    {
                        try
                        {
                            // スプラッシュウィンドウを表示
                            if (_splashWindow.Visibility != Visibility.Visible)
                            {
                                _splashWindow.Visibility = Visibility.Visible;
                            }
                            
                            // ウィンドウがまだ表示されていない場合のみShow()を呼ぶ
                            if (!_splashWindow.IsLoaded)
                            {
                                _splashWindow.Show();
                            }
                            
                            _splashWindow.UpdateLayout();
                        }
                        catch (InvalidOperationException)
                        {
                            // ウィンドウが既に閉じられている場合は無視
                        }
                    }
                }), DispatcherPriority.Loaded);
            }), DispatcherPriority.Loaded);

            await _host.StartAsync();

            // ホスト起動後にもう一度テーマを確認して適用（確実に適用するため）
            _ = Dispatcher.BeginInvoke(new System.Action(() =>
            {
                LoadAndApplyThemeOnStartup();
                UpdateThemeResources();
            }), DispatcherPriority.Loaded);
        }

        /// <summary>
        /// 起動時に保存されたテーマを読み込んで適用します
        /// </summary>
        private void LoadAndApplyThemeOnStartup()
        {
            try
            {
                // 設定ファイルから直接読み込む（WindowSettingsServiceを待たない）
                var appDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "FastExplorer");
                var settingsFilePath = Path.Combine(appDataPath, "window_settings.json");

                ApplicationTheme themeToApply = ApplicationTheme.Light;
                
                if (File.Exists(settingsFilePath))
                {
                    try
                    {
                        var json = File.ReadAllText(settingsFilePath);
                        var settings = System.Text.Json.JsonSerializer.Deserialize<Services.WindowSettings>(json);
                        
                        if (settings != null && !string.IsNullOrEmpty(settings.Theme))
                        {
                            themeToApply = settings.Theme == "Dark" ? ApplicationTheme.Dark : ApplicationTheme.Light;
                        }
                    }
                    catch
                    {
                        // JSONの読み込みに失敗した場合はデフォルトのテーマを使用
                    }
                }

                // ThemesDictionaryのThemeプロパティを先に更新（ApplicationThemeManager.Apply()の前に）
                if (Current.Resources is ResourceDictionary mainDictionary)
                {
                    var mergedDictionaries = mainDictionary.MergedDictionaries;
                    var themesDict = mergedDictionaries
                        .OfType<ResourceDictionary>()
                        .FirstOrDefault(rd => rd.GetType().Name == "ThemesDictionary");
                    
                    if (themesDict != null)
                    {
                        var themeProperty = themesDict.GetType().GetProperty("Theme");
                        if (themeProperty != null)
                        {
                            var themeType = themeProperty.PropertyType;
                            var themeValue = Enum.Parse(themeType, themeToApply == ApplicationTheme.Dark ? "Dark" : "Light");
                            themeProperty.SetValue(themesDict, themeValue);
                        }
                    }
                }

                // テーマを適用
                ApplicationThemeManager.Apply(themeToApply);
                
                // リソースディクショナリーも即座に更新
                UpdateThemeResources();
            }
            catch
            {
                // エラーハンドリング：デフォルトのテーマを使用
                ApplicationThemeManager.Apply(ApplicationTheme.Light);
            }
        }

        /// <summary>
        /// テーマに応じてリソースディクショナリーを更新します
        /// </summary>
        private void UpdateThemeResources()
        {
            UpdateThemeResourcesInternal();
        }

        /// <summary>
        /// テーマに応じてリソースディクショナリーを更新します（静的メソッド）
        /// </summary>
        public static void UpdateThemeResourcesInternal()
        {
            if (Current is App app && app._darkThemeResources != null)
            {
                if (Current.Resources is ResourceDictionary mainDictionary)
                {
                    var mergedDictionaries = mainDictionary.MergedDictionaries;

                    // 既存のダークテーマリソースを削除
                    var existingDarkTheme = mergedDictionaries
                        .OfType<ResourceDictionary>()
                        .FirstOrDefault(rd => rd.Source?.OriginalString?.Contains("DarkThemeResources.xaml") == true);

                    var isDark = ApplicationThemeManager.GetAppTheme() == ApplicationTheme.Dark;
                    var needsUpdate = false;

                    if (isDark && existingDarkTheme == null)
                    {
                        // ダークモードだがリソースが追加されていない
                        needsUpdate = true;
                    }
                    else if (!isDark && existingDarkTheme != null)
                    {
                        // ライトモードだがリソースが残っている
                        needsUpdate = true;
                    }

                    // 必要な場合のみ更新（ちらつきを防ぐため）
                    if (needsUpdate)
                    {
                        if (existingDarkTheme != null)
                        {
                            mergedDictionaries.Remove(existingDarkTheme);
                        }

                        // ダークモードの場合は追加
                        if (isDark)
                        {
                            mergedDictionaries.Add(app._darkThemeResources);
                        }
                    }
                }
            }

            // すべてのThemedSvgIconインスタンスにブラシを再適用（遅延実行でちらつきを防ぐ）
            Current.Dispatcher.BeginInvoke(new System.Action(() =>
            {
                FastExplorer.Controls.ThemedSvgIcon.RefreshAllInstances();
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>
        /// アプリケーションが閉じられるときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">終了イベント引数</param>
        private async void OnExit(object sender, ExitEventArgs e)
        {
            // 終了時に現在のテーマを保存
            SaveCurrentTheme();

            await _host.StopAsync();

            _host.Dispose();
        }

        /// <summary>
        /// 現在のテーマを保存します
        /// </summary>
        private void SaveCurrentTheme()
        {
            try
            {
                var windowSettingsService = Services.GetService(typeof(WindowSettingsService)) as WindowSettingsService;
                if (windowSettingsService != null)
                {
                    var currentTheme = ApplicationThemeManager.GetAppTheme();
                    var settings = windowSettingsService.GetSettings();
                    settings.Theme = currentTheme == ApplicationTheme.Light ? "Light" : "Dark";
                    windowSettingsService.SaveSettings(settings);
                }
            }
            catch
            {
                // エラーハンドリング：保存に失敗してもアプリケーションの終了は続行
            }
        }

        /// <summary>
        /// アプリケーションによってスローされたが処理されていない例外が発生したときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">未処理例外イベント引数</param>
        private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            // For more info see https://docs.microsoft.com/en-us/dotnet/api/system.windows.application.dispatcherunhandledexception?view=windowsdesktop-6.0
        }
    }
}
