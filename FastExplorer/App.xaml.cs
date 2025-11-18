using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
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
                if (basePath != null && basePath.Length > 0)
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
            // テーマを先に適用（起動時の高速化のため、同期的に実行）
            _darkThemeResources = new ResourceDictionary
            {
                Source = new Uri("pack://application:,,,/Resources/DarkThemeResources.xaml", UriKind.Absolute)
            };
            LoadAndApplyThemeOnStartup();
            
            // リソース更新は遅延実行（起動を最速化）
            _ = Dispatcher.BeginInvoke(new System.Action(() =>
            {
                UpdateThemeResources();
            }), DispatcherPriority.Background);

            // ウィンドウを表示
            await _host.StartAsync();
        }

        /// <summary>
        /// 起動時に保存されたテーマを読み込んで適用します
        /// </summary>
        private void LoadAndApplyThemeOnStartup()
        {
            try
            {
                // 設定ファイルから直接読み込む（WindowSettingsServiceを待たない）
                // パスをキャッシュして高速化
                var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var settingsFilePath = Path.Combine(localAppData, "FastExplorer", "window_settings.json");

                ApplicationTheme themeToApply = ApplicationTheme.Unknown;
                Services.WindowSettings? settings = null;
                
                if (File.Exists(settingsFilePath))
                {
                    try
                    {
                        var json = File.ReadAllText(settingsFilePath);
                        settings = System.Text.Json.JsonSerializer.Deserialize<Services.WindowSettings>(json);
                        
                        var themeStr = settings?.Theme;
                        if (themeStr != null && themeStr.Length > 0)
                        {
                            // switch式を最適化（文字列比較を高速化）
                            if (themeStr == "Dark")
                                themeToApply = ApplicationTheme.Dark;
                            else if (themeStr == "Light")
                                themeToApply = ApplicationTheme.Light;
                            // それ以外はUnknown（システムテーマ）
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
                    // LINQを避けて高速化（直接ループで検索）
                    ResourceDictionary? themesDict = null;
                    for (int i = 0; i < mergedDictionaries.Count; i++)
                    {
                        var dict = mergedDictionaries[i];
                        if (dict != null && dict.GetType().Name == "ThemesDictionary")
                        {
                            themesDict = dict;
                            break;
                        }
                    }
                    
                    if (themesDict != null)
                    {
                        var themeProperty = themesDict.GetType().GetProperty("Theme");
                        if (themeProperty != null)
                        {
                            var themeType = themeProperty.PropertyType;
                            string themeName;
                            
                            // ApplicationTheme.Unknownの場合は、システムテーマを取得
                            if (themeToApply == ApplicationTheme.Unknown)
                            {
                                var systemTheme = ApplicationThemeManager.GetSystemTheme();
                                themeName = systemTheme == SystemTheme.Dark ? "Dark" : "Light";
                            }
                            else
                            {
                                themeName = themeToApply == ApplicationTheme.Dark ? "Dark" : "Light";
                            }
                            
                            var themeValue = Enum.Parse(themeType, themeName);
                            themeProperty.SetValue(themesDict, themeValue);
                        }
                    }
                }

                // テーマを適用
                ApplicationThemeManager.Apply(themeToApply);

                // 保存されたテーマカラーを適用（リソース更新はOnStartupで遅延実行）
                var themeColorCode = settings?.ThemeColorCode;
                if (themeColorCode != null && themeColorCode.Length > 0)
                {
                    ApplyThemeColorOnStartup(settings!);
                }
            }
            catch
            {
                // エラーハンドリング：デフォルトのテーマ（システムテーマ）を使用
                ApplicationThemeManager.Apply(ApplicationTheme.Unknown);
            }
        }

        /// <summary>
        /// 起動時に保存されたテーマカラーを適用します
        /// </summary>
        /// <param name="settings">ウィンドウ設定</param>
        private void ApplyThemeColorOnStartup(Services.WindowSettings settings)
        {
            ApplyThemeColorFromSettings(settings);
        }

        /// <summary>
        /// 設定からテーマカラーを適用します（公開メソッド）
        /// </summary>
        /// <param name="settings">ウィンドウ設定</param>
        public static void ApplyThemeColorFromSettings(Services.WindowSettings settings)
        {
            try
            {
                if (Application.Current.Resources is ResourceDictionary mainDictionary)
                {
                    // メインカラーを適用
                    var mainColor = (Color)ColorConverter.ConvertFromString(settings.ThemeColorCode ?? "#F5F5F5");
                    var mainBrush = new SolidColorBrush(mainColor);
                    
                    // セカンダリカラーを適用
                    var secondaryColor = (Color)ColorConverter.ConvertFromString(settings.ThemeSecondaryColorCode ?? "#FCFCFC");
                    var secondaryBrush = new SolidColorBrush(secondaryColor);

                    // リソースを更新
                    if (mainDictionary.Contains("ApplicationBackgroundBrush"))
                    {
                        mainDictionary.Remove("ApplicationBackgroundBrush");
                    }
                    mainDictionary["ApplicationBackgroundBrush"] = mainBrush;

                    if (mainDictionary.Contains("TabAndNavigationBackgroundBrush"))
                    {
                        mainDictionary.Remove("TabAndNavigationBackgroundBrush");
                    }
                    mainDictionary["TabAndNavigationBackgroundBrush"] = secondaryBrush;

                    // アクセントカラー（タブとステータスバー用）を更新
                    if (mainDictionary.Contains("AccentFillColorDefaultBrush"))
                    {
                        mainDictionary.Remove("AccentFillColorDefaultBrush");
                    }
                    mainDictionary["AccentFillColorDefaultBrush"] = mainBrush;

                    // アクセントカラー（セカンダリ、ホバー時など）を更新
                    var accentSecondaryColor = Color.FromRgb(
                        (byte)Math.Max(0, mainColor.R - 20),
                        (byte)Math.Max(0, mainColor.G - 20),
                        (byte)Math.Max(0, mainColor.B - 20));
                    var accentSecondaryBrush = new SolidColorBrush(accentSecondaryColor);
                    if (mainDictionary.Contains("AccentFillColorSecondaryBrush"))
                    {
                        mainDictionary.Remove("AccentFillColorSecondaryBrush");
                    }
                    mainDictionary["AccentFillColorSecondaryBrush"] = accentSecondaryBrush;

                    // ステータスバーの文字色を背景色に応じて設定
                    var luminance = (0.299 * mainColor.R + 0.587 * mainColor.G + 0.114 * mainColor.B) / 255.0;
                    var statusBarTextColor = luminance > 0.5 ? Colors.Black : Colors.White;
                    var statusBarTextBrush = new SolidColorBrush(statusBarTextColor);
                    if (mainDictionary.Contains("StatusBarTextBrush"))
                    {
                        mainDictionary.Remove("StatusBarTextBrush");
                    }
                    mainDictionary["StatusBarTextBrush"] = statusBarTextBrush;

                    // ウィンドウが表示された後に背景色を更新
                    Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                    {
                        // すべてのウィンドウの背景色を更新
                        foreach (Window window in Application.Current.Windows)
                        {
                            if (window != null)
                            {
                                // ウィンドウの背景色を直接設定
                                window.Background = mainBrush;

                                // FluentWindowの場合は、Backgroundプロパティも更新
                                if (window is Wpf.Ui.Controls.FluentWindow fluentWindow)
                                {
                                    fluentWindow.Background = mainBrush;
                                }

                                // ウィンドウ内のNavigationViewの背景色も更新
                                var navigationView = FindVisualChild<Wpf.Ui.Controls.NavigationView>(window);
                                if (navigationView != null)
                                {
                                    navigationView.Background = secondaryBrush;
                                }

                                // ウィンドウのリソースを無効化
                                if (window is System.Windows.FrameworkElement fe)
                                {
                                    fe.InvalidateProperty(System.Windows.FrameworkElement.StyleProperty);
                                    fe.InvalidateProperty(System.Windows.Controls.Control.BackgroundProperty);
                                }

                                // ウィンドウのレイアウトを更新してDynamicResourceを再評価
                                window.UpdateLayout();
                                window.InvalidateVisual();
                            }
                        }
                    }), System.Windows.Threading.DispatcherPriority.Loaded);
                }
            }
            catch
            {
                // エラーハンドリング：デフォルトのテーマカラーを使用
            }
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

                    // 既存のダークテーマリソースを検索（参照比較とSource比較の両方を使用）
                    // LINQを避けて高速化（直接ループで検索）
                    ResourceDictionary? existingDarkTheme = null;
                    var darkThemeResources = app._darkThemeResources;
                    for (int i = 0; i < mergedDictionaries.Count; i++)
                    {
                        var dict = mergedDictionaries[i];
                        if (dict == darkThemeResources)
                        {
                            // 参照が一致する場合
                            existingDarkTheme = dict;
                            break;
                        }
                        var source = dict.Source?.OriginalString;
                        if (source != null && source.Contains("DarkThemeResources.xaml"))
                        {
                            // Sourceが一致する場合
                            existingDarkTheme = dict;
                            break;
                        }
                    }

                    var isDark = ApplicationThemeManager.GetAppTheme() == ApplicationTheme.Dark;

                    // リソースディクショナリーを更新（必要な場合のみ）
                    bool resourceChanged = false;
                    if (isDark)
                    {
                        // ダークモードの場合
                        if (existingDarkTheme == null)
                        {
                            // リソースが存在しない場合は追加
                            mergedDictionaries.Add(app._darkThemeResources);
                            resourceChanged = true;
                        }
                        else if (existingDarkTheme != app._darkThemeResources)
                        {
                            // 異なるインスタンスが存在する場合は置き換え
                            mergedDictionaries.Remove(existingDarkTheme);
                            mergedDictionaries.Add(app._darkThemeResources);
                            resourceChanged = true;
                        }
                    }
                    else
                    {
                        // ライトモードの場合
                        if (existingDarkTheme != null)
                        {
                            // リソースが存在する場合は削除
                            mergedDictionaries.Remove(existingDarkTheme);
                            resourceChanged = true;
                        }
                    }

                    // IconBrushリソースを直接更新してDynamicResourceの再評価を強制
                    // リソースを一度削除して再追加することで、DynamicResourceの再評価を確実にトリガー
                    try
                    {
                        var newColor = isDark ? Color.FromRgb(255, 255, 255) : Color.FromRgb(0, 0, 0);
                        
                        // 既存のIconBrushリソースを削除
                        if (mainDictionary.Contains("IconBrush"))
                        {
                            mainDictionary.Remove("IconBrush");
                        }
                        
                        // 新しいIconBrushリソースを作成して追加
                        var newBrush = new SolidColorBrush(newColor);
                        mainDictionary["IconBrush"] = newBrush;
                        resourceChanged = true;
                    }
                    catch
                    {
                        // エラーが発生した場合は無視
                    }

                    // リソースが変更された場合のみ、ウィンドウを更新
                    if (resourceChanged)
                    {
                        // リソースディクショナリーを更新した後、すべてのウィンドウのリソースを更新
                        // DynamicResourceの再評価を強制するため、ウィンドウを更新
                        // 起動時の高速化のため、Background優先度で遅延実行
                        var isStartup = Current.Windows.Count == 0; // ウィンドウが存在しない場合は起動時
                        Current.Dispatcher.BeginInvoke(new System.Action(() =>
                        {
                            // すべてのウィンドウのリソースを更新
                            foreach (System.Windows.Window window in Current.Windows)
                            {
                                if (window != null)
                                {
                                    // ウィンドウのビジュアルを無効化してDynamicResourceの再評価を強制
                                    window.InvalidateVisual();
                                    
                                    // 起動時以外の場合のみ、レイアウト更新と再帰的なリソース無効化（起動時の高速化）
                                    if (!isStartup)
                                    {
                                        window.UpdateLayout();
                                        InvalidateResourcesRecursive(window);
                                    }
                                }
                            }
                        }), System.Windows.Threading.DispatcherPriority.Background);
                    }
                }
            }

            // すべてのThemedSvgIconインスタンスにブラシを再適用（遅延実行でちらつきを防ぐ）
            // 起動時の高速化のため、常にBackground優先度で実行（起動を最速化）
            Current.Dispatcher.BeginInvoke(new System.Action(() =>
            {
                FastExplorer.Controls.ThemedSvgIcon.RefreshAllInstances();
            }), System.Windows.Threading.DispatcherPriority.Background);
        }


        /// <summary>
        /// 再帰的に要素内のすべてのDependencyObjectのリソースを無効化します
        /// </summary>
        /// <param name="element">開始要素</param>
        private static void InvalidateResourcesRecursive(System.Windows.DependencyObject element)
        {
            if (element == null)
                return;

            // TextBlockの場合、Foregroundプロパティを無効化
            if (element is System.Windows.Controls.TextBlock textBlock)
            {
                // DynamicResourceの再評価を強制
                textBlock.InvalidateProperty(System.Windows.Controls.TextBlock.ForegroundProperty);
            }
            // Controlの場合、Foregroundプロパティを無効化
            else if (element is System.Windows.Controls.Control control)
            {
                control.InvalidateProperty(System.Windows.Controls.Control.ForegroundProperty);
            }

            // FrameworkElementの場合、Styleプロパティも無効化
            if (element is System.Windows.FrameworkElement frameworkElement)
            {
                frameworkElement.InvalidateProperty(System.Windows.FrameworkElement.StyleProperty);
            }

            // 子要素を再帰的に処理
            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(element); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(element, i);
                InvalidateResourcesRecursive(child);
            }
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
                    settings.Theme = currentTheme switch
                    {
                        ApplicationTheme.Light => "Light",
                        ApplicationTheme.Dark => "Dark",
                        _ => "System" // ApplicationTheme.Unknownの場合は"System"として保存
                    };
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
