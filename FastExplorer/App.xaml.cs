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
        private ResourceDictionary? _darkThemeResources = null;

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
            //_darkThemeResources = new ResourceDictionary
            //{
            //    Source = new Uri("pack://application:,,,/Resources/DarkThemeResources.xaml", UriKind.Absolute)
            //};
            //LoadAndApplyThemeOnStartup();
            // リソース更新は既にLoadAndApplyThemeOnStartup内で実行されているため、ここでは実行しない

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

                // 保存されたテーマカラーを先に適用（リソース更新の前に）
                var themeColorCode = settings?.ThemeColorCode;
                if (themeColorCode != null && themeColorCode.Length > 0)
                {
                    ApplyThemeColorOnStartup(settings!);
                }

                // リソースを即座に更新（ウィンドウ表示前に確実に適用するため）
                UpdateThemeResourcesInternal();
            }
            catch
            {
                // エラーハンドリング：デフォルトのテーマ（システムテーマ）を使用
                ApplicationThemeManager.Apply(ApplicationTheme.Unknown);
                // エラー時もリソースを更新
                UpdateThemeResourcesInternal();
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
        /// <param name="precomputedColors">事前計算済みの色（nullの場合は計算する）</param>
        public static void ApplyThemeColorFromSettings(Services.WindowSettings settings, (Color mainColor, Color secondaryColor)? precomputedColors = null)
        {
            try
            {
                if (Application.Current.Resources is ResourceDictionary mainDictionary)
                {
                    Color mainColor;
                    Color secondaryColor;
                    
                    // 事前計算済みの色を使用するか、計算する
                    if (precomputedColors.HasValue)
                    {
                        mainColor = precomputedColors.Value.mainColor;
                        secondaryColor = precomputedColors.Value.secondaryColor;
                    }
                    else
                    {
                        // 高速なカスタム変換を使用
                        mainColor = Helpers.FastColorConverter.ParseHexColor(settings.ThemeColorCode ?? "#F5F5F5");
                        secondaryColor = Helpers.FastColorConverter.ParseHexColor(settings.ThemeSecondaryColorCode ?? "#FCFCFC");
                    }
                    
                    var mainBrush = new SolidColorBrush(mainColor);
                    var secondaryBrush = new SolidColorBrush(secondaryColor);

                    // リソースを更新（高速化：Contains/Removeを削除して直接インデクサーで上書き）
                    mainDictionary["ApplicationBackgroundBrush"] = mainBrush;
                    mainDictionary["TabAndNavigationBackgroundBrush"] = secondaryBrush;

                    // アクセントカラー（タブとステータスバー用）を更新
                    mainDictionary["AccentFillColorDefaultBrush"] = mainBrush;

                    // アクセントカラー（セカンダリ、ホバー時など）を更新（計算を最適化）
                    var r1 = mainColor.R;
                    var g1 = mainColor.G;
                    var b1 = mainColor.B;
                    var accentSecondaryColor = Color.FromRgb(
                        (byte)(r1 > 20 ? r1 - 20 : 0),
                        (byte)(g1 > 20 ? g1 - 20 : 0),
                        (byte)(b1 > 20 ? b1 - 20 : 0));
                    var accentSecondaryBrush = new SolidColorBrush(accentSecondaryColor);
                    mainDictionary["AccentFillColorSecondaryBrush"] = accentSecondaryBrush;

                    // コントロールの背景色（タブの非選択時など）を更新
                    mainDictionary["ControlFillColorDefaultBrush"] = secondaryBrush;

                    // コントロールの背景色（セカンダリ、ホバー時など）を更新（計算を最適化）
                    var r2 = secondaryColor.R;
                    var g2 = secondaryColor.G;
                    var b2 = secondaryColor.B;
                    var controlSecondaryColor = Color.FromRgb(
                        (byte)(r2 < 245 ? r2 + 10 : 255),
                        (byte)(g2 < 245 ? g2 + 10 : 255),
                        (byte)(b2 < 245 ? b2 + 10 : 255));
                    var controlSecondaryBrush = new SolidColorBrush(controlSecondaryColor);
                    mainDictionary["ControlFillColorSecondaryBrush"] = controlSecondaryBrush;

                    // ステータスバーの文字色を背景色に応じて設定（輝度計算を最適化）
                    var luminance = (0.299 * mainColor.R + 0.587 * mainColor.G + 0.114 * mainColor.B) * 0.00392156862745098; // 1/255を事前計算
                    var statusBarTextColor = luminance > 0.5 ? Colors.Black : Colors.White;
                    var statusBarTextBrush = new SolidColorBrush(statusBarTextColor);
                    mainDictionary["StatusBarTextBrush"] = statusBarTextBrush;

                    // 起動時はウィンドウがまだ作成されていない可能性があるため、
                    // ウィンドウの背景色更新はMainWindowのコンストラクタとLoadedイベントで行う
                    // ここではリソースのみ更新（ちらつきを防ぐため、ウィンドウ更新処理は削除）
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
                    // 高速化：Contains/Removeを削除して直接インデクサーで上書き
                    try
                    {
                        var newColor = isDark ? Color.FromRgb(255, 255, 255) : Color.FromRgb(0, 0, 0);
                        // 新しいIconBrushリソースを作成して追加（直接インデクサーで上書き）
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
                        var isStartup = Current.Windows.Count == 0; // ウィンドウが存在しない場合は起動時
                        
                        if (isStartup)
                        {
                            // 起動時は、ウィンドウが表示された後にリソースを再評価する
                            // MainWindowのLoadedイベントで処理されるため、ここでは何もしない
                        }
                        else
                        {
                            // 起動時以外は、即座にリソースを更新
                            Current.Dispatcher.BeginInvoke(new System.Action(() =>
                            {
                                // すべてのウィンドウのリソースを更新
                                foreach (System.Windows.Window window in Current.Windows)
                                {
                                    if (window != null)
                                    {
                                        // ウィンドウのビジュアルを無効化してDynamicResourceの再評価を強制
                                        window.InvalidateVisual();
                                        window.UpdateLayout();
                                        InvalidateResourcesRecursive(window);
                                    }
                                }
                            }), System.Windows.Threading.DispatcherPriority.Background);
                        }
                    }
                }
            }

            // すべてのThemedSvgIconインスタンスにブラシを再適用（遅延実行でちらつきを防ぐ）
            // 起動時の高速化のため、常にBackground優先度で実行（起動を最速化）
            // デリゲートのメモリアロケーションを削減
            Current.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new System.Action(() =>
            {
                FastExplorer.Controls.ThemedSvgIcon.RefreshAllInstances();
            }));
        }


        /// <summary>
        /// 再帰的に要素内のすべてのDependencyObjectのリソースを無効化します
        /// </summary>
        /// <param name="element">開始要素</param>
        public static void InvalidateResourcesRecursive(System.Windows.DependencyObject element)
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
