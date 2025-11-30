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
using Wpf.Ui.Controls;
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

        // コマンドライン引数で指定されたパスを保存
        private static string? _startupPath = null;
        // 単一タブモードで起動するかどうか（タブを移動する場合）
        private static bool _isSingleTabMode = false;
        // ウィンドウの位置（ドロップ位置）
        private static System.Windows.Point? _windowPosition = null;

        /// <summary>
        /// アプリケーションが読み込まれるときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">スタートアップイベント引数</param>
        private async void OnStartup(object sender, StartupEventArgs e)
        {
            // コマンドライン引数を処理
            if (e.Args != null && e.Args.Length > 0)
            {
                var firstArg = e.Args[0];
                if (!string.IsNullOrEmpty(firstArg))
                {
                    // "--single-tab"フラグをチェック
                    if (firstArg == "--single-tab")
                    {
                        _isSingleTabMode = true;
                        
                        // 引数を順番に処理
                        for (int i = 1; i < e.Args.Length; i++)
                        {
                            var arg = e.Args[i];
                            if (string.IsNullOrEmpty(arg))
                                continue;
                            
                            // "--position"フラグをチェック
                            if (arg == "--position" && i + 1 < e.Args.Length)
                            {
                                var positionArg = e.Args[i + 1];
                                if (!string.IsNullOrEmpty(positionArg))
                                {
                                    // 位置を解析（例: "100,200"）
                                    var parts = positionArg.Split(',');
                                    if (parts.Length == 2 &&
                                        double.TryParse(parts[0], out double x) &&
                                        double.TryParse(parts[1], out double y))
                                    {
                                        _windowPosition = new System.Windows.Point(x, y);
                                    }
                                }
                                i++; // 次の引数も処理済み
                            }
                            else if (string.IsNullOrEmpty(_startupPath))
                            {
                                // パス引数（最初の非フラグ引数）
                                arg = arg.Trim('"');
                                if (!string.IsNullOrEmpty(arg))
                                {
                                    _startupPath = arg;
                                }
                            }
                        }
                    }
                    else
                    {
                        // 引用符を削除
                        firstArg = firstArg.Trim('"');
                        if (!string.IsNullOrEmpty(firstArg))
                        {
                            _startupPath = firstArg;
                        }
                    }
                }
            }

            // テーマを先に適用（起動時の高速化のため、同期的に実行）
            _darkThemeResources = new ResourceDictionary
            {
                Source = new Uri("pack://application:,,,/Resources/DarkThemeResources.xaml", UriKind.Absolute)
            };
            LoadAndApplyThemeOnStartup();
            // リソース更新は既にLoadAndApplyThemeOnStartup内で実行されているため、ここでは実行しない

            // ウィンドウを表示
            await _host.StartAsync();
        }

        /// <summary>
        /// 起動時に指定されたパスを取得します
        /// </summary>
        /// <returns>起動時に指定されたパス、指定されていない場合はnull</returns>
        public static string? GetStartupPath()
        {
            return _startupPath;
        }

        /// <summary>
        /// 単一タブモードで起動するかどうかを取得します
        /// </summary>
        /// <returns>単一タブモードの場合はtrue、それ以外の場合はfalse</returns>
        public static bool IsSingleTabMode()
        {
            return _isSingleTabMode;
        }

        /// <summary>
        /// ウィンドウの位置を取得します
        /// </summary>
        /// <returns>ウィンドウの位置（指定されていない場合はnull）</returns>
        public static System.Windows.Point? GetWindowPosition()
        {
            return _windowPosition;
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
                // ライトモードの場合のみテーマカラーを適用
                if (themeToApply == ApplicationTheme.Light)
                {
                    var themeColorCode = settings?.ThemeColorCode;
                    if (themeColorCode != null && themeColorCode.Length > 0)
                    {
                        ApplyThemeColorOnStartup(settings!);
                    }
                }
                else if (themeToApply == ApplicationTheme.Dark)
                {
                    // ダークモードの場合は、デフォルトのダークテーマカラーにリセット
                    ResetToDefaultDarkThemeColors();
                }
                else
                {
                    // システムテーマ（Unknown）の場合は、システムのテーマに応じて処理
                    var systemTheme = ApplicationThemeManager.GetSystemTheme();
                    if (systemTheme == SystemTheme.Light)
                    {
                        var themeColorCode = settings?.ThemeColorCode;
                        if (themeColorCode != null && themeColorCode.Length > 0)
                        {
                            ApplyThemeColorOnStartup(settings!);
                        }
                    }
                    else
                    {
                        // システムがダークモードの場合は、デフォルトのダークテーマカラーにリセット
                        ResetToDefaultDarkThemeColors();
                    }
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
            // 色計算を1回だけ実行（高速なカスタム変換を使用）
            var mainColor = Helpers.FastColorConverter.ParseHexColor(settings.ThemeColorCode ?? "#F5F5F5");
            var secondaryColor = Helpers.FastColorConverter.ParseHexColor(settings.ThemeSecondaryColorCode ?? "#FCFCFC");
            var mainBrush = new System.Windows.Media.SolidColorBrush(mainColor);
            var secondaryBrush = new System.Windows.Media.SolidColorBrush(secondaryColor);
            
            // リソースを更新（計算済みの色を渡して重複計算を回避）
            ApplyThemeColorFromSettings(settings, (mainColor, secondaryColor));
            
            // カラーを選択したときと同じ挙動でテーマを復元
            // すべてのウィンドウの背景色を直接更新（ウィンドウが作成された後に実行）
            Current.Dispatcher.BeginInvoke(new System.Action(() =>
            {
                // すべてのウィンドウの背景色を更新
                foreach (Window window in Current.Windows)
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

                        // タブとListViewの選択中の色を更新するため、スタイルを無効化してDynamicResourceの再評価を強制
                        Views.Windows.MainWindow.InvalidateTabAndListViewStyles(window);
                    }
                }
            }), System.Windows.Threading.DispatcherPriority.Loaded);
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
                // 現在のテーマがライトモードでない場合は、テーマカラーを適用しない
                var currentTheme = ApplicationThemeManager.GetAppTheme();
                if (currentTheme == ApplicationTheme.Dark)
                {
                    // ダークモードの場合は、デフォルトのダークテーマカラーにリセット
                    ResetToDefaultDarkThemeColors();
                    return;
                }
                else if (currentTheme == ApplicationTheme.Unknown)
                {
                    // システムテーマ（Unknown）の場合は、システムのテーマを確認
                    var systemTheme = ApplicationThemeManager.GetSystemTheme();
                    if (systemTheme != SystemTheme.Light)
                    {
                        // システムがダークモードの場合は、デフォルトのダークテーマカラーにリセット
                        ResetToDefaultDarkThemeColors();
                        return;
                    }
                }

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

                    // ステータスバーの背景色を更新（ペインの一部として表示、ControlFillColorDefaultBrushと同じ色）
                    mainDictionary["StatusBarBackgroundBrush"] = secondaryBrush;

                    // ステータスバーの文字色を背景色に応じて設定（輝度計算を最適化）
                    var luminance = (0.299 * mainColor.R + 0.587 * mainColor.G + 0.114 * mainColor.B) * 0.00392156862745098; // 1/255を事前計算
                    var statusBarTextColor = luminance > 0.5 ? Colors.Black : Colors.White;
                    var statusBarTextBrush = new SolidColorBrush(statusBarTextColor);
                    mainDictionary["StatusBarTextBrush"] = statusBarTextBrush;

                    // 分割ペイン用の背景色を更新（選択されたテーマカラーのThirdColorCodeを使用）
                    Color unfocusedPaneColor;
                    if (!string.IsNullOrEmpty(settings.ThemeThirdColorCode))
                    {
                        // ThirdColorCodeが設定されている場合はそれを使用
                        unfocusedPaneColor = Helpers.FastColorConverter.ParseHexColor(settings.ThemeThirdColorCode);
                    }
                    else
                    {
                        // フォールバック：ThemeColorの計算メソッドを使用して薄い色を生成
                        unfocusedPaneColor = Models.ThemeColor.CalculateLightColor(mainColor);
                    }
                    var unfocusedPaneBrush = new SolidColorBrush(unfocusedPaneColor);
                    mainDictionary["UnfocusedPaneBackgroundBrush"] = unfocusedPaneBrush;

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
        /// デフォルトのダークテーマカラーにリセットします
        /// </summary>
        public static void ResetToDefaultDarkThemeColors()
        {
            try
            {
                if (Application.Current.Resources is ResourceDictionary mainDictionary)
                {
                    // WPF-UIのデフォルトのダークテーマカラーを使用
                    // 一般的なダークテーマの背景色
                    var darkMainColor = Color.FromRgb(0x1E, 0x1E, 0x1E); // #1E1E1E
                    var darkSecondaryColor = Color.FromRgb(0x25, 0x25, 0x26); // #252526
                    var darkMainBrush = new SolidColorBrush(darkMainColor);
                    var darkSecondaryBrush = new SolidColorBrush(darkSecondaryColor);

                    // リソースを更新
                    mainDictionary["ApplicationBackgroundBrush"] = darkMainBrush;
                    mainDictionary["TabAndNavigationBackgroundBrush"] = darkSecondaryBrush;

                    // アクセントカラー（タブとステータスバー用）を更新
                    mainDictionary["AccentFillColorDefaultBrush"] = darkMainBrush;

                    // アクセントカラー（セカンダリ、ホバー時など）を更新
                    var accentSecondaryColor = Color.FromRgb(
                        (byte)(darkMainColor.R > 20 ? darkMainColor.R - 20 : 0),
                        (byte)(darkMainColor.G > 20 ? darkMainColor.G - 20 : 0),
                        (byte)(darkMainColor.B > 20 ? darkMainColor.B - 20 : 0));
                    var accentSecondaryBrush = new SolidColorBrush(accentSecondaryColor);
                    mainDictionary["AccentFillColorSecondaryBrush"] = accentSecondaryBrush;

                    // コントロールの背景色（タブの非選択時など）を更新
                    mainDictionary["ControlFillColorDefaultBrush"] = darkSecondaryBrush;

                    // コントロールの背景色（セカンダリ、ホバー時など）を更新
                    var controlSecondaryColor = Color.FromRgb(
                        (byte)(darkSecondaryColor.R < 245 ? darkSecondaryColor.R + 10 : 255),
                        (byte)(darkSecondaryColor.G < 245 ? darkSecondaryColor.G + 10 : 255),
                        (byte)(darkSecondaryColor.B < 245 ? darkSecondaryColor.B + 10 : 255));
                    var controlSecondaryBrush = new SolidColorBrush(controlSecondaryColor);
                    mainDictionary["ControlFillColorSecondaryBrush"] = controlSecondaryBrush;

                    // ステータスバーの背景色を更新
                    mainDictionary["StatusBarBackgroundBrush"] = darkSecondaryBrush;

                    // ステータスバーの文字色（ダークモードでは白）
                    var statusBarTextBrush = new SolidColorBrush(Colors.White);
                    mainDictionary["StatusBarTextBrush"] = statusBarTextBrush;

                    // 分割ペイン用の背景色を更新
                    var unfocusedPaneColor = Color.FromRgb(0x2D, 0x2D, 0x30); // #2D2D30
                    var unfocusedPaneBrush = new SolidColorBrush(unfocusedPaneColor);
                    mainDictionary["UnfocusedPaneBackgroundBrush"] = unfocusedPaneBrush;

                    // テキスト色を白に設定（ダークモード）
                    var textPrimaryBrush = new SolidColorBrush(Colors.White);
                    var textSecondaryBrush = new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC)); // #CCCCCC
                    mainDictionary["TextFillColorPrimaryBrush"] = textPrimaryBrush;
                    mainDictionary["TextFillColorSecondaryBrush"] = textSecondaryBrush;
                }
            }
            catch
            {
                // エラーハンドリング：デフォルトのダークテーマカラーの適用に失敗した場合は無視
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
            if (Current is App app)
            {
                // _darkThemeResourcesが初期化されていない場合は初期化
                if (app._darkThemeResources == null)
                {
                    app._darkThemeResources = new ResourceDictionary
                    {
                        Source = new Uri("pack://application:,,,/Resources/DarkThemeResources.xaml", UriKind.Absolute)
                    };
                }

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
                        // ダークモードの場合：既存のダークテーマリソースを削除してから追加することで、確実に優先されるようにする
                        if (existingDarkTheme != null)
                        {
                            // 既存のリソースを削除
                            mergedDictionaries.Remove(existingDarkTheme);
                        }
                        // ダークテーマのリソースを最後に追加（後から追加されたリソースが優先される）
                        mergedDictionaries.Add(app._darkThemeResources);
                        resourceChanged = true;
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

                    // テキスト色を更新（ダークモード時は白に設定）
                    try
                    {
                        if (isDark)
                        {
                            // ダークモード時はテキスト色を白に設定
                            var textPrimaryBrush = new SolidColorBrush(Colors.White);
                            var textSecondaryBrush = new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC)); // #CCCCCC
                            mainDictionary["TextFillColorPrimaryBrush"] = textPrimaryBrush;
                            mainDictionary["TextFillColorSecondaryBrush"] = textSecondaryBrush;
                            resourceChanged = true;
                        }
                        else
                        {
                            // ライトモード時はテキスト色を黒に設定
                            var textPrimaryBrush = new SolidColorBrush(Colors.Black);
                            var textSecondaryBrush = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66)); // #666666
                            mainDictionary["TextFillColorPrimaryBrush"] = textPrimaryBrush;
                            mainDictionary["TextFillColorSecondaryBrush"] = textSecondaryBrush;
                            resourceChanged = true;
                        }
                    }
                    catch
                    {
                        // エラーが発生した場合は無視
                    }

                    // 分割ペイン用の背景色を更新（テーマ切り替え時、現在のテーマカラーに基づいて計算）
                    try
                    {
                        if (isDark)
                        {
                            // ダークモードの場合は、ダークモード用の色を優先
                            var unfocusedPaneColor = Color.FromRgb(0x2D, 0x2D, 0x30); // #2D2D30
                            var unfocusedPaneBrush = new SolidColorBrush(unfocusedPaneColor);
                            mainDictionary["UnfocusedPaneBackgroundBrush"] = unfocusedPaneBrush;
                            resourceChanged = true;
                        }
                        else
                        {
                            // ライトモードの場合は、現在のテーマカラーを取得（保存されたThirdColorCodeを使用）
                            try
                            {
                                var windowSettingsService = Services.GetService(typeof(WindowSettingsService)) as WindowSettingsService;
                                if (windowSettingsService != null)
                                {
                                    var currentSettings = windowSettingsService.GetSettings();
                                    if (!string.IsNullOrEmpty(currentSettings.ThemeThirdColorCode))
                                    {
                                        // ThirdColorCodeが設定されている場合はそれを使用
                                        var unfocusedPaneColor = Helpers.FastColorConverter.ParseHexColor(currentSettings.ThemeThirdColorCode);
                                        var unfocusedPaneBrush = new SolidColorBrush(unfocusedPaneColor);
                                        mainDictionary["UnfocusedPaneBackgroundBrush"] = unfocusedPaneBrush;
                                        resourceChanged = true;
                                    }
                                    else
                                    {
                                        // フォールバック：ThemeColorの計算メソッドを使用して現在のテーマカラーから計算
                                        var currentMainBrush = mainDictionary["ApplicationBackgroundBrush"] as SolidColorBrush;
                                        if (currentMainBrush != null)
                                        {
                                            var mainColor = currentMainBrush.Color;
                                            var unfocusedPaneColor = Models.ThemeColor.CalculateLightColor(mainColor);
                                            var unfocusedPaneBrush = new SolidColorBrush(unfocusedPaneColor);
                                            mainDictionary["UnfocusedPaneBackgroundBrush"] = unfocusedPaneBrush;
                                            resourceChanged = true;
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // エラーが発生した場合は無視
                            }
                        }
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
            // GetChildrenCount()を一度だけ呼び出してキャッシュ（パフォーマンス向上）
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(element);
            for (int i = 0; i < childrenCount; i++)
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
