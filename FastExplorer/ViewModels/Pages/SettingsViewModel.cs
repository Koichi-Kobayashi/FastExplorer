using FastExplorer.Controls;
using FastExplorer.Models;
using FastExplorer.Services;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using Wpf.Ui.Abstractions.Controls;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

namespace FastExplorer.ViewModels.Pages
{
    /// <summary>
    /// 設定ページのViewModel
    /// </summary>
    public partial class SettingsViewModel : ObservableObject, INavigationAware
    {
        private bool _isInitialized = false;
        private readonly WindowSettingsService _windowSettingsService;

        /// <summary>
        /// <see cref="SettingsViewModel"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="windowSettingsService">ウィンドウ設定サービス</param>
        public SettingsViewModel(WindowSettingsService windowSettingsService)
        {
            _windowSettingsService = windowSettingsService;
        }

        /// <summary>
        /// アプリケーションのバージョンを取得または設定します
        /// </summary>
        [ObservableProperty]
        private string _appVersion = String.Empty;

        /// <summary>
        /// 現在のテーマを取得または設定します
        /// </summary>
        [ObservableProperty]
        private ApplicationTheme _currentTheme = ApplicationTheme.Unknown;

        /// <summary>
        /// テーマカラーのコレクションを取得または設定します
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<ThemeColor> _themeColors = new();

        /// <summary>
        /// ページにナビゲートされたときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedToAsync()
        {
            if (!_isInitialized)
            {
                InitializeViewModel();
            }
            else
            {
                // 既に初期化済みの場合でも、現在のテーマを更新
                CurrentTheme = ApplicationThemeManager.GetAppTheme();
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// ページから離れるときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedFromAsync() => Task.CompletedTask;

        /// <summary>
        /// ViewModelを初期化します
        /// </summary>
        private void InitializeViewModel()
        {
            CurrentTheme = ApplicationThemeManager.GetAppTheme();
            AppVersion = $"UiDesktopApp1 - {GetAssemblyVersion()}";

            // テーマカラーリストを初期化
            InitializeThemeColors();

            _isInitialized = true;
        }

        /// <summary>
        /// テーマカラーリストを初期化します
        /// </summary>
        private void InitializeThemeColors()
        {
            ThemeColors.Clear();

            var colors = new[]
            {
                new ThemeColor { Name = "既定", NameEn = "Default", ColorCode = "#F5F5F5", SecondaryColorCode = "#FCFCFC", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F5F5F5")) },
                new ThemeColor { Name = "イエロー ゴールド", NameEn = "Yellow Gold", ColorCode = "#FCEECA", SecondaryColorCode = "#FEFAEF", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FCEECA")) },
                new ThemeColor { Name = "オレンジ ブライト", NameEn = "Orange Bright", ColorCode = "#FADDCC", SecondaryColorCode = "#FEF5F0", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FADDCC")) },
                new ThemeColor { Name = "ブリック レッド", NameEn = "Brick Red", ColorCode = "#F3D4D5", SecondaryColorCode = "#FBF2F2", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F3D4D5")) },
                new ThemeColor { Name = "モダン レッド", NameEn = "Modern Red", ColorCode = "#FCD7D7", SecondaryColorCode = "#FEF3F3", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FCD7D7")) },
                new ThemeColor { Name = "レッド", NameEn = "Red", ColorCode = "#F8CADC", SecondaryColorCode = "#FDEFF5", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8CADC")) },
                new ThemeColor { Name = "ローズ ブライト", NameEn = "Rose Bright", ColorCode = "#F8CADC", SecondaryColorCode = "#FDEFF5", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8CADC")) },
                new ThemeColor { Name = "ブルー", NameEn = "Blue", ColorCode = "#CAE2F4", SecondaryColorCode = "#EFF6FC", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAE2F4")) },
                new ThemeColor { Name = "アイリス パステル", NameEn = "Iris Pastel", ColorCode = "#E4DEEE", SecondaryColorCode = "#F7F5FA", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E4DEEE")) },
                new ThemeColor { Name = "バイオレット レッド ライト", NameEn = "Violet Red Light", ColorCode = "#EDD8F0", SecondaryColorCode = "#FAF3FB", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EDD8F0")) },
                new ThemeColor { Name = "クール ブルー ブライト", NameEn = "Cool Blue Bright", ColorCode = "#CAE8EF", SecondaryColorCode = "#EFF8FA", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAE8EF")) },
                new ThemeColor { Name = "シーフォーム", NameEn = "Seafoam", ColorCode = "#CAEEF0", SecondaryColorCode = "#EFFAFB", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAEEF0")) },
                new ThemeColor { Name = "ミント ライト", NameEn = "Mint Light", ColorCode = "#CAEDE7", SecondaryColorCode = "#EFFAF8", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAEDE7")) },
                new ThemeColor { Name = "グレー", NameEn = "Gray", ColorCode = "#E2E1E1", SecondaryColorCode = "#F6F6F6", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E1E1")) },
                new ThemeColor { Name = "グリーン", NameEn = "Green", ColorCode = "#CDE2CD", SecondaryColorCode = "#F0F6F0", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CDE2CD")) },
                new ThemeColor { Name = "オーバーキャスト", NameEn = "Overcast", ColorCode = "#E1E1E1", SecondaryColorCode = "#F6F6F6", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E1E1E1")) },
                new ThemeColor { Name = "ストーム", NameEn = "Storm", ColorCode = "#D9D9D8", SecondaryColorCode = "#F4F4F3", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D9D9D8")) },
                new ThemeColor { Name = "ブルー グレー", NameEn = "Blue Gray", ColorCode = "#DFE2E3", SecondaryColorCode = "#F5F6F7", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DFE2E3")) },
                new ThemeColor { Name = "グレー ダーク", NameEn = "Gray Dark", ColorCode = "#D9DADB", SecondaryColorCode = "#F4F4F4", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D9DADB")) },
                new ThemeColor { Name = "カモフラージュ", NameEn = "Camouflage", ColorCode = "#E3E1DD", SecondaryColorCode = "#F7F6F5", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E3E1DD")) }
            };

            foreach (var color in colors)
            {
                ThemeColors.Add(color);
            }
        }

        /// <summary>
        /// アセンブリのバージョンを取得します
        /// </summary>
        /// <returns>アセンブリのバージョン文字列</returns>
        private string GetAssemblyVersion()
        {
            return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString()
                ?? String.Empty;
        }

        /// <summary>
        /// テーマを変更します
        /// </summary>
        /// <param name="parameter">テーマパラメータ（"theme_light"またはその他）</param>
        [RelayCommand]
        private void OnChangeTheme(string parameter)
        {
            ApplicationTheme newTheme;
            switch (parameter)
            {
                case "theme_light":
                    if (CurrentTheme == ApplicationTheme.Light)
                        break;

                    newTheme = ApplicationTheme.Light;
                    ApplyTheme(newTheme);
                    break;

                default:
                    if (CurrentTheme == ApplicationTheme.Dark)
                        break;

                    newTheme = ApplicationTheme.Dark;
                    ApplyTheme(newTheme);
                    break;
            }
        }
        
        /// <summary>
        /// テーマを適用します（起動時と同じロジック）
        /// </summary>
        private void ApplyTheme(ApplicationTheme theme)
        {
            try
            {
                // ThemesDictionaryのThemeプロパティを先に更新（ApplicationThemeManager.Apply()の前に）
                if (System.Windows.Application.Current.Resources is System.Windows.ResourceDictionary mainDictionary)
                {
                    var mergedDictionaries = mainDictionary.MergedDictionaries;
                    var themesDict = mergedDictionaries
                        .OfType<System.Windows.ResourceDictionary>()
                        .FirstOrDefault(rd => rd.GetType().Name == "ThemesDictionary");
                    
                    if (themesDict != null)
                    {
                        var themeProperty = themesDict.GetType().GetProperty("Theme");
                        if (themeProperty != null)
                        {
                            var themeType = themeProperty.PropertyType;
                            // ApplicationTheme.Unknownの場合は、システムテーマを取得
                            if (theme == ApplicationTheme.Unknown)
                            {
                                var systemTheme = ApplicationThemeManager.GetSystemTheme();
                                var themeValue = Enum.Parse(themeType, systemTheme == SystemTheme.Dark ? "Dark" : "Light");
                                themeProperty.SetValue(themesDict, themeValue);
                            }
                            else
                            {
                                var themeValue = Enum.Parse(themeType, theme == ApplicationTheme.Dark ? "Dark" : "Light");
                                themeProperty.SetValue(themesDict, themeValue);
                            }
                        }
                    }
                }

                // テーマを適用
                ApplicationThemeManager.Apply(theme);
                CurrentTheme = theme;

                // テーマ設定を保存
                SaveTheme(theme);
                
                // リソースディクショナリーも即座に更新（起動時と同じ）
                App.UpdateThemeResourcesInternal();
                
                // すべてのThemedSvgIconインスタンスを即座に更新（リアルタイム反映のため）
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    ThemedSvgIcon.RefreshAllInstances();
                }), System.Windows.Threading.DispatcherPriority.Render);
            }
            catch
            {
                // エラーハンドリング：デフォルトのテーマ（システムテーマ）を使用
                ApplicationThemeManager.Apply(ApplicationTheme.Unknown);
            }
        }

        /// <summary>
        /// テーマ設定を保存します
        /// </summary>
        /// <param name="theme">保存するテーマ</param>
        private void SaveTheme(ApplicationTheme theme)
        {
            var settings = _windowSettingsService.GetSettings();
            settings.Theme = theme switch
            {
                ApplicationTheme.Light => "Light",
                ApplicationTheme.Dark => "Dark",
                _ => "System" // ApplicationTheme.Unknownの場合は"System"として保存
            };
            _windowSettingsService.SaveSettings(settings);
        }

        /// <summary>
        /// テーマカラーを選択します
        /// </summary>
        /// <param name="themeColor">選択されたテーマカラー</param>
        [RelayCommand]
        private void SelectThemeColor(ThemeColor? themeColor)
        {
            if (themeColor == null)
                return;

            try
            {
                // テーマカラーをリソースに適用
                if (Application.Current.Resources is ResourceDictionary mainDictionary)
                {
                    // メインカラーを適用
                    var mainColor = (Color)ColorConverter.ConvertFromString(themeColor.ColorCode);
                    var mainBrush = new SolidColorBrush(mainColor);
                    
                    // セカンダリカラーを適用
                    var secondaryColor = (Color)ColorConverter.ConvertFromString(themeColor.SecondaryColorCode);
                    var secondaryBrush = new SolidColorBrush(secondaryColor);

                    // メインのリソースディクショナリーに直接追加（MergedDictionariesより優先される）
                    // リソースを一度削除して再追加することで、DynamicResourceの再評価を強制
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
                    // メインカラーを少し濃くした色を使用
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
                    // 背景が明るい場合は黒、暗い場合は白
                    var luminance = (0.299 * mainColor.R + 0.587 * mainColor.G + 0.114 * mainColor.B) / 255.0;
                    var statusBarTextColor = luminance > 0.5 ? Colors.Black : Colors.White;
                    var statusBarTextBrush = new SolidColorBrush(statusBarTextColor);
                    if (mainDictionary.Contains("StatusBarTextBrush"))
                    {
                        mainDictionary.Remove("StatusBarTextBrush");
                    }
                    mainDictionary["StatusBarTextBrush"] = statusBarTextBrush;

                    // テーマカラーを保存
                    var settings = _windowSettingsService.GetSettings();
                    settings.ThemeColorName = themeColor.Name;
                    settings.ThemeColorCode = themeColor.ColorCode;
                    settings.ThemeSecondaryColorCode = themeColor.SecondaryColorCode;
                    _windowSettingsService.SaveSettings(settings);

                    System.Diagnostics.Debug.WriteLine($"Theme color applied: {themeColor.Name} - Main: {themeColor.ColorCode}, Secondary: {themeColor.SecondaryColorCode}");

                    // すべてのウィンドウの背景色を直接更新
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

                                // ステータスバーの背景色とテキスト色を直接更新
                                // StatusBarはBorderで、x:Name="StatusBar"が設定されている
                                var statusBar = FindVisualChildByName<System.Windows.Controls.Border>(window, "StatusBar");
                                if (statusBar != null)
                                {
                                    statusBar.Background = mainBrush;
                                }
                                
                                // StatusBarTextはTextBlockで、x:Name="StatusBarText"が設定されている
                                var statusBarText = FindVisualChildByName<System.Windows.Controls.TextBlock>(window, "StatusBarText");
                                if (statusBarText != null)
                                {
                                    statusBarText.Foreground = statusBarTextBrush;
                                }

                                // ウィンドウのリソースを無効化
                                if (window is System.Windows.FrameworkElement fe)
                                {
                                    fe.InvalidateProperty(System.Windows.FrameworkElement.StyleProperty);
                                    fe.InvalidateProperty(System.Windows.Controls.Control.BackgroundProperty);
                                }

                                // ウィンドウのレイアウトを更新してDynamicResourceを再評価
                                window.UpdateLayout();
                                // ビジュアルを無効化してDynamicResourceの再評価を強制
                                window.InvalidateVisual();
                                
                                // ウィンドウ内のすべての要素のDynamicResourceを再評価
                                InvalidateResourcesRecursive(window);
                            }
                        }
                    }), System.Windows.Threading.DispatcherPriority.Render);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SelectThemeColor failed: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
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
        /// ビジュアルツリー内の指定された名前と型の子要素を検索します
        /// </summary>
        /// <typeparam name="T">検索する型</typeparam>
        /// <param name="parent">親要素</param>
        /// <param name="name">要素の名前</param>
        /// <returns>見つかった要素、見つからない場合はnull</returns>
        private static T? FindVisualChildByName<T>(System.Windows.DependencyObject parent, string name) where T : System.Windows.FrameworkElement
        {
            if (parent == null)
                return null;

            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T t && t.Name == name)
                {
                    return t;
                }

                var childOfChild = FindVisualChildByName<T>(child, name);
                if (childOfChild != null)
                {
                    return childOfChild;
                }
            }

            return null;
        }

        /// <summary>
        /// 再帰的に要素内のすべてのDependencyObjectのリソースを無効化します
        /// </summary>
        /// <param name="element">開始要素</param>
        private static void InvalidateResourcesRecursive(System.Windows.DependencyObject element)
        {
            if (element == null)
                return;

            // FrameworkElementの場合、Styleプロパティを無効化
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
    }
}
