using FastExplorer.Controls;
using FastExplorer.Models;
using FastExplorer.Services;
using FastExplorer.Views.Pages;
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
        
        // 型をキャッシュ（パフォーマンス向上）
        private static readonly Type ExplorerPageViewModelType = typeof(ExplorerPageViewModel);
        
        // リフレクション結果をキャッシュ（パフォーマンス向上）
        private static readonly string ThemesDictionaryTypeName = "ThemesDictionary";
        private static System.Reflection.PropertyInfo? _cachedThemeProperty;

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
        /// 分割ペインが有効かどうかを取得または設定します
        /// </summary>
        [ObservableProperty]
        private bool _isSplitPaneEnabled;

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
        public Task OnNavigatedFromAsync()
        {
            // 分割ペインの設定を保存
            var settings = _windowSettingsService.GetSettings();
            settings.IsSplitPaneEnabled = IsSplitPaneEnabled;
            _windowSettingsService.SaveSettings(settings);
            return Task.CompletedTask;
        }

        /// <summary>
        /// IsSplitPaneEnabledが変更されたときに呼び出されます
        /// </summary>
        partial void OnIsSplitPaneEnabledChanged(bool value)
        {
            // 設定を即座に保存
            var settings = _windowSettingsService.GetSettings();
            settings.IsSplitPaneEnabled = value;
            _windowSettingsService.SaveSettings(settings);
            
            // ExplorerPageViewModelを更新（無限ループを防ぐため、値が異なる場合のみ更新）
            // 型をキャッシュ（パフォーマンス向上）
            var explorerPageViewModel = App.Services.GetService(ExplorerPageViewModelType) as ExplorerPageViewModel;
            if (explorerPageViewModel != null && explorerPageViewModel.IsSplitPaneEnabled != value)
            {
                explorerPageViewModel.IsSplitPaneEnabled = value;
            }
        }

        /// <summary>
        /// ViewModelを初期化します
        /// </summary>
        private void InitializeViewModel()
        {
            CurrentTheme = ApplicationThemeManager.GetAppTheme();
            AppVersion = $"UiDesktopApp1 - {GetAssemblyVersion()}";

            // テーマカラーリストを初期化
            InitializeThemeColors();

            // 分割ペインの設定を読み込む
            var settings = _windowSettingsService.GetSettings();
            IsSplitPaneEnabled = settings.IsSplitPaneEnabled;

            _isInitialized = true;
        }

        /// <summary>
        /// テーマカラーリストを初期化します
        /// </summary>
        private void InitializeThemeColors()
        {
            ThemeColors.Clear();

            // ThemeColorクラスから定義済みのテーマカラーを取得
            var colors = Models.ThemeColor.GetDefaultThemeColors();

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
                    // LINQクエリを直接ループに置き換え（メモリ割り当てを削減）
                    System.Windows.ResourceDictionary? themesDict = null;
                    foreach (var dict in mergedDictionaries)
                    {
                        if (dict is System.Windows.ResourceDictionary rd && rd.GetType().Name == ThemesDictionaryTypeName)
                        {
                            themesDict = rd;
                            break;
                        }
                    }
                    
                    if (themesDict != null)
                    {
                        // リフレクション結果をキャッシュ（パフォーマンス向上）
                        if (_cachedThemeProperty == null)
                        {
                            _cachedThemeProperty = themesDict.GetType().GetProperty("Theme");
                        }
                        var themeProperty = _cachedThemeProperty;
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
                
                // テーマに応じてテーマカラーを適用またはリセット
                if (theme == ApplicationTheme.Light)
                {
                    // ライトモードの場合は、保存されたテーマカラーを適用
                    var settings = _windowSettingsService.GetSettings();
                    var themeColorCode = settings.ThemeColorCode;
                    if (!string.IsNullOrEmpty(themeColorCode))
                    {
                        App.ApplyThemeColorFromSettings(settings);
                    }
                }
                else if (theme == ApplicationTheme.Dark)
                {
                    // ダークモードの場合は、デフォルトのダークテーマカラーにリセット
                    App.ResetToDefaultDarkThemeColors();
                }
                else
                {
                    // システムテーマ（Unknown）の場合は、システムのテーマに応じて処理
                    var systemTheme = ApplicationThemeManager.GetSystemTheme();
                    if (systemTheme == SystemTheme.Light)
                    {
                        var settings = _windowSettingsService.GetSettings();
                        var themeColorCode = settings.ThemeColorCode;
                        if (!string.IsNullOrEmpty(themeColorCode))
                        {
                            App.ApplyThemeColorFromSettings(settings);
                        }
                    }
                    else
                    {
                        // システムがダークモードの場合は、デフォルトのダークテーマカラーにリセット
                        App.ResetToDefaultDarkThemeColors();
                    }
                }
                
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
                // 現在のテーマがライトモードでない場合は、ライトモードに変更してからテーマカラーを適用
                var currentTheme = ApplicationThemeManager.GetAppTheme();
                if (currentTheme != ApplicationTheme.Light)
                {
                    // ダークモードの場合は、ライトモードに変更してからテーマカラーを適用
                    ApplyTheme(ApplicationTheme.Light);
                }

                // テーマカラーを保存
                var settings2 = _windowSettingsService.GetSettings();
                settings2.ThemeColorName = themeColor.Name;
                settings2.ThemeColorCode = themeColor.ColorCode;
                settings2.ThemeSecondaryColorCode = themeColor.SecondaryColorCode;
                settings2.ThemeThirdColorCode = themeColor.ThirdColorCode;
                _windowSettingsService.SaveSettings(settings2);

                // 色計算を1回だけ実行（高速なカスタム変換を使用）
                var mainColor = Helpers.FastColorConverter.ParseHexColor(themeColor.ColorCode);
                var secondaryColor = Helpers.FastColorConverter.ParseHexColor(themeColor.SecondaryColorCode);
                var mainBrush = new SolidColorBrush(mainColor);
                var secondaryBrush = new SolidColorBrush(secondaryColor);
                // 輝度計算を最適化（定数を事前計算）
                var luminance = (0.299 * mainColor.R + 0.587 * mainColor.G + 0.114 * mainColor.B) * 0.00392156862745098; // 1/255を事前計算
                var statusBarTextColor = luminance > 0.5 ? Colors.Black : Colors.White;
                var statusBarTextBrush = new SolidColorBrush(statusBarTextColor);

                // リソースを更新（計算済みの色を渡して重複計算を回避）
                App.ApplyThemeColorFromSettings(settings2, (mainColor, secondaryColor));

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

                            // ステータスバーは各タブに移動したため、MainWindowからの参照は不要

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
                }), System.Windows.Threading.DispatcherPriority.Render);
            }
            catch (Exception)
            {
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
