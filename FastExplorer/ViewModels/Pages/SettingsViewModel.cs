using FastExplorer.Controls;
using FastExplorer.Models;
using FastExplorer.Services;
using FastExplorer.Views.Pages;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;
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
        #region フィールド

        private bool _isInitialized = false;
        private bool _isLoadingSettings = false; // 設定を読み込み中かどうか
        private readonly WindowSettingsService _windowSettingsService;
        
        // 型をキャッシュ（パフォーマンス向上）
        private static readonly Type ExplorerPageViewModelType = typeof(ExplorerPageViewModel);
        
        // リフレクション結果をキャッシュ（パフォーマンス向上）
        private static readonly string ThemesDictionaryTypeName = "ThemesDictionary";
        private static System.Reflection.PropertyInfo? _cachedThemeProperty;

        // アセンブリのバージョンをキャッシュ（パフォーマンス向上）
        private static string? _cachedAssemblyVersion;

        #endregion

        #region コンストラクタ

        /// <summary>
        /// <see cref="SettingsViewModel"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="windowSettingsService">ウィンドウ設定サービス</param>
        public SettingsViewModel(WindowSettingsService windowSettingsService)
        {
            _windowSettingsService = windowSettingsService;
        }

        #endregion

        #region プロパティ

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
    /// 現在選択されているテーマカラーを取得または設定します
    /// </summary>
    [ObservableProperty]
    private ThemeColor? _selectedThemeColor;

    /// <summary>
    /// 選択中のテーマカラーのブラシを取得または設定します（設定ウィンドウのアクセント用）
    /// </summary>
    [ObservableProperty]
    private System.Windows.Media.SolidColorBrush _selectedAccentBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 120, 212)); // デフォルトの青色

    /// <summary>
    /// 選択中のテーマカラーの薄いブラシを取得または設定します（背景用）
    /// </summary>
    [ObservableProperty]
    private System.Windows.Media.SolidColorBrush _selectedAccentLightBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 0, 120, 212)); // デフォルトの薄い青色（透明度40）

    /// <summary>
    /// 分割ペインが有効かどうかを取得または設定します
    /// </summary>
    [ObservableProperty]
    private bool _isSplitPaneEnabled;

    /// <summary>
    /// 背景画像のファイルパスを取得または設定します
    /// </summary>
    [ObservableProperty]
    private string? _backgroundImagePath;

    /// <summary>
    /// 背景画像の不透明度を取得または設定します（0.0～1.0）
    /// </summary>
    [ObservableProperty]
    private double _backgroundImageOpacity = 1.0;

    /// <summary>
    /// 背景画像の調整方法を取得または設定します
    /// </summary>
    [ObservableProperty]
    private BackgroundImageStretch _backgroundImageStretch = BackgroundImageStretch.FitToWindow;

    /// <summary>
    /// 背景画像の垂直方向の配置を取得または設定します
    /// </summary>
    [ObservableProperty]
    private BackgroundImageAlignment _backgroundImageVerticalAlignment = BackgroundImageAlignment.Center;

    /// <summary>
    /// 背景画像の水平方向の配置を取得または設定します
    /// </summary>
    [ObservableProperty]
    private BackgroundImageAlignment _backgroundImageHorizontalAlignment = BackgroundImageAlignment.Center;

        #endregion

        #region ナビゲーション

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
                // 選択中のテーマカラーも更新
                UpdateSelectedThemeColor();
                
                // 背景画像の設定を再読み込み（設定画面を開くたびに最新の設定を反映）
                _isLoadingSettings = true; // 設定読み込み中フラグを設定
                try
                {
                    var settings = _windowSettingsService.GetSettings();
                    
                    // プロパティを設定（値が同じ場合でも変更通知を発火）
                    var newPath = settings.BackgroundImagePath;
                    if (BackgroundImagePath != newPath)
                    {
                        BackgroundImagePath = newPath;
                    }
                    OnPropertyChanged(nameof(BackgroundImagePath));
                    
                    var newOpacity = settings.BackgroundImageOpacity;
                    if (newOpacity <= 0)
                    {
                        newOpacity = 1.0; // デフォルト値
                    }
                    if (Math.Abs(BackgroundImageOpacity - newOpacity) > 0.001)
                    {
                        BackgroundImageOpacity = newOpacity;
                    }
                    OnPropertyChanged(nameof(BackgroundImageOpacity));
                    
                    // enumプロパティを更新し、文字列プロパティの変更通知も発火
                    BackgroundImageStretch stretchValue;
                    if (!string.IsNullOrEmpty(settings.BackgroundImageStretch) && Enum.TryParse<BackgroundImageStretch>(settings.BackgroundImageStretch, out stretchValue))
                    {
                        if (BackgroundImageStretch != stretchValue)
                        {
                            BackgroundImageStretch = stretchValue;
                        }
                    }
                    else
                    {
                        stretchValue = BackgroundImageStretch.FitToWindow;
                        if (BackgroundImageStretch != stretchValue)
                        {
                            BackgroundImageStretch = stretchValue;
                        }
                    }
                    OnPropertyChanged(nameof(BackgroundImageStretch));
                    OnPropertyChanged(nameof(BackgroundImageStretchString));
                    
                    BackgroundImageAlignment vAlignValue;
                    if (!string.IsNullOrEmpty(settings.BackgroundImageVerticalAlignment) && Enum.TryParse<BackgroundImageAlignment>(settings.BackgroundImageVerticalAlignment, out vAlignValue))
                    {
                        if (BackgroundImageVerticalAlignment != vAlignValue)
                        {
                            BackgroundImageVerticalAlignment = vAlignValue;
                        }
                    }
                    else
                    {
                        vAlignValue = BackgroundImageAlignment.Center;
                        if (BackgroundImageVerticalAlignment != vAlignValue)
                        {
                            BackgroundImageVerticalAlignment = vAlignValue;
                        }
                    }
                    OnPropertyChanged(nameof(BackgroundImageVerticalAlignment));
                    OnPropertyChanged(nameof(BackgroundImageVerticalAlignmentString));
                    
                    BackgroundImageAlignment hAlignValue;
                    if (!string.IsNullOrEmpty(settings.BackgroundImageHorizontalAlignment) && Enum.TryParse<BackgroundImageAlignment>(settings.BackgroundImageHorizontalAlignment, out hAlignValue))
                    {
                        if (BackgroundImageHorizontalAlignment != hAlignValue)
                        {
                            BackgroundImageHorizontalAlignment = hAlignValue;
                        }
                    }
                    else
                    {
                        hAlignValue = BackgroundImageAlignment.Center;
                        if (BackgroundImageHorizontalAlignment != hAlignValue)
                        {
                            BackgroundImageHorizontalAlignment = hAlignValue;
                        }
                    }
                    OnPropertyChanged(nameof(BackgroundImageHorizontalAlignment));
                    OnPropertyChanged(nameof(BackgroundImageHorizontalAlignmentString));
                }
                finally
                {
                    _isLoadingSettings = false; // 設定読み込み完了
                }
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
            
            // 背景画像の設定を保存
            settings.BackgroundImagePath = BackgroundImagePath;
            settings.BackgroundImageOpacity = BackgroundImageOpacity;
            settings.BackgroundImageStretch = BackgroundImageStretch.ToString();
            settings.BackgroundImageVerticalAlignment = BackgroundImageVerticalAlignment.ToString();
            settings.BackgroundImageHorizontalAlignment = BackgroundImageHorizontalAlignment.ToString();
            
            _windowSettingsService.SaveSettings(settings);
            return Task.CompletedTask;
        }

        #endregion

        #region プロパティ変更ハンドラー

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
        /// 背景画像の設定が変更されたときに呼び出されます
        /// </summary>
        partial void OnBackgroundImagePathChanged(string? value)
        {
            if (!_isLoadingSettings)
            {
                SaveBackgroundImageSettings();
                ApplyBackgroundImage();
            }
        }

        /// <summary>
        /// 背景画像の不透明度が変更されたときに呼び出されます
        /// </summary>
        partial void OnBackgroundImageOpacityChanged(double value)
        {
            if (!_isLoadingSettings)
            {
                SaveBackgroundImageSettings();
                ApplyBackgroundImage();
            }
        }

        /// <summary>
        /// 背景画像の調整方法が変更されたときに呼び出されます
        /// </summary>
        partial void OnBackgroundImageStretchChanged(BackgroundImageStretch value)
        {
            OnPropertyChanged(nameof(BackgroundImageStretchString));
            if (!_isLoadingSettings)
            {
                SaveBackgroundImageSettings();
                ApplyBackgroundImage();
            }
        }

        /// <summary>
        /// 背景画像の垂直方向の配置が変更されたときに呼び出されます
        /// </summary>
        partial void OnBackgroundImageVerticalAlignmentChanged(BackgroundImageAlignment value)
        {
            OnPropertyChanged(nameof(BackgroundImageVerticalAlignmentString));
            if (!_isLoadingSettings)
            {
                SaveBackgroundImageSettings();
                ApplyBackgroundImage();
            }
        }

        /// <summary>
        /// 背景画像の水平方向の配置が変更されたときに呼び出されます
        /// </summary>
        partial void OnBackgroundImageHorizontalAlignmentChanged(BackgroundImageAlignment value)
        {
            OnPropertyChanged(nameof(BackgroundImageHorizontalAlignmentString));
            if (!_isLoadingSettings)
            {
                SaveBackgroundImageSettings();
                ApplyBackgroundImage();
            }
        }

        /// <summary>
        /// 背景画像の調整方法を文字列から設定します（ComboBox用）
        /// </summary>
        public string BackgroundImageStretchString
        {
            get => BackgroundImageStretch.ToString();
            set
            {
                if (Enum.TryParse<BackgroundImageStretch>(value, out var stretch))
                {
                    BackgroundImageStretch = stretch;
                    OnPropertyChanged(nameof(BackgroundImageStretchString));
                }
            }
        }

        /// <summary>
        /// 背景画像の垂直方向の配置を文字列から設定します（ComboBox用）
        /// </summary>
        public string BackgroundImageVerticalAlignmentString
        {
            get => BackgroundImageVerticalAlignment.ToString();
            set
            {
                if (Enum.TryParse<BackgroundImageAlignment>(value, out var alignment))
                {
                    BackgroundImageVerticalAlignment = alignment;
                    OnPropertyChanged(nameof(BackgroundImageVerticalAlignmentString));
                }
            }
        }

        /// <summary>
        /// 背景画像の水平方向の配置を文字列から設定します（ComboBox用）
        /// </summary>
        public string BackgroundImageHorizontalAlignmentString
        {
            get => BackgroundImageHorizontalAlignment.ToString();
            set
            {
                if (Enum.TryParse<BackgroundImageAlignment>(value, out var alignment))
                {
                    BackgroundImageHorizontalAlignment = alignment;
                    OnPropertyChanged(nameof(BackgroundImageHorizontalAlignmentString));
                }
            }
        }

        /// <summary>
        /// 背景画像の設定を保存します
        /// </summary>
        private void SaveBackgroundImageSettings()
        {
            var settings = _windowSettingsService.GetSettings();
            settings.BackgroundImagePath = BackgroundImagePath;
            settings.BackgroundImageOpacity = BackgroundImageOpacity;
            settings.BackgroundImageStretch = BackgroundImageStretch.ToString();
            settings.BackgroundImageVerticalAlignment = BackgroundImageVerticalAlignment.ToString();
            settings.BackgroundImageHorizontalAlignment = BackgroundImageHorizontalAlignment.ToString();
            _windowSettingsService.SaveSettings(settings);
        }

        /// <summary>
        /// 背景画像を適用します
        /// </summary>
        private void ApplyBackgroundImage()
        {
            Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
            {
                foreach (Window window in Application.Current.Windows)
                {
                    if (window is Views.Windows.MainWindow mainWindow)
                    {
                        mainWindow.ApplyBackgroundImage(BackgroundImagePath, BackgroundImageOpacity, BackgroundImageStretch, BackgroundImageVerticalAlignment, BackgroundImageHorizontalAlignment);
                    }
                }
            }), System.Windows.Threading.DispatcherPriority.Render);
        }

        #endregion

        #region 初期化

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
            
            // 背景画像の設定を読み込む
            _isLoadingSettings = true; // 設定読み込み中フラグを設定
            try
            {
                BackgroundImagePath = settings.BackgroundImagePath;
                // 不透明度を読み込む（0以下の場合はデフォルト値1.0を使用）
                var opacity = settings.BackgroundImageOpacity;
                if (opacity <= 0)
                {
                    opacity = 1.0;
                }
                BackgroundImageOpacity = opacity;
                
                // enumプロパティを更新し、文字列プロパティの変更通知も発火
                if (!string.IsNullOrEmpty(settings.BackgroundImageStretch) && Enum.TryParse<BackgroundImageStretch>(settings.BackgroundImageStretch, out var stretch))
                {
                    BackgroundImageStretch = stretch;
                }
                else
                {
                    BackgroundImageStretch = BackgroundImageStretch.FitToWindow;
                }
                OnPropertyChanged(nameof(BackgroundImageStretchString));
                
                if (!string.IsNullOrEmpty(settings.BackgroundImageVerticalAlignment) && Enum.TryParse<BackgroundImageAlignment>(settings.BackgroundImageVerticalAlignment, out var vAlign))
                {
                    BackgroundImageVerticalAlignment = vAlign;
                }
                else
                {
                    BackgroundImageVerticalAlignment = BackgroundImageAlignment.Center;
                }
                OnPropertyChanged(nameof(BackgroundImageVerticalAlignmentString));
                
                if (!string.IsNullOrEmpty(settings.BackgroundImageHorizontalAlignment) && Enum.TryParse<BackgroundImageAlignment>(settings.BackgroundImageHorizontalAlignment, out var hAlign))
                {
                    BackgroundImageHorizontalAlignment = hAlign;
                }
                else
                {
                    BackgroundImageHorizontalAlignment = BackgroundImageAlignment.Center;
                }
                OnPropertyChanged(nameof(BackgroundImageHorizontalAlignmentString));
            }
            finally
            {
                _isLoadingSettings = false; // 設定読み込み完了
            }
            
            // 現在選択されているテーマカラーを設定
            UpdateSelectedThemeColor();

            _isInitialized = true;
        }

        #endregion

        #region テーマカラー管理

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
        /// 現在選択されているテーマカラーを更新します
        /// </summary>
        private void UpdateSelectedThemeColor()
        {
            var settings = _windowSettingsService.GetSettings();
            var themeColorCode = settings.ThemeColorCode;
            
            if (!string.IsNullOrEmpty(themeColorCode))
            {
                // 保存されているテーマカラーと一致するものを選択
                SelectedThemeColor = ThemeColors.FirstOrDefault(tc => tc.ColorCode == themeColorCode);
                
                // アクセントブラシを更新
                if (SelectedThemeColor != null && SelectedThemeColor.Color is System.Windows.Media.SolidColorBrush brush)
                {
                    SelectedAccentBrush = brush;
                    // 薄いブラシも作成（透明度40）
                    var color = brush.Color;
                    SelectedAccentLightBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(40, color.R, color.G, color.B));
                }
                else
                {
                    // デフォルトの青色を設定
                    SelectedAccentBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 120, 212));
                    SelectedAccentLightBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 0, 120, 212));
                }
            }
            else
            {
                SelectedThemeColor = null;
                // デフォルトの青色を設定
                SelectedAccentBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 120, 212));
                SelectedAccentLightBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 0, 120, 212));
            }
        }

        /// <summary>
        /// アセンブリのバージョンを取得します
        /// </summary>
        /// <returns>アセンブリのバージョン文字列</returns>
        private string GetAssemblyVersion()
        {
            // パフォーマンス最適化：アセンブリバージョンは変更されないのでキャッシュ
            return _cachedAssemblyVersion ??= 
                System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? String.Empty;
        }

        #endregion

        #region テーマ管理

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
                        
                        // 選択状態を更新
                        UpdateSelectedThemeColor();
                        
                        // MainWindowの背景色を再適用
                        var mainColor = Helpers.FastColorConverter.ParseHexColor(themeColorCode);
                        var secondaryColorCode = settings.ThemeSecondaryColorCode;
                        if (!string.IsNullOrEmpty(secondaryColorCode))
                        {
                            var secondaryColor = Helpers.FastColorConverter.ParseHexColor(secondaryColorCode);
                            var mainBrush = new SolidColorBrush(mainColor);
                            var secondaryBrush = new SolidColorBrush(secondaryColor);
                            
                            Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                            {
                                foreach (Window window in Application.Current.Windows)
                                {
                                    if (window is Views.Windows.MainWindow mainWindow)
                                    {
                                        mainWindow.Background = mainBrush;
                                        var navigationView = FindVisualChild<Wpf.Ui.Controls.NavigationView>(mainWindow);
                                        if (navigationView != null)
                                        {
                                            navigationView.Background = secondaryBrush;
                                        }
                                    }
                                }
                            }), System.Windows.Threading.DispatcherPriority.Render);
                        }
                    }
                }
                else if (theme == ApplicationTheme.Dark)
                {
                    // ダークモードの場合は、デフォルトのダークテーマカラーにリセット
                    App.ResetToDefaultDarkThemeColors();
                    // ダークモードでは選択状態を保持（視覚的な選択は表示するが、適用はしない）
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
                            UpdateSelectedThemeColor();
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
                
                // タブとListViewのスタイルを無効化してDynamicResourceの再評価を強制
                // Backgroundに変更して確実に実行されるようにする
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    foreach (Window window in Application.Current.Windows)
                    {
                        // MainWindowのスタイルを更新
                        if (window is Views.Windows.MainWindow)
                        {
                            Views.Windows.MainWindow.InvalidateTabAndListViewStyles(window);
                        }
                        // SettingsWindowの背景色をクリアしてDynamicResourceを使用するように戻す
                        else if (window is Views.Windows.SettingsWindow settingsWindow)
                        {
                            // 背景色をクリアして、DynamicResourceを使用するように戻す
                            settingsWindow.ClearValue(Window.BackgroundProperty);
                        }
                    }
                }), System.Windows.Threading.DispatcherPriority.Background);
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

        #endregion

        #region テーマカラー選択

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

                // 選択されたテーマカラーを更新
                SelectedThemeColor = themeColor;
                
                // アクセントブラシを更新（設定ウィンドウのUI用）
                if (themeColor.Color is System.Windows.Media.SolidColorBrush brush)
                {
                    SelectedAccentBrush = brush;
                    // 薄いブラシも作成（透明度40）
                    var color = brush.Color;
                    SelectedAccentLightBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(40, color.R, color.G, color.B));
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

                // MainWindowの背景色を直接更新（SettingsWindowには適用しない）
                Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    // MainWindowのみを対象にする
                    foreach (Window window in Application.Current.Windows)
                    {
                        // SettingsWindowは除外
                        if (window is Views.Windows.MainWindow mainWindow)
                        {
                            // ウィンドウの背景色を直接設定
                            mainWindow.Background = mainBrush;

                            // ウィンドウ内のNavigationViewの背景色も更新
                            var navigationView = FindVisualChild<Wpf.Ui.Controls.NavigationView>(mainWindow);
                            if (navigationView != null)
                            {
                                navigationView.Background = secondaryBrush;
                            }

                            // ウィンドウのリソースを無効化
                            mainWindow.InvalidateProperty(System.Windows.FrameworkElement.StyleProperty);
                            mainWindow.InvalidateProperty(System.Windows.Controls.Control.BackgroundProperty);

                            // タブとListViewの選択中の色を更新するため、スタイルを無効化してDynamicResourceの再評価を強制
                            Views.Windows.MainWindow.InvalidateTabAndListViewStyles(mainWindow);
                        }
                    }
                }), System.Windows.Threading.DispatcherPriority.Render);
            }
            catch (Exception)
            {
            }
        }

        #endregion

        #region 背景画像管理

        /// <summary>
        /// 背景画像を参照するコマンド
        /// </summary>
        [RelayCommand]
        private void BrowseBackgroundImage()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "画像ファイル|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tiff;*.webp|すべてのファイル|*.*",
                Title = "背景画像を選択"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                BackgroundImagePath = openFileDialog.FileName;
            }
        }

        /// <summary>
        /// 背景画像をクリアするコマンド
        /// </summary>
        [RelayCommand]
        private void ClearBackgroundImage()
        {
            BackgroundImagePath = null;
        }

        #endregion

        #region ビジュアルツリー操作

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

            // パフォーマンス最適化：GetChildrenCountを一度だけ呼び出す
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
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

            // パフォーマンス最適化：GetChildrenCountを一度だけ呼び出す
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
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

            // パフォーマンス最適化：GetChildrenCountを一度だけ呼び出す
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(element);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(element, i);
                InvalidateResourcesRecursive(child);
            }
        }

        #endregion
    }
}
