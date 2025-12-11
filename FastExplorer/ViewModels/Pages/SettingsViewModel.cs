using FastExplorer.Controls;
using FastExplorer.Helpers;
using FastExplorer.Models;
using FastExplorer.Services;
using FastExplorer.Views.Pages;
using System;
using System.Collections.Generic;
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
        private bool _isNavigatingFrom = false; // ナビゲーション離脱処理中かどうか
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

    /// <summary>
    /// 利用可能な言語のコレクションを取得または設定します
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<LanguageItem> _availableLanguages = new();

    /// <summary>
    /// 選択された言語を取得または設定します
    /// </summary>
    [ObservableProperty]
    private LanguageItem? _selectedLanguage;

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
                // 既に初期化済みの場合でも、言語設定を再読み込み
                InitializeLanguages();
                
                // 現在のテーマを更新
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
            // 重複呼び出しを防ぐ
            if (_isNavigatingFrom)
            {
                return Task.CompletedTask;
            }

            _isNavigatingFrom = true;
            try
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
                
                // 言語設定を保存
                if (SelectedLanguage != null)
                {
                    settings.LanguageCode = SelectedLanguage.Code;
                }
                
                _windowSettingsService.SaveSettings(settings);
            }
            finally
            {
                _isNavigatingFrom = false;
            }
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
        /// 選択された言語が変更されたときに呼び出されます
        /// </summary>
        partial void OnSelectedLanguageChanged(LanguageItem? value)
        {
            if (value != null && !_isLoadingSettings && !_isNavigatingFrom)
            {
                System.Diagnostics.Debug.WriteLine($"OnSelectedLanguageChanged called: {value.Code} - {value.Name}");
                
                // 現在の言語コードを取得（変更前）
                var currentLanguageCode = LocalizationHelper.CurrentLanguageCode;
                
                // 言語が実際に変更されたかどうかを確認
                if (currentLanguageCode == value.Code)
                {
                    System.Diagnostics.Debug.WriteLine("Language code is the same, skipping");
                    return;
                }
                
                // 設定を保存
                var settings = _windowSettingsService.GetSettings();
                settings.LanguageCode = value.Code;
                _windowSettingsService.SaveSettings(settings);

                // メッセージボックスを表示して、アプリを再起動するように促す
                Application.Current.Dispatcher.BeginInvoke(new System.Action(async () =>
                {
                    try
                    {
                        // 設定ウィンドウを取得
                        Views.Windows.SettingsWindow? settingsWindow = null;
                        foreach (Window window in Application.Current.Windows)
                        {
                            if (window is Views.Windows.SettingsWindow sw && sw.IsLoaded && sw.IsVisible)
                            {
                                settingsWindow = sw;
                                break;
                            }
                        }

                        // メッセージボックスを表示
                        var messageBox = new Wpf.Ui.Controls.MessageBox
                        {
                            Title = LocalizationHelper.GetString("LanguageChangedTitle", "言語設定が変更されました"),
                            Content = LocalizationHelper.GetString("LanguageChangedMessage", "言語設定を変更しました。変更を反映するには、アプリケーションを再起動してください。")
                        };

                        if (settingsWindow != null)
                        {
                            messageBox.Owner = settingsWindow;
                        }

                        await messageBox.ShowDialogAsync();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"メッセージボックス表示エラー: {ex.Message}");
                        System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                    }
                }), System.Windows.Threading.DispatcherPriority.Normal);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"OnSelectedLanguageChanged skipped: value={value?.Code}, _isLoadingSettings={_isLoadingSettings}, _isNavigatingFrom={_isNavigatingFrom}");
            }
        }

        /// <summary>
        /// メインウィンドウのUIをリフレッシュします（設定ウィンドウが開いている場合に使用）
        /// </summary>
        private void RefreshMainWindowUI()
        {
            // MainWindowViewModelのプロパティを更新
            var mainWindowViewModel = App.Services.GetService(typeof(ViewModels.Windows.MainWindowViewModel)) as ViewModels.Windows.MainWindowViewModel;
            if (mainWindowViewModel != null)
            {
                mainWindowViewModel.ApplicationTitle = LocalizationHelper.GetString("ApplicationTitle", "FastExplorer");
                mainWindowViewModel.StatusBarText = LocalizationHelper.GetString("Ready", "準備完了");
                
                // ホームメニューアイテムのContentを更新
                // MainWindowViewModelの_homeMenuItemを直接更新
                var homeMenuItemField = typeof(ViewModels.Windows.MainWindowViewModel).GetField("_homeMenuItem", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                if (homeMenuItemField != null)
                {
                    var homeMenuItem = homeMenuItemField.GetValue(mainWindowViewModel) as Wpf.Ui.Controls.NavigationViewItem;
                    if (homeMenuItem != null)
                    {
                        homeMenuItem.Content = LocalizationHelper.GetString("Home", "ホーム");
                    }
                }
                
                // MenuItemsからも更新（念のため）
                foreach (var item in mainWindowViewModel.MenuItems)
                {
                    if (item is Wpf.Ui.Controls.NavigationViewItem navItem && navItem.Tag is string tag && tag == "HOME")
                    {
                        navItem.Content = LocalizationHelper.GetString("Home", "ホーム");
                        break;
                    }
                }
            }

            // ExplorerViewModelのStatusBarTextを更新
            var explorerPageViewModel = App.Services.GetService(ExplorerPageViewModelType) as ExplorerPageViewModel;
            if (explorerPageViewModel != null)
            {
                // ExplorerViewModelのStatusBarTextを更新するメソッドを呼び出す
                var explorerViewModelType = typeof(ViewModels.Pages.ExplorerViewModel);
                var statusBarTextProperty = explorerViewModelType.GetProperty("StatusBarText");
                if (statusBarTextProperty != null)
                {
                    // 各タブのViewModelのStatusBarTextを更新
                    if (explorerPageViewModel.IsSplitPaneEnabled)
                    {
                        foreach (var tab in explorerPageViewModel.LeftPaneTabs)
                        {
                            if (tab.ViewModel != null)
                            {
                                statusBarTextProperty.SetValue(tab.ViewModel, LocalizationHelper.GetString("Ready", "準備完了"));
                            }
                        }
                        foreach (var tab in explorerPageViewModel.RightPaneTabs)
                        {
                            if (tab.ViewModel != null)
                            {
                                statusBarTextProperty.SetValue(tab.ViewModel, LocalizationHelper.GetString("Ready", "準備完了"));
                            }
                        }
                    }
                    else
                    {
                        foreach (var tab in explorerPageViewModel.Tabs)
                        {
                            if (tab.ViewModel != null)
                            {
                                statusBarTextProperty.SetValue(tab.ViewModel, LocalizationHelper.GetString("Ready", "準備完了"));
                            }
                        }
                    }
                }
            }

            // メインウィンドウのLocalizationHelperExtensionを使用している要素を更新
            foreach (Window window in Application.Current.Windows)
            {
                try
                {
                    // 設定ウィンドウは除外
                    if (window is Views.Windows.SettingsWindow)
                    {
                        continue;
                    }

                    if (window is Views.Windows.MainWindow mainWindow && mainWindow.IsLoaded && mainWindow.IsVisible)
                    {
                        // ListViewの列ヘッダーを更新
                        UpdateListViewColumnHeaders(mainWindow);
                        
                        // その他の要素を更新（設定ウィンドウを除外）
                        UpdateLocalizedElements(mainWindow);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"メインウィンドウ更新エラー: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ListViewの列ヘッダーとExplorerPage内の要素を更新します
        /// </summary>
        /// <param name="window">ウィンドウ</param>
        private static void UpdateListViewColumnHeaders(System.Windows.Window window)
        {
            try
            {
                // ExplorerPage内のListViewを検索
                var explorerPage = FindVisualChild<Views.Pages.ExplorerPage>(window);
                if (explorerPage != null)
                {
                    // すべてのListViewを検索（分割ペインの場合、左右両方のペインにListViewがある）
                    var listViews = FindVisualChildren<System.Windows.Controls.ListView>(explorerPage);
                    foreach (var listView in listViews)
                    {
                        if (listView.View is System.Windows.Controls.GridView gridView)
                        {
                            // GridViewColumnのHeaderを更新
                            foreach (var column in gridView.Columns)
                            {
                                if (column.Header is string headerText)
                                {
                                    string? key = null;
                                    if (headerText == "名前" || headerText == "Name")
                                        key = "Name";
                                    else if (headerText == "サイズ" || headerText == "Size")
                                        key = "Size";
                                    else if (headerText == "種類" || headerText == "Type")
                                        key = "Type";
                                    else if (headerText == "更新日時" || headerText == "Last Modified")
                                        key = "LastModified";

                                    if (key != null)
                                    {
                                        column.Header = LocalizationHelper.GetString(key, headerText);
                                    }
                                }
                            }
                        }
                    }

                    // ExplorerPage内のボタンやその他の要素を更新
                    // ListViewをスキップせずに、ListViewの親要素から更新する
                    UpdateExplorerPageElements(explorerPage);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ListView列ヘッダー更新エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// ExplorerPage内の要素を更新します（ListView内の要素も含む）
        /// </summary>
        /// <param name="explorerPage">ExplorerPage</param>
        private static void UpdateExplorerPageElements(Views.Pages.ExplorerPage explorerPage)
        {
            try
            {
                // ExplorerPage内のすべての要素を更新
                // ListViewをスキップせずに、ListViewの親要素から更新する
                UpdateLocalizedElementsInExplorerPage(explorerPage);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExplorerPage要素更新エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// ExplorerPage内の要素を更新します（ListView内の要素も含む）
        /// </summary>
        /// <param name="element">開始要素</param>
        private static void UpdateLocalizedElementsInExplorerPage(System.Windows.DependencyObject element)
        {
            if (element == null)
                return;

            // ListView内の要素も更新するため、ListViewをスキップしない
            // ただし、ListViewItemはスキップ（リスト項目が消えるのを防ぐ）

            // TextBlockの場合、Textプロパティを更新
            if (element is System.Windows.Controls.TextBlock textBlock)
            {
                var textBinding = System.Windows.Data.BindingOperations.GetBinding(textBlock, System.Windows.Controls.TextBlock.TextProperty);
                if (textBinding == null)
                {
                    var currentText = textBlock.Text;
                    if (!string.IsNullOrEmpty(currentText))
                    {
                        // ToolTipも更新
                        if (textBlock.ToolTip is string toolTipText)
                        {
                            UpdateTextBlockToolTip(textBlock, toolTipText);
                        }
                    }
                }
            }

            // Buttonの場合、ContentとToolTipを更新
            if (element is System.Windows.Controls.Button button)
            {
                var contentBinding = System.Windows.Data.BindingOperations.GetBinding(button, System.Windows.Controls.Button.ContentProperty);
                if (contentBinding == null && button.Content is string contentText)
                {
                    UpdateButtonByContent(button, contentText);
                }
                if (button.ToolTip is string toolTipText)
                {
                    UpdateButtonToolTip(button, toolTipText);
                }
            }

            // Wpf.Ui.Controls.Buttonの場合も更新
            if (element is Wpf.Ui.Controls.Button uiButton)
            {
                if (uiButton.ToolTip is string toolTipText)
                {
                    UpdateButtonToolTip(uiButton, toolTipText);
                }
            }

            // ListViewItemはスキップ（リスト項目が消えるのを防ぐ）
            if (element is System.Windows.Controls.ListViewItem)
            {
                return;
            }

            // 子要素を再帰的に処理
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(element);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(element, i);
                UpdateLocalizedElementsInExplorerPage(child);
            }
        }

        /// <summary>
        /// テキストからキーを推測してTextBlockのToolTipを更新します
        /// </summary>
        private static void UpdateTextBlockToolTip(System.Windows.Controls.TextBlock textBlock, string toolTipText)
        {
            string? key = null;
            if (toolTipText == "パスを編集" || toolTipText == "Edit path")
                key = "EditPath";
            else if (toolTipText == "戻る" || toolTipText == "Go back")
                key = "GoBack";
            else if (toolTipText == "更新" || toolTipText == "Refresh")
                key = "Refresh";
            else if (toolTipText == "お気に入りに追加" || toolTipText == "Add to favorites")
                key = "AddToFavorites";
            else if (toolTipText == "パスをクリップボードにコピー" || toolTipText == "Copy path to clipboard")
                key = "CopyPathToClipboard";

            if (key != null)
            {
                textBlock.ToolTip = LocalizationHelper.GetString(key, toolTipText);
            }
        }

        /// <summary>
        /// コンテンツからキーを推測してButtonを更新します
        /// </summary>
        private static void UpdateButtonByContent(System.Windows.Controls.Button button, string contentText)
        {
            string? key = null;
            if (contentText == "戻る" || contentText == "Go back")
                key = "GoBack";
            else if (contentText == "更新" || contentText == "Refresh")
                key = "Refresh";

            if (key != null)
            {
                button.Content = LocalizationHelper.GetString(key, contentText);
            }
        }

        /// <summary>
        /// ToolTipからキーを推測してButtonのToolTipを更新します
        /// </summary>
        private static void UpdateButtonToolTip(System.Windows.Controls.Button button, string toolTipText)
        {
            string? key = null;
            if (toolTipText == "パスを編集" || toolTipText == "Edit path")
                key = "EditPath";
            else if (toolTipText == "戻る" || toolTipText == "Go back")
                key = "GoBack";
            else if (toolTipText == "更新" || toolTipText == "Refresh")
                key = "Refresh";
            else if (toolTipText == "お気に入りに追加" || toolTipText == "Add to favorites")
                key = "AddToFavorites";
            else if (toolTipText == "パスをクリップボードにコピー" || toolTipText == "Copy path to clipboard")
                key = "CopyPathToClipboard";

            if (key != null)
            {
                button.ToolTip = LocalizationHelper.GetString(key, toolTipText);
            }
        }

        /// <summary>
        /// ToolTipからキーを推測してWpf.Ui.Controls.ButtonのToolTipを更新します
        /// </summary>
        private static void UpdateButtonToolTip(Wpf.Ui.Controls.Button button, string toolTipText)
        {
            string? key = null;
            if (toolTipText == "パスを編集" || toolTipText == "Edit path")
                key = "EditPath";
            else if (toolTipText == "戻る" || toolTipText == "Go back")
                key = "GoBack";
            else if (toolTipText == "更新" || toolTipText == "Refresh")
                key = "Refresh";
            else if (toolTipText == "お気に入りに追加" || toolTipText == "Add to favorites")
                key = "AddToFavorites";
            else if (toolTipText == "パスをクリップボードにコピー" || toolTipText == "Copy path to clipboard")
                key = "CopyPathToClipboard";

            if (key != null)
            {
                button.ToolTip = LocalizationHelper.GetString(key, toolTipText);
            }
        }

        /// <summary>
        /// ビジュアルツリーから指定された型のすべての子要素を検索します
        /// </summary>
        private static System.Collections.Generic.IEnumerable<T> FindVisualChildren<T>(System.Windows.DependencyObject parent) where T : System.Windows.DependencyObject
        {
            if (parent == null)
                yield break;

            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T t)
                {
                    yield return t;
                }

                foreach (var childOfChild in FindVisualChildren<T>(child))
                {
                    yield return childOfChild;
                }
            }
        }

        /// <summary>
        /// UIをリフレッシュします（言語変更時に呼び出される）
        /// </summary>
        private void RefreshUI()
        {
            // ナビゲーション離脱処理中はリフレッシュしない（設定ウィンドウが閉じられる前後で問題が発生するのを防ぐ）
            if (_isNavigatingFrom)
            {
                System.Diagnostics.Debug.WriteLine("RefreshUI skipped: _isNavigatingFrom is true");
                return;
            }

            // 設定ウィンドウが開いているかどうかを確認（RefreshUIを呼ぶ前にチェック）
            bool isSettingsWindowOpen = false;
            foreach (Window window in Application.Current.Windows)
            {
                if (window is Views.Windows.SettingsWindow settingsWindow && settingsWindow.IsLoaded && settingsWindow.IsVisible)
                {
                    isSettingsWindowOpen = true;
                    System.Diagnostics.Debug.WriteLine($"Settings window is open: {settingsWindow.Title}");
                    break;
                }
            }

            // 設定ウィンドウが開いている場合は、設定ウィンドウの更新をスキップする
            // 設定ウィンドウ内で言語を変更した場合、設定ウィンドウ自体は既に正しい言語で表示されているため
            if (isSettingsWindowOpen)
            {
                System.Diagnostics.Debug.WriteLine("RefreshUI: Settings window is open, skipping settings window update");
            }

            // アプリケーションタイトルを更新
            Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
            {
                // MainWindowViewModelのプロパティを更新
                var mainWindowViewModel = App.Services.GetService(typeof(ViewModels.Windows.MainWindowViewModel)) as ViewModels.Windows.MainWindowViewModel;
                if (mainWindowViewModel != null)
                {
                    mainWindowViewModel.ApplicationTitle = LocalizationHelper.GetString("ApplicationTitle", "FastExplorer");
                    mainWindowViewModel.StatusBarText = LocalizationHelper.GetString("Ready", "準備完了");
                    
                    // ホームメニューアイテムのContentを更新
                    foreach (var item in mainWindowViewModel.MenuItems)
                    {
                        if (item is Wpf.Ui.Controls.NavigationViewItem navItem && navItem.Tag is string tag && tag == "HOME")
                        {
                            navItem.Content = LocalizationHelper.GetString("Home", "ホーム");
                            break;
                        }
                    }
                }

                // ExplorerViewModelのStatusBarTextを更新
                var explorerPageViewModel = App.Services.GetService(ExplorerPageViewModelType) as ExplorerPageViewModel;
                if (explorerPageViewModel != null)
                {
                    // ExplorerViewModelのStatusBarTextを更新するメソッドを呼び出す
                    var explorerViewModelType = typeof(ViewModels.Pages.ExplorerViewModel);
                    var statusBarTextProperty = explorerViewModelType.GetProperty("StatusBarText");
                    if (statusBarTextProperty != null)
                    {
                        // 各タブのViewModelのStatusBarTextを更新
                        if (explorerPageViewModel.IsSplitPaneEnabled)
                        {
                            foreach (var tab in explorerPageViewModel.LeftPaneTabs)
                            {
                                if (tab.ViewModel != null)
                                {
                                    statusBarTextProperty.SetValue(tab.ViewModel, LocalizationHelper.GetString("Ready", "準備完了"));
                                }
                            }
                            foreach (var tab in explorerPageViewModel.RightPaneTabs)
                            {
                                if (tab.ViewModel != null)
                                {
                                    statusBarTextProperty.SetValue(tab.ViewModel, LocalizationHelper.GetString("Ready", "準備完了"));
                                }
                            }
                        }
                        else
                        {
                            foreach (var tab in explorerPageViewModel.Tabs)
                            {
                                if (tab.ViewModel != null)
                                {
                                    statusBarTextProperty.SetValue(tab.ViewModel, LocalizationHelper.GetString("Ready", "準備完了"));
                                }
                            }
                        }
                    }
                }

                // 設定ウィンドウのテキストを明示的に更新（開いている場合のみ）
                // ただし、設定ウィンドウ内で言語を変更した場合は更新をスキップする
                // （設定ウィンドウ自体は既に正しい言語で表示されているため、更新すると問題が発生する可能性がある）
                if (isSettingsWindowOpen)
                {
                    System.Diagnostics.Debug.WriteLine("RefreshUI: Skipping settings window update (language changed within settings window)");
                    // 設定ウィンドウ内で言語を変更した場合は、設定ウィンドウの更新をスキップ
                    // 設定ウィンドウは既に正しい言語で表示されているため
                }

                // その他のウィンドウは、LocalizationHelperExtensionを使用している要素のみを更新
                // タブやListViewなどの重要なUI要素のスタイルは無効化しない
                foreach (Window window in Application.Current.Windows)
                {
                    try
                    {
                        if (window != null && !(window is Views.Windows.SettingsWindow) && window.IsLoaded && window.IsVisible)
                        {
                            UpdateLocalizedElements(window);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"ウィンドウ更新エラー: {ex.Message}");
                    }
                }
            }), System.Windows.Threading.DispatcherPriority.Normal);
        }

        /// <summary>
        /// 設定ウィンドウのテキストを明示的に更新します
        /// </summary>
        /// <param name="settingsWindow">設定ウィンドウ</param>
        private static void UpdateSettingsWindowTexts(Views.Windows.SettingsWindow settingsWindow)
        {
            try
            {
                // 設定ウィンドウが閉じられている場合は処理しない
                if (settingsWindow == null || !settingsWindow.IsLoaded || !settingsWindow.IsVisible)
                {
                    System.Diagnostics.Debug.WriteLine("Settings window is not loaded or not visible, skipping update");
                    return;
                }

                System.Diagnostics.Debug.WriteLine("Updating settings window texts");

                // 設定ウィンドウのタイトルを更新
                settingsWindow.Title = LocalizationHelper.GetString("Settings", "設定");

                // TitleBarのタイトルを更新
                if (settingsWindow.TitleBar != null)
                {
                    settingsWindow.TitleBar.Title = LocalizationHelper.GetString("Settings", "設定");
                }

                // ナビゲーションボタンのContentを更新
                UpdateButtonContent(settingsWindow, "GeneralButton", "General", "全般");
                UpdateButtonContent(settingsWindow, "AppearanceButton", "Appearance", "外観");
                UpdateButtonContent(settingsWindow, "LayoutButton", "Layout", "レイアウト");
                UpdateButtonContent(settingsWindow, "AboutButton", "AboutApplication", "FastExplorer について");

                // 設定画面のすべてのTextBlockとComboBoxItemを更新
                UpdateSettingsPageTexts(settingsWindow);
                
                System.Diagnostics.Debug.WriteLine("Settings window texts updated successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"設定ウィンドウテキスト更新エラー: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 設定ページのテキストを更新します
        /// </summary>
        /// <param name="settingsWindow">設定ウィンドウ</param>
        private static void UpdateSettingsPageTexts(Views.Windows.SettingsWindow settingsWindow)
        {
            try
            {
                // すべてのTextBlockを検索して更新
                UpdateTextBlocksRecursive(settingsWindow);
                
                // すべてのComboBoxItemを検索して更新
                UpdateComboBoxItemsRecursive(settingsWindow);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"設定ページテキスト更新エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 再帰的にTextBlockを検索して更新します
        /// </summary>
        /// <param name="element">開始要素</param>
        private static void UpdateTextBlocksRecursive(System.Windows.DependencyObject element)
        {
            if (element == null)
                return;

            // 重要なUI要素（TabControl、ListViewなど）は完全にスキップ
            if (element is System.Windows.Controls.TabControl || 
                element is System.Windows.Controls.ListView ||
                element is System.Windows.Controls.ListViewItem ||
                element is System.Windows.Controls.TabItem ||
                element is System.Windows.Controls.Primitives.TabPanel ||
                element is System.Windows.Controls.GridView ||
                element is System.Windows.Controls.GridViewColumn ||
                element is System.Windows.Controls.GridViewColumnHeader)
            {
                return;
            }

            // TextBlockの場合、Textプロパティを更新
            if (element is System.Windows.Controls.TextBlock textBlock)
            {
                // バインディングがない場合（MarkupExtensionを使用している場合）、明示的に更新
                var textBinding = System.Windows.Data.BindingOperations.GetBinding(textBlock, System.Windows.Controls.TextBlock.TextProperty);
                if (textBinding == null)
                {
                    // 現在のテキストからキーを推測して更新
                    var currentText = textBlock.Text;
                    if (!string.IsNullOrEmpty(currentText))
                    {
                        UpdateTextBlockByText(textBlock, currentText);
                    }
                }
                else
                {
                    // バインディングがある場合でも、MarkupExtensionの場合は再評価が必要
                    // Textプロパティを無効化して再評価を強制
                    textBlock.InvalidateProperty(System.Windows.Controls.TextBlock.TextProperty);
                }
            }

            // 子要素を再帰的に処理
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(element);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(element, i);
                UpdateTextBlocksRecursive(child);
            }
        }

        /// <summary>
        /// テキストからキーを推測してTextBlockを更新します
        /// </summary>
        /// <param name="textBlock">TextBlock</param>
        /// <param name="currentText">現在のテキスト</param>
        private static void UpdateTextBlockByText(System.Windows.Controls.TextBlock textBlock, string currentText)
        {
            // 日本語のテキストからキーを推測
            string? key = null;
            if (currentText == "全般" || currentText == "General")
                key = "General";
            else if (currentText == "言語" || currentText == "Language")
                key = "Language";
            else if (currentText == "日付の形式" || currentText == "Date format")
                key = "DateFormat";
            else if (currentText == "例: 5秒前, 2020年12月31日" || 
                     currentText == "Example: 5 seconds ago, December 31, 2020" ||
                     currentText.StartsWith("例:") || 
                     currentText.StartsWith("Example:"))
                key = "DateFormatExample";
            else if (currentText == "起動時の設定" || currentText == "Startup settings")
                key = "StartupSettings";
            else if (currentText == "外観" || currentText == "Appearance")
                key = "Appearance";
            else if (currentText == "レイアウト" || currentText == "Layout")
                key = "Layout";
            else if (currentText == "操作" || currentText == "Operation")
                key = "Operation";
            else if (currentText == "タグ" || currentText == "Tags")
                key = "Tags";
            else if (currentText == "開発者向けツール" || currentText == "Developer Tools")
                key = "DeveloperTools";
            else if (currentText == "高度な設定" || currentText == "Advanced Settings")
                key = "AdvancedSettings";
            else if (currentText == "FastExplorer について" || currentText == "About FastExplorer")
                key = "AboutApplication";

            if (key != null)
            {
                textBlock.Text = LocalizationHelper.GetString(key, currentText);
            }
        }

        /// <summary>
        /// 再帰的にComboBoxItemを検索して更新します
        /// </summary>
        /// <param name="element">開始要素</param>
        private static void UpdateComboBoxItemsRecursive(System.Windows.DependencyObject element)
        {
            if (element == null)
                return;

            // ComboBoxItemの場合、Contentプロパティを更新
            if (element is System.Windows.Controls.ComboBoxItem comboBoxItem)
            {
                var contentBinding = System.Windows.Data.BindingOperations.GetBinding(comboBoxItem, System.Windows.Controls.ComboBoxItem.ContentProperty);
                if (contentBinding == null)
                {
                    var currentContent = comboBoxItem.Content?.ToString();
                    if (currentContent != null)
                    {
                        UpdateComboBoxItemByContent(comboBoxItem, currentContent);
                    }
                }
            }

            // 子要素を再帰的に処理
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(element);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(element, i);
                UpdateComboBoxItemsRecursive(child);
            }
        }

        /// <summary>
        /// コンテンツからキーを推測してComboBoxItemを更新します
        /// </summary>
        /// <param name="comboBoxItem">ComboBoxItem</param>
        /// <param name="currentContent">現在のコンテンツ</param>
        private static void UpdateComboBoxItemByContent(System.Windows.Controls.ComboBoxItem comboBoxItem, string currentContent)
        {
            string? key = null;
            if (currentContent == "アプリ" || currentContent == "App")
                key = "App";
            else if (currentContent == "最後の操作を続行" || currentContent == "Continue last operation")
                key = "ContinueLastOperation";

            if (key != null)
            {
                comboBoxItem.Content = LocalizationHelper.GetString(key, currentContent);
            }
        }

        /// <summary>
        /// ボタンのContentを更新します
        /// </summary>
        /// <param name="window">ウィンドウ</param>
        /// <param name="buttonName">ボタン名</param>
        /// <param name="key">翻訳キー</param>
        /// <param name="defaultValue">デフォルト値</param>
        private static void UpdateButtonContent(System.Windows.Window window, string buttonName, string key, string defaultValue)
        {
            try
            {
                var button = window.FindName(buttonName) as System.Windows.Controls.Button;
                if (button != null)
                {
                    button.Content = LocalizationHelper.GetString(key, defaultValue);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ボタンContent更新エラー ({buttonName}): {ex.Message}");
            }
        }

        /// <summary>
        /// ウィンドウ内のLocalizationHelperExtensionを使用している要素を更新します
        /// </summary>
        /// <param name="element">開始要素</param>
        private static void UpdateLocalizedElements(System.Windows.DependencyObject element)
        {
            if (element == null)
                return;

            // GridViewColumnのHeaderを更新（ListViewの列ヘッダー）
            if (element is System.Windows.Controls.GridViewColumn gridViewColumn)
            {
                if (gridViewColumn.Header is string headerText)
                {
                    // 列ヘッダーのテキストからキーを推測して更新
                    string? key = null;
                    if (headerText == "名前" || headerText == "Name")
                        key = "Name";
                    else if (headerText == "サイズ" || headerText == "Size")
                        key = "Size";
                    else if (headerText == "種類" || headerText == "Type")
                        key = "Type";
                    else if (headerText == "更新日時" || headerText == "Last Modified")
                        key = "LastModified";

                    if (key != null)
                    {
                        gridViewColumn.Header = LocalizationHelper.GetString(key, headerText);
                    }
                }
                // GridViewColumnは子要素を持たないため、ここで終了
                return;
            }

            // 重要なUI要素（TabControl、ListViewなど）は完全にスキップ（子要素も処理しない）
            if (element is System.Windows.Controls.TabControl || 
                element is System.Windows.Controls.ListView ||
                element is System.Windows.Controls.ListViewItem ||
                element is System.Windows.Controls.TabItem ||
                element is System.Windows.Controls.Primitives.TabPanel ||
                element is System.Windows.Controls.GridView ||
                element is System.Windows.Controls.GridViewColumnHeader)
            {
                // これらの要素は完全にスキップ（タブやListViewが消えるのを防ぐ）
                // ただし、GridViewColumnは上で処理済み
                return;
            }

            // TextBlockの場合、Textプロパティを更新
            if (element is System.Windows.Controls.TextBlock textBlock)
            {
                // TextプロパティがLocalizationHelperExtensionで設定されている場合、明示的に再設定する
                var textBinding = System.Windows.Data.BindingOperations.GetBinding(textBlock, System.Windows.Controls.TextBlock.TextProperty);
                if (textBinding == null)
                {
                    // バインディングがない場合、MarkupExtensionを再評価するためにプロパティを明示的に再設定
                    var currentText = textBlock.Text;
                    // 一時的に空にしてから、無効化することで再評価を強制
                    textBlock.Text = string.Empty;
                    textBlock.InvalidateProperty(System.Windows.Controls.TextBlock.TextProperty);
                    // 元の値に戻す（MarkupExtensionが再評価される）
                    textBlock.Text = currentText;
                }
            }
            // ContentControlの場合、Contentプロパティを更新
            else if (element is System.Windows.Controls.ContentControl contentControl && !(element is System.Windows.Controls.TabItem))
            {
                var contentBinding = System.Windows.Data.BindingOperations.GetBinding(contentControl, System.Windows.Controls.ContentControl.ContentProperty);
                if (contentBinding == null)
                {
                    // バインディングがない場合、MarkupExtensionを再評価するためにプロパティを明示的に再設定
                    var currentContent = contentControl.Content;
                    contentControl.Content = null;
                    contentControl.InvalidateProperty(System.Windows.Controls.ContentControl.ContentProperty);
                    contentControl.Content = currentContent;
                }
            }
            // Buttonの場合、Contentプロパティを更新
            else if (element is System.Windows.Controls.Button button)
            {
                var contentBinding = System.Windows.Data.BindingOperations.GetBinding(button, System.Windows.Controls.Button.ContentProperty);
                if (contentBinding == null)
                {
                    var currentContent = button.Content;
                    button.Content = null;
                    button.InvalidateProperty(System.Windows.Controls.Button.ContentProperty);
                    button.Content = currentContent;
                }
            }

            // 子要素を再帰的に処理
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(element);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(element, i);
                UpdateLocalizedElements(child);
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

            // 言語リストを初期化
            InitializeLanguages();

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
        /// 言語リストを初期化します
        /// </summary>
        private void InitializeLanguages()
        {
            // 無限ループを防ぐため、_isLoadingSettingsフラグを設定
            _isLoadingSettings = true;
            try
            {
                AvailableLanguages.Clear();

                // 利用可能な言語を追加
                // 言語名称は直接読み込む（CurrentLanguageCodeを変更しない）
                var jaName = LoadLanguageName("ja", "日本語");
                var jaLanguage = new LanguageItem("ja", jaName);
                System.Diagnostics.Debug.WriteLine($"日本語の言語名称: {jaName}");
                AvailableLanguages.Add(jaLanguage);

                var enName = LoadLanguageName("en", "English");
                var enLanguage = new LanguageItem("en", enName);
                System.Diagnostics.Debug.WriteLine($"英語の言語名称: {enName}");
                AvailableLanguages.Add(enLanguage);

                System.Diagnostics.Debug.WriteLine($"利用可能な言語数: {AvailableLanguages.Count}");
                foreach (var lang in AvailableLanguages)
                {
                    System.Diagnostics.Debug.WriteLine($"  - {lang.Code}: {lang.Name}");
                }

                // 現在の言語を設定から読み込んで選択
                var settings = _windowSettingsService.GetSettings();
                var currentLanguageCode = settings.LanguageCode ?? "ja";
                var selectedLang = AvailableLanguages.FirstOrDefault(l => l.Code == currentLanguageCode) ?? AvailableLanguages.First();
                
                // SelectedLanguageを設定する（_isLoadingSettingsがtrueの間はOnSelectedLanguageChangedが実行されない）
                SelectedLanguage = selectedLang;
                System.Diagnostics.Debug.WriteLine($"選択された言語: {SelectedLanguage?.Code} - {SelectedLanguage?.Name}");
            }
            finally
            {
                _isLoadingSettings = false;
            }
        }

        /// <summary>
        /// 指定された言語コードの言語名称を読み込みます（CurrentLanguageCodeを変更せずに）
        /// </summary>
        /// <param name="languageCode">言語コード</param>
        /// <param name="defaultName">デフォルト名称</param>
        /// <returns>言語名称</returns>
        private string LoadLanguageName(string languageCode, string defaultName)
        {
            try
            {
                // リソース名の候補を試す
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var resourceNames = new[]
                {
                    $"FastExplorer.lang.{languageCode}.json",
                    $"FastExplorer.lang\\{languageCode}.json",
                    $"FastExplorer.lang/{languageCode}.json"
                };

                foreach (var resourceName in resourceNames)
                {
                    var resourceStream = assembly.GetManifestResourceStream(resourceName);
                    if (resourceStream != null)
                    {
                        using var reader = new StreamReader(resourceStream);
                        var json = reader.ReadToEnd();
                        var languageDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                        if (languageDict != null && languageDict.TryGetValue("LanguageName", out var name) && !string.IsNullOrEmpty(name))
                        {
                            System.Diagnostics.Debug.WriteLine($"リソースから言語名称を読み込みました ({languageCode}): {name} (リソース名: {resourceName})");
                            return name;
                        }
                    }
                }

                // リソースから読み込めない場合は、ファイルシステムから読み込む
                var assemblyLocation = assembly.Location;
                var assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
                if (string.IsNullOrEmpty(assemblyDirectory))
                {
                    assemblyDirectory = AppDomain.CurrentDomain.BaseDirectory;
                }

                var langFilePath = Path.Combine(assemblyDirectory, "lang", $"{languageCode}.json");
                
                if (File.Exists(langFilePath))
                {
                    var json = File.ReadAllText(langFilePath);
                    var languageDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                    if (languageDict != null && languageDict.TryGetValue("LanguageName", out var name) && !string.IsNullOrEmpty(name))
                    {
                        System.Diagnostics.Debug.WriteLine($"ファイルから言語名称を読み込みました ({languageCode}): {name} (ファイルパス: {langFilePath})");
                        return name;
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"言語ファイルが見つかりません: {langFilePath}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"言語名称の読み込みエラー ({languageCode}): {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"スタックトレース: {ex.StackTrace}");
            }

            System.Diagnostics.Debug.WriteLine($"デフォルト名称を使用します ({languageCode}): {defaultName}");
            return defaultName;
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
                            }), System.Windows.Threading.DispatcherPriority.Background);
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
                // 優先度をBackgroundに変更して、UIの応答性を維持
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    ThemedSvgIcon.RefreshAllInstances();
                }), System.Windows.Threading.DispatcherPriority.Background);
                
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
                // 優先度をBackgroundに変更して、UIの応答性を維持
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
                }), System.Windows.Threading.DispatcherPriority.Background);
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
