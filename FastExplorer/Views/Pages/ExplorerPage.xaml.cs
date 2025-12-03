using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Runtime.InteropServices;
using FastExplorer.ViewModels.Pages;
using FastExplorer.Services;
using FastExplorer.Models;
using FastExplorer.Helpers;
using FastExplorer.ShellContextMenu;
using Wpf.Ui.Abstractions.Controls;
using Wpf.Ui.Controls;
using Wpf.Ui.Appearance;
using ListView = Wpf.Ui.Controls.ListView;
using ListViewItem = Wpf.Ui.Controls.ListViewItem;
using Button = Wpf.Ui.Controls.Button;
using TextBlock = Wpf.Ui.Controls.TextBlock;

namespace FastExplorer.Views.Pages
{
    /// <summary>
    /// エクスプローラーページを表すクラス
    /// </summary>
    public partial class ExplorerPage : UserControl, INavigableView<ExplorerPageViewModel>
    {
        /// <summary>
        /// エクスプローラーページのViewModelを取得します
        /// </summary>
        public ExplorerPageViewModel ViewModel { get; }

        // パフォーマンス最適化用のキャッシュ
        private System.Windows.Controls.ListView? _cachedLeftListView;
        private System.Windows.Controls.ListView? _cachedRightListView;
        private System.Windows.Controls.ListView? _cachedSinglePaneListView;
        private Brush? _cachedFocusedBackground;
        private Brush? _cachedUnfocusedBackground;
        private Services.WindowSettingsService? _cachedWindowSettingsService;
        private Brush? _cachedUnfocusedPaneBackgroundBrush;
        private Brush? _cachedControlFillColorDefaultBrush;
        // 前回のテーマ状態を追跡（テーマまたはテーマカラーが変更された場合のみ背景色キャッシュをクリア）
        private bool? _lastCachedIsDark = null;
        // 前回のテーマカラーを追跡（テーマカラーが変更された場合にキャッシュをクリア）
        private string? _lastCachedThemeColorCode = null;
        private string? _lastCachedSecondaryColorCode = null;
        private string? _lastCachedThirdColorCode = null;
        // TabControlのキャッシュ（ビジュアルツリー走査を削減）
        private System.Windows.Controls.TabControl? _cachedLeftTabControl;
        private System.Windows.Controls.TabControl? _cachedRightTabControl;
        private System.Windows.Controls.TabControl? _cachedSingleTabControl;

        // 選択タブのViewModelのPropertyChangedイベントハンドラーを追跡
        private System.ComponentModel.PropertyChangedEventHandler? _selectedTabViewModelPropertyChangedHandler;
        // 以前の選択タブを追跡
        private Models.ExplorerTab? _previousSelectedTab = null;
        // 分割ペイン用：左右のペインのタブのViewModelのPropertyChangedイベントハンドラーを追跡
        private System.ComponentModel.PropertyChangedEventHandler? _selectedLeftPaneTabViewModelPropertyChangedHandler;
        private System.ComponentModel.PropertyChangedEventHandler? _selectedRightPaneTabViewModelPropertyChangedHandler;
        // 以前の選択タブを追跡
        private Models.ExplorerTab? _previousSelectedLeftPaneTab = null;
        private Models.ExplorerTab? _previousSelectedRightPaneTab = null;

        /// <summary>
        /// <see cref="ExplorerPage"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">エクスプローラーページのViewModel</param>
        public ExplorerPage(ExplorerPageViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;

            InitializeComponent();

            // ActivePaneの変更を監視して、ListViewの背景色を更新
            ViewModel.PropertyChanged += ViewModel_PropertyChanged;

            // 選択タブの変更を監視して、タブのViewModelのIsHomePageプロパティの変更を監視
            ViewModel.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(ExplorerPageViewModel.SelectedLeftPaneTab) ||
                    e.PropertyName == nameof(ExplorerPageViewModel.SelectedRightPaneTab) ||
                    e.PropertyName == nameof(ExplorerPageViewModel.SelectedTab))
                {
                    SubscribeToSelectedTabViewModel();
                }
            };

            Loaded += ExplorerPage_Loaded;

            // 初期の選択タブのViewModelのPropertyChangedイベントを購読
            SubscribeToSelectedTabViewModel();
        }

        /// <summary>
        /// 現在の選択タブのViewModelのPropertyChangedイベントを購読します
        /// </summary>
        private void SubscribeToSelectedTabViewModel()
        {
            if (ViewModel.IsSplitPaneEnabled)
            {
                // 分割ペインの場合、左右両方のペインのタブのViewModelのPropertyChangedイベントを購読

                // 左ペインのタブのViewModelのPropertyChangedイベントを購読
                if (_previousSelectedLeftPaneTab?.ViewModel != null && _selectedLeftPaneTabViewModelPropertyChangedHandler != null)
                {
                    _previousSelectedLeftPaneTab.ViewModel.PropertyChanged -= _selectedLeftPaneTabViewModelPropertyChangedHandler;
                }

                var newLeftTab = ViewModel.SelectedLeftPaneTab;
                _previousSelectedLeftPaneTab = newLeftTab;

                if (newLeftTab?.ViewModel != null)
                {
                    _selectedLeftPaneTabViewModelPropertyChangedHandler = (s2, e2) =>
                    {
                        HandleTabViewModelPropertyChanged(newLeftTab, e2);
                    };
                    newLeftTab.ViewModel.PropertyChanged += _selectedLeftPaneTabViewModelPropertyChangedHandler;
                }

                // 右ペインのタブのViewModelのPropertyChangedイベントを購読
                if (_previousSelectedRightPaneTab?.ViewModel != null && _selectedRightPaneTabViewModelPropertyChangedHandler != null)
                {
                    _previousSelectedRightPaneTab.ViewModel.PropertyChanged -= _selectedRightPaneTabViewModelPropertyChangedHandler;
                }

                var newRightTab = ViewModel.SelectedRightPaneTab;
                _previousSelectedRightPaneTab = newRightTab;

                if (newRightTab?.ViewModel != null)
                {
                    _selectedRightPaneTabViewModelPropertyChangedHandler = (s2, e2) =>
                    {
                        HandleTabViewModelPropertyChanged(newRightTab, e2);
                    };
                    newRightTab.ViewModel.PropertyChanged += _selectedRightPaneTabViewModelPropertyChangedHandler;
                }
            }
            else
            {
                // 単一ペインの場合、選択タブのViewModelのPropertyChangedイベントを購読
                if (_previousSelectedTab?.ViewModel != null && _selectedTabViewModelPropertyChangedHandler != null)
                {
                    _previousSelectedTab.ViewModel.PropertyChanged -= _selectedTabViewModelPropertyChangedHandler;
                }

                var newTab = ViewModel.SelectedTab;
                _previousSelectedTab = newTab;

                if (newTab?.ViewModel != null)
                {
                    _selectedTabViewModelPropertyChangedHandler = (s2, e2) =>
                    {
                        HandleTabViewModelPropertyChanged(newTab, e2);
                    };
                    newTab.ViewModel.PropertyChanged += _selectedTabViewModelPropertyChangedHandler;
                }
            }
        }

        /// <summary>
        /// タブのViewModelのPropertyChangedイベントを処理します
        /// </summary>
        private void HandleTabViewModelPropertyChanged(Models.ExplorerTab newTab, System.ComponentModel.PropertyChangedEventArgs e2)
        {
            if (newTab?.ViewModel == null)
                return;

            if (e2.PropertyName == "IsHomePage" || e2.PropertyName == "CurrentPath")
            {
                // IsHomePageプロパティが変更された場合、背景色を更新
                if (e2.PropertyName == "IsHomePage")
                {
                    // IsHomePageがfalseになった場合（ホームタブから通常のディレクトリビューに切り替わった場合）
                    // そのタブが属するペインをアクティブにする
                    if (ViewModel.IsSplitPaneEnabled && newTab != null && newTab.ViewModel != null)
                    {
                        // IsHomePageがfalseになった場合のみアクティブにする
                        if (!newTab.ViewModel.IsHomePage)
                        {
                            // タブがどのペインに属しているかを判定
                            int? targetPane = null;
                            if (ViewModel.SelectedLeftPaneTab == newTab)
                            {
                                // 左ペインのタブがホームから通常のディレクトリビューに切り替わった場合
                                targetPane = 0;
                            }
                            else if (ViewModel.SelectedRightPaneTab == newTab)
                            {
                                // 右ペインのタブがホームから通常のディレクトリビューに切り替わった場合
                                targetPane = 2;
                            }

                            if (targetPane.HasValue)
                            {
                                // アクティブペインを設定
                                ViewModel.ActivePane = targetPane.Value;
                            }
                        }
                    }

                    // ListViewのキャッシュをクリアして、背景色を更新
                    _cachedLeftListView = null;
                    _cachedRightListView = null;
                    _cachedSinglePaneListView = null;

                    // UI構築完了後に背景色を更新（IsHomePageが変更された後、ListViewが表示されるまで少し時間がかかる可能性があるため）
                    UpdateListViewBackgroundColorsDelayed();
                }
                // CurrentPathプロパティが変更された場合（ホームタブでドライブをクリックした場合など）
                else if (e2.PropertyName == "CurrentPath")
                {
                    // ホームタブの状態でドライブをクリックした場合、そのタブが属するペインをアクティブにする
                    if (ViewModel.IsSplitPaneEnabled && newTab != null && newTab.ViewModel != null)
                    {
                        // CurrentPathが設定されている場合
                        if (!string.IsNullOrEmpty(newTab.ViewModel.CurrentPath))
                        {
                            // タブがどのペインに属しているかを判定
                            int? targetPane = null;
                            if (ViewModel.SelectedLeftPaneTab == newTab)
                            {
                                targetPane = 0;
                            }
                            else if (ViewModel.SelectedRightPaneTab == newTab)
                            {
                                targetPane = 2;
                            }

                            if (targetPane.HasValue)
                            {
                                // アクティブペインを設定
                                ViewModel.ActivePane = targetPane.Value;

                                // IsHomePageがfalseの場合（既にListViewが表示されている場合）、すぐに背景色を更新
                                if (!newTab.ViewModel.IsHomePage)
                                {
                                    _cachedLeftListView = null;
                                    _cachedRightListView = null;
                                    _cachedSinglePaneListView = null;
                                    UpdateListViewBackgroundColorsDelayed();
                                }
                                // IsHomePageがtrueの場合（ホームタブの状態）、IsHomePageがfalseになるまで待つ
                                // IsHomePageがfalseになったときに背景色が更新される（IsHomePage変更時の処理で対応）
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// ページが読み込まれたときに呼び出されます
        /// </summary>
        private void ExplorerPage_Loaded(object sender, RoutedEventArgs e)
        {
            // 初期状態の背景色を更新（UI構築完了後に実行）
            UpdateListViewBackgroundColorsDelayed();

            // 初期の選択タブのViewModelのPropertyChangedイベントを購読（Loaded後に実行）
            SubscribeToSelectedTabViewModel();
        }

        /// <summary>
        /// ViewModelのプロパティが変更されたときに呼び出されます
        /// </summary>
        private void ViewModel_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            // プロパティ名をキャッシュ（高速化）
            var propertyName = e.PropertyName;
            if (propertyName == nameof(ExplorerPageViewModel.ActivePane))
            {
                // ActivePaneが変更された場合、ListViewのキャッシュをクリアして、背景色を更新
                // キャッシュをクリアすることで、確実に最新のListViewを取得できるようにする
                _cachedLeftListView = null;
                _cachedRightListView = null;
                _cachedSinglePaneListView = null;

                // UIスレッドで背景色を更新
                // Dispatcher.BeginInvokeを使用して、UI更新後に確実に実行されるようにする
                Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    UpdateListViewBackgroundColors();
                }), System.Windows.Threading.DispatcherPriority.Loaded);

                // さらに、Render優先度でも実行（ListViewが表示された後に確実に更新されるようにする）
                Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    UpdateListViewBackgroundColors();
                }), System.Windows.Threading.DispatcherPriority.Render);
            }
            else if (propertyName == nameof(ExplorerPageViewModel.IsSplitPaneEnabled))
            {
                // 分割ペインの有効/無効が変更された場合はキャッシュをクリア
                _cachedLeftListView = null;
                _cachedRightListView = null;
                _cachedSinglePaneListView = null;
                _cachedLeftTabControl = null;
                _cachedRightTabControl = null;
                _cachedSingleTabControl = null;
                // ペインキャッシュもクリア（UI構造が変わるため）
                _paneCache.Clear();
                // UI構築完了後に背景色を更新（分割ペイン切り替え時）
                Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    UpdateListViewBackgroundColors();
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }


        /// <summary>
        /// ListViewの背景色を遅延更新します（複数の優先度で実行して確実に更新）
        /// </summary>
        private void UpdateListViewBackgroundColorsDelayed()
        {
            // Loaded優先度で実行
            Dispatcher.BeginInvoke(new System.Action(() =>
            {
                _cachedLeftListView = null;
                _cachedRightListView = null;
                _cachedSinglePaneListView = null;
                UpdateListViewBackgroundColors();
            }), System.Windows.Threading.DispatcherPriority.Loaded);

            // Render優先度でも実行（ListViewが表示された後に確実に更新されるようにする）
            Dispatcher.BeginInvoke(new System.Action(() =>
            {
                UpdateListViewBackgroundColors();
            }), System.Windows.Threading.DispatcherPriority.Render);
        }

        /// <summary>
        /// 左右のペインのListViewの背景色を更新します
        /// </summary>
        private void UpdateListViewBackgroundColors()
        {
            // 現在のテーマを確認（ダークモードかどうか）
            var isDark = ApplicationThemeManager.GetAppTheme() == ApplicationTheme.Dark;

            // WindowSettingsServiceをキャッシュ（パフォーマンス向上）
            if (_cachedWindowSettingsService == null)
            {
                _cachedWindowSettingsService = App.Services.GetService(typeof(Services.WindowSettingsService)) as Services.WindowSettingsService;
            }

            // 現在のテーマカラーを取得
            string? currentThemeColorCode = null;
            string? currentSecondaryColorCode = null;
            string? currentThirdColorCode = null;
            if (_cachedWindowSettingsService != null)
            {
                var settings = _cachedWindowSettingsService.GetSettings();
                currentThemeColorCode = settings.ThemeColorCode;
                currentSecondaryColorCode = settings.ThemeSecondaryColorCode;
                currentThirdColorCode = settings.ThemeThirdColorCode;
            }

            // テーマまたはテーマカラーが変更された場合、または初期表示時のみ背景色のキャッシュをクリア
            // 左右のリストビューをクリックしてフォーカスが移動しただけではキャッシュをクリアしない
            // これにより、テーマカラーが消える問題を防ぐ
            bool themeChanged = _lastCachedIsDark == null || _lastCachedIsDark.Value != isDark;
            bool themeColorChanged = _lastCachedThemeColorCode != currentThemeColorCode ||
                                     _lastCachedSecondaryColorCode != currentSecondaryColorCode ||
                                     _lastCachedThirdColorCode != currentThirdColorCode;

            if (themeChanged || themeColorChanged)
            {
                _lastCachedIsDark = isDark;
                _lastCachedThemeColorCode = currentThemeColorCode;
                _lastCachedSecondaryColorCode = currentSecondaryColorCode;
                _lastCachedThirdColorCode = currentThirdColorCode;
                _cachedFocusedBackground = null;
                _cachedUnfocusedBackground = null;
                _cachedControlFillColorDefaultBrush = null;
                _cachedUnfocusedPaneBackgroundBrush = null;
            }

            // フォーカスがないペインの背景色を取得（最適化：キャッシュを活用）
            // キャッシュが既に存在する場合は再計算をスキップ（パフォーマンス向上＆テーマカラー消失防止）
            if (_cachedUnfocusedBackground == null)
            {
                try
                {
                    if (isDark)
                    {
                        // ダークモードの場合は、ダークモード用の標準カラーを優先
                        var unfocusedPaneColor = Color.FromRgb(0x2D, 0x2D, 0x30); // #2D2D30
                        var unfocusedPaneBrush = new SolidColorBrush(unfocusedPaneColor);
                        unfocusedPaneBrush.Freeze();
                        _cachedUnfocusedBackground = unfocusedPaneBrush;
                    }
                    else
                    {
                        // ライトモードの場合は、ThirdColorCodeから取得
                        if (_cachedWindowSettingsService != null)
                        {
                            var settings = _cachedWindowSettingsService.GetSettings();
                            var thirdColorCode = settings.ThemeThirdColorCode;
                            if (!string.IsNullOrEmpty(thirdColorCode))
                            {
                                // ThirdColorCodeが設定されている場合はそれを使用
                                var thirdColor = Helpers.FastColorConverter.ParseHexColor(thirdColorCode);
                                var thirdColorBrush = new SolidColorBrush(thirdColor);
                                thirdColorBrush.Freeze();
                                _cachedUnfocusedBackground = thirdColorBrush;
                            }
                            else
                            {
                                // ThirdColorCodeが設定されていない場合はUnfocusedPaneBackgroundBrushリソースを使用（キャッシュを活用）
                                if (_cachedUnfocusedPaneBackgroundBrush == null)
                                {
                                    _cachedUnfocusedPaneBackgroundBrush = FindResource("UnfocusedPaneBackgroundBrush") as Brush;
                                }
                                _cachedUnfocusedBackground = _cachedUnfocusedPaneBackgroundBrush ??
                                    CreateDefaultUnfocusedBrush();
                            }
                        }
                        else
                        {
                            // WindowSettingsServiceが取得できない場合はUnfocusedPaneBackgroundBrushリソースを使用（キャッシュを活用）
                            if (_cachedUnfocusedPaneBackgroundBrush == null)
                            {
                                _cachedUnfocusedPaneBackgroundBrush = FindResource("UnfocusedPaneBackgroundBrush") as Brush;
                            }
                            _cachedUnfocusedBackground = _cachedUnfocusedPaneBackgroundBrush ??
                                CreateDefaultUnfocusedBrush();
                        }
                    }
                }
                catch
                {
                    // エラーが発生した場合はデフォルト値を使用
                    _cachedUnfocusedBackground = CreateDefaultUnfocusedBrush();
                }
            }

            // フォーカスがあるペインの背景色を取得（最適化：キャッシュを活用）
            // キャッシュが既に存在する場合は再計算をスキップ（パフォーマンス向上＆テーマカラー消失防止）
            if (_cachedFocusedBackground == null)
            {
                try
                {
                    if (isDark)
                    {
                        // ダークモードの場合は、ダークモード用の標準カラーを優先
                        var focusedPaneColor = Color.FromRgb(0x25, 0x25, 0x26); // #252526
                        var focusedPaneBrush = new SolidColorBrush(focusedPaneColor);
                        focusedPaneBrush.Freeze();
                        _cachedFocusedBackground = focusedPaneBrush;
                    }
                    else
                    {
                        // ライトモードの場合は、SecondaryColorCodeから取得
                        if (_cachedWindowSettingsService != null)
                        {
                            var settings = _cachedWindowSettingsService.GetSettings();
                            var secondaryColorCode = settings.ThemeSecondaryColorCode;
                            if (!string.IsNullOrEmpty(secondaryColorCode))
                            {
                                // SecondaryColorCodeが設定されている場合はそれを使用
                                var secondaryColor = Helpers.FastColorConverter.ParseHexColor(secondaryColorCode);
                                var secondaryColorBrush = new SolidColorBrush(secondaryColor);
                                secondaryColorBrush.Freeze();
                                _cachedFocusedBackground = secondaryColorBrush;
                            }
                            else
                            {
                                // SecondaryColorCodeが設定されていない場合はControlFillColorDefaultBrushリソースを使用（キャッシュを活用）
                                if (_cachedControlFillColorDefaultBrush == null)
                                {
                                    _cachedControlFillColorDefaultBrush = FindResource("ControlFillColorDefaultBrush") as Brush;
                                }
                                _cachedFocusedBackground = _cachedControlFillColorDefaultBrush ??
                                    CreateDefaultFocusedBrush();
                            }
                        }
                        else
                        {
                            // WindowSettingsServiceが取得できない場合はControlFillColorDefaultBrushリソースを使用（キャッシュを活用）
                            if (_cachedControlFillColorDefaultBrush == null)
                            {
                                _cachedControlFillColorDefaultBrush = FindResource("ControlFillColorDefaultBrush") as Brush;
                            }
                            _cachedFocusedBackground = _cachedControlFillColorDefaultBrush ??
                                CreateDefaultFocusedBrush();
                        }
                    }
                }
                catch
                {
                    // エラーが発生した場合はデフォルト値を使用
                    _cachedFocusedBackground = CreateDefaultFocusedBrush();
                }
            }

            // ViewModelプロパティをキャッシュ（パフォーマンス向上）
            var viewModel = ViewModel;
            var isSplitPaneEnabled = viewModel.IsSplitPaneEnabled;

            if (isSplitPaneEnabled)
            {
                // 分割ペインモードの場合
                // ListViewの参照を取得またはキャッシュから取得（最適化：一度だけ検索）
                // キャッシュがnullの場合は再検索（IsHomePageが変更された後など）
                if (_cachedLeftListView == null)
                {
                    _cachedLeftListView = FindListViewInPane(0);
                }
                if (_cachedRightListView == null)
                {
                    _cachedRightListView = FindListViewInPane(2);
                }

                // 背景色を更新
                // ActivePaneが-1（未設定）の場合は、左ペインをフォーカスあり（SecondaryColorCode）、右ペインをフォーカスなし（ThirdColorCode）にする
                var activePane = viewModel.ActivePane;
                if (activePane == -1)
                {
                    // 起動時や分割直後など、ActivePaneが未設定の場合は左ペインをフォーカスありにする
                    activePane = 0;
                }

                // 背景色を事前に決定（条件分岐を削減、nullチェックも含める）
                // アクティブなペインには薄い色（_cachedFocusedBackground）、非アクティブなペインには濃い色（_cachedUnfocusedBackground）を設定
                var leftBackground = activePane == 0 ? _cachedFocusedBackground : _cachedUnfocusedBackground;
                var rightBackground = activePane == 2 ? _cachedFocusedBackground : _cachedUnfocusedBackground;

                // 左ペインの背景色を更新
                if (_cachedLeftListView == null)
                {
                    _cachedLeftListView = FindListViewInPane(0);
                }
                if (_cachedLeftListView != null && leftBackground != null)
                {
                    _cachedLeftListView.Background = leftBackground;
                }

                // 右ペインの背景色を更新
                if (_cachedRightListView == null)
                {
                    _cachedRightListView = FindListViewInPane(2);
                }
                if (_cachedRightListView != null && rightBackground != null)
                {
                    _cachedRightListView.Background = rightBackground;
                }

                // 分割ペインモードの場合、ホームページScrollViewerの背景色も更新
                // 左ペインと右ペインのそれぞれでホームページScrollViewerを検索
                UpdateHomePageScrollViewerBackground(0, leftBackground);
                UpdateHomePageScrollViewerBackground(2, rightBackground);
            }
            else
            {
                // 単一ペインモードの場合、ListViewの背景色をSecondaryColorCodeに設定
                // キャッシュから取得（高速化）
                if (_cachedSinglePaneListView == null)
                {
                    _cachedSinglePaneListView = FindListViewInSinglePane();
                }

                // 背景色がnullでない場合、強制的に設定（キャッシュが古い可能性があるため）
                if (_cachedSinglePaneListView != null && _cachedFocusedBackground != null)
                {
                    _cachedSinglePaneListView.Background = _cachedFocusedBackground;
                }
                else if (_cachedSinglePaneListView == null && _cachedFocusedBackground != null)
                {
                    // ListViewが見つからない場合は、再検索を試みる
                    _cachedSinglePaneListView = FindListViewInSinglePane();
                    if (_cachedSinglePaneListView != null)
                    {
                        _cachedSinglePaneListView.Background = _cachedFocusedBackground;
                    }
                }

                // 単一ペインモードの場合、ホームページScrollViewerの背景色も更新
                // SinglePaneHomePageScrollViewerはDataTemplate内にあるため、ビジュアルツリーを走査して検索
                var singlePaneTabControl = _cachedSingleTabControl ?? FindChild<System.Windows.Controls.TabControl>(this, null);
                if (singlePaneTabControl != null && _cachedSingleTabControl == null)
                {
                    _cachedSingleTabControl = singlePaneTabControl;
                }
                if (singlePaneTabControl != null)
                {
                    var singlePaneHomePageScrollViewer = FindHomePageScrollViewerInTabControl(singlePaneTabControl);
                    if (singlePaneHomePageScrollViewer != null && _cachedFocusedBackground != null && singlePaneHomePageScrollViewer.Background != _cachedFocusedBackground)
                    {
                        singlePaneHomePageScrollViewer.Background = _cachedFocusedBackground;
                    }
                }
            }
        }

        /// <summary>
        /// 指定されたペインのホームページScrollViewerの背景色を更新します
        /// </summary>
        /// <param name="paneIndex">ペインのインデックス（0=左ペイン、2=右ペイン）</param>
        /// <param name="background">設定する背景色</param>
        private void UpdateHomePageScrollViewerBackground(int paneIndex, Brush? background)
        {
            if (background == null)
                return;

            try
            {
                // 指定されたペインのTabControlを取得
                System.Windows.Controls.TabControl? tabControl = null;
                if (paneIndex == 0)
                {
                    tabControl = _cachedLeftTabControl ?? FindTabControlInPane(0);
                    if (tabControl != null && _cachedLeftTabControl == null)
                    {
                        _cachedLeftTabControl = tabControl;
                    }
                }
                else if (paneIndex == 2)
                {
                    tabControl = _cachedRightTabControl ?? FindTabControlInPane(2);
                    if (tabControl != null && _cachedRightTabControl == null)
                    {
                        _cachedRightTabControl = tabControl;
                    }
                }

                if (tabControl == null)
                    return;

                // TabControl内のホームページScrollViewerを検索
                var scrollViewer = FindHomePageScrollViewerInTabControl(tabControl);
                if (scrollViewer != null && scrollViewer.Background != background)
                {
                    scrollViewer.Background = background;
                }
            }
            catch
            {
                // エラーが発生した場合は無視
            }
        }

        /// <summary>
        /// TabControl内のホームページScrollViewerを検索します
        /// </summary>
        /// <param name="tabControl">検索対象のTabControl</param>
        /// <returns>ホームページScrollViewer、見つからない場合はnull</returns>
        private System.Windows.Controls.ScrollViewer? FindHomePageScrollViewerInTabControl(System.Windows.Controls.TabControl tabControl)
        {
            // ビジュアルツリーを走査してScrollViewerを検索
            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(tabControl); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(tabControl, i);
                var scrollViewer = FindHomePageScrollViewerRecursive(child as System.Windows.DependencyObject);
                if (scrollViewer != null)
                    return scrollViewer;
            }
            return null;
        }

        /// <summary>
        /// 再帰的にホームページScrollViewerを検索します
        /// </summary>
        /// <param name="element">検索開始要素</param>
        /// <returns>ホームページScrollViewer、見つからない場合はnull</returns>
        private System.Windows.Controls.ScrollViewer? FindHomePageScrollViewerRecursive(System.Windows.DependencyObject? element)
        {
            if (element == null)
                return null;

            // ScrollViewerを確認
            if (element is System.Windows.Controls.ScrollViewer scrollViewer)
            {
                // VisibilityがIsHomePageにバインドされているScrollViewerを探す
                // または、名前がSinglePaneHomePageScrollViewerまたはSplitPaneHomePageScrollViewerのScrollViewerを探す
                if (scrollViewer.Name == "SinglePaneHomePageScrollViewer" ||
                    scrollViewer.Name == "SplitPaneHomePageScrollViewer" ||
                    (scrollViewer.Visibility == System.Windows.Visibility.Visible &&
                     BindingOperations.GetBinding(scrollViewer, System.Windows.UIElement.VisibilityProperty) != null))
                {
                    return scrollViewer;
                }
            }

            // 子要素を再帰的に検索
            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(element); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(element, i);
                var result = FindHomePageScrollViewerRecursive(child);
                if (result != null)
                    return result;
            }

            return null;
        }

        /// <summary>
        /// 指定されたペインのTabControlを検索します（キャッシュ付き）
        /// </summary>
        /// <param name="pane">ペイン番号（0=左、2=右）</param>
        /// <returns>TabControl、見つからない場合はnull</returns>
        private System.Windows.Controls.TabControl? FindTabControlInPane(int pane)
        {
            // キャッシュを確認
            if (pane == 0)
            {
                if (_cachedLeftTabControl != null)
                {
                    return _cachedLeftTabControl;
                }
            }
            else if (pane == 2)
            {
                if (_cachedRightTabControl != null)
                {
                    return _cachedRightTabControl;
                }
            }

            // キャッシュにない場合は検索
            // 分割ペインのTabControlを検索（Grid.Columnで判定）
            var tabControl = FindChild<System.Windows.Controls.TabControl>(this, tc =>
            {
                var column = Grid.GetColumn(tc);
                return column == pane;
            });

            // キャッシュに保存
            if (pane == 0)
            {
                _cachedLeftTabControl = tabControl;
            }
            else if (pane == 2)
            {
                _cachedRightTabControl = tabControl;
            }

            return tabControl;
        }

        /// <summary>
        /// 単一ペインモードのListViewを検索します（キャッシュ付き）
        /// </summary>
        /// <returns>ListView、見つからない場合はnull</returns>
        private System.Windows.Controls.ListView? FindListViewInSinglePane()
        {
            // キャッシュを確認
            if (_cachedSinglePaneListView != null)
            {
                return _cachedSinglePaneListView;
            }

            // TabControlのキャッシュを確認
            System.Windows.Controls.TabControl? tabControl = _cachedSingleTabControl;
            if (tabControl == null)
            {
                // 通常モードのTabControlを検索
                tabControl = FindChild<System.Windows.Controls.TabControl>(this, tc =>
                {
                    // 分割ペインのTabControlではないことを確認（Grid.Columnが設定されていない）
                    var column = Grid.GetColumn(tc);
                    return column == 0 && !ViewModel.IsSplitPaneEnabled; // 通常モードのTabControl
                });

                if (tabControl == null)
                {
                    // より広範囲に検索（Grid.Columnが設定されていないTabControlを探す）
                    tabControl = FindChild<System.Windows.Controls.TabControl>(this, null);
                    if (tabControl != null)
                    {
                        // 分割ペインのTabControlでないことを確認
                        var column = Grid.GetColumn(tabControl);
                        if (column != 0 && column != 2)
                        {
                            _cachedSingleTabControl = tabControl;
                        }
                        else
                        {
                            tabControl = null;
                        }
                    }
                }
                else
                {
                    _cachedSingleTabControl = tabControl;
                }
            }

            if (tabControl == null)
                return null;

            // TabControl内のListViewを検索
            var listView = FindChild<System.Windows.Controls.ListView>(tabControl, null);
            _cachedSinglePaneListView = listView;
            return listView;
        }

        /// <summary>
        /// 指定されたペイン内のListViewを検索します（キャッシュ付き）
        /// </summary>
        /// <param name="pane">ペイン番号（0=左、2=右）</param>
        /// <returns>ListView、見つからない場合はnull</returns>
        private System.Windows.Controls.ListView? FindListViewInPane(int pane)
        {
            // キャッシュを確認
            System.Windows.Controls.TabControl? tabControl = null;
            if (pane == 0)
            {
                if (_cachedLeftTabControl != null)
                {
                    tabControl = _cachedLeftTabControl;
                }
                else if (_cachedLeftListView != null)
                {
                    // ListViewから親のTabControlを取得（キャッシュ）
                    var parent = VisualTreeHelper.GetParent(_cachedLeftListView);
                    while (parent != null && !(parent is System.Windows.Controls.TabControl))
                    {
                        parent = VisualTreeHelper.GetParent(parent);
                    }
                    if (parent is System.Windows.Controls.TabControl tc)
                    {
                        _cachedLeftTabControl = tc;
                        tabControl = tc;
                    }
                }
            }
            else if (pane == 2)
            {
                if (_cachedRightTabControl != null)
                {
                    tabControl = _cachedRightTabControl;
                }
                else if (_cachedRightListView != null)
                {
                    // ListViewから親のTabControlを取得（キャッシュ）
                    var parent = VisualTreeHelper.GetParent(_cachedRightListView);
                    while (parent != null && !(parent is System.Windows.Controls.TabControl))
                    {
                        parent = VisualTreeHelper.GetParent(parent);
                    }
                    if (parent is System.Windows.Controls.TabControl tc)
                    {
                        _cachedRightTabControl = tc;
                        tabControl = tc;
                    }
                }
            }

            // キャッシュにない場合は検索
            if (tabControl == null)
            {
                // 分割ペインのTabControlを検索（Grid.Columnで判定）
                tabControl = FindChild<System.Windows.Controls.TabControl>(this, tc =>
                {
                    var column = Grid.GetColumn(tc);
                    return column == pane;
                });

                // キャッシュに保存
                if (pane == 0)
                {
                    _cachedLeftTabControl = tabControl;
                }
                else if (pane == 2)
                {
                    _cachedRightTabControl = tabControl;
                }
            }

            if (tabControl == null)
                return null;

            // TabControl内のListViewを検索（キャッシュを確認）
            if (pane == 0 && _cachedLeftListView != null)
            {
                return _cachedLeftListView;
            }
            if (pane == 2 && _cachedRightListView != null)
            {
                return _cachedRightListView;
            }

            var listView = FindChild<System.Windows.Controls.ListView>(tabControl, null);

            // キャッシュに保存
            if (pane == 0)
            {
                _cachedLeftListView = listView;
            }
            else if (pane == 2)
            {
                _cachedRightListView = listView;
            }

            return listView;
        }

        /// <summary>
        /// 指定された型の子要素を検索します（最適化版：最大深度制限付き）
        /// </summary>
        private T? FindChild<T>(DependencyObject parent, Func<T, bool>? predicate) where T : DependencyObject
        {
            if (parent == null)
                return null;

            return FindChildInternal<T>(parent, predicate, 0, 20); // 最大20階層まで
        }

        /// <summary>
        /// 指定された型の子要素を検索します（内部実装、最大深度制限付き）
        /// </summary>
        private T? FindChildInternal<T>(DependencyObject parent, Func<T, bool>? predicate, int depth, int maxDepth) where T : DependencyObject
        {
            if (parent == null || depth >= maxDepth)
                return null;

            // 子要素の数を一度だけ取得してキャッシュ（パフォーマンス向上）
            var childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is T t && (predicate == null || predicate(t)))
                {
                    return t;
                }

                var result = FindChildInternal<T>(child, predicate, depth + 1, maxDepth);
                if (result != null)
                    return result;
            }

            return null;
        }

        /// <summary>
        /// デフォルトのフォーカスなし背景ブラシを作成します
        /// </summary>
        private static Brush CreateDefaultUnfocusedBrush()
        {
            var brush = new SolidColorBrush(Color.FromRgb(0xFE, 0xEB, 0xEB));
            brush.Freeze();
            return brush;
        }

        /// <summary>
        /// デフォルトのフォーカスあり背景ブラシを作成します
        /// </summary>
        private static Brush CreateDefaultFocusedBrush()
        {
            var brush = new SolidColorBrush(Color.FromArgb(30, 128, 128, 128)); // 薄いグレー
            brush.Freeze();
            return brush;
        }

        /// <summary>
        /// アドレスバーでキーが押されたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">キーイベント引数</param>
        private void AddressBar_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var textBox = sender as System.Windows.Controls.TextBox;
                if (textBox != null && ViewModel.SelectedTab != null)
                {
                    ViewModel.SelectedTab.ViewModel.NavigateToPathCommand.Execute(textBox.Text);
                }
            }
        }

        /// <summary>
        /// リストビューでマウスがダブルクリックされたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void ListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // リネームタイマーをキャンセル
            CancelRenameTimer();
            // リネームをキャンセル
            if (_isRenaming)
            {
                CancelRename();
            }

            // リネーム直後（500ms以内）のダブルクリックは無視
            if ((DateTime.Now - _renameCompletedTime).TotalMilliseconds < 500)
            {
                e.Handled = true;
                return;
            }

            var listView = sender as System.Windows.Controls.ListView;
            if (listView == null)
                return;

            Models.ExplorerTab? targetTab = null;

            if (ViewModel.IsSplitPaneEnabled)
            {
                // 分割ペインモードの場合、クリックされたListViewがどのペインに属しているかを判定
                var pane = GetPaneForElement(listView);
                if (pane == 0)
                {
                    // 左ペイン
                    targetTab = ViewModel.SelectedLeftPaneTab;
                    // ActivePaneを更新
                    ViewModel.ActivePane = 0;
                }
                else if (pane == 2)
                {
                    // 右ペイン
                    targetTab = ViewModel.SelectedRightPaneTab;
                    // ActivePaneを更新
                    ViewModel.ActivePane = 2;
                }
                else
                {
                    // 判定できない場合は、GetActiveTab()を使用
                    targetTab = GetActiveTab();
                }
            }
            else
            {
                // 通常モード
                targetTab = ViewModel.SelectedTab;
            }

            if (targetTab == null)
                return;

            // クリックされた位置がListViewItem上かどうかを確認（ビジュアルツリー走査を最適化）
            System.Windows.Controls.ListViewItem? clickedItem = null;
            if (e.OriginalSource is DependencyObject source)
            {
                DependencyObject? current = source;
                while (current != null && current != listView)
                {
                    // 型チェックを一度に実行（パフォーマンス向上）
                    if (current is System.Windows.Controls.GridViewColumnHeader)
                    {
                        // ヘッダー上でクリックされた場合は処理しない
                        e.Handled = true;
                        return;
                    }
                    if (current is System.Windows.Controls.ListViewItem item)
                    {
                        // ListViewItem上でクリックされた場合は処理を続行
                        clickedItem = item;
                        break;
                    }
                    current = VisualTreeHelper.GetParent(current);
                }
            }

            // ListViewItemが見つからなかった場合は空領域
            if (clickedItem == null)
            {
                // 空領域（アイテムがない場所）をダブルクリックした場合は親フォルダーに移動
                targetTab.ViewModel.NavigateToParentCommand.Execute(null);
                e.Handled = true;
                return;
            }

            // クリックされたアイテムを取得
            var clickedDataItem = clickedItem.DataContext as Models.FileSystemItem;
            if (clickedDataItem == null)
            {
                // DataContextがFileSystemItemでない場合は処理しない
                e.Handled = true;
                return;
            }

            // ディレクトリの場合は、新しいディレクトリに移動した後にスクロール位置を0に戻す
            if (clickedDataItem.IsDirectory)
            {
                targetTab.ViewModel.NavigateToItemCommand.Execute(clickedDataItem);

                // 少し遅延してからスクロール位置を0に戻す（ItemsSourceが更新されるのを待つ）
                Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    if (listView != null)
                    {
                        var scrollViewer = GetScrollViewer(listView);
                        if (scrollViewer != null)
                        {
                            scrollViewer.ScrollToTop();
                        }
                    }
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }
            else
            {
                // ファイルの場合は、ファイルを開く
                targetTab.ViewModel.NavigateToItemCommand.Execute(clickedDataItem);
            }
        }

        /// <summary>
        /// DependencyObjectからScrollViewerを取得します（キャッシュ付き）
        /// </summary>
        /// <param name="element">要素</param>
        /// <returns>ScrollViewer、見つからない場合はnull</returns>
        private System.Windows.Controls.ScrollViewer? GetScrollViewer(System.Windows.DependencyObject element)
        {
            if (element == null)
                return null;

            // ListViewの場合はキャッシュを確認
            System.Windows.Controls.ListView? listView = null;
            if (element is System.Windows.Controls.ListView lv)
            {
                listView = lv;
                if (_scrollViewerCache.TryGetValue(listView, out var cached))
                {
                    return cached;
                }
            }

            if (element is System.Windows.Controls.ScrollViewer scrollViewer)
            {
                // ListViewの場合はキャッシュに保存
                if (listView != null)
                {
                    _scrollViewerCache[listView] = scrollViewer;
                }
                return scrollViewer;
            }

            // 子要素の数を一度だけ取得してキャッシュ（パフォーマンス向上）
            var childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(element);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(element, i);
                var result = GetScrollViewer(child);
                if (result != null)
                {
                    // ListViewの場合はキャッシュに保存
                    if (listView != null)
                    {
                        _scrollViewerCache[listView] = result;
                    }
                    return result;
                }
            }

            // ListViewの場合はnullをキャッシュに保存（再走査を避ける）
            if (listView != null)
            {
                _scrollViewerCache[listView] = null;
            }

            return null;
        }

        /// <summary>
        /// リストビューでフォーカスが取得されたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ルーティングイベント引数</param>
        private void ListView_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!ViewModel.IsSplitPaneEnabled)
                return;

            var listView = sender as System.Windows.Controls.ListView;
            if (listView == null)
                return;

            // フォーカスが取得されたListViewがどのペインに属しているかを判定
            var pane = GetPaneForElement(listView);
            if (pane == 0 || pane == 2)
            {
                // ViewModelのActivePaneプロパティを更新（これによりPropertyChangedが発火し、背景色が更新される）
                var previousActivePane = ViewModel.ActivePane;
                ViewModel.ActivePane = pane;

                // ActivePaneが変更されなかった場合（既に同じ値の場合）、手動で背景色を更新
                if (previousActivePane == pane)
                {
                    _cachedLeftListView = null;
                    _cachedRightListView = null;
                    _cachedSinglePaneListView = null;
                    UpdateListViewBackgroundColorsDelayed();
                }
            }
        }

        /// <summary>
        /// リストビューでマウスボタンが押される前に呼び出されます（Previewイベント）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void ListView_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (!ViewModel.IsSplitPaneEnabled)
                return;

            var listView = sender as System.Windows.Controls.ListView;
            if (listView == null)
                return;

            // クリックされたListViewがどのペインに属しているかを判定
            var pane = GetPaneForElement(listView);
            if (pane == 0 || pane == 2)
            {
                // ViewModelのActivePaneプロパティを更新（これによりPropertyChangedが発火し、背景色が更新される）
                var previousActivePane = ViewModel.ActivePane;
                ViewModel.ActivePane = pane;

                // ActivePaneが変更されなかった場合（既に同じ値の場合）、手動で背景色を更新
                if (previousActivePane == pane)
                {
                    _cachedLeftListView = null;
                    _cachedRightListView = null;
                    _cachedSinglePaneListView = null;
                    UpdateListViewBackgroundColorsDelayed();
                }
            }
        }

        /// <summary>
        /// リストビューでキーが押されたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">キーイベント引数</param>
        private void ListView_KeyDown(object sender, KeyEventArgs e)
        {
            var listView = sender as System.Windows.Controls.ListView;
            if (listView == null)
                return;

            // F2キーでリネームを開始
            if (e.Key == Key.F2)
            {
                StartRename(listView);
                e.Handled = true;
                return;
            }

            if (e.Key != Key.Back)
                return;
            if (listView == null)
                return;

            Models.ExplorerTab? targetTab = null;

            if (ViewModel.IsSplitPaneEnabled)
            {
                // 分割ペインモードの場合、フォーカスがあるListViewがどのペインに属しているかを判定
                var pane = GetPaneForElement(listView);
                if (pane == 0)
                {
                    // 左ペイン
                    targetTab = ViewModel.SelectedLeftPaneTab;
                }
                else if (pane == 2)
                {
                    // 右ペイン
                    targetTab = ViewModel.SelectedRightPaneTab;
                }
                else
                {
                    // 判定できない場合は、GetActiveTab()を使用
                    targetTab = GetActiveTab();
                }
            }
            else
            {
                // 通常モード
                targetTab = ViewModel.SelectedTab;
            }

            if (targetTab != null)
            {
                targetTab.ViewModel.NavigateToParentCommand.Execute(null);
                e.Handled = true;
            }
        }

        /// <summary>
        /// リストビューでマウスボタンが押されたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void ListView_MouseButtonDown(object sender, MouseButtonEventArgs e)
        {
            var listView = sender as System.Windows.Controls.ListView;
            if (listView == null)
                return;

            Models.ExplorerTab? targetTab = null;

            if (ViewModel.IsSplitPaneEnabled)
            {
                // 分割ペインモードの場合、クリックされたListViewがどのペインに属しているかを判定
                var pane = GetPaneForElement(listView);
                if (pane == 0)
                {
                    // 左ペイン
                    targetTab = ViewModel.SelectedLeftPaneTab;
                    // ActivePaneを更新（必ずアクティブにする）
                    var previousActivePane = ViewModel.ActivePane;
                    ViewModel.ActivePane = 0;

                    // ActivePaneが変更されなかった場合（既に同じ値の場合）、手動で背景色を更新
                    if (previousActivePane == 0)
                    {
                        _cachedLeftListView = null;
                        _cachedRightListView = null;
                        _cachedSinglePaneListView = null;
                        UpdateListViewBackgroundColorsDelayed();
                    }

                    // ListViewにフォーカスを設定
                    listView.Focus();
                }
                else if (pane == 2)
                {
                    // 右ペイン
                    targetTab = ViewModel.SelectedRightPaneTab;
                    // ActivePaneを更新（必ずアクティブにする）
                    var previousActivePane = ViewModel.ActivePane;
                    ViewModel.ActivePane = 2;

                    // ActivePaneが変更されなかった場合（既に同じ値の場合）、手動で背景色を更新
                    if (previousActivePane == 2)
                    {
                        _cachedLeftListView = null;
                        _cachedRightListView = null;
                        _cachedSinglePaneListView = null;
                        UpdateListViewBackgroundColorsDelayed();
                    }

                    // ListViewにフォーカスを設定
                    listView.Focus();
                }
                else
                {
                    // 判定できない場合は、GetActiveTab()を使用
                    targetTab = GetActiveTab();
                    // ListViewにフォーカスを設定
                    listView.Focus();
                }
            }
            else
            {
                // 通常モード
                targetTab = ViewModel.SelectedTab;
                // ListViewにフォーカスを設定
                listView.Focus();
            }

            if (targetTab == null)
                return;

            // マウスのバックボタン（XButton1）が押された場合
            if (e.ChangedButton == MouseButton.XButton1)
            {
                targetTab.ViewModel.NavigateToParentCommand.Execute(null);
                e.Handled = true;
            }
            // マウスの進むボタン（XButton2）が押された場合
            else if (e.ChangedButton == MouseButton.XButton2)
            {
                targetTab.ViewModel.NavigateForwardCommand.Execute(null);
                e.Handled = true;
            }
        }

        /// <summary>
        /// 現在アクティブなタブを取得します（分割ペインモードも考慮）
        /// </summary>
        /// <returns>現在アクティブなタブ、見つからない場合はnull</returns>
        private Models.ExplorerTab? GetActiveTab()
        {
            if (ViewModel.IsSplitPaneEnabled)
            {
                // 分割ペインモードの場合、右ペインを優先（フォーカスがある方）
                return ViewModel.SelectedRightPaneTab ?? ViewModel.SelectedLeftPaneTab;
            }
            else
            {
                return ViewModel.SelectedTab;
            }
        }

        /// <summary>
        /// ピン留めフォルダーがクリックされたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void PinnedFolder_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var border = sender as System.Windows.Controls.Border;
            if (border?.DataContext is Models.FavoriteItem favorite)
            {
                Models.ExplorerTab? targetTab = null;

                if (ViewModel.IsSplitPaneEnabled)
                {
                    // 分割ペインモードの場合、クリックされた要素がどのペインに属しているかを判定
                    var element = border as FrameworkElement;
                    var pane = GetPaneForElement(element);
                    if (pane == 0)
                    {
                        // 左ペイン
                        targetTab = ViewModel.SelectedLeftPaneTab;
                    }
                    else if (pane == 2)
                    {
                        // 右ペイン
                        targetTab = ViewModel.SelectedRightPaneTab;
                    }
                    else
                    {
                        // 判定できない場合は、GetActiveTab()を使用
                        targetTab = GetActiveTab();
                    }
                }
                else
                {
                    // 通常モード
                    targetTab = ViewModel.SelectedTab;
                }

                if (targetTab != null)
                {
                    targetTab.ViewModel.NavigateToPathCommand.Execute(favorite.Path);
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// ドライブがクリックされたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void Drive_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var border = sender as System.Windows.Controls.Border;
            if (border?.DataContext is Models.DriveInfoModel drive)
            {
                int? pane = null;

                if (ViewModel.IsSplitPaneEnabled)
                {
                    // 分割ペインモードの場合、クリックされた要素がどのペインに属しているかを判定
                    // GetPaneForElementではなく、TabControlを直接検索して確実に判定
                    var element = border as FrameworkElement;
                    var tabControl = FindAncestor<System.Windows.Controls.TabControl>(element);
                    if (tabControl != null)
                    {
                        var column = Grid.GetColumn(tabControl);
                        if (column == 0 || column == 2)
                        {
                            pane = column;
                        }
                    }

                    // TabControlが見つからない場合は、GetPaneForElementを使用（フォールバック）
                    if (!pane.HasValue)
                    {
                        var paneValue = GetPaneForElement(element);
                        if (paneValue == 0 || paneValue == 2)
                        {
                            pane = paneValue;
                        }
                    }
                }

                // ViewModelのコマンドを呼び出し
                ViewModel.NavigateToDriveCommand.Execute((drive.Path, pane));
                e.Handled = true;
            }
        }

        /// <summary>
        /// 要素がどのペインに属しているかを取得します（分割ペインモードの場合、キャッシュ付き）
        /// </summary>
        /// <param name="element">要素</param>
        /// <returns>左ペインの場合は0、右ペインの場合は2、判定できない場合は-1</returns>
        private int GetPaneForElement(FrameworkElement? element)
        {
            if (element == null)
                return -1;

            // ViewModelプロパティを一度だけ取得してキャッシュ（高速化）
            if (!ViewModel.IsSplitPaneEnabled)
                return -1;

            // キャッシュを確認（ただし、タブ移動後はキャッシュが無効化されている可能性があるため、TabControlを直接検索する方法も併用）
            if (_paneCache.TryGetValue(element, out var cachedPane))
            {
                // キャッシュされた値が有効かどうかを確認するため、TabControlを直接検索して検証
                var tabControl = FindAncestor<System.Windows.Controls.TabControl>(element);
                if (tabControl != null)
                {
                    var column = Grid.GetColumn(tabControl);
                    if (column == cachedPane && (column == 0 || column == 2))
                    {
                        // キャッシュが有効な場合はそのまま返す
                        return cachedPane;
                    }
                    // キャッシュが無効な場合は更新
                    if (column == 0 || column == 2)
                    {
                        _paneCache[element] = column;
                        return column;
                    }
                }
            }

            // キャッシュがない、または無効な場合は、TabControlを直接検索
            var tabControlDirect = FindAncestor<System.Windows.Controls.TabControl>(element);
            if (tabControlDirect != null)
            {
                var column = Grid.GetColumn(tabControlDirect);
                if (column == 0 || column == 2)
                {
                    // キャッシュに保存
                    _paneCache[element] = column;
                    return column;
                }
            }

            // 見つからなかった場合もキャッシュに保存（再走査を避ける）
            _paneCache[element] = -1;
            return -1;
        }

        /// <summary>
        /// 最近使用したファイルがクリックされたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void RecentFile_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var border = sender as System.Windows.Controls.Border;
            if (border?.DataContext is Models.FileSystemItem fileItem)
            {
                Models.ExplorerTab? targetTab = null;

                if (ViewModel.IsSplitPaneEnabled)
                {
                    // 分割ペインモードの場合、クリックされた要素がどのペインに属しているかを判定
                    var element = border as FrameworkElement;
                    var pane = GetPaneForElement(element);
                    if (pane == 0)
                    {
                        // 左ペイン
                        targetTab = ViewModel.SelectedLeftPaneTab;
                    }
                    else if (pane == 2)
                    {
                        // 右ペイン
                        targetTab = ViewModel.SelectedRightPaneTab;
                    }
                    else
                    {
                        // 判定できない場合は、GetActiveTab()を使用
                        targetTab = GetActiveTab();
                    }
                }
                else
                {
                    // 通常モード
                    targetTab = ViewModel.SelectedTab;
                }

                if (targetTab != null)
                {
                    if (fileItem.IsDirectory)
                    {
                        targetTab.ViewModel.NavigateToPathCommand.Execute(fileItem.FullPath);
                    }
                    else
                    {
                        targetTab.ViewModel.NavigateToItemCommand.Execute(fileItem);
                    }
                    e.Handled = true;
                }
            }
        }


        // ドラッグ&ドロップ用の変数
        private Point _dragStartPoint;
        private bool _isDragging = false;
        private FileSystemItem? _draggedItem = null;

        // タブのドラッグ&ドロップ用の変数
        private Point _tabDragStartPoint;
        private ExplorerTab? _draggedTab;
        private System.Windows.Controls.TabItem? _draggedTabItem; // ドラッグ中のTabItemを保持
        private bool _isTabDragging; // タブのドラッグ操作が進行中かどうか

        // ビジュアルツリー走査の結果をキャッシュ（パフォーマンス向上）
        private readonly Dictionary<FrameworkElement, int> _paneCache = new();
        private readonly Dictionary<System.Windows.Controls.ListView, System.Windows.Controls.ScrollViewer?> _scrollViewerCache = new();
        private readonly Dictionary<System.Windows.Controls.TabItem, System.Windows.Controls.TabControl?> _tabControlCache = new();

        // リネーム用の変数
        private FileSystemItem? _renamingItem = null;
        private System.Windows.Controls.TextBox? _renameTextBox = null;
        private System.Windows.Controls.TextBlock? _renameTextBlock = null;
        private System.Windows.Controls.ListViewItem? _renamingListViewItem = null;
        private System.Windows.Threading.DispatcherTimer? _renameClickTimer = null;
        private bool _isRenaming = false;
        private FileSystemItem? _lastClickedItem = null;
        private DateTime _lastClickTime = DateTime.MinValue;
        private Models.ExplorerTab? _renamingTab = null;
        private DateTime _renameCompletedTime = DateTime.MinValue;
        private int _activePaneBeforeRename = -1; // リネーム開始前のアクティブペーン

        /// <summary>
        /// ListViewItemでマウスが押されたときに呼び出されます（ドラッグ開始の検出）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void ListViewItem_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragStartPoint = e.GetPosition(null);
            _isDragging = false;
            _draggedItem = null;

            // リネームタイマーをキャンセル
            CancelRenameTimer();

            // ダブルクリックの場合はリネーム処理をスキップ
            if (e.ClickCount >= 2)
            {
                _lastClickedItem = null;
                _lastClickTime = DateTime.MinValue;
                return;
            }

            // System.Windows.Controls.ListViewItemを使用（エイリアスのListViewItemはWpf.Ui.Controls.ListViewItem）
            if (sender is System.Windows.Controls.ListViewItem listViewItem && listViewItem.DataContext is FileSystemItem item)
            {
                _draggedItem = item;

                // リネーム中の場合は何もしない
                if (_isRenaming)
                    return;

                // ListViewItemの親ListViewを取得してタブを判定
                Models.ExplorerTab? targetTab = null;
                var parentListView = FindParent<System.Windows.Controls.ListView>(listViewItem);
                if (parentListView != null && ViewModel.IsSplitPaneEnabled)
                {
                    var pane = GetPaneForElement(parentListView);
                    if (pane == 0)
                    {
                        targetTab = ViewModel.SelectedLeftPaneTab;
                        ViewModel.ActivePane = 0;
                    }
                    else if (pane == 2)
                    {
                        targetTab = ViewModel.SelectedRightPaneTab;
                        ViewModel.ActivePane = 2;
                    }
                    else
                    {
                        targetTab = GetActiveTab();
                    }
                    
                    // 親ListViewにフォーカスを設定
                    parentListView.Focus();
                }
                else
                {
                    targetTab = ViewModel.SelectedTab ?? GetActiveTab();
                    if (parentListView != null)
                    {
                        parentListView.Focus();
                    }
                }

                // 既に選択されているアイテムをクリックした場合、リネームタイマーを開始
                if (targetTab?.ViewModel.SelectedItem == item && _lastClickedItem == item)
                {
                    var now = DateTime.Now;
                    var timeSinceLastClick = (now - _lastClickTime).TotalMilliseconds;

                    // ダブルクリックの間隔より長く、合理的な時間内であればリネームを開始
                    // ダブルクリック時間はシステム設定から取得（デフォルト約500ms）
                    var doubleClickTime = GetDoubleClickTime();
                    if (timeSinceLastClick > doubleClickTime && timeSinceLastClick < 2000)
                    {
                        StartRenameTimer(listViewItem, item, targetTab);
                    }
                }

                _lastClickedItem = item;
                _lastClickTime = DateTime.Now;
            }
        }

        /// <summary>
        /// リネームタイマーを開始します
        /// </summary>
        /// <param name="listViewItem">リネーム対象のListViewItem</param>
        /// <param name="item">リネーム対象のFileSystemItem</param>
        /// <param name="tab">リネーム対象のタブ</param>
        private void StartRenameTimer(System.Windows.Controls.ListViewItem listViewItem, FileSystemItem item, Models.ExplorerTab? tab)
        {
            CancelRenameTimer();

            // リネーム開始前のアクティブペーンを記録（クリック時点で記録）
            _activePaneBeforeRename = ViewModel.IsSplitPaneEnabled ? ViewModel.ActivePane : -1;

            _renameClickTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(500)
            };

            _renameClickTimer.Tick += (s, args) =>
            {
                _renameClickTimer?.Stop();
                _renameClickTimer = null;

                // リネーム中でなく、ドラッグ中でもなければリネームを開始
                if (!_isRenaming && !_isDragging)
                {
                    StartRenameForItem(listViewItem, item, tab);
                }
            };

            _renameClickTimer.Start();
        }

        /// <summary>
        /// リネームタイマーをキャンセルします
        /// </summary>
        private void CancelRenameTimer()
        {
            if (_renameClickTimer != null)
            {
                _renameClickTimer.Stop();
                _renameClickTimer = null;
            }
        }

        /// <summary>
        /// ListViewItemでマウスが移動したときに呼び出されます（ドラッグ開始の判定）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスイベント引数</param>
        private void ListViewItem_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (_draggedItem == null)
                return;

            if (e.LeftButton != MouseButtonState.Pressed)
            {
                _isDragging = false;
                return;
            }

            var currentPoint = e.GetPosition(null);
            var diff = _dragStartPoint - currentPoint;

            // ドラッグ開始の閾値（5ピクセル以上移動した場合）
            if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
            {
                if (!_isDragging)
                {
                    _isDragging = true;
                    var dataObject = new DataObject();
                    dataObject.SetData("FileSystemItem", _draggedItem);
                    DragDrop.DoDragDrop(sender as DependencyObject ?? this, dataObject, DragDropEffects.Move);
                    _isDragging = false;
                    _draggedItem = null;
                }
            }
        }

        /// <summary>
        /// ListViewでドラッグオーバーされたときに呼び出されます（ホバー表示）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ドラッグイベント引数</param>
        private void ListView_DragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent("FileSystemItem"))
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            var activeTab = GetActiveTab();
            if (activeTab == null)
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            // マウス位置にあるListViewItemを取得
            var listView = sender as ListView;
            if (listView == null)
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            var point = e.GetPosition(listView);
            var item = GetItemAtPoint(listView, point);

            // フォルダーの上にホバーしている場合のみ移動を許可
            if (item is FileSystemItem fileItem && fileItem.IsDirectory)
            {
                e.Effects = DragDropEffects.Move;
                e.Handled = true;
            }
            else
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
            }
        }

        /// <summary>
        /// ListViewでドロップされたときに呼び出されます（ファイル移動）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ドラッグイベント引数</param>
        private void ListView_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent("FileSystemItem"))
                return;

            var draggedItem = e.Data.GetData("FileSystemItem") as FileSystemItem;
            if (draggedItem == null)
                return;

            var activeTab = GetActiveTab();
            if (activeTab == null)
                return;

            var listView = sender as ListView;
            if (listView == null)
                return;

            // マウス位置にあるListViewItemを取得
            var point = e.GetPosition(listView);
            var dropTarget = GetItemAtPoint(listView, point);

            // ドロップ先がフォルダーの場合のみ移動
            if (dropTarget is FileSystemItem targetItem && targetItem.IsDirectory)
            {
                // 同じフォルダーへの移動は無視（ReadOnlySpanを使用してメモリ割り当てを削減）
                var draggedPath = draggedItem.FullPath.AsSpan();
                var targetPath = targetItem.FullPath.AsSpan();
                if (draggedPath.CompareTo(targetPath, StringComparison.OrdinalIgnoreCase) == 0)
                    return;

                // 親フォルダーへの移動も無視（無限ループを防ぐ）
                var parentPath = System.IO.Path.GetDirectoryName(draggedItem.FullPath);
                if (parentPath != null)
                {
                    var parentPathSpan = parentPath.AsSpan();
                    if (parentPathSpan.CompareTo(targetPath, StringComparison.OrdinalIgnoreCase) == 0)
                        return;
                }

                // FileSystemServiceを取得して移動を実行
                var fileSystemService = App.Services.GetService(typeof(FileSystemService)) as FileSystemService;
                if (fileSystemService != null)
                {
                    try
                    {
                        if (fileSystemService.MoveItem(draggedItem.FullPath, targetItem.FullPath))
                        {
                            // 移動成功後、現在のタブを更新
                            activeTab.ViewModel.RefreshCommand.Execute(null);
                        }
                    }
                    catch
                    {
                        // エラーハンドリング
                    }
                }
            }

            e.Handled = true;
        }

        /// <summary>
        /// 指定されたポイントにあるListViewItemのDataContextを取得します
        /// </summary>
        /// <param name="listView">ListView</param>
        /// <param name="point">ポイント</param>
        /// <returns>DataContext、見つからない場合はnull</returns>
        private object? GetItemAtPoint(ListView listView, Point point)
        {
            var hitTestResult = VisualTreeHelper.HitTest(listView, point);
            if (hitTestResult == null)
                return null;

            var dependencyObject = hitTestResult.VisualHit;
            while (dependencyObject != null && dependencyObject != listView)
            {
                if (dependencyObject is ListViewItem listViewItem)
                {
                    return listViewItem.DataContext;
                }
                dependencyObject = VisualTreeHelper.GetParent(dependencyObject);
            }

            return null;
        }

        /// <summary>
        /// ListViewItemでマウス右ボタンが離されたときに呼び出されます（右クリックメニュー表示）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void ListViewItem_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {

            // DataContextからファイルパスを取得
            if (sender is ListViewItem listViewItem && listViewItem.DataContext is FileSystemItem fileItem)
            {
                var filePath = fileItem.FullPath;

                if (!string.IsNullOrEmpty(filePath) && (System.IO.Directory.Exists(filePath) || System.IO.File.Exists(filePath)))
                {
                    // イベントを処理済みとしてマーク
                    e.Handled = true;

                    // 画面上の座標をスクリーン座標に変換
                    var point = e.GetPosition(this);
                    var screenPoint = PointToScreen(point);

                    // ウィンドウハンドルを取得
                    var window = Window.GetWindow(this);
                    var hWnd = window != null ? new WindowInteropHelper(window).Handle : IntPtr.Zero;

                    // ShellContextMenuでOS標準メニューを表示
                    var scm = new FastExplorer.ShellContextMenu.ShellContextMenuService();
                    scm.ShowContextMenu(new[] { filePath }, hWnd, (int)screenPoint.X, (int)screenPoint.Y);
                }
            }
        }

        /// <summary>
        /// ListViewでマウス右ボタンが離されたときに呼び出されます（空領域の右クリックメニュー表示）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void ListView_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {

            // ListViewItem上でクリックされた場合は処理しない
            if (e.OriginalSource is DependencyObject source)
            {
                DependencyObject? current = source;
                while (current != null)
                {
                    if (current is ListViewItem)
                    {
                        return;
                    }
                    current = VisualTreeHelper.GetParent(current);
                }
            }

            // イベントを処理済みとしてマーク
            e.Handled = true;

            // 現在のタブを取得
            Models.ExplorerTab? targetTab = null;

            if (ViewModel.IsSplitPaneEnabled)
            {
                // 分割ペインモードの場合、クリックされたListViewがどのペインに属しているかを判定
                if (sender is System.Windows.Controls.ListView listView)
                {
                    var pane = GetPaneForElement(listView);
                    if (pane == 0)
                    {
                        targetTab = ViewModel.SelectedLeftPaneTab;
                    }
                    else if (pane == 2)
                    {
                        targetTab = ViewModel.SelectedRightPaneTab;
                    }
                    else
                    {
                        targetTab = GetActiveTab();
                    }
                }
            }
            else
            {
                // 通常モード
                targetTab = ViewModel.SelectedTab;
            }

            if (targetTab == null)
            {
                return;
            }

            // 現在のパスを取得
            var path = targetTab.ViewModel?.CurrentPath;

            // パスが空の場合はホームディレクトリを使用
            if (string.IsNullOrEmpty(path))
            {
                path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            }

            if (!string.IsNullOrEmpty(path) && (System.IO.Directory.Exists(path) || System.IO.File.Exists(path)))
            {
                // 画面上の座標をスクリーン座標に変換
                var point = e.GetPosition(this);
                var screenPoint = PointToScreen(point);

                // ウィンドウハンドルを取得
                var window = Window.GetWindow(this);
                var hWnd = window != null ? new WindowInteropHelper(window).Handle : IntPtr.Zero;

                // ShellContextMenuでOS標準メニューを表示
                var scm = new FastExplorer.ShellContextMenu.ShellContextMenuService();
                scm.ShowContextMenu(new[] { path }, hWnd, (int)screenPoint.X, (int)screenPoint.Y);
            }
        }

        /// <summary>
        /// GridViewColumnHeaderがクリックされたときに呼び出されます（ソート処理）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ルーティングイベント引数</param>
        private void GridViewColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not System.Windows.Controls.GridViewColumnHeader header)
                return;

            // クリックされたListViewを取得
            var listView = FindAncestor<System.Windows.Controls.ListView>(header);
            if (listView == null)
                return;

            // ターゲットタブを取得（最適化：早期リターン）
            Models.ExplorerTab? targetTab;

            if (ViewModel.IsSplitPaneEnabled)
            {
                var pane = GetPaneForElement(listView);
                targetTab = pane switch
                {
                    0 => ViewModel.SelectedLeftPaneTab,
                    2 => ViewModel.SelectedRightPaneTab,
                    _ => GetActiveTab()
                };
            }
            else
            {
                targetTab = ViewModel.SelectedTab;
            }

            if (targetTab?.ViewModel == null)
                return;

            // ヘッダーテキストから列名を取得
            var columnName = GetColumnNameFromHeader(header);
            if (!string.IsNullOrEmpty(columnName))
            {
                targetTab.ViewModel.SortByColumn(columnName);
            }
        }

        /// <summary>
        /// GridViewColumnHeaderがダブルクリックされたときに呼び出されます
        /// リサイズハンドル上でのダブルクリックの場合は列幅を自動調整
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void GridViewColumnHeader_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is not System.Windows.Controls.GridViewColumnHeader header)
            {
                e.Handled = true;
                return;
            }

            var targetColumn = GetResizeHandleColumn(header, e.GetPosition(header));
            if (targetColumn != null)
            {
                AutoSizeColumn(header, targetColumn);
                e.Handled = true;
                return;
            }

            // リサイズハンドル以外でのダブルクリックは無効化
            e.Handled = true;
        }

        /// <summary>
        /// GridViewColumnHeaderがダブルクリックされる前に呼び出されます（Previewイベント）
        /// リサイズハンドル上でのダブルクリックの場合は列幅を自動調整し、親要素への伝播を防ぐ
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void GridViewColumnHeader_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is not System.Windows.Controls.GridViewColumnHeader header)
                return;

            var targetColumn = GetResizeHandleColumn(header, e.GetPosition(header));
            if (targetColumn != null)
            {
                // Previewイベントでは非同期で処理（UIスレッドのブロックを防ぐ）
                var columnToResize = targetColumn;
                Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    AutoSizeColumn(header, columnToResize);
                }), System.Windows.Threading.DispatcherPriority.Loaded);
                e.Handled = true;
            }
            // リサイズハンドルでない場合は、通常のイベントを発火させる（Handledしない）
        }

        /// <summary>
        /// クリック位置からリサイズハンドルに対応する列を取得します（最適化：共通メソッド）
        /// </summary>
        /// <param name="header">GridViewColumnHeader</param>
        /// <param name="clickPosition">クリック位置</param>
        /// <returns>リサイズハンドルに対応する列、見つからない場合はnull</returns>
        private System.Windows.Controls.GridViewColumn? GetResizeHandleColumn(
            System.Windows.Controls.GridViewColumnHeader header,
            Point clickPosition)
        {
            if (header == null)
                return null;

            var headerWidth = header.ActualWidth;
            // リサイズハンドルの検出範囲（ピクセル単位）
            // GridViewColumnHeaderの列境界線付近でクリックされた場合にリサイズハンドルとして認識する範囲
            // 通常のリサイズハンドルは約5ピクセルだが、クリックしやすくするため8ピクセルに設定
            const double resizeHandleWidth = 8.0;

            // 右端のリサイズハンドル（この列の右端）
            if (clickPosition.X >= headerWidth - resizeHandleWidth && clickPosition.X <= headerWidth)
            {
                return header.Column;
            }

            // 左端のリサイズハンドル（前の列の右端、つまりこの列の左端）
            if (clickPosition.X >= 0 && clickPosition.X <= resizeHandleWidth)
            {
                // ListViewを取得
                var listView = FindAncestor<System.Windows.Controls.ListView>(header);
                if (listView?.View is System.Windows.Controls.GridView gridView)
                {
                    var currentColumn = header.Column;
                    if (currentColumn != null)
                    {
                        var columns = gridView.Columns;
                        var currentIndex = columns.IndexOf(currentColumn);
                        if (currentIndex > 0)
                        {
                            return columns[currentIndex - 1];
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// 列幅を自動調整します（ヘッダーと内容の両方を考慮）
        /// </summary>
        /// <param name="header">GridViewColumnHeader</param>
        /// <param name="column">GridViewColumn</param>
        private void AutoSizeColumn(System.Windows.Controls.GridViewColumnHeader header, System.Windows.Controls.GridViewColumn column)
        {
            if (header == null || column == null)
                return;

            // ListViewを取得
            var listView = FindAncestor<System.Windows.Controls.ListView>(header);
            if (listView == null)
                return;

            // ターゲットタブを取得
            Models.ExplorerTab? targetTab = null;

            if (ViewModel.IsSplitPaneEnabled)
            {
                var pane = GetPaneForElement(listView);
                targetTab = pane switch
                {
                    0 => ViewModel.SelectedLeftPaneTab,
                    2 => ViewModel.SelectedRightPaneTab,
                    _ => GetActiveTab()
                };
            }
            else
            {
                targetTab = ViewModel.SelectedTab;
            }

            if (targetTab?.ViewModel?.Items == null)
                return;

            // ヘッダーテキストの幅を計算
            double maxWidth = 0.0;

            // ヘッダーの幅を測定
            if (header.Content is string headerText)
            {
                var textBlock = new TextBlock
                {
                    Text = headerText,
                    FontFamily = header.FontFamily,
                    FontSize = header.FontSize,
                    FontWeight = header.FontWeight,
                    FontStyle = header.FontStyle,
                    FontStretch = header.FontStretch
                };
                textBlock.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                maxWidth = Math.Max(maxWidth, textBlock.DesiredSize.Width);
            }
            else if (header.Content is FrameworkElement headerElement)
            {
                headerElement.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                maxWidth = Math.Max(maxWidth, headerElement.DesiredSize.Width);
            }

            // 列の内容の幅を測定
            var items = targetTab.ViewModel.Items;
            var itemsCount = items.Count;
            if (itemsCount == 0)
            {
                // アイテムがない場合はヘッダーの幅のみを使用
                column.Width = Math.Max(50.0, maxWidth + 20.0);
                return;
            }

            var cellTemplate = column.CellTemplate;
            var columnHeader = column.Header as string;
            var isNameColumn = columnHeader == "名前";

            // ListViewのフォントプロパティをキャッシュ（パフォーマンス向上）
            var fontFamily = listView.FontFamily;
            var fontSize = listView.FontSize;
            var fontWeight = listView.FontWeight;
            var fontStyle = listView.FontStyle;
            var fontStretch = listView.FontStretch;

            // 名前列の定数
            const double iconWidth = 20.0;
            const double iconMargin = 12.0;

            if (cellTemplate != null)
            {
                // セルテンプレートを使用して内容の幅を測定
                foreach (var item in items)
                {
                    if (item is FileSystemItem fileItem)
                    {
                        double itemWidth;

                        if (isNameColumn)
                        {
                            // 名前列の場合は、SymbolIcon + マージン + テキストの幅を計算
                            var textBlock = new TextBlock
                            {
                                Text = fileItem.Name,
                                FontFamily = fontFamily,
                                FontSize = fontSize,
                                FontWeight = fontWeight,
                                FontStyle = fontStyle,
                                FontStretch = fontStretch
                            };
                            textBlock.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                            itemWidth = iconWidth + iconMargin + textBlock.DesiredSize.Width;
                        }
                        else
                        {
                            // その他の列はContentPresenterで測定
                            var contentPresenter = new ContentPresenter
                            {
                                Content = fileItem,
                                ContentTemplate = cellTemplate
                            };
                            contentPresenter.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                            itemWidth = contentPresenter.DesiredSize.Width;
                        }

                        maxWidth = Math.Max(maxWidth, itemWidth);
                    }
                }
            }
            else
            {
                // セルテンプレートがない場合は、データを直接測定
                foreach (var item in items)
                {
                    if (item is FileSystemItem fileItem)
                    {
                        string text = GetTextForColumn(fileItem, column);
                        if (!string.IsNullOrEmpty(text))
                        {
                            var textBlock = new TextBlock
                            {
                                Text = text,
                                FontFamily = fontFamily,
                                FontSize = fontSize,
                                FontWeight = fontWeight,
                                FontStyle = fontStyle,
                                FontStretch = fontStretch
                            };
                            textBlock.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                            maxWidth = Math.Max(maxWidth, textBlock.DesiredSize.Width);
                        }
                    }
                }
            }

            // パディングとマージンを考慮（余裕を持たせる）
            const double padding = 20.0; // 左右のパディング
            maxWidth += padding;

            // 最小幅と最大幅を設定
            const double minWidth = 50.0;
            const double maxColumnWidth = 500.0;
            maxWidth = Math.Max(minWidth, Math.Min(maxWidth, maxColumnWidth));

            // 列幅を設定
            column.Width = maxWidth;
        }

        /// <summary>
        /// 列に応じたテキストを取得します
        /// </summary>
        /// <param name="item">FileSystemItem</param>
        /// <param name="column">GridViewColumn</param>
        /// <returns>列に表示するテキスト</returns>
        private static string GetTextForColumn(FileSystemItem item, System.Windows.Controls.GridViewColumn column)
        {
            // 列のヘッダーから判定（簡易実装）
            // より正確には、CellTemplateの内容を解析する必要がある
            if (column.Header is string headerText)
            {
                return headerText switch
                {
                    "名前" => item.Name,
                    "サイズ" => item.FormattedSize,
                    "種類" => item.Extension,
                    "更新日時" => item.FormattedDate,
                    _ => item.Name
                };
            }
            return item.Name;
        }

        /// <summary>
        /// ヘッダーから列名を取得します（最適化：定数を使用）
        /// </summary>
        /// <param name="header">GridViewColumnHeader</param>
        /// <returns>列名（"Name", "Size", "Extension", "LastModified"）</returns>
        private static string GetColumnNameFromHeader(System.Windows.Controls.GridViewColumnHeader header)
        {
            if (header?.Content is not string headerText)
                return string.Empty;

            // 定数を使用してメモリ割り当てを削減
            const string NameHeader = "名前";
            const string SizeHeader = "サイズ";
            const string ExtensionHeader = "種類";
            const string LastModifiedHeader = "更新日時";

            return headerText switch
            {
                NameHeader => "Name",
                SizeHeader => "Size",
                ExtensionHeader => "Extension",
                LastModifiedHeader => "LastModified",
                _ => string.Empty
            };
        }

        /// <summary>
        /// 指定された型の親要素を検索します
        /// </summary>
        /// <typeparam name="T">検索する型</typeparam>
        /// <param name="element">開始要素</param>
        /// <returns>見つかった親要素、見つからない場合はnull</returns>
        private T? FindAncestor<T>(DependencyObject element) where T : DependencyObject
        {
            var current = element;
            while (current != null)
            {
                if (current is T ancestor)
                {
                    return ancestor;
                }
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }

        /// <summary>
        /// パスコピーボタンがクリックされたときに呼び出されます（現在のタブのパスをクリップボードにコピー）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ルーティングイベント引数</param>
        private void CopyPathButton_Click(object sender, RoutedEventArgs e)
        {
            // ボタンのDataContextから直接タブを取得（最適化：親要素を辿る前にボタン自体のDataContextを確認）
            Models.ExplorerTab? tab = null;

            if (sender is FrameworkElement buttonElement)
            {
                // まずボタン自体のDataContextを確認
                tab = buttonElement.DataContext as Models.ExplorerTab;

                // DataContextが見つからない場合、親要素を辿ってExplorerTabのDataContextを取得（最大5階層まで）
                if (tab == null)
                {
                    var current = VisualTreeHelper.GetParent(buttonElement);
                    int depth = 0;
                    const int maxDepth = 5; // 最大深度を制限してパフォーマンス向上
                    while (current != null && depth < maxDepth)
                    {
                        if (current is FrameworkElement element && element.DataContext is Models.ExplorerTab explorerTab)
                        {
                            tab = explorerTab;
                            break;
                        }
                        current = VisualTreeHelper.GetParent(current);
                        depth++;
                    }
                }
            }

            // DataContextから取得できない場合は、選択されているタブを使用（フォールバック）
            if (tab == null)
            {
                var viewModel = ViewModel;
                if (viewModel.IsSplitPaneEnabled)
                {
                    var activePane = viewModel.ActivePane;
                    tab = activePane == 0 ? viewModel.SelectedLeftPaneTab
                        : activePane == 2 ? viewModel.SelectedRightPaneTab
                        : viewModel.SelectedLeftPaneTab ?? viewModel.SelectedRightPaneTab;
                }
                else
                {
                    tab = viewModel.SelectedTab;
                }
            }

            // パスを取得してクリップボードにコピー（最適化：null合体演算子を使用）
            var path = tab?.ViewModel?.CurrentPath;
            if (string.IsNullOrEmpty(path))
            {
                path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            }

            if (!string.IsNullOrEmpty(path))
            {
                try
                {
                    System.Windows.Clipboard.SetText(path);
                }
                catch
                {
                    // クリップボードへのコピーに失敗した場合は何もしない
                }
            }
        }

        /// <summary>
        /// タブのButtonがクリックされたときに呼び出されます（タブのパスをクリップボードにコピー）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ルーティングイベント引数</param>
        private void TabArea_ButtonClick(object sender, RoutedEventArgs e)
        {
            // 閉じるボタン上でクリックされた場合は処理しない
            // 最適化：Tag?.ToString()の代わりに直接比較（メモリ割り当てを削減）
            if (sender is Button button)
            {
                var tag = button.Tag;
                // 文字列の場合は直接比較、それ以外の場合はToString()を呼び出し
                if (tag is string tagString && tagString == "CloseButton")
                {
                    return;
                }
            }

            // タブのDataContextを取得（パターンマッチングを使用）
            Models.ExplorerTab? tab = null;
            if (sender is FrameworkElement element)
            {
                // DataContextを直接取得（DataTemplate内のButtonはDataContextを継承している）
                tab = element.DataContext as Models.ExplorerTab;

                // DataContextが取得できない場合は、親要素を探す
                if (tab == null)
                {
                    var parent = VisualTreeHelper.GetParent(element) as FrameworkElement;
                    tab = parent?.DataContext as Models.ExplorerTab;
                }
            }

            if (tab != null)
            {
                // タブのパスを取得（プロパティアクセスをキャッシュ）
                var path = tab.ViewModel?.CurrentPath;
                if (string.IsNullOrEmpty(path))
                {
                    // パスが空の場合はホームディレクトリを使用
                    path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                }

                if (!string.IsNullOrEmpty(path))
                {
                    try
                    {
                        System.Windows.Clipboard.SetText(path);
                    }
                    catch
                    {
                        // クリップボードへのコピーに失敗した場合は何もしない
                    }
                }
            }
        }

        /// <summary>
        /// タブエリアでマウス右ボタンが離されたときに呼び出されます（タブの右クリックメニュー表示）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void TabArea_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            // 閉じるボタン上でクリックされた場合は処理しない
            if (e.OriginalSource is DependencyObject source)
            {
                DependencyObject? current = source;
                while (current != null)
                {
                    if (current is Button button && button.Tag?.ToString() == "CloseButton")
                    {
                        // 閉じるボタン上でクリックされた場合は処理しない
                        return;
                    }
                    current = VisualTreeHelper.GetParent(current);
                }
            }

            // タブのDataContextを取得
            Models.ExplorerTab? tab = null;
            if (sender is FrameworkElement element)
            {
                // DataContextを直接取得（DataTemplate内のButtonはDataContextを継承している）
                tab = element.DataContext as Models.ExplorerTab;

                // DataContextが取得できない場合は、親要素を探す
                if (tab == null)
                {
                    var parent = VisualTreeHelper.GetParent(element) as FrameworkElement;
                    tab = parent?.DataContext as Models.ExplorerTab;
                }
            }

            // TabItemを取得してコンテキストメニューを表示
            if (sender is System.Windows.Controls.TabItem tabItem && tab != null)
            {
                // タブを選択状態にする（右クリックしたタブを選択）
                if (ViewModel.IsSplitPaneEnabled)
                {
                    var leftPaneTabs = ViewModel.LeftPaneTabs;
                    var rightPaneTabs = ViewModel.RightPaneTabs;

                    if (leftPaneTabs.Contains(tab))
                    {
                        ViewModel.SelectedLeftPaneTab = tab;
                        ViewModel.ActivePane = 0;
                    }
                    else if (rightPaneTabs.Contains(tab))
                    {
                        ViewModel.SelectedRightPaneTab = tab;
                        ViewModel.ActivePane = 2;
                    }
                }
                else
                {
                    ViewModel.SelectedTab = tab;
                }

                // コンテキストメニューを表示（XAMLで設定されたコンテキストメニューを使用）
                if (tabItem.ContextMenu != null)
                {
                    // DataContextを設定してバインディングを有効にする
                    tabItem.ContextMenu.DataContext = tab;
                    tabItem.ContextMenu.Tag = tab; // タブをTagに保存して後で使用
                    tabItem.ContextMenu.IsOpen = true;
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// タブのコンテキストメニューが開かれたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ルーティングイベント引数</param>
        private void TabContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.ContextMenu contextMenu && contextMenu.Tag is Models.ExplorerTab tab)
            {
                // コンテキストメニューの各MenuItemにタブを保存
                foreach (var item in contextMenu.Items)
                {
                    if (item is System.Windows.Controls.MenuItem menuItem)
                    {
                        menuItem.Tag = tab;
                    }
                }
            }
        }

        /// <summary>
        /// タブのコンテキストメニューのアイテムがクリックされたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ルーティングイベント引数</param>
        private void TabContextMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.MenuItem menuItem && menuItem.Tag is Models.ExplorerTab tab)
            {
                var menuItemName = menuItem.Name;

                switch (menuItemName)
                {
                    case "DuplicateTabMenuItem":
                        ViewModel.DuplicateTabCommand.Execute(tab);
                        break;
                    case "DuplicateTabToNewWindowMenuItem":
                        ViewModel.DuplicateTabToNewWindowCommand.Execute(tab);
                        break;
                    case "MoveTabToNewWindowMenuItem":
                        ViewModel.MoveTabToNewWindowCommand.Execute(tab);
                        break;
                    case "CloseTabsToLeftMenuItem":
                        ViewModel.CloseTabsToLeftCommand.Execute(tab);
                        break;
                    case "CloseTabsToRightMenuItem":
                        ViewModel.CloseTabsToRightCommand.Execute(tab);
                        break;
                    case "CloseOtherTabsMenuItem":
                        ViewModel.CloseOtherTabsCommand.Execute(tab);
                        break;
                }
            }
        }

        /// <summary>
        /// TabItemでマウスが押されたときに呼び出されます（タブのドラッグ開始の検出）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスボタンイベント引数</param>
        private void TabItem_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // 閉じるボタン上でクリックされた場合は処理しない
            if (IsCloseButton(e.OriginalSource))
                return;

            if (sender is not System.Windows.Controls.TabItem tabItem || tabItem.DataContext is not ExplorerTab tab)
                return;

            _tabDragStartPoint = e.GetPosition(null);
            _draggedTab = tab;
            _draggedTabItem = tabItem;
            _isTabDragging = false;
            
            // ドラッグを開始するために、マウスキャプチャを設定
            tabItem.CaptureMouse();
        }

        /// <summary>
        /// 指定された要素が閉じるボタンかどうかを判定します
        /// </summary>
        private static bool IsCloseButton(object? source)
        {
            if (source is not DependencyObject depObj)
                return false;

            DependencyObject? current = depObj;
            while (current != null)
            {
                if (current is Button button)
                {
                    object? tag = button.Tag;
                    if (tag != null && tag.ToString() == "CloseButton")
                        return true;
                }
                current = VisualTreeHelper.GetParent(current);
            }
            return false;
        }

        /// <summary>
        /// TabItemでマウスが移動したときに呼び出されます（タブのドラッグ開始の判定）
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">マウスイベント引数</param>
        private void TabItem_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            // 早期リターン: マウスボタンが離されている場合は処理しない
            if (e.LeftButton != MouseButtonState.Pressed)
            {
                ReleaseMouseCapture();
                return;
            }

            // 早期リターン: ドラッグ条件を満たさない場合は処理しない
            if (_draggedTabItem == null || _isTabDragging)
                return;

            // 閉じるボタン上でのドラッグを防ぐ
            if (IsCloseButton(e.OriginalSource))
                return;

            Point currentPoint = e.GetPosition(null);
            double deltaX = currentPoint.X - _tabDragStartPoint.X;
            double deltaY = currentPoint.Y - _tabDragStartPoint.Y;

            // ドラッグ開始の閾値チェック（Math.Absを避けて絶対値比較を最適化）
            double minDragDistance = SystemParameters.MinimumHorizontalDragDistance;
            if ((deltaX > minDragDistance || deltaX < -minDragDistance) ||
                (deltaY > minDragDistance || deltaY < -minDragDistance))
            {
                _isTabDragging = true;

                // TabControl_Dropが呼ばれたかどうかを追跡
                bool dropHandled = false;
                DragEventHandler? dropHandler = null;
                dropHandler = (s, e) =>
                {
                    dropHandled = true;
                };

                // すべてのTabControlのDropイベントを監視
                var tabControls = new List<System.Windows.Controls.TabControl>();
                FindVisualChildrenRecursive(this, tabControls);

                foreach (var tabControl in tabControls)
                {
                    tabControl.Drop += dropHandler;
                }

                // ウィンドウを取得
                var window = Window.GetWindow(_draggedTabItem);
                var windowHandle = window != null ? new WindowInteropHelper(window).Handle : IntPtr.Zero;

                // ウィンドウ外へのドロップを検出するためのフラグ
                bool droppedOutsideWindow = false;
                // 最後に検出されたマウス位置（ウィンドウ外判定の精度向上のため）
                System.Windows.Point? lastMousePosition = null;

                // GiveFeedbackイベントでマウス位置を追跡（デフォルトの禁止マークを表示するため、特別な処理は不要）
                GiveFeedbackEventHandler? giveFeedbackHandler = null;
                if (_draggedTabItem != null && windowHandle != IntPtr.Zero)
                {
                    giveFeedbackHandler = (s, e) =>
                    {
                        // 現在のマウス位置を取得（ウィンドウ外判定のため）
                        POINT mousePos;
                        if (GetCursorPos(out mousePos))
                        {
                            lastMousePosition = new System.Windows.Point(mousePos.X, mousePos.Y);

                            // ウィンドウの境界を取得（毎回取得して最新の状態を反映）
                            RECT windowRect;
                            if (GetWindowRect(windowHandle, out windowRect))
                            {
                                // マウス位置がウィンドウ外にあるかどうかをチェック
                                // 境界を含めない（境界上はウィンドウ内とみなす）
                                bool isOutsideWindow = mousePos.X < windowRect.Left || mousePos.X > windowRect.Right ||
                                                       mousePos.Y < windowRect.Top || mousePos.Y > windowRect.Bottom;

                                // ウィンドウ外にいる場合、デフォルトの禁止マークを表示（UseDefaultCursors = trueでデフォルト動作）
                                // ウィンドウ内にいる場合も、デフォルトのカーソルを使用
                                e.UseDefaultCursors = true;
                            }
                            else
                            {
                                // ウィンドウの境界取得に失敗した場合、デフォルトカーソルを使用
                                e.UseDefaultCursors = true;
                            }
                        }
                        else
                        {
                            // マウス位置の取得に失敗した場合、デフォルトカーソルを使用
                            e.UseDefaultCursors = true;
                        }
                    };
                    _draggedTabItem.GiveFeedback += giveFeedbackHandler;
                }

                // QueryContinueDragイベントでウィンドウ外へのドロップを検出
                QueryContinueDragEventHandler? queryContinueDragHandler = null;
                if (_draggedTabItem != null && windowHandle != IntPtr.Zero)
                {
                    queryContinueDragHandler = (s, e) =>
                    {
                        // ドラッグ中にマウス位置を追跡
                        if (e.Action == DragAction.Continue || e.Action == DragAction.Drop)
                        {
                            // 現在のマウス位置を取得
                            POINT mousePos;
                            if (GetCursorPos(out mousePos))
                            {
                                lastMousePosition = new System.Windows.Point(mousePos.X, mousePos.Y);

                                // ウィンドウの境界を取得（毎回取得して最新の状態を反映）
                                RECT windowRect;
                                if (GetWindowRect(windowHandle, out windowRect))
                                {
                                    // マウス位置がウィンドウ外にあるかどうかをチェック
                                    // 境界を含めない（境界上はウィンドウ内とみなす）
                                    bool isOutsideWindow = mousePos.X < windowRect.Left || mousePos.X > windowRect.Right ||
                                                          mousePos.Y < windowRect.Top || mousePos.Y > windowRect.Bottom;
                                    
                                    if (isOutsideWindow)
                                    {
                                        if (e.Action == DragAction.Drop)
                                        {
                                            droppedOutsideWindow = true;
                                        }
                                        // カーソルは変更しない（GiveFeedbackで禁止マークをポップアップ表示）
                                    }
                                }
                            }
                        }
                    };
                    _draggedTabItem.QueryContinueDrag += queryContinueDragHandler;
                }

                try
                {
                    var dragResult = DragDrop.DoDragDrop(_draggedTabItem, _draggedTabItem, DragDropEffects.Move);

                    // ウィンドウ外にドロップされた場合、またはTabControl_Dropが呼ばれなかった場合
                    if (!dropHandled && _draggedTab != null)
                    {
                        // ドロップ時のマウス位置を再取得して確認
                        System.Windows.Point? dropPosition = null;
                        bool isOutsideWindow = droppedOutsideWindow; // QueryContinueDragで検出された値を優先

                        // dragResultがNoneの場合、タブがどこにもドロップされなかった可能性が高い
                        // この場合、デスクトップにドロップされたとみなす
                        bool dragResultIsNone = dragResult == DragDropEffects.None;

                        if (windowHandle != IntPtr.Zero)
                        {
                            POINT mousePos;
                            if (GetCursorPos(out mousePos))
                            {
                                RECT windowRect;
                                if (GetWindowRect(windowHandle, out windowRect))
                                {
                                    // マウス位置がウィンドウ外にあるかどうかを再確認
                                    // ただし、QueryContinueDragで既に検出されている場合はそれを優先
                                    bool currentIsOutsideWindow = mousePos.X < windowRect.Left || mousePos.X > windowRect.Right ||
                                                                  mousePos.Y < windowRect.Top || mousePos.Y > windowRect.Bottom;

                                    // QueryContinueDragで検出されていない場合のみ、現在のマウス位置で判定
                                    if (!droppedOutsideWindow)
                                    {
                                        isOutsideWindow = currentIsOutsideWindow;
                                    }
                                    
                                    if (isOutsideWindow)
                                    {
                                        dropPosition = new System.Windows.Point(mousePos.X, mousePos.Y);
                                    }
                                }
                            }
                        }

                        // ドロップが処理されなかった場合、デスクトップにドロップされたかどうかを判定
                        bool isDroppedOnDesktop = false;

                        // まず、droppedOutsideWindowがTrueかどうかを確認（最優先）
                        // QueryContinueDragで一度ウィンドウ外に出たことを検出した場合、デスクトップにドロップされたとみなす
                        if (droppedOutsideWindow)
                        {
                            isDroppedOnDesktop = true;
                        }
                        // dragResultがNoneで、dropHandledがfalseの場合、
                        // タブがどこにもドロップされなかったことを意味するため、デスクトップにドロップされたとみなす
                        else if (dragResultIsNone)
                        {
                            // ウィンドウ外にドロップされたかどうかを再確認（ただし、必須ではない）
                            if (isOutsideWindow)
                            {
                                // ウィンドウ外にドロップされた場合、確実にデスクトップにドロップされたとみなす
                                isDroppedOnDesktop = true;
                            }
                            else
                            {
                                // ウィンドウ外でなくても、dragResultがNoneでdropHandledがfalseの場合、
                                // ドロップ後にマウスがウィンドウ内に戻った可能性があるため、デスクトップにドロップされたとみなす
                                // ただし、lastMousePositionを使用して最後のマウス位置を確認
                                if (lastMousePosition.HasValue)
                                {
                                    // lastMousePositionがウィンドウ外にあるかどうかを確認
                                    RECT windowRect;
                                    if (GetWindowRect(windowHandle, out windowRect))
                                    {
                                        bool lastPosOutsideWindow = lastMousePosition.Value.X < windowRect.Left ||
                                                                   lastMousePosition.Value.X > windowRect.Right ||
                                                                   lastMousePosition.Value.Y < windowRect.Top ||
                                                                   lastMousePosition.Value.Y > windowRect.Bottom;

                                        if (lastPosOutsideWindow)
                                        {
                                            isDroppedOnDesktop = true;
                                        }
                                        else
                                        {
                                            // lastMousePositionもウィンドウ内の場合、デスクトップではない可能性が高い
                                        }
                                    }
                                    else
                                    {
                                        // ウィンドウ矩形が取得できない場合、念のためデスクトップにドロップされたとみなす
                                        isDroppedOnDesktop = true;
                                    }
                                }
                                else
                                {
                                    // lastMousePositionがない場合、念のためデスクトップにドロップされたとみなす
                                    isDroppedOnDesktop = true;
                                }
                            }
                        }
                        else if (isOutsideWindow)
                        {
                            // ウィンドウ外にドロップされた場合（droppedOutsideWindowは既に上でチェック済み）
                            if (dropPosition.HasValue)
                            {
                                // ドロップ位置の下にあるウィンドウを取得
                                POINT dropPoint = new POINT { X = (int)dropPosition.Value.X, Y = (int)dropPosition.Value.Y };
                                IntPtr windowUnderPoint = WindowFromPoint(dropPoint);
                                
                                // デスクトップのウィンドウハンドルを取得
                                IntPtr desktopWindow = GetDesktopWindow();
                                
                                // ウィンドウがデスクトップまたはその子ウィンドウであるかを確認（参考情報として）
                                bool isDesktopOrChild = IsDesktopOrChildWindow(windowUnderPoint, desktopWindow);

                                // ドロップ位置の下にあるウィンドウが現在のアプリケーションのウィンドウでない場合
                                if (windowUnderPoint != windowHandle)
                                {
                                    // ウィンドウ外にドロップされ、かつ現在のアプリケーションのウィンドウでない場合
                                    // （droppedOutsideWindowは既に上でチェック済みなので、ここではfalse）

                                    if (dragResult == DragDropEffects.None)
                                    {
                                        // ウィンドウ外にドロップされ、dragResultがNoneの場合、デスクトップにドロップされたとみなす
                                        // isDesktopOrChildの結果に関わらず、これらの条件が満たされればデスクトップと判定
                                        isDroppedOnDesktop = true;
                                    }
                                    else
                                    {
                                        // dragResultがNoneでない場合、別のアプリケーションがドロップを受け取った可能性がある
                                        // ただし、windowUnderPointがデスクトップまたはその子ウィンドウである場合は、デスクトップにドロップされたとみなす
                                        isDroppedOnDesktop = isDesktopOrChild;
                                    }
                                }
                                else
                                {
                                    // 現在のアプリケーションのウィンドウの上にドロップされた場合は、デスクトップではない
                                    isDroppedOnDesktop = false;
                                }
                            }
                            else
                            {
                                // dropPositionが取得できない場合でも、ウィンドウ外にドロップされた場合の判定
                                // （droppedOutsideWindowは既に上でチェック済みなので、ここではfalse）
                                if (dragResult == DragDropEffects.None)
                                {
                                    isDroppedOnDesktop = true;
                                }
                                else
                                {
                                    isDroppedOnDesktop = false;
                                }
                            }
                        }
                        
                        // デスクトップにドロップされた場合のみ、新しいウィンドウを作成
                        if (isDroppedOnDesktop)
                        {
                            try
                            {
                                // デスクトップにドロップされた場合、新しいウィンドウを作成（ドロップ位置を指定）
                                // dropPositionがnullの場合は、lastMousePositionを使用、それもnullの場合は位置を指定しない
                                ViewModel.MoveTabToNewWindow(_draggedTab, dropPosition ?? lastMousePosition);
                                
                                // デスクトップドロップ完了後、変数をクリア
                                _draggedTab = null;
                                _draggedTabItem = null;
                                _isTabDragging = false;
                            }
                            catch (System.Exception)
                            {
                                // 例外が発生した場合、位置を指定せずに新しいウィンドウを作成
                                try
                                {
                                    ViewModel.MoveTabToNewWindow(_draggedTab, null);
                                    
                                    // 例外が発生しても、変数をクリア
                                    _draggedTab = null;
                                    _draggedTabItem = null;
                                    _isTabDragging = false;
                                }
                                catch (System.Exception)
                                {
                                    // 失敗した場合も変数をクリア
                                    _draggedTab = null;
                                    _draggedTabItem = null;
                                    _isTabDragging = false;
                                }
                            }
                        }
                    }
                }
                finally
                {
                    // Dropイベントを削除
                    foreach (var tabControl in tabControls)
                    {
                        tabControl.Drop -= dropHandler;
                    }

                    // GiveFeedbackイベントを削除
                    if (_draggedTabItem != null && giveFeedbackHandler != null)
                    {
                        _draggedTabItem.GiveFeedback -= giveFeedbackHandler;
                    }

                    // QueryContinueDragイベントを削除
                    if (_draggedTabItem != null && queryContinueDragHandler != null)
                    {
                        _draggedTabItem.QueryContinueDrag -= queryContinueDragHandler;
                    }
                }

                _isTabDragging = false;
            }
        }

        /// <summary>
        /// TabItemでマウスボタンが離されたときに呼び出されます
        /// </summary>
        private void TabItem_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            // 閉じるボタン上でクリックされた場合は処理しない
            if (IsCloseButton(e.OriginalSource))
                return;

            ReleaseMouseCapture();

            // ドラッグが開始されなかった場合（単なるクリック）は、タブを選択
            if (!_isTabDragging && _draggedTab != null && sender is System.Windows.Controls.TabItem tabItem && tabItem.DataContext is ExplorerTab tab)
            {
                Point currentPoint = e.GetPosition(null);
                double deltaX = currentPoint.X - _tabDragStartPoint.X;
                double deltaY = currentPoint.Y - _tabDragStartPoint.Y;
                double minDragDistance = SystemParameters.MinimumHorizontalDragDistance;

                // 単なるクリックの場合、タブを選択（Math.Absを避けて絶対値比較を最適化）
                if (deltaX >= -minDragDistance && deltaX <= minDragDistance &&
                    deltaY >= -minDragDistance && deltaY <= minDragDistance)
                {
                    // ViewModelプロパティを一度だけ取得してキャッシュ（高速化）
                    var isSplitPaneEnabled = ViewModel.IsSplitPaneEnabled;
                    if (isSplitPaneEnabled)
                    {
                        // TabControlの親要素を検索（高速化：キャッシュを活用）
                        if (!_tabControlCache.TryGetValue(tabItem, out var tabControl))
                        {
                            tabControl = FindAncestor<System.Windows.Controls.TabControl>(tabItem);
                            _tabControlCache[tabItem] = tabControl;
                        }
                        if (tabControl != null)
                        {
                            int column = Grid.GetColumn(tabControl);
                            if (column == 0)
                            {
                                ViewModel.SelectedLeftPaneTab = tab;
                            }
                            else if (column == 2)
                            {
                                ViewModel.SelectedRightPaneTab = tab;
                            }
                        }
                    }
                    else
                    {
                        ViewModel.SelectedTab = tab;
                    }
                }
            }

            // 変数をクリア（ドロップが完了していない場合のみ）
            if (!_isTabDragging)
            {
                _draggedTab = null;
                _draggedTabItem = null;
            }
            _isTabDragging = false;
        }

        /// <summary>
        /// マウスキャプチャを解放します
        /// </summary>
        private new void ReleaseMouseCapture()
        {
            if (_draggedTabItem != null)
            {
                _draggedTabItem.ReleaseMouseCapture();
            }
        }

        /// <summary>
        /// TabControlでドラッグオーバーされたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ドラッグイベント引数</param>
        private void TabControl_DragOver(object sender, DragEventArgs e)
        {
            if (_draggedTabItem == null)
            {
                e.Effects = DragDropEffects.None;
                return;
            }

            e.Effects = DragDropEffects.Move;
            e.Handled = true;
        }

        /// <summary>
        /// TabControlでドロップされたときに呼び出されます（タブの並び替え）
        /// wpfuiの実装を参考に、HitTestを使ってドロップ位置のTabItemを検出
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ドラッグイベント引数</param>
        private void TabControl_Drop(object sender, DragEventArgs e)
        {
            try
            {
                // 早期リターン: 必要な変数がnullの場合は処理しない
                if (_draggedTabItem == null || _draggedTab == null || sender is not System.Windows.Controls.TabControl dropTabControl)
                    return;

                // ViewModelプロパティを一度だけ取得してキャッシュ（高速化）
                var isSplitPaneEnabled = ViewModel.IsSplitPaneEnabled;

                // ドロップ先のコレクションを取得（dropColumnをキャッシュして再利用）
                ObservableCollection<ExplorerTab>? dropTabs;
                int dropColumn = -1;
                if (isSplitPaneEnabled)
                {
                    dropColumn = Grid.GetColumn(dropTabControl);
                    dropTabs = dropColumn == 0 ? ViewModel.LeftPaneTabs
                        : dropColumn == 2 ? ViewModel.RightPaneTabs
                        : null;
                    if (dropTabs == null)
                        return;
                }
                else
                {
                    dropTabs = ViewModel.Tabs;
                }

                // ドラッグ元のコレクションを取得（高速化：キャッシュされたTabControlを使用）
                ObservableCollection<ExplorerTab>? sourceTabs;
                if (isSplitPaneEnabled)
                {
                    // ドラッグ元のTabControlを特定（高速化：キャッシュを活用）
                    System.Windows.Controls.TabControl? sourceTabControl = null;
                    if (_draggedTabItem != null && !_tabControlCache.TryGetValue(_draggedTabItem, out sourceTabControl))
                    {
                        sourceTabControl = FindAncestor<System.Windows.Controls.TabControl>(_draggedTabItem);
                        if (_draggedTabItem != null)
                        {
                            _tabControlCache[_draggedTabItem] = sourceTabControl;
                        }
                    }
                    if (sourceTabControl == null)
                        return;

                    int sourceColumn = Grid.GetColumn(sourceTabControl);
                    sourceTabs = sourceColumn == 0 ? ViewModel.LeftPaneTabs
                        : sourceColumn == 2 ? ViewModel.RightPaneTabs
                        : null;
                    if (sourceTabs == null)
                        return;
                }
                else
                {
                    sourceTabs = ViewModel.Tabs;
                }

                // ペイン間での移動か、同じペイン内での移動かを判定
                bool isCrossPaneMove = sourceTabs != dropTabs;

                // HitTestを一度だけ実行して結果を再利用（高速化）
                Point dropPosition = e.GetPosition(dropTabControl);
                HitTestResult? hitTestResult = VisualTreeHelper.HitTest(dropTabControl, dropPosition);

                // ドロップ位置のTabItemを取得（高速化：一度だけ実行）
                System.Windows.Controls.TabItem? targetTabItem = null;
                ExplorerTab? targetTab = null;
                if (hitTestResult?.VisualHit != null)
                {
                    targetTabItem = FindParent<System.Windows.Controls.TabItem>(hitTestResult.VisualHit);
                    if (targetTabItem != null && targetTabItem.DataContext is ExplorerTab tab)
                    {
                        targetTab = tab;
                    }
                }

                // HitTestが失敗した場合、ドロップ位置から挿入位置を計算（フォールバック）
                if (targetTab == null && dropTabs.Count > 0)
                {
                    // TabPanelを取得してタブの位置を計算
                    var tabPanel = FindChild<System.Windows.Controls.Primitives.TabPanel>(dropTabControl, null);
                    if (tabPanel != null)
                    {
                        // 各TabItemの位置を確認して、ドロップ位置に最も近いタブを探す
                        double minDistance = double.MaxValue;
                        System.Windows.Controls.TabItem? closestTabItem = null;

                        for (int i = 0; i < dropTabControl.Items.Count; i++)
                        {
                            if (dropTabControl.ItemContainerGenerator.ContainerFromIndex(i) is System.Windows.Controls.TabItem item)
                            {
                                var itemPosition = item.TransformToAncestor(dropTabControl).Transform(new Point(0, 0));
                                var itemCenterX = itemPosition.X + item.ActualWidth / 2;
                                var distance = Math.Abs(dropPosition.X - itemCenterX);

                                if (distance < minDistance)
                                {
                                    minDistance = distance;
                                    closestTabItem = item;
                                }
                            }
                        }

                        if (closestTabItem != null && closestTabItem.DataContext is ExplorerTab closestTab)
                        {
                            targetTabItem = closestTabItem;
                            targetTab = closestTab;

                            // ドロップ位置がタブの右側にある場合は、次の位置に挿入
                            var closestPosition = closestTabItem.TransformToAncestor(dropTabControl).Transform(new Point(0, 0));
                            var closestCenterX = closestPosition.X + closestTabItem.ActualWidth / 2;
                            if (dropPosition.X > closestCenterX)
                            {
                                // 右側にドロップされた場合は、次のタブの位置を探す
                                int closestIndex = dropTabs.IndexOf(closestTab);
                                if (closestIndex >= 0 && closestIndex < dropTabs.Count - 1)
                                {
                                    targetTab = dropTabs[closestIndex + 1];
                                }
                            }
                        }
                    }
                }

                // コレクション変更を高速化するため、UI更新を最小限に抑制
                if (isCrossPaneMove)
                {
                    // ペイン間での移動: ドラッグ元から削除して、ドロップ先に追加
                    int sourceIndex = sourceTabs.IndexOf(_draggedTab);
                    if (sourceIndex < 0)
                        return;

                    int insertIndex = dropTabs.Count; // デフォルトは最後に追加

                    // ターゲットタブが見つかった場合は、その位置に挿入
                    if (targetTab != null)
                    {
                        int targetIndex = dropTabs.IndexOf(targetTab);
                        if (targetIndex >= 0)
                        {
                            insertIndex = targetIndex;
                        }
                    }

                    // コレクションの変更を実行（同期的に実行して、SelectedItemが正しく設定されるようにする）
                    sourceTabs.RemoveAt(sourceIndex);
                    dropTabs.Insert(insertIndex, _draggedTab);

                    // ペイン間でタブを移動した場合、移動先のペインをアクティブにする（SelectedItemを設定する前に実行）
                    // これにより、背景色が正しく更新される
                    if (isSplitPaneEnabled && dropColumn >= 0)
                    {
                        // ActivePaneを先に設定（背景色の更新を確実にするため）
                        ViewModel.ActivePane = dropColumn;
                        // ListViewのキャッシュをクリア（タブ移動後、ListViewが再構築される可能性があるため）
                        _cachedLeftListView = null;
                        _cachedRightListView = null;
                        // 背景色を明示的に更新（UI構築完了後に実行、PropertyChangedイベントに依存せず確実に更新するため）
                        Dispatcher.BeginInvoke(new System.Action(() =>
                        {
                            UpdateListViewBackgroundColors();
                        }), System.Windows.Threading.DispatcherPriority.Normal);
                    }

                    // SelectedItemを同期的に設定（ListViewが表示されるようにする）
                    dropTabControl.SelectedItem = _draggedTab;

                    // ViewModelの選択タブも明示的に更新（タブ移動後のペイン判定を正しくするため）
                    // dropColumnは既に取得済み（上記の処理で使用）なので再利用
                    if (isSplitPaneEnabled && dropColumn >= 0)
                    {
                        // switch式を使用してパフォーマンス向上
                        switch (dropColumn)
                        {
                            case 0:
                                ViewModel.SelectedLeftPaneTab = _draggedTab;
                                break;
                            case 2:
                                ViewModel.SelectedRightPaneTab = _draggedTab;
                                break;
                        }
                    }
                }
                else
                {
                    // 同じペイン内での移動
                    int draggedIndex = dropTabs.IndexOf(_draggedTab);
                    if (draggedIndex < 0)
                        return;

                    // ターゲットタブが見つからない場合は、ドロップ位置から挿入位置を決定
                    if (targetTab == null || targetTab == _draggedTab)
                    {
                        // ドロップ位置が最後のタブより右側にある場合は、最後に追加
                        if (dropTabs.Count > 0)
                        {
                            var lastTab = dropTabs[dropTabs.Count - 1];
                            if (dropTabControl.ItemContainerGenerator.ContainerFromItem(lastTab) is System.Windows.Controls.TabItem lastItem)
                            {
                                var lastPosition = lastItem.TransformToAncestor(dropTabControl).Transform(new Point(0, 0));
                                if (dropPosition.X > lastPosition.X + lastItem.ActualWidth)
                                {
                                    // 最後のタブより右側にドロップされた場合は、最後に移動
                                    int lastIndex = dropTabs.Count - 1;
                                    if (draggedIndex != lastIndex)
                                    {
                                        // Moveメソッドを使用してUI更新を1回に削減
                                        dropTabs.Move(draggedIndex, lastIndex);
                                        dropTabControl.SelectedItem = _draggedTab;
                                    }
                                    return;
                                }
                            }
                        }
                        // それ以外の場合は移動しない（既に正しい位置にある可能性がある）
                        return;
                    }

                    // ターゲットタブがドラッグ中のタブと同じ場合は移動しない
                    if (targetTabItem == _draggedTabItem)
                        return;

                    // タブの並び替えを直接実行（最適化: Moveメソッドを使用してUI更新を1回に削減）
                    int targetIndex = dropTabs.IndexOf(targetTab);
                    if (targetIndex < 0 || draggedIndex == targetIndex)
                        return;

                    // コレクションの変更を実行（Moveメソッドで1回のUI更新のみ）
                    // ObservableCollection.Moveは、RemoveAtとInsertを1回の操作として実行し、
                    // CollectionChangedイベントを1回だけ発火するため、パフォーマンスが向上
                    dropTabs.Move(draggedIndex, targetIndex);

                    // SelectedItemを同期的に設定（ListViewが表示されるようにする）
                    // TwoWayバインディングにより、ViewModelも自動的に更新され、TabItemのIsSelectedも自動的に更新される
                    dropTabControl.SelectedItem = _draggedTab;
                }

                // ドロップ完了後に変数をクリア
                _draggedTab = null;
                _draggedTabItem = null;
                _isTabDragging = false;

                // タブが移動したため、ペインキャッシュをクリア（ドライブ要素などのペイン判定を正しく更新するため）
                // 即座にクリアして、UI更新後にも再クリア（ビジュアルツリーが更新されるのを待つ）
                _paneCache.Clear();

                // UI更新後に再度キャッシュをクリア（ビジュアルツリーの更新を確実に反映）
                // デリゲートをキャッシュしてメモリ割り当てを削減
                Dispatcher.BeginInvoke(
                    new System.Action(() => _paneCache.Clear()),
                    System.Windows.Threading.DispatcherPriority.Loaded);
            }
            catch (System.Exception)
            {
                // ドロップ完了後に変数をクリア（例外が発生しても状態をリセット）
                _draggedTab = null;
                _draggedTabItem = null;
                _isTabDragging = false;
            }
        }

        /// <summary>
        /// ビジュアルツリーから指定された型の親要素を検索します（最適化版）
        /// </summary>
        private static T? FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject? current = child;
            while (current != null)
            {
                current = VisualTreeHelper.GetParent(current);
                if (current is T parent)
                {
                    return parent;
                }
            }
            return null;
        }

        /// <summary>
        /// ビジュアルツリーを再帰的に走査して、指定された型の子要素をすべて見つけます
        /// </summary>
        /// <typeparam name="T">検索する型</typeparam>
        /// <param name="parent">親要素</param>
        /// <param name="results">結果を格納するリスト</param>
        private void FindVisualChildrenRecursive<T>(DependencyObject parent, List<T> results) where T : DependencyObject
        {
            if (parent == null)
                return;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T t)
                {
                    results.Add(t);
                }
                FindVisualChildrenRecursive(child, results);
            }
        }

        /// <summary>
        /// TabControl内で指定されたDataContextに対応するTabItemを検索します
        /// </summary>
        /// <param name="tabControl">TabControl</param>
        /// <param name="dataContext">検索するDataContext</param>
        /// <returns>見つかったTabItem、見つからない場合はnull</returns>
        private System.Windows.Controls.TabItem? FindTabItemByDataContext(System.Windows.Controls.TabControl tabControl, object dataContext)
        {
            if (tabControl == null || dataContext == null)
                return null;

            // TabControlのItemsを走査して、DataContextが一致するTabItemを検索（Itemsコレクションの方が効率的）
            foreach (var item in tabControl.Items)
            {
                if (item is System.Windows.Controls.TabItem tabItem && tabItem.DataContext == dataContext)
                    return tabItem;
            }

            // Itemsコレクションで見つからない場合、ビジュアルツリーを走査
            return FindChild<System.Windows.Controls.TabItem>(tabControl, ti => ti.DataContext == dataContext);
        }

        /// <summary>
        /// TabControlでタブが追加されようとしたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">タブ追加イベント引数</param>
        private void TabControl_TabAdding(object sender, TabAddingEventArgs e)
        {
            if (sender is not System.Windows.Controls.TabControl tabControl)
                return;

            // どのTabControlから呼ばれたかを判定
            var column = Grid.GetColumn(tabControl);
            if (ViewModel.IsSplitPaneEnabled)
            {
                if (column == 0)
                {
                    // 左ペイン
                    ViewModel.CreateNewLeftPaneTabCommand.Execute(null);
                }
                else if (column == 2)
                {
                    // 右ペイン
                    ViewModel.CreateNewRightPaneTabCommand.Execute(null);
                }
            }
            else
            {
                // 通常モード
                ViewModel.CreateNewTabCommand.Execute(null);
            }

            // イベントをキャンセル（WPFUI標準のタブ追加を防ぐ）
            e.Cancel = true;
        }

        /// <summary>
        /// TabControlでタブが閉じられようとしたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">タブ閉じるイベント引数</param>
        private void TabControl_TabClosing(object sender, TabClosingEventArgs e)
        {
            if (e.TabItem?.DataContext is ExplorerTab tab)
            {
                // ViewModelのCloseTabCommandを実行
                ViewModel.CloseTabCommand.Execute(tab);
                // イベントをキャンセル（WPFUI標準のタブ閉じる処理を防ぐ）
                e.Cancel = true;
            }
        }

        // Win32 APIの定義
        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        /// <summary>
        /// 指定されたウィンドウがデスクトップまたはその子ウィンドウであるかを確認します
        /// </summary>
        /// <param name="hWnd">確認するウィンドウハンドル</param>
        /// <param name="desktopWindow">デスクトップのウィンドウハンドル</param>
        /// <returns>デスクトップまたはその子ウィンドウである場合はtrue、それ以外はfalse</returns>
        private static bool IsDesktopOrChildWindow(IntPtr hWnd, IntPtr desktopWindow)
        {
            if (hWnd == IntPtr.Zero)
                return false;

            // デスクトップウィンドウ自体である場合
            if (hWnd == desktopWindow)
                return true;

            // 親ウィンドウを再帰的に確認して、デスクトップの子ウィンドウであるかを確認
            IntPtr currentWindow = hWnd;
            const int maxDepth = 10; // 無限ループを防ぐための最大深度
            int depth = 0;

            while (depth < maxDepth)
            {
                IntPtr parent = GetParent(currentWindow);
                if (parent == IntPtr.Zero)
                {
                    // 親ウィンドウがない場合、デスクトップの子ウィンドウではない
                    break;
                }

                if (parent == desktopWindow)
                {
                    // 親ウィンドウがデスクトップである場合、デスクトップの子ウィンドウである
                    return true;
                }

                currentWindow = parent;
                depth++;
            }

            return false;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SetCursor(IntPtr hCursor);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr LoadCursor(IntPtr hInstance, IntPtr lpCursorName);

        [DllImport("user32.dll")]
        private static extern uint GetDoubleClickTime();

        private static readonly IntPtr IDC_HAND = new IntPtr(32649); // Windows標準の手カーソル

        #region リネーム機能

        /// <summary>
        /// 選択中のアイテムのリネームを開始します
        /// </summary>
        private void StartRename()
        {
            var activeTab = GetActiveTab();
            if (activeTab == null)
                return;

            var selectedItem = activeTab.ViewModel.SelectedItem;
            if (selectedItem == null)
                return;

            // アクティブなListViewを取得
            var listView = GetActiveListView();
            if (listView == null)
                return;

            // 選択中のListViewItemを取得
            var listViewItem = listView.ItemContainerGenerator.ContainerFromItem(selectedItem) as System.Windows.Controls.ListViewItem;
            if (listViewItem == null)
                return;

            StartRenameForItem(listViewItem, selectedItem, activeTab);
        }

        /// <summary>
        /// 指定されたListViewで選択中のアイテムのリネームを開始します
        /// </summary>
        /// <param name="listView">対象のListView</param>
        private void StartRename(System.Windows.Controls.ListView listView)
        {
            if (listView == null)
                return;

            // リネーム開始前のアクティブペーンを記録（F2キー時点で記録）
            _activePaneBeforeRename = ViewModel.IsSplitPaneEnabled ? ViewModel.ActivePane : -1;

            // ListViewがどのペインに属しているかを判定してタブを取得
            Models.ExplorerTab? targetTab = null;

            if (ViewModel.IsSplitPaneEnabled)
            {
                var pane = GetPaneForElement(listView);
                if (pane == 0)
                {
                    targetTab = ViewModel.SelectedLeftPaneTab;
                }
                else if (pane == 2)
                {
                    targetTab = ViewModel.SelectedRightPaneTab;
                }
                else
                {
                    targetTab = GetActiveTab();
                }
            }
            else
            {
                targetTab = ViewModel.SelectedTab;
            }

            if (targetTab == null)
                return;

            var selectedItem = targetTab.ViewModel.SelectedItem;
            if (selectedItem == null)
                return;

            // 選択中のListViewItemを取得
            var listViewItem = listView.ItemContainerGenerator.ContainerFromItem(selectedItem) as System.Windows.Controls.ListViewItem;
            if (listViewItem == null)
                return;

            StartRenameForItem(listViewItem, selectedItem, targetTab);
        }

        /// <summary>
        /// 指定されたアイテムのリネームを開始します
        /// </summary>
        /// <param name="listViewItem">リネーム対象のListViewItem</param>
        /// <param name="item">リネーム対象のFileSystemItem</param>
        /// <param name="tab">リネーム対象のタブ（nullの場合はアクティブなタブを使用）</param>
        private void StartRenameForItem(System.Windows.Controls.ListViewItem listViewItem, FileSystemItem item, Models.ExplorerTab? tab = null)
        {
            if (_isRenaming)
                return;

            // TextBlockとTextBoxを検索
            var textBlock = FindVisualChild<System.Windows.Controls.TextBlock>(listViewItem, "FileNameTextBlock");
            var textBox = FindVisualChild<System.Windows.Controls.TextBox>(listViewItem, "FileNameTextBox");

            if (textBlock == null || textBox == null)
                return;

            // リネーム開始前のアクティブペーンは、StartRenameTimerまたはStartRenameで既に記録済み

            _isRenaming = true;
            _renamingItem = item;
            _renameTextBox = textBox;
            _renameTextBlock = textBlock;
            _renamingListViewItem = listViewItem;
            _renamingTab = tab ?? GetActiveTab();

            // TextBlockを非表示にし、TextBoxを表示
            textBlock.Visibility = Visibility.Collapsed;
            textBox.Visibility = Visibility.Visible;
            textBox.Text = item.Name;

            // ファイルの場合、拡張子を除いた部分を選択
            if (!item.IsDirectory && item.Name.Contains('.'))
            {
                var lastDotIndex = item.Name.LastIndexOf('.');
                textBox.Focus();
                textBox.Select(0, lastDotIndex);
            }
            else
            {
                textBox.Focus();
                textBox.SelectAll();
            }

            // イベントハンドラーを追加
            textBox.KeyDown += RenameTextBox_KeyDown;
            textBox.LostFocus += RenameTextBox_LostFocus;
        }

        /// <summary>
        /// リネームを確定します
        /// </summary>
        private async void CommitRename()
        {
            if (!_isRenaming || _renamingItem == null || _renameTextBox == null)
                return;

            var newName = _renameTextBox.Text.Trim();
            var oldName = _renamingItem.Name;
            var oldPath = _renamingItem.FullPath;
            var isDirectory = _renamingItem.IsDirectory;
            var renamingTab = _renamingTab;

            // リネームを開始したペインを記録
            var originalActivePane = ViewModel.ActivePane;

            // 名前が変更されていない場合はキャンセル
            if (string.IsNullOrWhiteSpace(newName) || newName == oldName)
            {
                CancelRename();
                return;
            }

            // 無効な文字が含まれていないかチェック
            var invalidChars = System.IO.Path.GetInvalidFileNameChars();
            if (newName.IndexOfAny(invalidChars) >= 0)
            {
                await ShowRenameErrorMessageAsync("ファイル名に使用できない文字が含まれています。");
                _renameTextBox?.Focus();
                return;
            }

            try
            {
                var directory = System.IO.Path.GetDirectoryName(oldPath);
                if (directory == null)
                {
                    CancelRename();
                    return;
                }

                var newPath = System.IO.Path.Combine(directory, newName);

                // 同じ名前のファイル/フォルダーが存在するかチェック
                if ((System.IO.File.Exists(newPath) || System.IO.Directory.Exists(newPath)) && 
                    !string.Equals(oldPath, newPath, StringComparison.OrdinalIgnoreCase))
                {
                    await ShowRenameErrorMessageAsync("同じ名前のファイルまたはフォルダーが既に存在します。");
                    _renameTextBox?.Focus();
                    return;
                }

                // リネーム実行
                if (isDirectory)
                {
                    System.IO.Directory.Move(oldPath, newPath);
                }
                else
                {
                    System.IO.File.Move(oldPath, newPath);
                }

                // リネーム開始前のアクティブペーンを保存
                var paneToRestore = _activePaneBeforeRename;
                var shouldMaintainPane = ViewModel.IsSplitPaneEnabled && (paneToRestore == 0 || paneToRestore == 2);

                // UIの更新
                EndRename();

                // フォルダーの名前を変更した場合、他のタブのパスを先に更新（パス更新を優先）
                if (isDirectory)
                {
                    UpdateTabPathsAfterRename(oldPath, newPath, renamingTab, paneToRestore);
                }

                // リネームを開始したタブをリフレッシュ（アクティブペーンを保持）
                // パス更新の後に実行して、バックスペースの反応を改善
                if (renamingTab != null)
                {
                    var tabToRefresh = renamingTab;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        // RefreshCommandの実行前後でアクティブペーンを保持
                        if (shouldMaintainPane)
                        {
                            ViewModel.ActivePane = paneToRestore;
                        }
                        
                        tabToRefresh.ViewModel.RefreshCommand.Execute(null);
                        
                        // RefreshCommandの実行後、再度アクティブペーンとフォーカスを設定
                        if (shouldMaintainPane)
                        {
                            ViewModel.ActivePane = paneToRestore;
                            var listView = FindListViewInPane(paneToRestore);
                            listView?.Focus();
                        }
                    }), System.Windows.Threading.DispatcherPriority.Loaded);
                }
            }
            catch (Exception ex)
            {
                await ShowRenameErrorMessageAsync($"名前の変更に失敗しました: {ex.Message}");
                _renameTextBox?.Focus();
            }
        }

        /// <summary>
        /// フォルダー名の変更後、影響を受けるタブのパスを更新します
        /// </summary>
        /// <param name="oldPath">変更前のフォルダーパス</param>
        /// <param name="newPath">変更後のフォルダーパス</param>
        /// <param name="excludeTab">更新から除外するタブ（既にリフレッシュ済みのタブ）</param>
        /// <param name="activePaneToMaintain">維持するアクティブペーン</param>
        private void UpdateTabPathsAfterRename(string oldPath, string newPath, Models.ExplorerTab? excludeTab, int activePaneToMaintain)
        {
            // すべてのタブを取得
            var allTabs = ViewModel.IsSplitPaneEnabled
                ? new List<Models.ExplorerTab>(ViewModel.LeftPaneTabs.Count + ViewModel.RightPaneTabs.Count)
                : new List<Models.ExplorerTab>(ViewModel.Tabs.Count);

            if (ViewModel.IsSplitPaneEnabled)
            {
                allTabs.AddRange(ViewModel.LeftPaneTabs);
                allTabs.AddRange(ViewModel.RightPaneTabs);
            }
            else
            {
                allTabs.AddRange(ViewModel.Tabs);
            }

            // パス更新が必要なタブを収集
            var tabsToRefresh = new List<Models.ExplorerTab>();
            var separator = System.IO.Path.DirectorySeparatorChar;

            foreach (var tab in allTabs)
            {
                // 既にリフレッシュ済みのタブは除外
                if (tab == excludeTab)
                    continue;

                var currentPath = tab.ViewModel.CurrentPath;
                if (string.IsNullOrEmpty(currentPath))
                    continue;

                // パスが変更されたフォルダー自体、またはその配下の場合
                if (currentPath.Equals(oldPath, StringComparison.OrdinalIgnoreCase) ||
                    currentPath.StartsWith(oldPath + separator, StringComparison.OrdinalIgnoreCase))
                {
                    // 新しいパスに更新（CurrentPathプロパティを直接設定してナビゲーションイベントを回避）
                    // パス更新を先に完了させて、バックスペースが正しく動作するようにする
                    var updatedPath = newPath + currentPath.Substring(oldPath.Length);
                    tab.ViewModel.CurrentPath = updatedPath;
                    tabsToRefresh.Add(tab);
                }
            }

            // パス更新が必要なタブがない場合は終了
            if (tabsToRefresh.Count == 0)
                return;

            // RefreshCommandの実行は遅延実行（パス更新は即座に完了しているため、バックスペースが正しく動作する）
            var paneToRestore = activePaneToMaintain;
            var shouldMaintainPane = ViewModel.IsSplitPaneEnabled && (paneToRestore == 0 || paneToRestore == 2);
            
            // パス更新は完了しているため、RefreshCommandの実行を遅延実行してUIスレッドをブロックしない
            Dispatcher.BeginInvoke(new Action(() =>
            {
                // RefreshCommandの実行前後でアクティブペーンを保持
                if (shouldMaintainPane)
                {
                    ViewModel.ActivePane = paneToRestore;
                }

                // すべてのタブを一括でリフレッシュ
                foreach (var tab in tabsToRefresh)
                {
                    tab.ViewModel.RefreshCommand.Execute(null);
                }

                // RefreshCommandの実行後、再度アクティブペーンを設定
                if (shouldMaintainPane)
                {
                    ViewModel.ActivePane = paneToRestore;
                }
            }), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        /// <summary>
        /// リネーム後にフォーカスを復元します
        /// </summary>
        /// <param name="renamingListView">リネームを開始したListView</param>
        /// <param name="originalActivePane">リネーム開始時のアクティブペイン（未使用、互換性のため残す）</param>
        private void RestoreFocusAfterRename(System.Windows.Controls.ListView? renamingListView, int originalActivePane)
        {
            // ListViewが属するペインをアクティブにする
            if (ViewModel.IsSplitPaneEnabled && renamingListView != null)
            {
                var pane = GetPaneForElement(renamingListView);
                if (pane == 0 || pane == 2)
                {
                    ViewModel.ActivePane = pane;
                }
            }

            // ListViewにフォーカスを戻す（即座に実行）
            if (renamingListView != null)
            {
                renamingListView.Focus();
            }
        }

        /// <summary>
        /// リネームエラーメッセージを表示します
        /// </summary>
        /// <param name="message">表示するメッセージ</param>
        private async Task ShowRenameErrorMessageAsync(string message)
        {
            var messageBox = new Wpf.Ui.Controls.MessageBox
            {
                Title = "名前の変更",
                Content = message,
                CloseButtonText = "OK",
                Owner = Window.GetWindow(this)
            };
            await messageBox.ShowDialogAsync();
        }

        /// <summary>
        /// リネームをキャンセルします
        /// </summary>
        private void CancelRename()
        {
            EndRename();
        }

        /// <summary>
        /// リネームモードを終了します
        /// </summary>
        private void EndRename()
        {
            if (_renameTextBox != null)
            {
                _renameTextBox.KeyDown -= RenameTextBox_KeyDown;
                _renameTextBox.LostFocus -= RenameTextBox_LostFocus;
                _renameTextBox.Visibility = Visibility.Collapsed;
            }

            if (_renameTextBlock != null)
            {
                _renameTextBlock.Visibility = Visibility.Visible;
            }

            _isRenaming = false;
            _renamingItem = null;
            _renameTextBox = null;
            _renameTextBlock = null;
            _renamingListViewItem = null;
            _renamingTab = null;
            _activePaneBeforeRename = -1;

            // クリック追跡をリセット（ダブルクリック誤動作防止）
            _lastClickedItem = null;
            _lastClickTime = DateTime.MinValue;

            // リネーム完了時刻を記録（直後のダブルクリック誤動作防止）
            _renameCompletedTime = DateTime.Now;
        }

        /// <summary>
        /// リネームTextBoxでキーが押されたときのハンドラー
        /// </summary>
        private void RenameTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CommitRename();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                CancelRename();
                e.Handled = true;
            }
        }

        /// <summary>
        /// リネームTextBoxがフォーカスを失ったときのハンドラー
        /// </summary>
        private void RenameTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            // 少し遅延させて、別の操作（例：エラーダイアログ）によるフォーカス喪失を処理
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_isRenaming && _renameTextBox != null && !_renameTextBox.IsFocused)
                {
                    CommitRename();
                }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>
        /// 指定された名前の子要素を検索します
        /// </summary>
        private T? FindVisualChild<T>(DependencyObject parent, string name) where T : FrameworkElement
        {
            if (parent == null)
                return null;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild && typedChild.Name == name)
                {
                    return typedChild;
                }

                var result = FindVisualChild<T>(child, name);
                if (result != null)
                    return result;
            }

            return null;
        }

        /// <summary>
        /// アクティブなListViewを取得します
        /// </summary>
        private System.Windows.Controls.ListView? GetActiveListView()
        {
            if (ViewModel.IsSplitPaneEnabled)
            {
                if (ViewModel.ActivePane == 0)
                {
                    return _cachedLeftListView ?? FindListViewInPane(0);
                }
                else
                {
                    return _cachedRightListView ?? FindListViewInPane(2);
                }
            }
            else
            {
                // 単一ペインモードの場合、ビジュアルツリーからListViewを検索
                if (_cachedSinglePaneListView == null)
                {
                    _cachedSinglePaneListView = FindVisualChild<System.Windows.Controls.ListView>(this, "FileListView");
                }
                return _cachedSinglePaneListView;
            }
        }

        #endregion
    }
}

