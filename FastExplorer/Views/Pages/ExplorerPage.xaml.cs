using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using FastExplorer.ViewModels.Pages;
using FastExplorer.Services;
using FastExplorer.Models;
using FastExplorer.Helpers;
using FastExplorer.ShellContextMenu;
using Wpf.Ui.Abstractions.Controls;

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
        // TabControlのキャッシュ（ビジュアルツリー走査を削減）
        private System.Windows.Controls.TabControl? _cachedLeftTabControl;
        private System.Windows.Controls.TabControl? _cachedRightTabControl;
        private System.Windows.Controls.TabControl? _cachedSingleTabControl;

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
            Loaded += ExplorerPage_Loaded;
        }

        /// <summary>
        /// ページが読み込まれたときに呼び出されます
        /// </summary>
        private void ExplorerPage_Loaded(object sender, RoutedEventArgs e)
        {
            // 初期状態の背景色を更新（UI構築完了後に実行、1回のみで十分）
            Dispatcher.BeginInvoke(new System.Action(() =>
            {
                UpdateListViewBackgroundColors();
            }), System.Windows.Threading.DispatcherPriority.Loaded);
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
                // ActivePaneが変更された場合のみ更新
                UpdateListViewBackgroundColors();
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
        /// 左右のペインのListViewの背景色を更新します
        /// </summary>
        private void UpdateListViewBackgroundColors()
        {
            // WindowSettingsServiceをキャッシュ（パフォーマンス向上）
            if (_cachedWindowSettingsService == null)
            {
                _cachedWindowSettingsService = App.Services.GetService(typeof(Services.WindowSettingsService)) as Services.WindowSettingsService;
            }

            // フォーカスがないペインの背景色をThirdColorCodeから取得（最適化：キャッシュを活用）
            try
            {
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
                            (_cachedUnfocusedBackground ??= CreateDefaultUnfocusedBrush());
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
                        (_cachedUnfocusedBackground ??= CreateDefaultUnfocusedBrush());
                }
            }
            catch
            {
                // エラーが発生した場合は既存のキャッシュを使用、またはデフォルト値
                _cachedUnfocusedBackground ??= CreateDefaultUnfocusedBrush();
            }

            // フォーカスがあるペインの背景色をSecondaryColorCodeから取得（最適化：キャッシュを活用）
            try
            {
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
                            (_cachedFocusedBackground ??= CreateDefaultFocusedBrush());
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
                        (_cachedFocusedBackground ??= CreateDefaultFocusedBrush());
                }
            }
            catch
            {
                // エラーが発生した場合は既存のキャッシュを使用、またはデフォルト値
                _cachedFocusedBackground ??= CreateDefaultFocusedBrush();
            }

            // ViewModelプロパティをキャッシュ（パフォーマンス向上）
            var viewModel = ViewModel;
            var isSplitPaneEnabled = viewModel.IsSplitPaneEnabled;
            
            if (isSplitPaneEnabled)
            {
                // 分割ペインモードの場合
                // ListViewの参照を取得またはキャッシュから取得（最適化：一度だけ検索）
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
                var leftBackground = activePane == 0 ? _cachedFocusedBackground : _cachedUnfocusedBackground;
                var rightBackground = activePane == 2 ? _cachedFocusedBackground : _cachedUnfocusedBackground;

                // 左ペインの背景色を更新（nullチェックと値の変更チェックで不要な更新を回避）
                if (_cachedLeftListView != null && leftBackground != null && _cachedLeftListView.Background != leftBackground)
                {
                    _cachedLeftListView.Background = leftBackground;
                }

                // 右ペインの背景色を更新（nullチェックと値の変更チェックで不要な更新を回避）
                if (_cachedRightListView != null && rightBackground != null && _cachedRightListView.Background != rightBackground)
                {
                    _cachedRightListView.Background = rightBackground;
                }
            }
            else
            {
                // 単一ペインモードの場合、ListViewの背景色をSecondaryColorCodeに設定
                // キャッシュから取得（高速化）
                var singlePaneListView = _cachedSinglePaneListView ?? FindListViewInSinglePane();
                if (singlePaneListView != null && _cachedFocusedBackground != null && singlePaneListView.Background != _cachedFocusedBackground)
                {
                    // SecondaryColorCodeを取得（既に_cachedFocusedBackgroundに設定されている）
                    // 背景色を強制的に更新（値が変更されている場合のみ）
                    singlePaneListView.Background = _cachedFocusedBackground;
                }
            }
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
                    if (current is ListViewItem)
                    {
                        // ListViewItem上でクリックされた場合は処理を続行
                        break;
                    }
                    current = VisualTreeHelper.GetParent(current);
                }
                
                // ListViewItemが見つからなかった場合は空領域
                if (current == null || !(current is ListViewItem))
                {
                    // 空領域（アイテムがない場所）をダブルクリックした場合は親フォルダーに移動
                    targetTab.ViewModel.NavigateToParentCommand.Execute(null);
                    e.Handled = true;
                    return;
                }
            }
            else
            {
                // OriginalSourceがDependencyObjectでない場合は空領域として扱う
                targetTab.ViewModel.NavigateToParentCommand.Execute(null);
                e.Handled = true;
                return;
            }

            var selectedItem = targetTab.ViewModel.SelectedItem;
            
            // ディレクトリの場合は、新しいディレクトリに移動した後にスクロール位置を0に戻す
            if (selectedItem != null && selectedItem.IsDirectory)
            {
                targetTab.ViewModel.NavigateToItemCommand.Execute(selectedItem);
                
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
                targetTab.ViewModel.NavigateToItemCommand.Execute(selectedItem);
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
                ViewModel.ActivePane = pane;
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
                ViewModel.ActivePane = pane;
            }
        }

        /// <summary>
        /// リストビューでキーが押されたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">キーイベント引数</param>
        private void ListView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Back)
                return;

            var listView = sender as System.Windows.Controls.ListView;
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

            if (sender is ListViewItem listViewItem && listViewItem.DataContext is FileSystemItem item)
            {
                _draggedItem = item;
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
            System.Diagnostics.Debug.WriteLine("[ExplorerPage] ListViewItem_MouseRightButtonUp呼び出されました");

            // DataContextからファイルパスを取得
            if (sender is ListViewItem listViewItem && listViewItem.DataContext is FileSystemItem fileItem)
            {
                var filePath = fileItem.FullPath;
                System.Diagnostics.Debug.WriteLine($"[ExplorerPage] ListViewItem右クリック: path={filePath}");

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
            System.Diagnostics.Debug.WriteLine("[ExplorerPage] ListView_MouseRightButtonUp呼び出されました");

            // ListViewItem上でクリックされた場合は処理しない
            if (e.OriginalSource is DependencyObject source)
            {
                DependencyObject? current = source;
                while (current != null)
                {
                    if (current is ListViewItem)
                    {
                        System.Diagnostics.Debug.WriteLine("[ExplorerPage] ListViewItem上でクリックされたため、ListView_MouseRightButtonUpをスキップ");
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
                System.Diagnostics.Debug.WriteLine("[ExplorerPage] ターゲットタブがnullです（空領域）");
                return;
            }

            // 現在のパスを取得
            var path = targetTab.ViewModel?.CurrentPath;

            // パスが空の場合はホームディレクトリを使用
            if (string.IsNullOrEmpty(path))
            {
                path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                System.Diagnostics.Debug.WriteLine($"[ExplorerPage] ListView右クリック（空領域、ホーム）: path={path}");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[ExplorerPage] ListView右クリック（空領域）: path={path}");
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

            // イベントを処理済みとしてマーク（左クリックイベントが発火しないようにする）
            e.Handled = true;

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

            if (tab != null)
            {
                // タブのパスを取得
                var path = tab.ViewModel?.CurrentPath;

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
                DragDrop.DoDragDrop(_draggedTabItem, _draggedTabItem, DragDropEffects.Move);
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
            // 早期リターン: 必要な変数がnullの場合は処理しない
            if (_draggedTabItem == null || _draggedTab == null || sender is not System.Windows.Controls.TabControl dropTabControl)
                return;

            // ViewModelプロパティを一度だけ取得してキャッシュ（高速化）
            var isSplitPaneEnabled = ViewModel.IsSplitPaneEnabled;

            // ドロップ先のコレクションを取得（dropColumnをキャッシュして再利用）
            ObservableCollection<ExplorerTab> dropTabs;
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
            ObservableCollection<ExplorerTab> sourceTabs;
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
    }
}

