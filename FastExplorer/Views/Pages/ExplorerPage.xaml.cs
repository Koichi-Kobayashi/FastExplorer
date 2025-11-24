using System;
using System.Collections.Generic;
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

        /// <summary>
        /// <see cref="ExplorerPage"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">エクスプローラーページのViewModel</param>
        public ExplorerPage(ExplorerPageViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;

            InitializeComponent();
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

            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(element); i++)
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
                // ViewModelのActivePaneプロパティを更新
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
                // ViewModelのActivePaneプロパティを更新
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
                    targetTab.ViewModel.NavigateToPathCommand.Execute(drive.Path);
                e.Handled = true;
            }
            }
        }

        /// <summary>
        /// 要素がどのペインに属しているかを取得します（分割ペインモードの場合、キャッシュ付き）
        /// </summary>
        /// <param name="element">要素</param>
        /// <returns>左ペインの場合は0、右ペインの場合は2、判定できない場合は-1</returns>
        private int GetPaneForElement(FrameworkElement? element)
        {
            if (element == null || !ViewModel.IsSplitPaneEnabled)
                return -1;
            
            // キャッシュを確認
            if (_paneCache.TryGetValue(element, out var cachedPane))
            {
                return cachedPane;
            }
            
            // 親要素をたどって、Grid.Column="0"（左ペイン）またはGrid.Column="2"（右ペイン）を探す
            var current = element;
            while (current != null)
            {
                // TabControlを探す
                if (current is System.Windows.Controls.TabControl tabControl)
                {
                    // TabControl自体のGrid.Columnを取得
                    var column = System.Windows.Controls.Grid.GetColumn(tabControl);
                    if (column == 0 || column == 2)
                    {
                        // キャッシュに保存
                        _paneCache[element] = column;
                        return column;
                    }
                }
                
                current = VisualTreeHelper.GetParent(current) as FrameworkElement;
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

        // ビジュアルツリー走査の結果をキャッシュ（パフォーマンス向上）
        private readonly Dictionary<FrameworkElement, int> _paneCache = new();
        private readonly Dictionary<System.Windows.Controls.ListView, System.Windows.Controls.ScrollViewer?> _scrollViewerCache = new();

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
                    if (current is Button)
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
            if (sender is FrameworkElement element && element.DataContext is Models.ExplorerTab tab)
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
    }
}

