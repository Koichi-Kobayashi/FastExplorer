using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using FastExplorer.ViewModels.Pages;
using FastExplorer.Services;
using FastExplorer.Models;
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
            var activeTab = GetActiveTab();
            if (activeTab != null)
            {
                var listView = sender as System.Windows.Controls.ListView;
                var selectedItem = activeTab.ViewModel.SelectedItem;
                
                // ディレクトリの場合は、新しいディレクトリに移動した後にスクロール位置を0に戻す
                if (selectedItem != null && selectedItem.IsDirectory)
                {
                    activeTab.ViewModel.NavigateToItemCommand.Execute(selectedItem);
                    
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
                    activeTab.ViewModel.NavigateToItemCommand.Execute(selectedItem);
                }
            }
        }

        /// <summary>
        /// DependencyObjectからScrollViewerを取得します
        /// </summary>
        /// <param name="element">要素</param>
        /// <returns>ScrollViewer、見つからない場合はnull</returns>
        private static System.Windows.Controls.ScrollViewer? GetScrollViewer(System.Windows.DependencyObject element)
        {
            if (element == null)
                return null;

            if (element is System.Windows.Controls.ScrollViewer scrollViewer)
                return scrollViewer;

            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(element); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(element, i);
                var result = GetScrollViewer(child);
                if (result != null)
                    return result;
            }

            return null;
        }

        /// <summary>
        /// リストビューでキーが押されたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">キーイベント引数</param>
        private void ListView_KeyDown(object sender, KeyEventArgs e)
        {
            var activeTab = GetActiveTab();
            if (e.Key == Key.Back && activeTab != null)
            {
                activeTab.ViewModel.NavigateToParentCommand.Execute(null);
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
            var activeTab = GetActiveTab();
            if (activeTab == null)
                return;

            // マウスのバックボタン（XButton1）が押された場合
            if (e.ChangedButton == MouseButton.XButton1)
            {
                activeTab.ViewModel.NavigateToParentCommand.Execute(null);
                e.Handled = true;
            }
            // マウスの進むボタン（XButton2）が押された場合
            else if (e.ChangedButton == MouseButton.XButton2)
            {
                activeTab.ViewModel.NavigateForwardCommand.Execute(null);
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
        /// 要素がどのペインに属しているかを取得します（分割ペインモードの場合）
        /// </summary>
        /// <param name="element">要素</param>
        /// <returns>左ペインの場合は0、右ペインの場合は2、判定できない場合は-1</returns>
        private int GetPaneForElement(FrameworkElement? element)
        {
            if (element == null || !ViewModel.IsSplitPaneEnabled)
                return -1;
            
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
                        return column;
                    }
                }
                
                current = VisualTreeHelper.GetParent(current) as FrameworkElement;
            }
            
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

        /// <summary>
        /// コンテキストメニューが開かれたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ルーティングイベント引数</param>
        private void ContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.ContextMenu contextMenu && ViewModel != null)
            {
                // すべてのMenuItemを確認
                foreach (var item in contextMenu.Items)
                {
                    if (item is System.Windows.Controls.MenuItem menuItem)
                    {
                        // アイコンで分割ペインメニューアイテムを識別
                        if (menuItem.Icon is Wpf.Ui.Controls.SymbolIcon symbolIcon && symbolIcon.Symbol == Wpf.Ui.Controls.SymbolRegular.SplitHorizontal24)
                        {
                            // メニューテキストを更新
                            var converter = new Helpers.BooleanToSplitPaneMenuTextConverter();
                            var header = converter.Convert(ViewModel.IsSplitPaneEnabled, typeof(string), null, System.Globalization.CultureInfo.CurrentCulture) ?? "分割ペインを有効にする";
                            menuItem.Header = header;
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 分割ペインメニューアイテムがクリックされたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ルーティングイベント引数</param>
        private void SplitPaneMenuItem_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.ToggleSplitPaneCommand.Execute(null);
        }

        // ドラッグ&ドロップ用の変数
        private Point _dragStartPoint;
        private bool _isDragging = false;
        private FileSystemItem? _draggedItem = null;

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
                // 同じフォルダーへの移動は無視
                if (string.Equals(draggedItem.FullPath, targetItem.FullPath, StringComparison.OrdinalIgnoreCase))
                    return;

                // 親フォルダーへの移動も無視（無限ループを防ぐ）
                var parentPath = System.IO.Path.GetDirectoryName(draggedItem.FullPath);
                if (string.Equals(parentPath, targetItem.FullPath, StringComparison.OrdinalIgnoreCase))
                    return;

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
    }
}

