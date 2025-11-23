using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using FastExplorer.ViewModels.Pages;
using FastExplorer.Services;
using FastExplorer.Models;
using FastExplorer.Helpers;
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

            // クリックされた位置がListViewItem上かどうかを確認（ビジュアルツリー走査を避けるため、OriginalSourceを使用）
            bool isOnItem = false;
            if (e.OriginalSource is DependencyObject source)
            {
                DependencyObject? current = source;
                while (current != null && current != listView)
                {
                    if (current is ListViewItem)
                    {
                        isOnItem = true;
                        break;
                    }
                    current = VisualTreeHelper.GetParent(current);
                }
            }

            // 空領域（アイテムがない場所）をダブルクリックした場合は親フォルダーに移動
            if (!isOnItem)
            {
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
                    var scm = new ShellContextMenu();
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
                var scm = new ShellContextMenu();
                scm.ShowContextMenu(new[] { path }, hWnd, (int)screenPoint.X, (int)screenPoint.Y);
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
                    var scm = new ShellContextMenu();
                    scm.ShowContextMenu(new[] { path }, hWnd, (int)screenPoint.X, (int)screenPoint.Y);
                }
            }
        }
    }
}

