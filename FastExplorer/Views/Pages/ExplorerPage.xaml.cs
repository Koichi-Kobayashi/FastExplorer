using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using FastExplorer.ViewModels.Pages;
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
    }
}

