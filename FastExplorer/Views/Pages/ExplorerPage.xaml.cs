using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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
            if (ViewModel.SelectedTab != null)
            {
                var listView = sender as System.Windows.Controls.ListView;
                var selectedItem = ViewModel.SelectedTab.ViewModel.SelectedItem;
                
                // ディレクトリの場合は、新しいディレクトリに移動した後にスクロール位置を0に戻す
                if (selectedItem != null && selectedItem.IsDirectory)
                {
                    ViewModel.SelectedTab.ViewModel.NavigateToItemCommand.Execute(selectedItem);
                    
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
                    ViewModel.SelectedTab.ViewModel.NavigateToItemCommand.Execute(selectedItem);
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
            if (e.Key == Key.Back && ViewModel.SelectedTab != null)
            {
                ViewModel.SelectedTab.ViewModel.NavigateToParentCommand.Execute(null);
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
            if (ViewModel.SelectedTab == null)
                return;

            // マウスのバックボタン（XButton1）が押された場合
            if (e.ChangedButton == MouseButton.XButton1)
            {
                ViewModel.SelectedTab.ViewModel.NavigateToParentCommand.Execute(null);
                e.Handled = true;
            }
            // マウスの進むボタン（XButton2）が押された場合
            else if (e.ChangedButton == MouseButton.XButton2)
            {
                ViewModel.SelectedTab.ViewModel.NavigateForwardCommand.Execute(null);
                e.Handled = true;
            }
        }
    }
}

