using System.Windows;
using System.Windows.Controls;
using FastExplorer.ViewModels.Pages;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

namespace FastExplorer.Views.Windows
{
    /// <summary>
    /// 設定ウィンドウを表すクラス
    /// </summary>
    public partial class SettingsWindow : FluentWindow
    {
        /// <summary>
        /// 設定ページのViewModelを取得します
        /// </summary>
        public SettingsViewModel ViewModel { get; }

        /// <summary>
        /// <see cref="SettingsWindow"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">設定ページのViewModel</param>
        public SettingsWindow(SettingsViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;

            InitializeComponent();
            
            // ウィンドウ読み込み時にViewModelを初期化
            Loaded += async (s, e) =>
            {
                await ViewModel.OnNavigatedToAsync();
                // システムテーマの監視を設定
                SystemThemeWatcher.Watch(this);
            };
            
            // ウィンドウを閉じる時にViewModelのクリーンアップ
            Closing += async (s, e) =>
            {
                await ViewModel.OnNavigatedFromAsync();
            };
        }

        /// <summary>
        /// ナビゲーションボタンがクリックされたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ルーティングイベント引数</param>
        private void NavigationButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Wpf.Ui.Controls.Button button || button.Tag is not string tag)
                return;

            // すべてのボタンの外観をSecondaryにリセット
            GeneralButton.Appearance = ControlAppearance.Secondary;
            AppearanceButton.Appearance = ControlAppearance.Secondary;
            LayoutButton.Appearance = ControlAppearance.Secondary;
            OperationButton.Appearance = ControlAppearance.Secondary;
            TagButton.Appearance = ControlAppearance.Secondary;
            DevToolsButton.Appearance = ControlAppearance.Secondary;
            AdvancedButton.Appearance = ControlAppearance.Secondary;
            AboutButton.Appearance = ControlAppearance.Secondary;

            // クリックされたボタンをPrimaryに設定
            button.Appearance = ControlAppearance.Primary;

            // TODO: タグに応じてコンテンツエリアを切り替える実装
            // 現在は全般ページのみ表示
        }
    }
}
