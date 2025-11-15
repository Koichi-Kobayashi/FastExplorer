using FastExplorer.ViewModels.Pages;
using Wpf.Ui.Abstractions.Controls;

namespace FastExplorer.Views.Pages
{
    /// <summary>
    /// 設定ページを表すクラス
    /// </summary>
    public partial class SettingsPage : INavigableView<SettingsViewModel>
    {
        /// <summary>
        /// 設定ページのViewModelを取得します
        /// </summary>
        public SettingsViewModel ViewModel { get; }

        /// <summary>
        /// <see cref="SettingsPage"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">設定ページのViewModel</param>
        public SettingsPage(SettingsViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;

            InitializeComponent();
        }
    }
}
