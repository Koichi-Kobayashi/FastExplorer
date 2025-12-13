using System.Windows.Controls;
using FastExplorer.ViewModels.Pages;

namespace FastExplorer.Views.Pages.SettingsPage
{
    /// <summary>
    /// 全般設定ページを表すクラス
    /// </summary>
    public partial class GeneralSettingsPage : UserControl
    {
        /// <summary>
        /// 設定ページのViewModelを取得または設定します
        /// </summary>
        public SettingsViewModel ViewModel { get; }

        /// <summary>
        /// <see cref="GeneralSettingsPage"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">設定ページのViewModel</param>
        public GeneralSettingsPage(SettingsViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;
            InitializeComponent();
        }
    }
}

