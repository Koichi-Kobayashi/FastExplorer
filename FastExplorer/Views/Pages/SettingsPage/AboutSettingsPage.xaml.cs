using System.Windows.Controls;
using FastExplorer.ViewModels.Pages;

namespace FastExplorer.Views.Pages.SettingsPage
{
    /// <summary>
    /// FastExploreについてのページを表すクラス
    /// </summary>
    public partial class AboutSettingsPage : UserControl
    {
        /// <summary>
        /// 設定ページのViewModelを取得または設定します
        /// </summary>
        public SettingsViewModel ViewModel { get; }

        /// <summary>
        /// <see cref="AboutSettingsPage"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">設定ページのViewModel</param>
        public AboutSettingsPage(SettingsViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;
            InitializeComponent();
        }
    }
}

