using System.Windows.Controls;
using FastExplorer.ViewModels.Pages;

namespace FastExplorer.Views.Pages.SettingsPage
{
    /// <summary>
    /// 外観設定ページを表すクラス
    /// </summary>
    public partial class AppearanceSettingsPage : UserControl
    {
        /// <summary>
        /// 設定ページのViewModelを取得または設定します
        /// </summary>
        public SettingsViewModel ViewModel { get; }

        /// <summary>
        /// <see cref="AppearanceSettingsPage"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">設定ページのViewModel</param>
        public AppearanceSettingsPage(SettingsViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;
            InitializeComponent();
        }
    }
}

