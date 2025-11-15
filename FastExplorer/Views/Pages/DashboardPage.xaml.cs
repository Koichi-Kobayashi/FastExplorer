using FastExplorer.ViewModels.Pages;
using Wpf.Ui.Abstractions.Controls;

namespace FastExplorer.Views.Pages
{
    /// <summary>
    /// ダッシュボードページを表すクラス
    /// </summary>
    public partial class DashboardPage : INavigableView<DashboardViewModel>
    {
        /// <summary>
        /// ダッシュボードページのViewModelを取得します
        /// </summary>
        public DashboardViewModel ViewModel { get; }

        /// <summary>
        /// <see cref="DashboardPage"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">ダッシュボードページのViewModel</param>
        public DashboardPage(DashboardViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;

            InitializeComponent();
        }
    }
}
