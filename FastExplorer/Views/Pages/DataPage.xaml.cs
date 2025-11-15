using FastExplorer.ViewModels.Pages;
using Wpf.Ui.Abstractions.Controls;

namespace FastExplorer.Views.Pages
{
    /// <summary>
    /// データページを表すクラス
    /// </summary>
    public partial class DataPage : INavigableView<DataViewModel>
    {
        /// <summary>
        /// データページのViewModelを取得します
        /// </summary>
        public DataViewModel ViewModel { get; }

        /// <summary>
        /// <see cref="DataPage"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">データページのViewModel</param>
        public DataPage(DataViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;

            InitializeComponent();
        }
    }
}
