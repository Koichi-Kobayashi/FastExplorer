namespace FastExplorer.Models
{
    /// <summary>
    /// エクスプローラーのタブを表すクラス
    /// </summary>
    public partial class ExplorerTab : ObservableObject
    {
        /// <summary>
        /// タブの一意のIDを取得または設定します
        /// </summary>
        public string Id { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// タブのタイトルを取得または設定します
        /// </summary>
        [ObservableProperty]
        private string _title = string.Empty;

        /// <summary>
        /// 現在表示しているパスを取得または設定します
        /// </summary>
        public string CurrentPath { get; set; } = string.Empty;

        /// <summary>
        /// タブに関連付けられたViewModelを取得または設定します
        /// </summary>
        public ViewModels.Pages.ExplorerViewModel ViewModel { get; set; } = null!;
    }
}

