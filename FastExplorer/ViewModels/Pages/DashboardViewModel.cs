namespace FastExplorer.ViewModels.Pages
{
    /// <summary>
    /// ダッシュボードページのViewModel
    /// </summary>
    public partial class DashboardViewModel : ObservableObject
    {
        #region プロパティ

        /// <summary>
        /// カウンターの値を取得または設定します
        /// </summary>
        [ObservableProperty]
        private int _counter = 0;

        #endregion

        #region コマンド

        /// <summary>
        /// カウンターをインクリメントします
        /// </summary>
        [RelayCommand]
        private void OnCounterIncrement()
        {
            Counter++;
        }

        #endregion
    }
}
