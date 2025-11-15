namespace FastExplorer.ViewModels.Pages
{
    /// <summary>
    /// ダッシュボードページのViewModel
    /// </summary>
    public partial class DashboardViewModel : ObservableObject
    {
        /// <summary>
        /// カウンターの値を取得または設定します
        /// </summary>
        [ObservableProperty]
        private int _counter = 0;

        /// <summary>
        /// カウンターをインクリメントします
        /// </summary>
        [RelayCommand]
        private void OnCounterIncrement()
        {
            Counter++;
        }
    }
}
