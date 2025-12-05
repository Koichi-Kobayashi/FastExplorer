namespace FastExplorer.Models
{
    /// <summary>
    /// 背景画像の調整方法を表す列挙型
    /// </summary>
    public enum BackgroundImageStretch
    {
        /// <summary>
        /// ウィンドウのサイズに合わせる
        /// </summary>
        FitToWindow,

        /// <summary>
        /// ウィンドウを埋める
        /// </summary>
        Fill,

        /// <summary>
        /// タイル状に繰り返す
        /// </summary>
        Tile,

        /// <summary>
        /// 元のサイズで表示
        /// </summary>
        None
    }
}
