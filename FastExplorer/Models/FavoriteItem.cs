namespace FastExplorer.Models
{
    /// <summary>
    /// お気に入りアイテムを表すクラス
    /// </summary>
    public class FavoriteItem
    {
        /// <summary>
        /// お気に入りの一意のIDを取得または設定します
        /// </summary>
        public string Id { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// お気に入りの表示名を取得または設定します
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// お気に入りのパスを取得または設定します
        /// </summary>
        public string Path { get; set; } = string.Empty;
    }
}


