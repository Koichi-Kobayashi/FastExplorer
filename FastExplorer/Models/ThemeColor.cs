using System.Windows.Media;

namespace FastExplorer.Models
{
    /// <summary>
    /// テーマカラーを表すクラス
    /// </summary>
    public class ThemeColor
    {
        /// <summary>
        /// 名前（日本語）
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 名前（英語）
        /// </summary>
        public string NameEn { get; set; } = string.Empty;

        /// <summary>
        /// カラーコード（メイン）
        /// </summary>
        public string ColorCode { get; set; } = "#FFFFFF";

        /// <summary>
        /// カラーコード（セカンダリ）
        /// </summary>
        public string SecondaryColorCode { get; set; } = "#FFFFFF";

        /// <summary>
        /// カラーブラシ
        /// </summary>
        public Brush Color { get; set; } = new SolidColorBrush(Colors.White);
    }
}
