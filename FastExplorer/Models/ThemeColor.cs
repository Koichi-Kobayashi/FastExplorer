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
        /// カラーコード（サード：分割ペイン用の背景色）
        /// </summary>
        public string ThirdColorCode { get; set; } = "#FFFFFF";

        /// <summary>
        /// カラーブラシ
        /// </summary>
        public Brush Color { get; set; } = new SolidColorBrush(Colors.White);

        /// <summary>
        /// テーマカラーから薄い色を計算します（白とのブレンド）
        /// </summary>
        /// <param name="mainColor">メインカラー</param>
        /// <param name="blendFactor">白とのブレンド係数（0.0=元の色、1.0=白）。デフォルトは0.7</param>
        /// <returns>計算された薄い色</returns>
        public static Color CalculateLightColor(Color mainColor, double blendFactor = 0.7)
        {
            return Color.FromRgb(
                (byte)(mainColor.R + (255 - mainColor.R) * blendFactor),
                (byte)(mainColor.G + (255 - mainColor.G) * blendFactor),
                (byte)(mainColor.B + (255 - mainColor.B) * blendFactor));
        }
    }
}
