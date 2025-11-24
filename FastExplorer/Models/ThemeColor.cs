using System;
using System.Collections.Generic;
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

        // パフォーマンス最適化：定義済みテーマカラーのキャッシュ
        private static ThemeColor[]? _cachedDefaultThemeColors;

        /// <summary>
        /// テーマカラーから薄い色を計算します（白とのブレンド）
        /// </summary>
        /// <param name="mainColor">メインカラー</param>
        /// <param name="blendFactor">白とのブレンド係数（0.0=元の色、1.0=白）。デフォルトは0.7</param>
        /// <returns>計算された薄い色</returns>
        public static System.Windows.Media.Color CalculateLightColor(System.Windows.Media.Color mainColor, double blendFactor = 0.7)
        {
            return System.Windows.Media.Color.FromRgb(
                (byte)(mainColor.R + (255 - mainColor.R) * blendFactor),
                (byte)(mainColor.G + (255 - mainColor.G) * blendFactor),
                (byte)(mainColor.B + (255 - mainColor.B) * blendFactor));
        }

        /// <summary>
        /// 定義済みのテーマカラーのリストを取得します（キャッシュ付き）
        /// </summary>
        /// <returns>テーマカラーの配列</returns>
        public static ThemeColor[] GetDefaultThemeColors()
        {
            // キャッシュが存在する場合はそれを返す
            if (_cachedDefaultThemeColors != null)
            {
                return _cachedDefaultThemeColors;
            }

            // 初回のみ配列を作成してキャッシュ
            _cachedDefaultThemeColors = new[]
            {
                CreateThemeColor("既定", "Default", "#F5F5F5", "#FCFCFC", "#F5F5F5"),
                CreateThemeColor("イエロー ゴールド", "Yellow Gold", "#FCEECA", "#FEFAEF", "#FEF9E8"),
                CreateThemeColor("オレンジ ブライト", "Orange Bright", "#FADDCC", "#FEF5F0", "#FEF0E5"),
                CreateThemeColor("ブリック レッド", "Brick Red", "#F3D4D5", "#FBF2F2", "#FEEBEB"),
                CreateThemeColor("モダン レッド", "Modern Red", "#FCD7D7", "#FEF3F3", "#FEEBEB"),
                CreateThemeColor("レッド", "Red", "#F8CADC", "#FDEFF5", "#FEEBEB"),
                CreateThemeColor("ローズ ブライト", "Rose Bright", "#F8CADC", "#FDEFF5", "#FEEBEB"),
                CreateThemeColor("ブルー", "Blue", "#CAE2F4", "#EFF6FC", "#E8F2FA"),
                CreateThemeColor("アイリス パステル", "Iris Pastel", "#E4DEEE", "#F7F5FA", "#F2EEF7"),
                CreateThemeColor("バイオレット レッド ライト", "Violet Red Light", "#EDD8F0", "#FAF3FB", "#F5EBF7"),
                CreateThemeColor("クール ブルー ブライト", "Cool Blue Bright", "#CAE8EF", "#EFF8FA", "#E8F4F8"),
                CreateThemeColor("シーフォーム", "Seafoam", "#CAEEF0", "#EFFAFB", "#E8F7F8"),
                CreateThemeColor("ミント ライト", "Mint Light", "#CAEDE7", "#EFFAF8", "#E8F6F3"),
                CreateThemeColor("グレー", "Gray", "#E2E1E1", "#F6F6F6", "#F0F0F0"),
                CreateThemeColor("グリーン", "Green", "#CDE2CD", "#F0F6F0", "#E8F2E8"),
                CreateThemeColor("オーバーキャスト", "Overcast", "#E1E1E1", "#F6F6F6", "#F0F0F0"),
                CreateThemeColor("ストーム", "Storm", "#D9D9D8", "#F4F4F3", "#EDEDEC"),
                CreateThemeColor("ブルー グレー", "Blue Gray", "#DFE2E3", "#F5F6F7", "#F0F1F2"),
                CreateThemeColor("グレー ダーク", "Gray Dark", "#D9DADB", "#F4F4F4", "#EDEDED"),
                CreateThemeColor("カモフラージュ", "Camouflage", "#E3E1DD", "#F7F6F5", "#F2F1EF")
                };

            return _cachedDefaultThemeColors;
        }

        /// <summary>
        /// テーマカラーを作成します（ヘルパーメソッド）
        /// </summary>
        private static ThemeColor CreateThemeColor(string name, string nameEn, string colorCode, string secondaryColorCode, string thirdColorCode)
        {
            // FastColorConverterを使用して高速化（ColorConverter.ConvertFromStringより高速）
            var color = Helpers.FastColorConverter.ParseHexColor(colorCode);
            return new ThemeColor
            {
                Name = name,
                NameEn = nameEn,
                ColorCode = colorCode,
                SecondaryColorCode = secondaryColorCode,
                ThirdColorCode = thirdColorCode,
                Color = new SolidColorBrush(color)
            };
        }
    }
}
