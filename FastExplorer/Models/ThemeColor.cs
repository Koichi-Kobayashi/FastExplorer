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
        public static System.Windows.Media.Color CalculateLightColor(System.Windows.Media.Color mainColor, double blendFactor = 0.7)
        {
            return System.Windows.Media.Color.FromRgb(
                (byte)(mainColor.R + (255 - mainColor.R) * blendFactor),
                (byte)(mainColor.G + (255 - mainColor.G) * blendFactor),
                (byte)(mainColor.B + (255 - mainColor.B) * blendFactor));
        }

        /// <summary>
        /// 定義済みのテーマカラーのリストを取得します
        /// </summary>
        /// <returns>テーマカラーのリスト</returns>
        public static IReadOnlyList<ThemeColor> GetDefaultThemeColors()
        {
            return new[]
            {
                new ThemeColor { Name = "既定", NameEn = "Default", ColorCode = "#F5F5F5", SecondaryColorCode = "#FCFCFC", ThirdColorCode = "#F5F5F5", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F5F5F5")) },
                new ThemeColor { Name = "イエロー ゴールド", NameEn = "Yellow Gold", ColorCode = "#FCEECA", SecondaryColorCode = "#FEFAEF", ThirdColorCode = "#FEF9E8", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FCEECA")) },
                new ThemeColor { Name = "オレンジ ブライト", NameEn = "Orange Bright", ColorCode = "#FADDCC", SecondaryColorCode = "#FEF5F0", ThirdColorCode = "#FEF0E5", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FADDCC")) },
                new ThemeColor { Name = "ブリック レッド", NameEn = "Brick Red", ColorCode = "#F3D4D5", SecondaryColorCode = "#FBF2F2", ThirdColorCode = "#FEEBEB", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F3D4D5")) },
                new ThemeColor { Name = "モダン レッド", NameEn = "Modern Red", ColorCode = "#FCD7D7", SecondaryColorCode = "#FEF3F3", ThirdColorCode = "#FEEBEB", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FCD7D7")) },
                new ThemeColor { Name = "レッド", NameEn = "Red", ColorCode = "#F8CADC", SecondaryColorCode = "#FDEFF5", ThirdColorCode = "#FEEBEB", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8CADC")) },
                new ThemeColor { Name = "ローズ ブライト", NameEn = "Rose Bright", ColorCode = "#F8CADC", SecondaryColorCode = "#FDEFF5", ThirdColorCode = "#FEEBEB", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8CADC")) },
                new ThemeColor { Name = "ブルー", NameEn = "Blue", ColorCode = "#CAE2F4", SecondaryColorCode = "#EFF6FC", ThirdColorCode = "#E8F2FA", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAE2F4")) },
                new ThemeColor { Name = "アイリス パステル", NameEn = "Iris Pastel", ColorCode = "#E4DEEE", SecondaryColorCode = "#F7F5FA", ThirdColorCode = "#F2EEF7", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E4DEEE")) },
                new ThemeColor { Name = "バイオレット レッド ライト", NameEn = "Violet Red Light", ColorCode = "#EDD8F0", SecondaryColorCode = "#FAF3FB", ThirdColorCode = "#F5EBF7", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EDD8F0")) },
                new ThemeColor { Name = "クール ブルー ブライト", NameEn = "Cool Blue Bright", ColorCode = "#CAE8EF", SecondaryColorCode = "#EFF8FA", ThirdColorCode = "#E8F4F8", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAE8EF")) },
                new ThemeColor { Name = "シーフォーム", NameEn = "Seafoam", ColorCode = "#CAEEF0", SecondaryColorCode = "#EFFAFB", ThirdColorCode = "#E8F7F8", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAEEF0")) },
                new ThemeColor { Name = "ミント ライト", NameEn = "Mint Light", ColorCode = "#CAEDE7", SecondaryColorCode = "#EFFAF8", ThirdColorCode = "#E8F6F3", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAEDE7")) },
                new ThemeColor { Name = "グレー", NameEn = "Gray", ColorCode = "#E2E1E1", SecondaryColorCode = "#F6F6F6", ThirdColorCode = "#F0F0F0", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E1E1")) },
                new ThemeColor { Name = "グリーン", NameEn = "Green", ColorCode = "#CDE2CD", SecondaryColorCode = "#F0F6F0", ThirdColorCode = "#E8F2E8", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CDE2CD")) },
                new ThemeColor { Name = "オーバーキャスト", NameEn = "Overcast", ColorCode = "#E1E1E1", SecondaryColorCode = "#F6F6F6", ThirdColorCode = "#F0F0F0", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E1E1E1")) },
                new ThemeColor { Name = "ストーム", NameEn = "Storm", ColorCode = "#D9D9D8", SecondaryColorCode = "#F4F4F3", ThirdColorCode = "#EDEDEC", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D9D9D8")) },
                new ThemeColor { Name = "ブルー グレー", NameEn = "Blue Gray", ColorCode = "#DFE2E3", SecondaryColorCode = "#F5F6F7", ThirdColorCode = "#F0F1F2", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DFE2E3")) },
                new ThemeColor { Name = "グレー ダーク", NameEn = "Gray Dark", ColorCode = "#D9DADB", SecondaryColorCode = "#F4F4F4", ThirdColorCode = "#EDEDED", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D9DADB")) },
                new ThemeColor { Name = "カモフラージュ", NameEn = "Camouflage", ColorCode = "#E3E1DD", SecondaryColorCode = "#F7F6F5", ThirdColorCode = "#F2F1EF", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E3E1DD")) }
            };
        }
    }
}
