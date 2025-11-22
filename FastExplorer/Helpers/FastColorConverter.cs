using System;
using System.Windows.Media;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// 高速な色変換を行うヘルパークラス
    /// ColorConverter.ConvertFromStringより高速（文字列解析を最適化）
    /// </summary>
    public static class FastColorConverter
    {
        /// <summary>
        /// 16進数文字列をColorに変換します（最適化版）
        /// #RRGGBB または #AARRGGBB 形式をサポート
        /// </summary>
        public static Color ParseHexColor(string hex)
        {
            if (string.IsNullOrEmpty(hex))
                return Colors.Transparent;

            // #を削除
            var startIndex = hex[0] == '#' ? 1 : 0;
            var length = hex.Length - startIndex;

            if (length == 6)
            {
                // #RRGGBB形式（不透明度は255）
                return Color.FromRgb(
                    ParseHexByte(hex, startIndex),
                    ParseHexByte(hex, startIndex + 2),
                    ParseHexByte(hex, startIndex + 4));
            }
            else if (length == 8)
            {
                // #AARRGGBB形式
                return Color.FromArgb(
                    ParseHexByte(hex, startIndex),
                    ParseHexByte(hex, startIndex + 2),
                    ParseHexByte(hex, startIndex + 4),
                    ParseHexByte(hex, startIndex + 6));
            }

            // フォールバック：標準のColorConverterを使用
            return (Color)System.Windows.Media.ColorConverter.ConvertFromString(hex);
        }

        /// <summary>
        /// 2文字の16進数をbyteに変換（インライン最適化）
        /// </summary>
        private static byte ParseHexByte(string hex, int startIndex)
        {
            // 高速化：文字列操作を最小限に
            var c1 = hex[startIndex];
            var c2 = hex[startIndex + 1];

            // 0-9, A-F, a-f を直接数値に変換（分岐を最小化）
            int v1 = c1 >= '0' && c1 <= '9' ? c1 - '0' :
                     c1 >= 'A' && c1 <= 'F' ? c1 - 'A' + 10 :
                     c1 >= 'a' && c1 <= 'f' ? c1 - 'a' + 10 : 0;

            int v2 = c2 >= '0' && c2 <= '9' ? c2 - '0' :
                     c2 >= 'A' && c2 <= 'F' ? c2 - 'A' + 10 :
                     c2 >= 'a' && c2 <= 'f' ? c2 - 'a' + 10 : 0;

            return (byte)((v1 << 4) | v2);
        }
    }
}

