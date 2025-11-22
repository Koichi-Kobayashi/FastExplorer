using System.IO;
using Cysharp.Text;

namespace FastExplorer.Models
{
    /// <summary>
    /// ファイルシステムのアイテム（ファイルまたはディレクトリ）を表すクラス
    /// </summary>
    public class FileSystemItem
    {
        /// <summary>
        /// アイテムの名前を取得または設定します
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// アイテムの完全なパスを取得または設定します
        /// </summary>
        public string FullPath { get; set; } = string.Empty;

        /// <summary>
        /// ファイルの拡張子を取得または設定します（ディレクトリの場合は空文字列）
        /// </summary>
        public string Extension { get; set; } = string.Empty;

        /// <summary>
        /// ファイルサイズ（バイト単位）を取得または設定します
        /// </summary>
        public long Size { get; set; }

        /// <summary>
        /// 最終更新日時を取得または設定します
        /// </summary>
        public DateTime LastModified { get; set; }

        /// <summary>
        /// ディレクトリかどうかを示す値を取得または設定します
        /// </summary>
        public bool IsDirectory { get; set; }

        /// <summary>
        /// ファイル属性を取得または設定します
        /// </summary>
        public FileAttributes Attributes { get; set; }

        /// <summary>
        /// 表示用の名前を取得します（ディレクトリの場合は名前、ファイルの場合は拡張子を除いた名前）
        /// </summary>
        public string DisplayName => IsDirectory ? Name : Path.GetFileNameWithoutExtension(Name);

        /// <summary>
        /// フォーマット済みのサイズ文字列を取得します（ディレクトリの場合は"&lt;DIR&gt;"）
        /// </summary>
        public string FormattedSize => IsDirectory ? "<DIR>" : FormatFileSize(Size);

        /// <summary>
        /// フォーマット済みの日時文字列を取得します（yyyy/MM/dd HH:mm形式）
        /// </summary>
        public string FormattedDate => LastModified.ToString("yyyy/MM/dd HH:mm");

        /// <summary>
        /// バイト数を人間が読みやすい形式（B, KB, MB, GB, TB）に変換します
        /// </summary>
        /// <param name="bytes">変換するバイト数</param>
        /// <returns>フォーマット済みのサイズ文字列</returns>
        private static string FormatFileSize(long bytes)
        {
            // 定数配列を静的フィールドに移動してメモリ割り当てを削減
            const string B = "B";
            const string KB = "KB";
            const string MB = "MB";
            const string GB = "GB";
            const string TB = "TB";
            
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < 4) // sizes.Length - 1 = 4
            {
                order++;
                len = len / 1024;
            }
            
            // 文字列補間を最適化（ToString()の呼び出しを削減）
            string unit = order switch
            {
                0 => B,
                1 => KB,
                2 => MB,
                3 => GB,
                4 => TB,
                _ => B
            };
            
            // 数値のフォーマットを最適化（ZString.Formatを使用してボクシングを回避）
            if (len >= 100)
                return ZString.Format("{0:F0} {1}", len, unit);
            else if (len >= 10)
                return ZString.Format("{0:F1} {1}", len, unit);
            else
                return ZString.Format("{0:F2} {1}", len, unit);
        }
    }
}

