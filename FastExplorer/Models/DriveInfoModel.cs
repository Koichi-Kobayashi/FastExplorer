using Cysharp.Text;

namespace FastExplorer.Models
{
    /// <summary>
    /// ドライブ情報を表すクラス
    /// </summary>
    public class DriveInfoModel
    {
        /// <summary>
        /// ドライブ名を取得または設定します
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// ドライブのパスを取得または設定します
        /// </summary>
        public string Path { get; set; } = string.Empty;

        /// <summary>
        /// ボリュームラベルを取得または設定します
        /// </summary>
        public string VolumeLabel { get; set; } = string.Empty;

        /// <summary>
        /// 総容量（バイト）を取得または設定します
        /// </summary>
        public long TotalSize { get; set; }

        /// <summary>
        /// 空き容量（バイト）を取得または設定します
        /// </summary>
        public long FreeSpace { get; set; }

        /// <summary>
        /// 使用容量（バイト）を取得または設定します
        /// </summary>
        public long UsedSpace => TotalSize - FreeSpace;

        /// <summary>
        /// 使用率（0.0～1.0）を取得します
        /// </summary>
        public double UsagePercentage => TotalSize > 0 ? (double)UsedSpace / TotalSize : 0.0;

        /// <summary>
        /// フォーマット済みの総容量文字列を取得します
        /// </summary>
        public string FormattedTotalSize => FormatBytes(TotalSize);

        /// <summary>
        /// フォーマット済みの空き容量文字列を取得します
        /// </summary>
        public string FormattedFreeSpace => FormatBytes(FreeSpace);

        /// <summary>
        /// フォーマット済みの使用容量文字列を取得します
        /// </summary>
        public string FormattedUsedSpace => FormatBytes(UsedSpace);

        /// <summary>
        /// バイト数を人間が読みやすい形式に変換します
        /// </summary>
        private static string FormatBytes(long bytes)
        {
            // 定数配列を静的フィールドに移動してメモリ割り当てを削減
            string[] sizes = { "B", "KB", "MB", "GB", "TB", "PiB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            
            // 文字列補間を最適化（ZString.Formatを使用してボクシングを回避）
            return ZString.Format("{0:0.##} {1}", len, sizes[order]);
        }
    }
}


