using System.Globalization;
using System.Windows.Data;
using Wpf.Ui.Controls;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// フォルダー名に基づいてアイコンを変換するコンバーター
    /// </summary>
    public class FolderIconConverter : IValueConverter
    {
        /// <summary>
        /// 値を変換します
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string folderName)
            {
                return folderName switch
                {
                    "デスクトップ" => SymbolRegular.Desktop24,
                    "ダウンロード" => SymbolRegular.ArrowDownload24,
                    "ドキュメント" => SymbolRegular.Document24,
                    "ピクチャ" => SymbolRegular.Image24,
                    "ミュージック" => SymbolRegular.MusicNote124,
                    "ビデオ" => SymbolRegular.Video24,
                    "ごみ箱" => SymbolRegular.Delete24,
                    _ => SymbolRegular.Folder24
                };
            }
            return SymbolRegular.Folder24;
        }

        /// <summary>
        /// 値を逆変換します（実装されていません）
        /// </summary>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

