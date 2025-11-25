using System.Globalization;
using System.Windows.Data;
using Wpf.Ui.Controls;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// SymbolRegularへのnull値変換を処理するコンバーター
    /// </summary>
    public class SymbolRegularConverter : IValueConverter
    {
        /// <summary>
        /// 値を変換します
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // SymbolRegular型の場合はそのまま返す
            if (value is SymbolRegular symbol)
            {
                return symbol;
            }

            // null値またはその他の場合は、パラメータで指定されたデフォルト値を使用
            // パラメータが文字列の場合は、それをSymbolRegularに変換
            if (parameter != null)
            {
                if (parameter is SymbolRegular defaultSymbol)
                {
                    return defaultSymbol;
                }
                
                // 文字列パラメータの場合（"Left"または"Right"）
                if (parameter is string paramStr)
                {
                    return paramStr switch
                    {
                        "Left" => SymbolRegular.ChevronLeft24,
                        "Right" => SymbolRegular.ChevronRight24,
                        _ => SymbolRegular.ChevronLeft24
                    };
                }
            }

            // デフォルト値
            return SymbolRegular.ChevronLeft24;
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

