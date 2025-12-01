using System;
using System.Globalization;
using System.Windows.Data;
using FastExplorer.Models;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// テーマカラーの選択状態を判定するコンバーター
    /// </summary>
    public class ThemeColorSelectionConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length != 2)
                return false;

            var currentColorCode = values[0] as string;
            var selectedThemeColor = values[1] as ThemeColor;

            if (string.IsNullOrEmpty(currentColorCode) || selectedThemeColor == null)
                return false;

            return currentColorCode == selectedThemeColor.ColorCode;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
