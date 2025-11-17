using FastExplorer.Controls;
using FastExplorer.Models;
using FastExplorer.Services;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;
using Wpf.Ui.Abstractions.Controls;
using Wpf.Ui.Appearance;

namespace FastExplorer.ViewModels.Pages
{
    /// <summary>
    /// 設定ページのViewModel
    /// </summary>
    public partial class SettingsViewModel : ObservableObject, INavigationAware
    {
        private bool _isInitialized = false;
        private readonly WindowSettingsService _windowSettingsService;

        /// <summary>
        /// <see cref="SettingsViewModel"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="windowSettingsService">ウィンドウ設定サービス</param>
        public SettingsViewModel(WindowSettingsService windowSettingsService)
        {
            _windowSettingsService = windowSettingsService;
        }

        /// <summary>
        /// アプリケーションのバージョンを取得または設定します
        /// </summary>
        [ObservableProperty]
        private string _appVersion = String.Empty;

        /// <summary>
        /// 現在のテーマを取得または設定します
        /// </summary>
        [ObservableProperty]
        private ApplicationTheme _currentTheme = ApplicationTheme.Unknown;

        /// <summary>
        /// テーマカラーのコレクションを取得または設定します
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<ThemeColor> _themeColors = new();

        /// <summary>
        /// ページにナビゲートされたときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedToAsync()
        {
            if (!_isInitialized)
            {
                InitializeViewModel();
            }
            else
            {
                // 既に初期化済みの場合でも、現在のテーマを更新
                CurrentTheme = ApplicationThemeManager.GetAppTheme();
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// ページから離れるときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedFromAsync() => Task.CompletedTask;

        /// <summary>
        /// ViewModelを初期化します
        /// </summary>
        private void InitializeViewModel()
        {
            CurrentTheme = ApplicationThemeManager.GetAppTheme();
            AppVersion = $"UiDesktopApp1 - {GetAssemblyVersion()}";

            // テーマカラーリストを初期化
            InitializeThemeColors();

            _isInitialized = true;
        }

        /// <summary>
        /// テーマカラーリストを初期化します
        /// </summary>
        private void InitializeThemeColors()
        {
            ThemeColors.Clear();

            var colors = new[]
            {
                new ThemeColor { Name = "既定", NameEn = "Default", ColorCode = "#F5F5F5", SecondaryColorCode = "#FCFCFC", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F5F5F5")) },
                new ThemeColor { Name = "イエロー ゴールド", NameEn = "Yellow Gold", ColorCode = "#FCEECA", SecondaryColorCode = "#FEFAEF", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FCEECA")) },
                new ThemeColor { Name = "オレンジ ブライト", NameEn = "Orange Bright", ColorCode = "#FADDCC", SecondaryColorCode = "#FEF5F0", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FADDCC")) },
                new ThemeColor { Name = "ブリック レッド", NameEn = "Brick Red", ColorCode = "#F3D4D5", SecondaryColorCode = "#FBF2F2", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F3D4D5")) },
                new ThemeColor { Name = "モダン レッド", NameEn = "Modern Red", ColorCode = "#FCD7D7", SecondaryColorCode = "#FEF3F3", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FCD7D7")) },
                new ThemeColor { Name = "レッド", NameEn = "Red", ColorCode = "#F8CADC", SecondaryColorCode = "#FDEFF5", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8CADC")) },
                new ThemeColor { Name = "ローズ ブライト", NameEn = "Rose Bright", ColorCode = "#F8CADC", SecondaryColorCode = "#FDEFF5", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8CADC")) },
                new ThemeColor { Name = "ブルー", NameEn = "Blue", ColorCode = "#CAE2F4", SecondaryColorCode = "#EFF6FC", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAE2F4")) },
                new ThemeColor { Name = "アイリス パステル", NameEn = "Iris Pastel", ColorCode = "#E4DEEE", SecondaryColorCode = "#F7F5FA", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E4DEEE")) },
                new ThemeColor { Name = "バイオレット レッド ライト", NameEn = "Violet Red Light", ColorCode = "#EDD8F0", SecondaryColorCode = "#FAF3FB", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EDD8F0")) },
                new ThemeColor { Name = "クール ブルー ブライト", NameEn = "Cool Blue Bright", ColorCode = "#CAE8EF", SecondaryColorCode = "#EFF8FA", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAE8EF")) },
                new ThemeColor { Name = "シーフォーム", NameEn = "Seafoam", ColorCode = "#CAEEF0", SecondaryColorCode = "#EFFAFB", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAEEF0")) },
                new ThemeColor { Name = "ミント ライト", NameEn = "Mint Light", ColorCode = "#CAEDE7", SecondaryColorCode = "#EFFAF8", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CAEDE7")) },
                new ThemeColor { Name = "グレー", NameEn = "Gray", ColorCode = "#E2E1E1", SecondaryColorCode = "#F6F6F6", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E1E1")) },
                new ThemeColor { Name = "グリーン", NameEn = "Green", ColorCode = "#CDE2CD", SecondaryColorCode = "#F0F6F0", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CDE2CD")) },
                new ThemeColor { Name = "オーバーキャスト", NameEn = "Overcast", ColorCode = "#E1E1E1", SecondaryColorCode = "#F6F6F6", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E1E1E1")) },
                new ThemeColor { Name = "ストーム", NameEn = "Storm", ColorCode = "#D9D9D8", SecondaryColorCode = "#F4F4F3", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D9D9D8")) },
                new ThemeColor { Name = "ブルー グレー", NameEn = "Blue Gray", ColorCode = "#DFE2E3", SecondaryColorCode = "#F5F6F7", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DFE2E3")) },
                new ThemeColor { Name = "グレー ダーク", NameEn = "Gray Dark", ColorCode = "#D9DADB", SecondaryColorCode = "#F4F4F4", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D9DADB")) },
                new ThemeColor { Name = "カモフラージュ", NameEn = "Camouflage", ColorCode = "#E3E1DD", SecondaryColorCode = "#F7F6F5", Color = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E3E1DD")) }
            };

            foreach (var color in colors)
            {
                ThemeColors.Add(color);
            }
        }

        /// <summary>
        /// アセンブリのバージョンを取得します
        /// </summary>
        /// <returns>アセンブリのバージョン文字列</returns>
        private string GetAssemblyVersion()
        {
            return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString()
                ?? String.Empty;
        }

        /// <summary>
        /// テーマを変更します
        /// </summary>
        /// <param name="parameter">テーマパラメータ（"theme_light"またはその他）</param>
        [RelayCommand]
        private void OnChangeTheme(string parameter)
        {
            ApplicationTheme newTheme;
            switch (parameter)
            {
                case "theme_light":
                    if (CurrentTheme == ApplicationTheme.Light)
                        break;

                    newTheme = ApplicationTheme.Light;
                    ApplyTheme(newTheme);
                    break;

                default:
                    if (CurrentTheme == ApplicationTheme.Dark)
                        break;

                    newTheme = ApplicationTheme.Dark;
                    ApplyTheme(newTheme);
                    break;
            }
        }
        
        /// <summary>
        /// テーマを適用します（起動時と同じロジック）
        /// </summary>
        private void ApplyTheme(ApplicationTheme theme)
        {
            try
            {
                // ThemesDictionaryのThemeプロパティを先に更新（ApplicationThemeManager.Apply()の前に）
                if (System.Windows.Application.Current.Resources is System.Windows.ResourceDictionary mainDictionary)
                {
                    var mergedDictionaries = mainDictionary.MergedDictionaries;
                    var themesDict = mergedDictionaries
                        .OfType<System.Windows.ResourceDictionary>()
                        .FirstOrDefault(rd => rd.GetType().Name == "ThemesDictionary");
                    
                    if (themesDict != null)
                    {
                        var themeProperty = themesDict.GetType().GetProperty("Theme");
                        if (themeProperty != null)
                        {
                            var themeType = themeProperty.PropertyType;
                            var themeValue = Enum.Parse(themeType, theme == ApplicationTheme.Dark ? "Dark" : "Light");
                            themeProperty.SetValue(themesDict, themeValue);
                        }
                    }
                }

                // テーマを適用
                ApplicationThemeManager.Apply(theme);
                CurrentTheme = theme;

                // テーマ設定を保存
                SaveTheme(theme);
                
                // リソースディクショナリーも即座に更新（起動時と同じ）
                App.UpdateThemeResourcesInternal();
                
                // すべてのThemedSvgIconインスタンスを即座に更新（リアルタイム反映のため）
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    ThemedSvgIcon.RefreshAllInstances();
                }), System.Windows.Threading.DispatcherPriority.Render);
            }
            catch
            {
                // エラーハンドリング：デフォルトのテーマを使用
                ApplicationThemeManager.Apply(ApplicationTheme.Light);
            }
        }

        /// <summary>
        /// テーマ設定を保存します
        /// </summary>
        /// <param name="theme">保存するテーマ</param>
        private void SaveTheme(ApplicationTheme theme)
        {
            var settings = _windowSettingsService.GetSettings();
            settings.Theme = theme == ApplicationTheme.Light ? "Light" : "Dark";
            _windowSettingsService.SaveSettings(settings);
        }
    }
}
