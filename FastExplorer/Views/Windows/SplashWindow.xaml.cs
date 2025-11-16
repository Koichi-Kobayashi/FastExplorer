using System.Windows;
using System.Windows.Media;
using Wpf.Ui.Appearance;

namespace FastExplorer.Views.Windows
{

    /// <summary>
    /// スプラッシュウィンドウを表すクラス
    /// </summary>
    public partial class SplashWindow : Window
    {
        /// <summary>
        /// <see cref="SplashWindow"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        public SplashWindow()
        {
            // 最初は非表示で作成（テーマ適用後に表示するため）
            Visibility = Visibility.Hidden;
            
            InitializeComponent();
            
            // テーマに応じて背景色を直接設定（DynamicResourceが解決される前に表示されるのを防ぐため）
            Loaded += SplashWindow_Loaded;
        }

        private void SplashWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // テーマに応じて背景色を設定
            ApplyThemeColors();
        }

        /// <summary>
        /// テーマに応じて色を適用します
        /// </summary>
        public void ApplyThemeColors()
        {
            var isDark = ApplicationThemeManager.GetAppTheme() == ApplicationTheme.Dark;
            
            // Borderの背景色を設定
            if (Content is System.Windows.Controls.Border border)
            {
                if (isDark)
                {
                    border.Background = new SolidColorBrush(Color.FromRgb(32, 32, 32)); // ダークモードの背景色
                }
                else
                {
                    border.Background = new SolidColorBrush(Color.FromRgb(243, 243, 243)); // ライトモードの背景色
                }
            }
        }

        /// <summary>
        /// ウィンドウが閉じられているかどうかを確認します
        /// </summary>
        public bool IsClosed()
        {
            // Windowが閉じられているかどうかを確認
            // IsLoadedがfalseの場合、ウィンドウは閉じられているか、まだ読み込まれていない
            return !IsLoaded;
        }
    }
}

