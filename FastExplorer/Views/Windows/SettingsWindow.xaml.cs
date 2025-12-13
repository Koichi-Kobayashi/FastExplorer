using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using FastExplorer.ViewModels.Pages;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

namespace FastExplorer.Views.Windows
{
    /// <summary>
    /// 設定ウィンドウを表すクラス
    /// </summary>
    public partial class SettingsWindow : FluentWindow
    {
        /// <summary>
        /// 設定ページのViewModelを取得します
        /// </summary>
        public SettingsViewModel ViewModel { get; }

        /// <summary>
        /// <see cref="SettingsWindow"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">設定ページのViewModel</param>
        public SettingsWindow(SettingsViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;

            InitializeComponent();
            
            // TitleBarの×ボタンイベントを処理
            TitleBar.CloseClicked += TitleBar_CloseClicked;
            
            // ウィンドウ読み込み時にViewModelを初期化
            Loaded += async (s, e) =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("SettingsWindow Loaded event started");
                    await ViewModel.OnNavigatedToAsync();
                    System.Diagnostics.Debug.WriteLine("SettingsWindow OnNavigatedToAsync completed");
                    
                    // システムテーマの監視を設定（テーマが"System"の場合のみ）
                    // SettingsWindowは設定画面なので、常にシステムテーマの変更を監視する必要はない
                    // ただし、設定画面自体のテーマはシステムテーマに追従させる
                    SystemThemeWatcher.Watch(this);
                    
                    // デフォルトで全般ボタンを選択状態にする
                    GeneralButton.Appearance = ControlAppearance.Primary;
                    
                    // ナビゲーションボタンのホバーイベントを設定
                    SetupNavigationButtonHover();
                    
                    System.Diagnostics.Debug.WriteLine("SettingsWindow Loaded event completed");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error in SettingsWindow Loaded event: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                    // エラーが発生してもウィンドウは表示する
                }
            };
            
            // ウィンドウを閉じる時にViewModelのクリーンアップ
            // Alt+F4など、TitleBar.CloseClicked以外で閉じられた場合にも設定を保存
            // _isNavigatingFromフラグで重複を防ぐ
            Closing += (s, e) =>
            {
                // 非同期処理はfire-and-forgetで実行（ウィンドウを閉じる処理をブロックしない）
                // _isNavigatingFromフラグで重複を防ぐ
                _ = ViewModel.OnNavigatedFromAsync();
            };

            // ウィンドウが閉じられた後にクリーンアップ
            Closed += (s, e) =>
            {
                // イベントハンドラーを解除してメモリリークを防ぐ
                TitleBar.CloseClicked -= TitleBar_CloseClicked;
            };
        }

        /// <summary>
        /// TitleBarの×ボタンがクリックされたときに呼び出されます
        /// </summary>
        private async void TitleBar_CloseClicked(object? sender, RoutedEventArgs e)
        {
            // 閉じる前に設定を保存（完了を待つ）
            await ViewModel.OnNavigatedFromAsync();
            Close();
        }

        /// <summary>
        /// ナビゲーションボタンのホバーイベントを設定します
        /// </summary>
        private void SetupNavigationButtonHover()
        {
            var buttons = new[] { GeneralButton, AppearanceButton, AboutButton };
            var darkThemeBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x2d, 0x2d, 0x2d)); // #2d2d2d
            
            foreach (var button in buttons)
            {
                // IsMouseOverプロパティの変更を監視
                var descriptor = System.ComponentModel.DependencyPropertyDescriptor.FromProperty(
                    System.Windows.UIElement.IsMouseOverProperty, 
                    typeof(System.Windows.UIElement));
                
                if (descriptor != null)
                {
                    descriptor.AddValueChanged(button, (s, e) =>
                    {
                        if (ViewModel.CurrentTheme == Wpf.Ui.Appearance.ApplicationTheme.Dark)
                        {
                            if (button.Appearance == ControlAppearance.Secondary)
                            {
                                // 複数回試行して確実に適用
                                Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() =>
                                {
                                    if (button.IsMouseOver && button.Appearance == ControlAppearance.Secondary)
                                    {
                                        button.SetCurrentValue(System.Windows.Controls.Control.BackgroundProperty, darkThemeBrush);
                                    }
                                }));
                                
                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
                                {
                                    if (button.IsMouseOver && button.Appearance == ControlAppearance.Secondary)
                                    {
                                        button.SetCurrentValue(System.Windows.Controls.Control.BackgroundProperty, darkThemeBrush);
                                    }
                                    else if (!button.IsMouseOver && button.Appearance == ControlAppearance.Secondary)
                                    {
                                        button.SetCurrentValue(System.Windows.Controls.Control.BackgroundProperty, System.Windows.Media.Brushes.Transparent);
                                    }
                                }));
                            }
                        }
                    });
                }
                
                // MouseEnter/MouseLeaveも設定（二重に設定）
                button.MouseEnter += (s, e) =>
                {
                    if (ViewModel.CurrentTheme == Wpf.Ui.Appearance.ApplicationTheme.Dark)
                    {
                        if (button.Appearance == ControlAppearance.Secondary)
                        {
                            button.SetCurrentValue(System.Windows.Controls.Control.BackgroundProperty, darkThemeBrush);
                        }
                    }
                };
                
                button.MouseLeave += (s, e) =>
                {
                    if (button.Appearance == ControlAppearance.Secondary)
                    {
                        button.SetCurrentValue(System.Windows.Controls.Control.BackgroundProperty, System.Windows.Media.Brushes.Transparent);
                    }
                };
            }
        }

        /// <summary>
        /// ナビゲーションボタンがクリックされたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="e">ルーティングイベント引数</param>
        private void NavigationButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Wpf.Ui.Controls.Button button || button.Tag is not string tag)
                return;

            // すべてのボタンの外観をSecondaryにリセット
            GeneralButton.Appearance = ControlAppearance.Secondary;
            AppearanceButton.Appearance = ControlAppearance.Secondary;
            AboutButton.Appearance = ControlAppearance.Secondary;

            // クリックされたボタンをPrimaryに設定
            button.Appearance = ControlAppearance.Primary;

            // すべてのページを非表示にする
            GeneralPage.Visibility = Visibility.Collapsed;
            AppearancePage.Visibility = Visibility.Collapsed;
            AboutPage.Visibility = Visibility.Collapsed;

            // タグに応じて適切なページを表示
            switch (tag)
            {
                case "General":
                    GeneralPage.Visibility = Visibility.Visible;
                    break;
                case "Appearance":
                    AppearancePage.Visibility = Visibility.Visible;
                    break;
                case "About":
                    AboutPage.Visibility = Visibility.Visible;
                    break;
                // その他のページは実装予定
                default:
                    GeneralPage.Visibility = Visibility.Visible;
                    break;
            }
        }
    }
}
