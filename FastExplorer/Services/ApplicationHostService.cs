using FastExplorer.Views.Pages;
using FastExplorer.Views.Windows;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Wpf.Ui;
using Wpf.Ui.Appearance;

namespace FastExplorer.Services
{
    /// <summary>
    /// アプリケーションの管理ホストを表すクラス
    /// </summary>
    public class ApplicationHostService : IHostedService
    {
        private readonly IServiceProvider _serviceProvider;

        private INavigationWindow? _navigationWindow;

        /// <summary>
        /// <see cref="ApplicationHostService"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="serviceProvider">サービスプロバイダー</param>
        public ApplicationHostService(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        /// <summary>
        /// アプリケーションホストがサービスを開始する準備ができたときに呼び出されます
        /// </summary>
        /// <param name="cancellationToken">開始プロセスが中止されたことを示すトークン</param>
        public async Task StartAsync(CancellationToken cancellationToken)
        {
            // テーマを確実に適用するため、ここでも適用を試みる
            ApplyThemeFromSettings();
            
            await HandleActivationAsync();
        }

        /// <summary>
        /// 設定からテーマを読み込んで適用します
        /// </summary>
        private void ApplyThemeFromSettings()
        {
            try
            {
                var windowSettingsService = _serviceProvider.GetService(typeof(WindowSettingsService)) as WindowSettingsService;
                if (windowSettingsService != null)
                {
                    var settings = windowSettingsService.GetSettings();
                    if (!string.IsNullOrEmpty(settings.Theme))
                    {
                        var theme = settings.Theme == "Dark" ? Wpf.Ui.Appearance.ApplicationTheme.Dark : Wpf.Ui.Appearance.ApplicationTheme.Light;
                        Wpf.Ui.Appearance.ApplicationThemeManager.Apply(theme);
                        
                        // リソースディクショナリーも更新
                        App.UpdateThemeResourcesInternal();

                        // 保存されたテーマカラーを適用
                        if (!string.IsNullOrEmpty(settings.ThemeColorCode))
                        {
                            App.ApplyThemeColorFromSettings(settings);
                        }
                    }
                }
            }
            catch
            {
                // エラーハンドリング：デフォルトのテーマを使用
            }
        }

        /// <summary>
        /// アプリケーションホストが正常なシャットダウンを実行しているときに呼び出されます
        /// </summary>
        /// <param name="cancellationToken">シャットダウンプロセスがもはや正常でないことを示すトークン</param>
        public async Task StopAsync(CancellationToken cancellationToken)
        {
            await Task.CompletedTask;
        }

        /// <summary>
        /// アクティベーション中にメインウィンドウを作成します
        /// </summary>
        private async Task HandleActivationAsync()
        {
            if (!Application.Current.Windows.OfType<MainWindow>().Any())
            {
                // テーマを確実に適用してからメインウィンドウを表示
                ApplyThemeFromSettings();
                
                // 少し待ってからテーマが完全に適用されるのを待つ
                await Task.Delay(50);
                
                _navigationWindow = (
                    _serviceProvider.GetService(typeof(INavigationWindow)) as INavigationWindow
                )!;
                
                // メインウィンドウを表示（MainWindow_LoadedでVisibilityがVisibleに設定される）
                _navigationWindow!.ShowWindow();

                _navigationWindow.Navigate(typeof(Views.Pages.ExplorerPage));
                
                // メインウィンドウが表示されたらスプラッシュウィンドウを閉じる
                _ = Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    var splashWindow = Application.Current.Windows.OfType<Views.Windows.SplashWindow>().FirstOrDefault();
                    if (splashWindow != null)
                    {
                        splashWindow.Close();
                    }
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }

            await Task.CompletedTask;
        }
    }
}
