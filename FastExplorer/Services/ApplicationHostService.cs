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
        
        // 型をキャッシュ（パフォーマンス向上）
        private static readonly Type ExplorerPageType = typeof(Views.Pages.ExplorerPage);

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
            // テーマは既にApp.xaml.csで適用されているため、ここでは適用しない（重複を避けて起動を高速化）
            await HandleActivationAsync();
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
            // LINQを避けて高速化（直接ループでチェック）
            bool hasMainWindow = false;
            var windows = Application.Current.Windows;
            for (int i = 0; i < windows.Count; i++)
            {
                if (windows[i] is MainWindow)
                {
                    hasMainWindow = true;
                    break;
                }
            }
            
            if (!hasMainWindow)
            {
                // テーマは既にApp.xaml.csで適用されているため、ここでは適用しない（重複を避ける）
                // メインウィンドウを取得
                _navigationWindow = (
                    _serviceProvider.GetService(typeof(INavigationWindow)) as INavigationWindow
                )!;
                
                // メインウィンドウを表示（MainWindow_LoadedでVisibilityがVisibleに設定される）
                _navigationWindow!.ShowWindow();

                _navigationWindow.Navigate(ExplorerPageType);
            }

            await Task.CompletedTask;
        }
    }
}
