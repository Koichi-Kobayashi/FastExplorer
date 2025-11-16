using FastExplorer.Views.Pages;
using FastExplorer.Views.Windows;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Wpf.Ui;

namespace FastExplorer.Services
{
    /// <summary>
    /// アプリケーションの管理ホストを表すクラス
    /// </summary>
    public class ApplicationHostService : IHostedService
    {
        private readonly IServiceProvider _serviceProvider;

        private INavigationWindow _navigationWindow;

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
            if (!Application.Current.Windows.OfType<MainWindow>().Any())
            {
                _navigationWindow = (
                    _serviceProvider.GetService(typeof(INavigationWindow)) as INavigationWindow
                )!;
                _navigationWindow!.ShowWindow();

                _navigationWindow.Navigate(typeof(Views.Pages.ExplorerPage));
            }

            await Task.CompletedTask;
        }
    }
}
