using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using FastExplorer.Models;
using FastExplorer.Services;
using Wpf.Ui;
using Wpf.Ui.Controls;

namespace FastExplorer.ViewModels.Windows
{
    /// <summary>
    /// メインウィンドウのViewModel
    /// </summary>
    public partial class MainWindowViewModel : ObservableObject
    {
        private readonly FavoriteService _favoriteService;
        private INavigationService? _navigationService;

        /// <summary>
        /// アプリケーションのタイトル
        /// </summary>
        [ObservableProperty]
        private string _applicationTitle = "FastExplorer";

        /// <summary>
        /// ナビゲーションメニューアイテム（お気に入り）
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<object> _menuItems = new();

        /// <summary>
        /// ステータスバーに表示するテキスト
        /// </summary>
        [ObservableProperty]
        private string _statusBarText = "準備完了";

        /// <summary>
        /// <see cref="MainWindowViewModel"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="favoriteService">お気に入りサービス</param>
        public MainWindowViewModel(FavoriteService favoriteService)
        {
            _favoriteService = favoriteService;
            LoadFavorites();
        }

        /// <summary>
        /// ナビゲーションサービスを設定します
        /// </summary>
        /// <param name="navigationService">ナビゲーションサービス</param>
        public void SetNavigationService(INavigationService navigationService)
        {
            _navigationService = navigationService;
        }

        /// <summary>
        /// お気に入りを読み込みます
        /// </summary>
        public void LoadFavorites()
        {
            MenuItems.Clear();
            
            // エクスプローラーページへのリンクを追加
            // TargetPageTypeを設定しないことで、Contentがクリアされないようにする
            var explorerItem = new NavigationViewItem()
            {
                Content = "ホーム",
                Icon = new SymbolIcon { Symbol = SymbolRegular.Folder24 },
                Tag = "HOME" // ホームアイテムを識別するためのTag
            };
            
            // Clickイベントを設定（ItemInvokedイベントが発火しない場合のフォールバック）
            explorerItem.Click += (s, e) =>
            {
                NavigateToHome();
            };
            
            MenuItems.Add(explorerItem);

            // お気に入りを追加
            var favorites = _favoriteService.GetFavorites();
            foreach (var favorite in favorites)
            {
                MenuItems.Add(CreateFavoriteNavigationItem(favorite));
            }
        }

        /// <summary>
        /// お気に入りからNavigationViewItemを作成します
        /// </summary>
        /// <param name="favorite">お気に入りアイテム</param>
        /// <returns>NavigationViewItem</returns>
        private NavigationViewItem CreateFavoriteNavigationItem(FavoriteItem favorite)
        {
            var item = new NavigationViewItem()
            {
                Content = favorite.Name,
                Icon = new SymbolIcon { Symbol = SymbolRegular.Folder24 },
                Tag = favorite.Path // パスをTagに保存
            };

            // クリックイベントを設定（ItemInvokedイベントが発火しない場合のフォールバック）
            item.Click += (s, e) =>
            {
                NavigateToFavorite(favorite.Path);
            };

            return item;
        }

        /// <summary>
        /// ホームページにナビゲートします
        /// </summary>
        private void NavigateToHome()
        {
            // エクスプローラーページにナビゲート
            if (_navigationService != null)
            {
                _navigationService.Navigate(typeof(Views.Pages.ExplorerPage));
                
                // 少し遅延してからホームページを表示（ページが読み込まれるのを待つ）
                Task.Delay(100).ContinueWith(_ =>
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        var explorerPageViewModel = App.Services.GetService(typeof(ViewModels.Pages.ExplorerPageViewModel)) as ViewModels.Pages.ExplorerPageViewModel;
                        if (explorerPageViewModel != null && explorerPageViewModel.SelectedTab != null)
                        {
                            // ホームページにナビゲート
                            explorerPageViewModel.SelectedTab.ViewModel.NavigateToHome();
                        }
                    });
                });
            }
        }

        /// <summary>
        /// お気に入りのパスにナビゲートします
        /// </summary>
        /// <param name="path">ナビゲートするパス</param>
        private void NavigateToFavorite(string path)
        {
            // エクスプローラーページにナビゲート
            if (_navigationService != null)
            {
                _navigationService.Navigate(typeof(Views.Pages.ExplorerPage));
                
                // 少し遅延してからパスを設定（ページが読み込まれるのを待つ）
                Task.Delay(100).ContinueWith(_ =>
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        var explorerPageViewModel = App.Services.GetService(typeof(ViewModels.Pages.ExplorerPageViewModel)) as ViewModels.Pages.ExplorerPageViewModel;
                        if (explorerPageViewModel != null && explorerPageViewModel.SelectedTab != null)
                        {
                            explorerPageViewModel.SelectedTab.ViewModel.NavigateToPathCommand.Execute(path);
                        }
                    });
                });
            }
        }

        /// <summary>
        /// お気に入りを追加します
        /// </summary>
        /// <param name="name">表示名</param>
        /// <param name="path">パス</param>
        public void AddFavorite(string name, string path)
        {
            _favoriteService.AddFavorite(name, path);
            LoadFavorites();
        }

        /// <summary>
        /// お気に入りを削除します
        /// </summary>
        /// <param name="id">削除するお気に入りのID</param>
        public void RemoveFavorite(string id)
        {
            _favoriteService.RemoveFavorite(id);
            LoadFavorites();
        }

        [ObservableProperty]
        private ObservableCollection<object> _footerMenuItems = new()
        {
            new NavigationViewItem()
            {
                Content = "Settings",
                Icon = new SymbolIcon { Symbol = SymbolRegular.Settings24 },
                TargetPageType = typeof(Views.Pages.SettingsPage)
            }
        };

        [ObservableProperty]
        private ObservableCollection<MenuItem> _trayMenuItems = new()
        {
            new MenuItem { Header = "Home", Tag = "tray_home" }
        };
    }
}
