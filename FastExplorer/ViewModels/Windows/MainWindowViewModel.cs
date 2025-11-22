using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
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
        
        // 型をキャッシュ（パフォーマンス向上）
        private static readonly Type ExplorerPageType = typeof(Views.Pages.ExplorerPage);
        private static readonly Type ExplorerPageViewModelType = typeof(ViewModels.Pages.ExplorerPageViewModel);
        
        // ViewModelをキャッシュ（パフォーマンス向上）
        private ViewModels.Pages.ExplorerPageViewModel? _cachedExplorerPageViewModel;
        
        // ホームアイテムをキャッシュ（パフォーマンス向上）
        private NavigationViewItem? _homeMenuItem;

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
            // ホームアイテムが存在しない場合は作成
            if (_homeMenuItem == null)
            {
                _homeMenuItem = new NavigationViewItem()
                {
                    Content = "ホーム",
                    Icon = new SymbolIcon { Symbol = SymbolRegular.Folder24 },
                    Tag = "HOME" // ホームアイテムを識別するためのTag
                };
                
                // Clickイベントを設定（ItemInvokedイベントが発火しない場合のフォールバック）
                _homeMenuItem.Click += (s, e) =>
                {
                    NavigateToHome();
                };
            }
            
            // 現在のお気に入りを取得（一度だけToList()を呼び出す）
            var currentFavorites = _favoriteService.GetFavorites();
            var currentFavoritesList = currentFavorites as IList<FavoriteItem> ?? currentFavorites.ToList();
            
            // 既存のメニューアイテムからお気に入りアイテムを抽出（ホームアイテムを除く）
            // ToList()を削減：直接反復処理を使用、容量を事前に確保
            var existingFavoriteItems = new List<NavigationViewItem>(MenuItems.Count);
            foreach (var item in MenuItems)
            {
                if (item is NavigationViewItem navItem && navItem.Tag is string tag && tag != "HOME")
                {
                    existingFavoriteItems.Add(navItem);
                }
            }
            
            // 既存のお気に入りパスを取得（HashSetを直接構築、容量を事前に確保）
            var existingPaths = new HashSet<string>(existingFavoriteItems.Count, StringComparer.OrdinalIgnoreCase);
            foreach (var item in existingFavoriteItems)
            {
                if (item.Tag is string path && !string.IsNullOrEmpty(path))
                {
                    existingPaths.Add(path);
                }
            }
            
            // 新しいお気に入りのパスを取得（HashSetを直接構築、容量を事前に確保）
            var newPaths = new HashSet<string>(currentFavoritesList.Count, StringComparer.OrdinalIgnoreCase);
            foreach (var favorite in currentFavoritesList)
            {
                newPaths.Add(favorite.Path);
            }
            
            // 削除されたお気に入りをメニューから削除（ToList()を削減）
            foreach (var item in existingFavoriteItems)
            {
                if (item.Tag is string tag && !newPaths.Contains(tag))
                {
                    MenuItems.Remove(item);
                }
            }
            
            // 新しく追加されたお気に入りをメニューに追加（ToList()を削減、容量を事前に確保）
            var itemsToAdd = new List<NavigationViewItem>(currentFavoritesList.Count);
            foreach (var favorite in currentFavoritesList)
            {
                if (!existingPaths.Contains(favorite.Path))
                {
                    itemsToAdd.Add(CreateFavoriteNavigationItem(favorite));
                }
            }
            
            // ホームアイテムが存在しない場合は先頭に追加
            if (!MenuItems.Contains(_homeMenuItem))
            {
                MenuItems.Insert(0, _homeMenuItem);
            }
            
            // 新しいお気に入りアイテムを追加（ホームアイテムの後）
            var homeIndex = MenuItems.IndexOf(_homeMenuItem);
            for (int i = 0; i < itemsToAdd.Count; i++)
            {
                MenuItems.Insert(homeIndex + 1 + i, itemsToAdd[i]);
            }
            
            // 既存アイテムの名前が変更された場合は更新（FirstOrDefault()を最適化）
            foreach (var favorite in currentFavoritesList)
            {
                NavigationViewItem? existingItem = null;
                foreach (var item in existingFavoriteItems)
                {
                    if (item.Tag is string tag && string.Equals(tag, favorite.Path, StringComparison.OrdinalIgnoreCase))
                    {
                        existingItem = item;
                        break; // 見つかったら早期終了
                    }
                }
                if (existingItem != null && existingItem.Content?.ToString() != favorite.Name)
                {
                    existingItem.Content = favorite.Name;
                }
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

            // コンテキストメニューを追加（右クリックで削除）
            var contextMenu = new System.Windows.Controls.ContextMenu();
            var deleteMenuItem = new System.Windows.Controls.MenuItem
            {
                Header = "削除",
                Icon = new SymbolIcon { Symbol = SymbolRegular.Delete24 }
            };
            deleteMenuItem.Click += (s, e) =>
            {
                RemoveFavoriteByPath(favorite.Path);
            };
            contextMenu.Items.Add(deleteMenuItem);
            item.ContextMenu = contextMenu;

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
                _navigationService.Navigate(ExplorerPageType);
                
                // ページが読み込まれるのを待ってからホームページを表示
                // DispatcherPriority.Loadedを使用することで、レイアウトが完了してから実行される
                Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    // ViewModelをキャッシュから取得（なければ取得してキャッシュ）
                    if (_cachedExplorerPageViewModel == null)
                    {
                        _cachedExplorerPageViewModel = App.Services.GetService(ExplorerPageViewModelType) as ViewModels.Pages.ExplorerPageViewModel;
                    }
                    
                    if (_cachedExplorerPageViewModel != null && _cachedExplorerPageViewModel.SelectedTab != null)
                    {
                        // ホームページにナビゲート
                        _cachedExplorerPageViewModel.SelectedTab.ViewModel.NavigateToHome();
                    }
                }), System.Windows.Threading.DispatcherPriority.Loaded);
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
                _navigationService.Navigate(ExplorerPageType);
                
                // ページが読み込まれるのを待ってからパスを設定
                // DispatcherPriority.Loadedを使用することで、レイアウトが完了してから実行される
                var pathCopy = path; // クロージャで使用するため変数に保存
                Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    // ViewModelをキャッシュから取得（なければ取得してキャッシュ）
                    if (_cachedExplorerPageViewModel == null)
                    {
                        _cachedExplorerPageViewModel = App.Services.GetService(ExplorerPageViewModelType) as ViewModels.Pages.ExplorerPageViewModel;
                    }
                    
                    if (_cachedExplorerPageViewModel != null && _cachedExplorerPageViewModel.SelectedTab != null)
                    {
                        _cachedExplorerPageViewModel.SelectedTab.ViewModel.NavigateToPathCommand.Execute(pathCopy);
                    }
                }), System.Windows.Threading.DispatcherPriority.Loaded);
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

        /// <summary>
        /// パスでお気に入りを削除します
        /// </summary>
        /// <param name="path">削除するお気に入りのパス</param>
        private void RemoveFavoriteByPath(string path)
        {
            _favoriteService.RemoveFavoriteByPath(path);
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
        private ObservableCollection<System.Windows.Controls.MenuItem> _trayMenuItems = new()
        {
            new System.Windows.Controls.MenuItem { Header = "Home", Tag = "tray_home" }
        };
    }
}
