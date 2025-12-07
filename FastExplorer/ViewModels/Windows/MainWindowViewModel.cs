using System;
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
        #region フィールド

        private readonly FavoriteService _favoriteService;
        private INavigationService? _navigationService;
        
        // 型をキャッシュ（パフォーマンス向上）
        private static readonly Type ExplorerPageType = typeof(Views.Pages.ExplorerPage);
        private static readonly Type ExplorerPageViewModelType = typeof(ViewModels.Pages.ExplorerPageViewModel);
        
        // ViewModelをキャッシュ（パフォーマンス向上）
        private ViewModels.Pages.ExplorerPageViewModel? _cachedExplorerPageViewModel;
        
        // ホームアイテムをキャッシュ（パフォーマンス向上）
        private NavigationViewItem? _homeMenuItem;
        
        // Application.Currentをキャッシュ（パフォーマンス向上）
        private static System.Windows.Application? _cachedApplication;

        #endregion

        #region プロパティ

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
            
            // 起動時の高速化のため、ホームアイテムのみ先に追加
            // お気に入りの読み込みはBackground優先度で遅延実行
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
                
                MenuItems.Insert(0, _homeMenuItem);
            }
            
            // お気に入りの読み込みを遅延実行（起動を高速化）
            var app = _cachedApplication ?? (_cachedApplication = System.Windows.Application.Current);
            app?.Dispatcher.BeginInvoke(new System.Action(() =>
            {
                LoadFavorites();
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>
        /// ナビゲーションサービスを設定します
        /// </summary>
        /// <param name="navigationService">ナビゲーションサービス</param>
        public void SetNavigationService(INavigationService navigationService)
        {
            _navigationService = navigationService;
        }

        #endregion

        #region お気に入り管理

        /// <summary>
        /// お気に入りを読み込みます
        /// </summary>
        public void LoadFavorites()
        {
            // ホームアイテムが存在しない場合は作成（コンストラクタで既に作成されている可能性がある）
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
            
            // 現在のお気に入りを取得（IListとして直接使用、ToList()を完全に回避）
            var currentFavoritesList = _favoriteService.GetFavorites();
            
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
            // Contains()とIndexOf()を1回の呼び出しに統合（パフォーマンス向上）
            var homeIndex = MenuItems.IndexOf(_homeMenuItem);
            if (homeIndex < 0)
            {
                MenuItems.Insert(0, _homeMenuItem);
                homeIndex = 0;
            }
            
            // 新しいお気に入りアイテムを追加（ホームアイテムの後）
            for (int i = 0; i < itemsToAdd.Count; i++)
            {
                MenuItems.Insert(homeIndex + 1 + i, itemsToAdd[i]);
            }
            
            // 既存アイテムの名前が変更された場合は更新（FirstOrDefault()を最適化）
            foreach (var favorite in currentFavoritesList)
            {
                NavigationViewItem? existingItem = null;
                // 文字列比較を最適化（ReadOnlySpanを使用してメモリ割り当てを削減）
                var favoritePathSpan = favorite.Path.AsSpan();
                foreach (var item in existingFavoriteItems)
                {
                    if (item.Tag is string tag)
                    {
                        var tagSpan = tag.AsSpan();
                        if (tagSpan.Length == favoritePathSpan.Length && 
                            tagSpan.CompareTo(favoritePathSpan, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            existingItem = item;
                            break; // 見つかったら早期終了
                        }
                    }
                }
                // ToString()を呼び出す前に型チェック（パフォーマンス向上）
                if (existingItem != null)
                {
                    var currentContent = existingItem.Content;
                    var currentName = currentContent is string str ? str : currentContent?.ToString();
                    if (currentName != favorite.Name)
                    {
                        existingItem.Content = favorite.Name;
                    }
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

        #endregion

        #region ナビゲーション

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
                // Application.Currentをキャッシュ（パフォーマンス向上）
                var app = _cachedApplication ?? (_cachedApplication = System.Windows.Application.Current);
                app?.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    // ViewModelをキャッシュから取得（なければ取得してキャッシュ）
                    if (_cachedExplorerPageViewModel == null)
                    {
                        _cachedExplorerPageViewModel = App.Services.GetService(ExplorerPageViewModelType) as ViewModels.Pages.ExplorerPageViewModel;
                    }
                    
                    if (_cachedExplorerPageViewModel != null)
                    {
                        // 分割ペインモードを考慮してホームにナビゲート
                        if (_cachedExplorerPageViewModel.NavigateToHomeInActivePaneCommand?.CanExecute(null) == true)
                        {
                            _cachedExplorerPageViewModel.NavigateToHomeInActivePaneCommand.Execute(null);
                        }
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
                // Application.Currentをキャッシュ（パフォーマンス向上）
                var app = _cachedApplication ?? (_cachedApplication = System.Windows.Application.Current);
                app?.Dispatcher.BeginInvoke(new System.Action(() =>
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

        #endregion

        #region お気に入り操作

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

        #endregion

        #region その他のプロパティ

        [ObservableProperty]
        private ObservableCollection<object> _footerMenuItems = new()
        {
            new NavigationViewItem()
            {
                Content = "Settings",
                Icon = new SymbolIcon { Symbol = SymbolRegular.Settings24 },
                Tag = "SETTINGS" // 設定ページを識別するためのTag
            }
        };

        [ObservableProperty]
        private ObservableCollection<System.Windows.Controls.MenuItem> _trayMenuItems = new()
        {
            new System.Windows.Controls.MenuItem { Header = "Home", Tag = "tray_home" }
        };

        #endregion
    }
}
