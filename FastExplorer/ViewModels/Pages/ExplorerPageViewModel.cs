using System.Collections.ObjectModel;
using System.IO;
using FastExplorer.Models;
using FastExplorer.Services;
using Wpf.Ui.Abstractions.Controls;

namespace FastExplorer.ViewModels.Pages
{
    /// <summary>
    /// エクスプローラーページのタブ管理を行うViewModel
    /// </summary>
    public partial class ExplorerPageViewModel : ObservableObject, INavigationAware
    {
        private readonly FileSystemService _fileSystemService;
        private readonly FavoriteService? _favoriteService;

        /// <summary>
        /// エクスプローラータブのコレクション
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<ExplorerTab> _tabs = new();

        /// <summary>
        /// 現在選択されているタブ
        /// </summary>
        [ObservableProperty]
        private ExplorerTab? _selectedTab;

        /// <summary>
        /// <see cref="ExplorerPageViewModel"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="fileSystemService">ファイルシステムサービス</param>
        /// <param name="favoriteService">お気に入りサービス</param>
        public ExplorerPageViewModel(FileSystemService fileSystemService, FavoriteService? favoriteService = null)
        {
            _fileSystemService = fileSystemService;
            _favoriteService = favoriteService;
        }

        /// <summary>
        /// ページにナビゲートされたときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedToAsync()
        {
            if (Tabs.Count == 0)
            {
                CreateNewTab();
            }
            return Task.CompletedTask;
        }

        /// <summary>
        /// ページから離れるときに呼び出されます
        /// </summary>
        /// <returns>完了を表すタスク</returns>
        public Task OnNavigatedFromAsync() => Task.CompletedTask;

        /// <summary>
        /// 新しいタブを作成します
        /// </summary>
        [RelayCommand]
        private void CreateNewTab()
        {
            var viewModel = new ExplorerViewModel(_fileSystemService);
            var tab = new ExplorerTab
            {
                Title = "PC",
                CurrentPath = string.Empty,
                ViewModel = viewModel
            };

            // CurrentPathが変更されたときにTitleとCurrentPathを更新
            viewModel.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(ExplorerViewModel.CurrentPath))
                {
                    tab.CurrentPath = viewModel.CurrentPath;
                    UpdateTabTitle(tab);
                    UpdateStatusBar();
                }
                else if (e.PropertyName == nameof(ExplorerViewModel.Items))
                {
                    UpdateStatusBar();
                }
            };

            // タブを追加する前に初期化を完了させる
            // これにより、タブが追加された直後にフォルダーをダブルクリックしても問題が発生しない
            tab.ViewModel.NavigateToDrives();
            
            Tabs.Add(tab);
            SelectedTab = tab;
            
            UpdateTabTitle(tab);
        }

        /// <summary>
        /// 指定されたタブを閉じます
        /// </summary>
        /// <param name="tab">閉じるタブ。nullの場合、またはタブが1つしかない場合は何もしません</param>
        [RelayCommand]
        private void CloseTab(ExplorerTab? tab)
        {
            if (tab == null || Tabs.Count <= 1)
                return;

            var index = Tabs.IndexOf(tab);
            Tabs.Remove(tab);

            if (SelectedTab == tab)
            {
                if (index > 0)
                    SelectedTab = Tabs[index - 1];
                else if (Tabs.Count > 0)
                    SelectedTab = Tabs[0];
            }
        }

        /// <summary>
        /// タブのタイトルを現在のパスに基づいて更新します
        /// </summary>
        /// <param name="tab">タイトルを更新するタブ</param>
        private void UpdateTabTitle(ExplorerTab tab)
        {
            if (string.IsNullOrEmpty(tab.ViewModel.CurrentPath))
            {
                tab.Title = "PC";
            }
            else
            {
                try
                {
                    var dirInfo = new DirectoryInfo(tab.ViewModel.CurrentPath);
                    tab.Title = dirInfo.Name;
                }
                catch
                {
                    tab.Title = Path.GetFileName(tab.ViewModel.CurrentPath) ?? "不明";
                }
            }
        }

        /// <summary>
        /// 現在のパスをお気に入りに追加します
        /// </summary>
        [RelayCommand]
        private void AddCurrentPathToFavorites()
        {
            if (SelectedTab == null || string.IsNullOrEmpty(SelectedTab.ViewModel.CurrentPath))
                return;

            var path = SelectedTab.ViewModel.CurrentPath;
            var name = Path.GetFileName(path) ?? path;
            
            _favoriteService?.AddFavorite(name, path);
            
            // MainWindowViewModelを更新
            var mainWindowViewModel = App.Services.GetService(typeof(ViewModels.Windows.MainWindowViewModel)) as ViewModels.Windows.MainWindowViewModel;
            mainWindowViewModel?.LoadFavorites();
        }

        /// <summary>
        /// SelectedTabが変更されたときに呼び出されます
        /// </summary>
        partial void OnSelectedTabChanged(ExplorerTab? value)
        {
            UpdateStatusBar();
        }

        /// <summary>
        /// ステータスバーのテキストを更新します
        /// </summary>
        private void UpdateStatusBar()
        {
            var mainWindowViewModel = App.Services.GetService(typeof(ViewModels.Windows.MainWindowViewModel)) as ViewModels.Windows.MainWindowViewModel;
            if (mainWindowViewModel == null)
                return;

            if (SelectedTab?.ViewModel != null)
            {
                var path = SelectedTab.ViewModel.CurrentPath;
                var itemCount = SelectedTab.ViewModel.Items.Count;
                
                if (string.IsNullOrEmpty(path))
                {
                    mainWindowViewModel.StatusBarText = $"パス: PC {itemCount}個の項目";
                }
                else
                {
                    mainWindowViewModel.StatusBarText = $"パス: {path} {itemCount}個の項目";
                }
            }
            else
            {
                mainWindowViewModel.StatusBarText = "準備完了";
            }
        }
    }
}

