using System;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Wpf.Ui.Controls;
using MenuItem = Wpf.Ui.Controls.MenuItem;

namespace FastExplorer.ShellContextMenu
{
    /// <summary>
    /// ListViewの空いている領域を右クリックしたときに表示するコンテキストメニュー
    /// Filesアプリと同じUIコンテキストメニューを表示します
    /// </summary>
    public partial class ListViewEmptyAreaContextMenu : ContextMenu
    {
        private readonly Action<string>? _sortByColumnAction;
        private readonly Action<string>? _setLayoutAction;
        private readonly Action<string, System.Windows.Controls.ListView>? _startRenameAction;
        private readonly System.Windows.Media.Brush _backgroundBrush;
        private readonly System.Windows.Media.Brush _borderBrush;
        private readonly System.Windows.Media.Brush _foregroundBrush;
        private readonly string? _currentPath;
        private readonly ICommand? _refreshCommand;
        private System.Windows.Controls.ListView? _listView;

        /// <summary>
        /// コンテキストメニューを構築します
        /// </summary>
        /// <param name="refreshCommand">最新の情報に更新コマンド</param>
        /// <param name="addToFavoritesCommand">サイドバーにピン留めコマンド</param>
        /// <param name="currentPath">現在のパス</param>
        /// <param name="sortByColumnAction">並べ替えアクション（列名を指定）</param>
        /// <param name="setLayoutAction">レイアウト設定アクション（レイアウト名を指定）</param>
        /// <param name="startRenameAction">新規作成後の名前変更開始アクション（作成したアイテムのパスとListViewを指定）</param>
        public ListViewEmptyAreaContextMenu(
            ICommand? refreshCommand = null,
            ICommand? addToFavoritesCommand = null,
            string? currentPath = null,
            Action<string>? sortByColumnAction = null,
            Action<string>? setLayoutAction = null,
            Action<string, System.Windows.Controls.ListView>? startRenameAction = null)
        {
            InitializeComponent();

            _sortByColumnAction = sortByColumnAction;
            _setLayoutAction = setLayoutAction;
            _startRenameAction = startRenameAction;
            _currentPath = currentPath;
            _refreshCommand = refreshCommand;

            // テーマカラーを取得（Application.Currentから取得）
            _backgroundBrush = Application.Current.TryFindResource("ApplicationBackgroundBrush") as System.Windows.Media.Brush
                ?? Application.Current.TryFindResource("ControlFillColorDefaultBrush") as System.Windows.Media.Brush
                ?? System.Windows.Media.Brushes.White;
            _borderBrush = Application.Current.TryFindResource("ControlStrokeColorDefaultBrush") as System.Windows.Media.Brush
                ?? System.Windows.Media.Brushes.LightGray;
            _foregroundBrush = Application.Current.TryFindResource("TextFillColorPrimaryBrush") as System.Windows.Media.Brush
                ?? System.Windows.Media.Brushes.Black;

            // ContextMenuの背景色とボーダー色を設定
            Background = _backgroundBrush;
            BorderBrush = _borderBrush;

            // 各要素のForegroundを設定
            LayoutIcon.Foreground = _foregroundBrush;
            LayoutText.Foreground = _foregroundBrush;
            SortIcon.Foreground = _foregroundBrush;
            SortText.Foreground = _foregroundBrush;
            SortMenuItem.Background = _backgroundBrush;
            GroupIcon.Foreground = _foregroundBrush;
            GroupText.Foreground = _foregroundBrush;
            RefreshIcon.Foreground = _foregroundBrush;
            RefreshText.Foreground = _foregroundBrush;
            NewIcon.Foreground = _foregroundBrush;
            NewText.Foreground = _foregroundBrush;
            PasteShortcutIcon.Foreground = _foregroundBrush;
            PasteShortcutText.Foreground = _foregroundBrush;
            PinToSidebarIcon.Foreground = _foregroundBrush;
            PinToSidebarText.Foreground = _foregroundBrush;
            PinToStartIcon.Foreground = _foregroundBrush;
            PinToStartText.Foreground = _foregroundBrush;
            OpenTerminalIcon.Foreground = _foregroundBrush;
            OpenTerminalText.Foreground = _foregroundBrush;
            LoadingIcon.Foreground = _foregroundBrush;
            LoadingText.Foreground = _foregroundBrush;

            // レイアウト（サブメニュー）の子項目を追加
            LayoutMenuItem.Items.Add(CreateLayoutSubItem("詳細", "Details"));
            LayoutMenuItem.Items.Add(CreateLayoutSubItem("カード", "Cards"));
            LayoutMenuItem.Items.Add(CreateLayoutSubItem("リスト", "List"));
            LayoutMenuItem.Items.Add(CreateLayoutSubItem("グリッド", "Grid"));
            LayoutMenuItem.Items.Add(CreateLayoutSubItem("列", "Columns"));
            LayoutMenuItem.Items.Add(CreateLayoutSubItem("適応", "Adaptive"));

            // 並べ替え（サブメニュー）の子項目にイベントハンドラーを設定
            SortByNameMenuItem.Click += (s, e) => _sortByColumnAction?.Invoke("Name");
            SortByDateModifiedMenuItem.Click += (s, e) => _sortByColumnAction?.Invoke("DateModified");
            SortByDateCreatedMenuItem.Click += (s, e) => _sortByColumnAction?.Invoke("DateCreated");
            SortByTypeMenuItem.Click += (s, e) => _sortByColumnAction?.Invoke("Type");
            SortBySizeMenuItem.Click += (s, e) => _sortByColumnAction?.Invoke("Size");
            SortAscendingMenuItem.Click += (s, e) => _sortByColumnAction?.Invoke("Ascending");
            SortDescendingMenuItem.Click += (s, e) => _sortByColumnAction?.Invoke("Descending");
            
            // 並び替えメニュー項目のスタイルを設定
            SortByNameMenuItem.Background = _backgroundBrush;
            SortByNameMenuItem.Foreground = _foregroundBrush;
            SortByDateModifiedMenuItem.Background = _backgroundBrush;
            SortByDateModifiedMenuItem.Foreground = _foregroundBrush;
            SortByDateCreatedMenuItem.Background = _backgroundBrush;
            SortByDateCreatedMenuItem.Foreground = _foregroundBrush;
            SortByTypeMenuItem.Background = _backgroundBrush;
            SortByTypeMenuItem.Foreground = _foregroundBrush;
            SortBySizeMenuItem.Background = _backgroundBrush;
            SortBySizeMenuItem.Foreground = _foregroundBrush;
            SortAscendingMenuItem.Background = _backgroundBrush;
            SortAscendingMenuItem.Foreground = _foregroundBrush;
            SortDescendingMenuItem.Background = _backgroundBrush;
            SortDescendingMenuItem.Foreground = _foregroundBrush;

            // コマンドの設定
            RefreshMenuItem.Command = refreshCommand;
            PinToSidebarMenuItem.Command = addToFavoritesCommand;

            // 新規作成メニューのクリックイベントを設定
            NewFolderMenuItem.Click += NewFolderMenuItem_Click;
            NewTextDocumentMenuItem.Click += NewTextDocumentMenuItem_Click;
            NewShortcutMenuItem.Click += NewShortcutMenuItem_Click;

            // OpenedイベントでListViewを保存
            this.Opened += (s, e) =>
            {
                if (PlacementTarget is System.Windows.Controls.ListView listView)
                {
                    _listView = listView;
                }
            };

            // サブメニューを持つMenuItemのMouseEnterイベントを設定
            SetupSubmenuMouseEnter(LayoutMenuItem);
            SetupSubmenuMouseEnter(SortMenuItem);
            SetupSubmenuMouseEnter(GroupMenuItem);
            SetupSubmenuMouseEnter(NewMenuItem);

            // サブメニューのスタイルを設定
            SetupSubmenuStyle(LayoutMenuItem);
            SetupSubmenuStyle(SortMenuItem);
            SetupSubmenuStyle(GroupMenuItem);
            SetupSubmenuStyle(NewMenuItem);
        }

        /// <summary>
        /// サブメニューを持つMenuItemのMouseEnterイベントを設定します
        /// </summary>
        /// <param name="menuItem">対象のMenuItem</param>
        private void SetupSubmenuMouseEnter(MenuItem menuItem)
        {
            if (menuItem == null || menuItem.Items.Count == 0)
                return;

            // MenuItem全体でマウスイベントを処理するように設定
            menuItem.MouseEnter += (s, e) =>
            {
                // サブメニューを開く
                menuItem.IsSubmenuOpen = true;
            };

            // Header内の要素とその子要素にもマウスイベントを設定
            if (menuItem.Header is System.Windows.FrameworkElement headerElement)
            {
                SetupMouseEnterForElement(headerElement, menuItem);
            }
        }

        /// <summary>
        /// 要素とその子要素にMouseEnterイベントを設定します
        /// </summary>
        /// <param name="element">対象の要素</param>
        /// <param name="menuItem">MenuItem</param>
        private void SetupMouseEnterForElement(System.Windows.FrameworkElement element, MenuItem menuItem)
        {
            if (element == null)
                return;

            // 要素自体にMouseEnterイベントを設定
            element.MouseEnter += (s, e) =>
            {
                menuItem.IsSubmenuOpen = true;
            };

            // 子要素にも再帰的に設定
            if (element is System.Windows.Controls.Panel panel)
            {
                foreach (var child in panel.Children)
                {
                    if (child is System.Windows.FrameworkElement childElement)
                    {
                        SetupMouseEnterForElement(childElement, menuItem);
                    }
                }
            }
        }

        /// <summary>
        /// サブメニューのスタイルを設定します（メインメニューと同じ見た目にする）
        /// </summary>
        /// <param name="menuItem">対象のMenuItem</param>
        private void SetupSubmenuStyle(MenuItem menuItem)
        {
            if (menuItem == null || menuItem.Items.Count == 0)
                return;

            menuItem.SubmenuOpened += (s, e) =>
            {
                // サブメニューが開いた後にスタイルを適用（少し遅延させる）
                Dispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(() =>
                {
                    ApplySubmenuStyle(menuItem);
                }));
            };
        }

        /// <summary>
        /// サブメニューのスタイルを適用します
        /// </summary>
        /// <param name="menuItem">対象のMenuItem</param>
        private void ApplySubmenuStyle(MenuItem menuItem)
        {
            // サブメニューのPopupを取得
            var popup = FindVisualChild<System.Windows.Controls.Primitives.Popup>(menuItem);
            if (popup != null && popup.Child is System.Windows.FrameworkElement popupChild)
            {
                // Popup内のBorderを探してスタイルを適用
                var border = FindVisualChild<System.Windows.Controls.Border>(popupChild);
                if (border != null)
                {
                    border.Background = _backgroundBrush;
                    border.BorderBrush = _borderBrush;
                    border.BorderThickness = new Thickness(1);
                    border.CornerRadius = new CornerRadius(8);
                    border.Effect = new System.Windows.Media.Effects.DropShadowEffect
                    {
                        BlurRadius = 18,
                        Opacity = 0.8,
                        ShadowDepth = 0,
                        Color = System.Windows.Media.Color.FromArgb(0x40, 0x00, 0x00, 0x00)
                    };
                }
            }
        }

        /// <summary>
        /// ビジュアルツリーから指定された型の子要素を検索します
        /// </summary>
        private T? FindVisualChild<T>(System.Windows.DependencyObject parent) where T : System.Windows.DependencyObject
        {
            if (parent == null)
                return null;

            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T result)
                    return result;

                var childOfChild = FindVisualChild<T>(child);
                if (childOfChild != null)
                    return childOfChild;
            }

            return null;
        }

        private MenuItem CreateLayoutSubItem(string text, string layoutName)
        {
            var item = new MenuItem { Header = text, Background = _backgroundBrush, Foreground = _foregroundBrush };
            item.Click += (s, e) => _setLayoutAction?.Invoke(layoutName);
            return item;
        }

        private MenuItem CreateSortSubItem(string text, string columnName)
        {
            var item = new MenuItem { Header = text, Background = _backgroundBrush, Foreground = _foregroundBrush };
            item.Click += (s, e) => _sortByColumnAction?.Invoke(columnName);
            return item;
        }

        /// <summary>
        /// 新規作成したアイテムの名前変更モードを開始します
        /// </summary>
        /// <param name="itemPath">作成したアイテムのパス</param>
        private void StartRenameForNewItem(string itemPath)
        {
            if (_listView == null || _startRenameAction == null)
                return;

            // リスト更新後に名前変更を開始するため、少し遅延させる
            Dispatcher.CurrentDispatcher.BeginInvoke(
                DispatcherPriority.Loaded,
                () => _startRenameAction(itemPath, _listView));
        }

        /// <summary>
        /// エラーダイアログを表示します
        /// </summary>
        private async System.Threading.Tasks.Task ShowErrorDialogAsync(string title, string message)
        {
            try
            {
                // MainWindowからRootContentDialogを取得
                var mainWindow = Application.Current.MainWindow;
                if (mainWindow != null)
                {
                    var contentPresenter = mainWindow.FindName("RootContentDialog") as ContentPresenter;
                    if (contentPresenter != null)
                    {
                        var contentDialog = new ContentDialog(contentPresenter)
                        {
                            Title = title,
                            Content = message,
                            CloseButtonText = "閉じる"
                        };

                        await contentDialog.ShowAsync();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"エラーダイアログ表示エラー: {ex.Message}");
                var messageBox = new Wpf.Ui.Controls.MessageBox
                {
                    Title = title,
                    Content = message,
                    CloseButtonText = "閉じる"
                };
                await messageBox.ShowDialogAsync();
            }
        }

        /// <summary>
        /// フォルダーを作成します
        /// </summary>
        private async void NewFolderMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentPath) || !Directory.Exists(_currentPath))
            {
                return;
            }

            try
            {
                // 新しいフォルダー名を生成（既存のフォルダー名と重複しないようにする）
                string newFolderName = "新しいフォルダー";
                string newFolderPath = Path.Combine(_currentPath, newFolderName);
                int counter = 1;

                while (Directory.Exists(newFolderPath))
                {
                    newFolderName = $"新しいフォルダー ({counter})";
                    newFolderPath = Path.Combine(_currentPath, newFolderName);
                    counter++;
                }

                // フォルダーを作成
                Directory.CreateDirectory(newFolderPath);

                // リストを更新
                _refreshCommand?.Execute(null);

                // 名前変更モードを開始
                StartRenameForNewItem(newFolderPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"フォルダー作成エラー: {ex.Message}");
                await ShowErrorDialogAsync("エラー", $"フォルダーの作成に失敗しました。\n{ex.Message}");
            }
        }

        /// <summary>
        /// テキストファイルを作成します
        /// </summary>
        private async void NewTextDocumentMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentPath) || !Directory.Exists(_currentPath))
            {
                return;
            }

            try
            {
                // 新しいテキストファイル名を生成（既存のファイル名と重複しないようにする）
                string newFileName = "新しいテキスト ドキュメント.txt";
                string newFilePath = Path.Combine(_currentPath, newFileName);
                int counter = 1;

                while (File.Exists(newFilePath))
                {
                    newFileName = $"新しいテキスト ドキュメント ({counter}).txt";
                    newFilePath = Path.Combine(_currentPath, newFileName);
                    counter++;
                }

                // テキストファイルを作成（空のファイル）
                File.Create(newFilePath).Dispose();

                // リストを更新
                _refreshCommand?.Execute(null);

                // 名前変更モードを開始
                StartRenameForNewItem(newFilePath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"テキストファイル作成エラー: {ex.Message}");
                await ShowErrorDialogAsync("エラー", $"テキストファイルの作成に失敗しました。\n{ex.Message}");
            }
        }

        /// <summary>
        /// ショートカットを作成します
        /// </summary>
        private async void NewShortcutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentPath) || !Directory.Exists(_currentPath))
            {
                return;
            }

            try
            {
                // ショートカットの作成先を選択するダイアログを表示
                var dialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "ショートカットの作成先を選択",
                    Filter = "すべてのファイル (*.*)|*.*",
                    CheckFileExists = true
                };

                if (dialog.ShowDialog() == true)
                {
                    string targetPath = dialog.FileName;
                    string shortcutName = Path.GetFileNameWithoutExtension(targetPath) + ".lnk";
                    string shortcutPath = Path.Combine(_currentPath, shortcutName);
                    int counter = 1;

                    // 既存のショートカット名と重複しないようにする
                    while (File.Exists(shortcutPath))
                    {
                        shortcutName = $"{Path.GetFileNameWithoutExtension(targetPath)} ({counter}).lnk";
                        shortcutPath = Path.Combine(_currentPath, shortcutName);
                        counter++;
                    }

                    // ショートカットを作成
                    CreateShortcut(shortcutPath, targetPath);

                    // リストを更新
                    _refreshCommand?.Execute(null);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ショートカット作成エラー: {ex.Message}");
                await ShowErrorDialogAsync("エラー", $"ショートカットの作成に失敗しました。\n{ex.Message}");
            }
        }

        /// <summary>
        /// ショートカット（.lnkファイル）を作成します
        /// </summary>
        private void CreateShortcut(string shortcutPath, string targetPath)
        {
            // WScript.Shell COMオブジェクトを使用してショートカットを作成
            var shellLinkType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellLinkType == null)
            {
                throw new InvalidOperationException("WScript.Shell COMオブジェクトを作成できませんでした。");
            }

            dynamic shell = Activator.CreateInstance(shellLinkType);
            dynamic shortcut = shell.CreateShortcut(shortcutPath);
            shortcut.TargetPath = targetPath;
            shortcut.Description = $"ショートカット: {targetPath}";
            shortcut.WorkingDirectory = Path.GetDirectoryName(targetPath) ?? "";
            shortcut.Save();
        }
    }
}

