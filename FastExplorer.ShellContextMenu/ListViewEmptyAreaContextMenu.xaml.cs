using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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
        private readonly System.Windows.Media.Brush _backgroundBrush;
        private readonly System.Windows.Media.Brush _borderBrush;
        private readonly System.Windows.Media.Brush _foregroundBrush;

        /// <summary>
        /// コンテキストメニューを構築します
        /// </summary>
        /// <param name="refreshCommand">最新の情報に更新コマンド</param>
        /// <param name="addToFavoritesCommand">サイドバーにピン留めコマンド</param>
        /// <param name="currentPath">現在のパス</param>
        /// <param name="sortByColumnAction">並べ替えアクション（列名を指定）</param>
        /// <param name="setLayoutAction">レイアウト設定アクション（レイアウト名を指定）</param>
        public ListViewEmptyAreaContextMenu(
            ICommand? refreshCommand = null,
            ICommand? addToFavoritesCommand = null,
            string? currentPath = null,
            Action<string>? sortByColumnAction = null,
            Action<string>? setLayoutAction = null)
        {
            InitializeComponent();

            _sortByColumnAction = sortByColumnAction;
            _setLayoutAction = setLayoutAction;

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

            // 並べ替え（サブメニュー）の子項目を追加
            SortMenuItem.Items.Add(CreateSortSubItem("名前", "Name"));
            SortMenuItem.Items.Add(CreateSortSubItem("変更日時", "DateModified"));
            SortMenuItem.Items.Add(CreateSortSubItem("作成日時", "DateCreated"));
            SortMenuItem.Items.Add(CreateSortSubItem("種類", "Type"));
            SortMenuItem.Items.Add(CreateSortSubItem("サイズ", "Size"));
            SortMenuItem.Items.Add(new Separator());
            SortMenuItem.Items.Add(CreateSortSubItem("昇順", "Ascending"));
            SortMenuItem.Items.Add(CreateSortSubItem("降順", "Descending"));

            // コマンドの設定
            RefreshMenuItem.Command = refreshCommand;
            PinToSidebarMenuItem.Command = addToFavoritesCommand;
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
    }
}

