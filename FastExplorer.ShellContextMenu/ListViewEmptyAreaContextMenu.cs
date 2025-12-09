using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Xml;
using Wpf.Ui.Controls;
using MenuItem = Wpf.Ui.Controls.MenuItem;
using TextBlock = System.Windows.Controls.TextBlock;

namespace FastExplorer.ShellContextMenu
{
    /// <summary>
    /// ListViewの空いている領域を右クリックしたときに表示するコンテキストメニュー
    /// Filesアプリと同じUIコンテキストメニューを表示します
    /// </summary>
    public class ListViewEmptyAreaContextMenu : ContextMenu
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

            // Filesアプリのような見た目にするためのスタイル設定（テーマカラーを使用）
            Background = _backgroundBrush;
            BorderBrush = _borderBrush;
            BorderThickness = new Thickness(1);
            Padding = new Thickness(4);
            MinWidth = 200;

            // レイアウト（サブメニュー）
            var layoutMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.Grid24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = _foregroundBrush },
                        new TextBlock { Text = "レイアウト", VerticalAlignment = VerticalAlignment.Center, FontSize = 14, Foreground = _foregroundBrush }
                    }
                },
                Padding = new Thickness(-8, 6, 8, 6),
                MinHeight = 32
            };
            layoutMenuItem.Items.Add(CreateLayoutSubItem("詳細", "Details"));
            layoutMenuItem.Items.Add(CreateLayoutSubItem("カード", "Cards"));
            layoutMenuItem.Items.Add(CreateLayoutSubItem("リスト", "List"));
            layoutMenuItem.Items.Add(CreateLayoutSubItem("グリッド", "Grid"));
            layoutMenuItem.Items.Add(CreateLayoutSubItem("列", "Columns"));
            layoutMenuItem.Items.Add(CreateLayoutSubItem("適応", "Adaptive"));
            Items.Add(layoutMenuItem);

            // ★ ContextMenu を角丸にする ControlTemplate を C# で適用
            var template = new ControlTemplate(typeof(ContextMenu));

#pragma warning disable 618
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(8));               // ← 角丸
            border.SetValue(Border.BackgroundProperty, _backgroundBrush);
            border.SetValue(Border.BorderBrushProperty, _borderBrush);
            border.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            border.SetValue(Border.PaddingProperty, new Thickness(4));
            border.SetValue(Border.SnapsToDevicePixelsProperty, true);

            // 影（Windows 11 らしさUP）
            border.SetValue(Border.EffectProperty, new System.Windows.Media.Effects.DropShadowEffect
            {
                Color = System.Windows.Media.Color.FromArgb(80, 0, 0, 0),
                BlurRadius = 16,
                ShadowDepth = 0,
                Opacity = 0.6
            });

            // メニュー内部
            var scrollViewer = new FrameworkElementFactory(typeof(ScrollViewer));
            scrollViewer.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
            scrollViewer.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Disabled);

            var itemsPresenter = new FrameworkElementFactory(typeof(ItemsPresenter));
            scrollViewer.AppendChild(itemsPresenter);

            border.AppendChild(scrollViewer);
            template.VisualTree = border;
#pragma warning restore 618

            this.Template = template;

            // 並べ替え（サブメニュー）
            var sortMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.ArrowSort24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = _foregroundBrush },
                        new TextBlock { Text = "並べ替え", VerticalAlignment = VerticalAlignment.Center, FontSize = 14, Foreground = _foregroundBrush }
                    }
                },
                Padding = new Thickness(-8, 6, 8, 6),
                MinHeight = 32
            };
            sortMenuItem.Background = _backgroundBrush;
            sortMenuItem.Items.Add(CreateSortSubItem("名前", "Name"));
            sortMenuItem.Items.Add(CreateSortSubItem("変更日時", "DateModified"));
            sortMenuItem.Items.Add(CreateSortSubItem("作成日時", "DateCreated"));
            sortMenuItem.Items.Add(CreateSortSubItem("種類", "Type"));
            sortMenuItem.Items.Add(CreateSortSubItem("サイズ", "Size"));
            sortMenuItem.Items.Add(new System.Windows.Controls.Separator());
            sortMenuItem.Items.Add(CreateSortSubItem("昇順", "Ascending"));
            sortMenuItem.Items.Add(CreateSortSubItem("降順", "Descending"));
            Items.Add(sortMenuItem);

            // グループで表示（サブメニュー）
            var groupMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.GroupList24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = _foregroundBrush },
                        new TextBlock { Text = "グループで表示", VerticalAlignment = VerticalAlignment.Center, FontSize = 14, Foreground = _foregroundBrush }
                    }
                },
                Padding = new Thickness(-8, 6, 8, 6),
                MinHeight = 32
            };
            groupMenuItem.Items.Add(new MenuItem { Header = "なし" });
            groupMenuItem.Items.Add(new MenuItem { Header = "名前" });
            groupMenuItem.Items.Add(new MenuItem { Header = "種類" });
            groupMenuItem.Items.Add(new MenuItem { Header = "変更日時" });
            Items.Add(groupMenuItem);

            // 最新の情報に更新
            var refreshMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.ArrowClockwise24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = _foregroundBrush },
                        new TextBlock { Text = "最新の情報に更新", VerticalAlignment = VerticalAlignment.Center, FontSize = 14, Foreground = _foregroundBrush },
                        new TextBlock { Text = "Ctrl+R", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(20, 0, 0, 0), Foreground = System.Windows.Media.Brushes.Gray, FontSize = 12 }
                    }
                },
                Command = refreshCommand,
                Padding = new Thickness(-8, 6, 8, 6),
                MinHeight = 32
            };
            Items.Add(refreshMenuItem);

            Items.Add(new System.Windows.Controls.Separator());

            // 新規作成（サブメニュー）
            var newMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.AddCircle24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = _foregroundBrush },
                        new TextBlock { Text = "新規作成", VerticalAlignment = VerticalAlignment.Center, FontSize = 14, Foreground = _foregroundBrush }
                    }
                },
                Padding = new Thickness(-8, 6, 8, 6),
                MinHeight = 32
            };
            newMenuItem.Items.Add(new MenuItem { Header = "フォルダー" });
            newMenuItem.Items.Add(new MenuItem { Header = "ショートカット" });
            newMenuItem.Items.Add(new MenuItem { Header = "テキスト ドキュメント" });
            Items.Add(newMenuItem);

            // ショートカットを貼り付け
            var pasteShortcutMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.Document24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = _foregroundBrush },
                        new TextBlock { Text = "ショートカットを貼り付け", VerticalAlignment = VerticalAlignment.Center, FontSize = 14, Foreground = _foregroundBrush }
                    }
                },
                Padding = new Thickness(-8, 6, 8, 6),
                MinHeight = 32
            };
            Items.Add(pasteShortcutMenuItem);

            // サイドバーにピン留めする
            var pinToSidebarMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.Star24 , Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = _foregroundBrush },
                        new TextBlock { Text = "サイドバーにピン留めする", VerticalAlignment = VerticalAlignment.Center, FontSize = 14, Foreground = _foregroundBrush }
                    }
                },
                Command = addToFavoritesCommand,
                Padding = new Thickness(-8, 6, 8, 6),
                MinHeight = 32
            };
            Items.Add(pinToSidebarMenuItem);

            // スタートメニューにピン留めする
            var pinToStartMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.Star24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = _foregroundBrush },
                        new TextBlock { Text = "スタートメニューにピン留めする", VerticalAlignment = VerticalAlignment.Center, FontSize = 14, Foreground = _foregroundBrush }
                    }
                },
                Padding = new Thickness(-8, 6, 8, 6),
                MinHeight = 32
            };
            Items.Add(pinToStartMenuItem);

            Items.Add(new System.Windows.Controls.Separator());

            // ターミナルで開く
            var openTerminalMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.Window24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = _foregroundBrush },
                        new TextBlock { Text = "ターミナルで開く", VerticalAlignment = VerticalAlignment.Center, FontSize = 14, Foreground = _foregroundBrush },
                        new TextBlock { Text = "Ctrl+@", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(20, 0, 0, 0), Foreground = System.Windows.Media.Brushes.Gray, FontSize = 12 }
                    }
                },
                Padding = new Thickness(-8, 6, 8, 6),
                MinHeight = 32
            };
            Items.Add(openTerminalMenuItem);

            // 読み込み中...（サブメニュー）- 通常は非表示、必要に応じて表示
            var loadingMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.MoreHorizontal24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = _foregroundBrush },
                        new TextBlock { Text = "読み込み中...", VerticalAlignment = VerticalAlignment.Center, FontSize = 14, Foreground = _foregroundBrush }
                    }
                },
                Visibility = Visibility.Collapsed // デフォルトでは非表示
            };
            Items.Add(loadingMenuItem);

            // ここまでで Items.Add(loadingMenuItem); まで終わっている想定

            // ---- ここから Template 差し替え ----
            const string templateXaml = @"
<ControlTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
                 xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
                 TargetType='ContextMenu'>
    <Border x:Name='BackgroundBorder'
            Background='{TemplateBinding Background}'
            BorderBrush='{TemplateBinding BorderBrush}'
            BorderThickness='{TemplateBinding BorderThickness}'
            CornerRadius='8'
            Padding='{TemplateBinding Padding}'
            SnapsToDevicePixels='True'
            Margin='4'>
        <Border.Effect>
            <DropShadowEffect Color='#40000000'
                              BlurRadius='18'
                              ShadowDepth='0'
                              Opacity='0.8' />
        </Border.Effect>
        <ScrollViewer CanContentScroll='True'
                      VerticalScrollBarVisibility='Auto'
                      HorizontalScrollBarVisibility='Disabled'>
            <ItemsPresenter KeyboardNavigation.DirectionalNavigation='Cycle' />
        </ScrollViewer>
    </Border>
</ControlTemplate>";

            using (var stringReader = new StringReader(templateXaml))
            using (var xmlReader = XmlReader.Create(stringReader))
            {
                this.Template = (ControlTemplate)XamlReader.Load(xmlReader);
            }

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
