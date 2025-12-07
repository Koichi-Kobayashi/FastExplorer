using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Wpf.Ui.Controls;
using MenuItem = Wpf.Ui.Controls.MenuItem;
using TextBlock = System.Windows.Controls.TextBlock;

namespace FastExplorer.ShellContextMenu
{
    /// <summary>
    /// ListViewの空いている領域を右クリックしたときに表示するコンテキストメニュー
    /// </summary>
    public class ListViewEmptyAreaContextMenu : ContextMenu
    {
        /// <summary>
        /// コンテキストメニューを構築します
        /// </summary>
        /// <param name="refreshCommand">最新の情報に更新コマンド</param>
        /// <param name="addToFavoritesCommand">サイドバーにピン留めコマンド</param>
        /// <param name="currentPath">現在のパス</param>
        public ListViewEmptyAreaContextMenu(
            ICommand? refreshCommand = null,
            ICommand? addToFavoritesCommand = null,
            string? currentPath = null)
        {
            // レイアウト
            var layoutMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.Grid24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0) },
                        new TextBlock { Text = "レイアウト", VerticalAlignment = VerticalAlignment.Center }
                    }
                }
            };
            layoutMenuItem.Items.Add(new MenuItem { Header = "詳細" });
            layoutMenuItem.Items.Add(new MenuItem { Header = "一覧" });
            layoutMenuItem.Items.Add(new MenuItem { Header = "タイル" });
            layoutMenuItem.Items.Add(new MenuItem { Header = "コンテンツ" });
            layoutMenuItem.Items.Add(new MenuItem { Header = "特大アイコン" });
            layoutMenuItem.Items.Add(new MenuItem { Header = "大アイコン" });
            layoutMenuItem.Items.Add(new MenuItem { Header = "中アイコン" });
            layoutMenuItem.Items.Add(new MenuItem { Header = "小アイコン" });

            // 並べ替え
            var sortMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.ArrowSort24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0) },
                        new TextBlock { Text = "並べ替え", VerticalAlignment = VerticalAlignment.Center }
                    }
                }
            };
            sortMenuItem.Items.Add(new MenuItem { Header = "名前" });
            sortMenuItem.Items.Add(new MenuItem { Header = "日付" });
            sortMenuItem.Items.Add(new MenuItem { Header = "種類" });
            sortMenuItem.Items.Add(new MenuItem { Header = "サイズ" });

            // グループで表示
            var groupMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.GroupList24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0) },
                        new TextBlock { Text = "グループで表示", VerticalAlignment = VerticalAlignment.Center }
                    }
                }
            };
            groupMenuItem.Items.Add(new MenuItem { Header = "なし" });
            groupMenuItem.Items.Add(new MenuItem { Header = "名前" });
            groupMenuItem.Items.Add(new MenuItem { Header = "日付" });
            groupMenuItem.Items.Add(new MenuItem { Header = "種類" });
            groupMenuItem.Items.Add(new MenuItem { Header = "サイズ" });

            // 最新の情報に更新
            var refreshMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.ArrowClockwise24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0) },
                        new TextBlock { Text = "最新の情報に更新", VerticalAlignment = VerticalAlignment.Center }
                    }
                },
                Command = refreshCommand,
                InputGestureText = "Ctrl+R"
            };

            Items.Add(layoutMenuItem);
            Items.Add(sortMenuItem);
            Items.Add(groupMenuItem);
            Items.Add(refreshMenuItem);
            Items.Add(new Separator());

            // 新規作成
            var newMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.AddCircle24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0) },
                        new TextBlock { Text = "新規作成", VerticalAlignment = VerticalAlignment.Center }
                    }
                }
            };
            newMenuItem.Items.Add(new MenuItem { Header = "フォルダー" });
            newMenuItem.Items.Add(new MenuItem { Header = "ショートカット" });
            newMenuItem.Items.Add(new MenuItem { Header = "テキスト ドキュメント" });

            // ショートカットを貼り付け
            var pasteShortcutMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.Copy24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0) },
                        new TextBlock { Text = "ショートカットを貼り付け", VerticalAlignment = VerticalAlignment.Center }
                    }
                }
            };

            // サイドバーにピン留めする
            var pinMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.Star24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0) },
                        new TextBlock { Text = "サイドバーにピン留めする", VerticalAlignment = VerticalAlignment.Center }
                    }
                },
                Command = addToFavoritesCommand
            };

            Items.Add(newMenuItem);
            Items.Add(pasteShortcutMenuItem);
            Items.Add(pinMenuItem);
            Items.Add(new Separator());

            // ターミナルで開く
            var terminalMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.Window24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0) },
                        new TextBlock { Text = "ターミナルで開く", VerticalAlignment = VerticalAlignment.Center }
                    }
                },
                InputGestureText = "Ctrl+@"
            };
            terminalMenuItem.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(currentPath) && Directory.Exists(currentPath))
                {
                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "wt.exe",
                            Arguments = $"-d \"{currentPath}\"",
                            UseShellExecute = true
                        });
                    }
                    catch
                    {
                        // wt.exeが見つからない場合は、cmd.exeを使用
                        try
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = "cmd.exe",
                                Arguments = $"/k cd /d \"{currentPath}\"",
                                UseShellExecute = true
                            });
                        }
                        catch
                        {
                            // エラーは無視
                        }
                    }
                }
            };

            Items.Add(terminalMenuItem);
            Items.Add(new Separator());

            // その他のオプションを表示
            var moreOptionsMenuItem = new MenuItem
            {
                Header = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new SymbolIcon { Symbol = SymbolRegular.MoreHorizontal24, Width = 16, Height = 16, Margin = new Thickness(-8, 0, 12, 0) },
                        new TextBlock { Text = "その他のオプションを表示", VerticalAlignment = VerticalAlignment.Center }
                    }
                }
            };
            moreOptionsMenuItem.Items.Add(new MenuItem { Header = "オプション" });

            Items.Add(moreOptionsMenuItem);
        }
    }
}
