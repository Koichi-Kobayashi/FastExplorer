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
        /// <param name="refreshCommand">最新の情報に更新コマンド（使用されません）</param>
        /// <param name="addToFavoritesCommand">サイドバーにピン留めコマンド（使用されません）</param>
        /// <param name="currentPath">現在のパス（使用されません）</param>
        public ListViewEmptyAreaContextMenu(
            ICommand? refreshCommand = null,
            ICommand? addToFavoritesCommand = null,
            string? currentPath = null)
        {
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

            Items.Add(newMenuItem);
        }
    }
}
