using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using FastExplorer.Models;
using FastExplorer.ViewModels.Pages;

namespace FastExplorer.Views.Pages.SettingsPage
{
    /// <summary>
    /// 外観設定ページを表すクラス
    /// </summary>
    public partial class AppearanceSettingsPage : UserControl
    {
        /// <summary>
        /// 設定ページのViewModelを取得または設定します
        /// </summary>
        public SettingsViewModel ViewModel { get; }

        private readonly List<Button> _colorButtons = new List<Button>();

        /// <summary>
        /// <see cref="AppearanceSettingsPage"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">設定ページのViewModel</param>
        public AppearanceSettingsPage(SettingsViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;
            InitializeComponent();

            // ViewModelのPropertyChangedイベントを監視
            ViewModel.PropertyChanged += ViewModel_PropertyChanged;
        }

        private void ThemeColorsItemsControl_Loaded(object sender, RoutedEventArgs e)
        {
            // レンダリングが完了した後にボタンを取得
            var itemsControl = sender as ItemsControl;
            if (itemsControl != null)
            {
                itemsControl.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    // すべてのボタンを取得
                    _colorButtons.Clear();
                    FindVisualChildren<Button>(itemsControl, _colorButtons);
                    // 初期状態を更新
                    UpdateCheckMarks();
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }

        private static void FindVisualChildren<T>(DependencyObject parent, List<T> results) where T : DependencyObject
        {
            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T result)
                {
                    results.Add(result);
                }
                FindVisualChildren<T>(child, results);
            }
        }

        private void ViewModel_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ViewModel.SelectedThemeColor))
            {
                UpdateCheckMarks();
            }
        }

        private void UpdateCheckMarks()
        {
            var selectedColorCode = ViewModel.SelectedThemeColor?.ColorCode;

            foreach (var button in _colorButtons)
            {
                if (button.Tag is ThemeColor themeColor)
                {
                    if (themeColor.ColorCode == selectedColorCode)
                    {
                        // 選択されている場合は✓マークを表示
                        var grid = new Grid();
                        var ellipse = new System.Windows.Shapes.Ellipse
                        {
                            Width = 24,
                            Height = 24,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Fill = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(128, 0, 0, 0))
                        };
                        var textBlock = new System.Windows.Controls.TextBlock
                        {
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            FontSize = 18,
                            FontWeight = FontWeights.Bold,
                            Foreground = System.Windows.Media.Brushes.White,
                            Text = "✓"
                        };
                        grid.Children.Add(ellipse);
                        grid.Children.Add(textBlock);
                        button.Content = grid;
                    }
                    else
                    {
                        // 選択されていない場合はContentをクリア
                        button.Content = null;
                    }
                }
            }
        }

        private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T result)
                {
                    return result;
                }
                var childOfChild = FindVisualChild<T>(child);
                if (childOfChild != null)
                {
                    return childOfChild;
                }
            }
            return null;
        }
    }
}

