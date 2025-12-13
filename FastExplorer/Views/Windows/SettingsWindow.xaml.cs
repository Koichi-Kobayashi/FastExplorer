using System.ComponentModel;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using FastExplorer.ViewModels.Pages;
using FastExplorer.Views.Pages.SettingsPage;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

namespace FastExplorer.Views.Windows
{
    /// <summary>
    /// 設定ウィンドウを表すクラス
    /// </summary>
    public partial class SettingsWindow : FluentWindow
    {
        /// <summary>
        /// 設定ページのViewModelを取得します
        /// </summary>
        public SettingsViewModel ViewModel { get; }

        private Type? _cachedArgsType;
        private PropertyInfo? _cachedInvokedItemContainerProperty;

        /// <summary>
        /// <see cref="SettingsWindow"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="viewModel">設定ページのViewModel</param>
        public SettingsWindow(SettingsViewModel viewModel)
        {
            ViewModel = viewModel;
            DataContext = this;

            InitializeComponent();
            
            // TitleBarの×ボタンイベントを処理
            TitleBar.CloseClicked += TitleBar_CloseClicked;
            
            // ウィンドウ読み込み時にViewModelを初期化
            Loaded += async (s, e) =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("SettingsWindow Loaded event started");
                    await ViewModel.OnNavigatedToAsync();
                    System.Diagnostics.Debug.WriteLine("SettingsWindow OnNavigatedToAsync completed");
                    
                    // システムテーマの監視を設定（テーマが"System"の場合のみ）
                    // SettingsWindowは設定画面なので、常にシステムテーマの変更を監視する必要はない
                    // ただし、設定画面自体のテーマはシステムテーマに追従させる
                    SystemThemeWatcher.Watch(this);
                    
                    // NavigationViewのイベントを明示的に接続（コードビハインドで再接続）
                    SettingsNavigationView.ItemInvoked += NavigationView_ItemInvoked;
                    System.Diagnostics.Debug.WriteLine("SettingsWindow: NavigationView ItemInvoked event handler connected");
                    
                    // NavigationViewItemにPreviewMouseLeftButtonDownイベントを追加（フォールバック）
                    // これにより、NavigationViewのイベントが発生しない場合でもページを切り替えられる
                    GeneralNavItem.PreviewMouseLeftButtonDown += (s, e) =>
                    {
                        System.Diagnostics.Debug.WriteLine("GeneralNavItem PreviewMouseLeftButtonDown");
                        NavigateToPage(typeof(GeneralSettingsPage));
                        e.Handled = true; // NavigationViewのイベントを阻止
                    };
                    AppearanceNavItem.PreviewMouseLeftButtonDown += (s, e) =>
                    {
                        System.Diagnostics.Debug.WriteLine("AppearanceNavItem PreviewMouseLeftButtonDown");
                        NavigateToPage(typeof(AppearanceSettingsPage));
                        e.Handled = true;
                    };
                    AboutNavItem.PreviewMouseLeftButtonDown += (s, e) =>
                    {
                        System.Diagnostics.Debug.WriteLine("AboutNavItem PreviewMouseLeftButtonDown");
                        NavigateToPage(typeof(AboutSettingsPage));
                        e.Handled = true;
                    };
                    
                    // デフォルトで全般ページを表示
                    NavigateToPage(typeof(GeneralSettingsPage));
                    
                    System.Diagnostics.Debug.WriteLine("SettingsWindow Loaded event completed");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error in SettingsWindow Loaded event: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                    // エラーが発生してもウィンドウは表示する
                }
            };
            
            // ウィンドウを閉じる時にViewModelのクリーンアップ
            // Alt+F4など、TitleBar.CloseClicked以外で閉じられた場合にも設定を保存
            // _isNavigatingFromフラグで重複を防ぐ
            Closing += (s, e) =>
            {
                // 非同期処理はfire-and-forgetで実行（ウィンドウを閉じる処理をブロックしない）
                // _isNavigatingFromフラグで重複を防ぐ
                _ = ViewModel.OnNavigatedFromAsync();
            };

            // ウィンドウが閉じられた後にクリーンアップ
            Closed += (s, e) =>
            {
                // イベントハンドラーを解除してメモリリークを防ぐ
                TitleBar.CloseClicked -= TitleBar_CloseClicked;
            };
        }

        /// <summary>
        /// TitleBarの×ボタンがクリックされたときに呼び出されます
        /// </summary>
        private void TitleBar_CloseClicked(object? sender, RoutedEventArgs e)
        {
            // イベントを処理済みとしてマーク
            e.Handled = true;
            
            // 閉じる処理はClosingイベントで行うため、ここでは単純に閉じる
            // 非同期処理はClosingイベントでfire-and-forgetで実行される
            Close();
        }

        /// <summary>
        /// NavigationViewのアイテムが選択されたときに呼び出されます
        /// </summary>
        /// <param name="sender">イベントの送信元</param>
        /// <param name="args">ナビゲーション選択変更イベント引数</param>
        private void NavigationView_ItemInvoked(object sender, object args)
        {
            // デバッグ: メソッドが呼び出されたことを確認
            System.Diagnostics.Debug.WriteLine("SettingsWindow NavigationView_ItemInvoked called");
            
            // リフレクションを使用してInvokedItemContainerプロパティにアクセス
            var argsType = args.GetType();
            
            // キャッシュされた型と一致しない場合は、プロパティを再取得
            if (!ReferenceEquals(_cachedArgsType, argsType))
            {
                _cachedArgsType = argsType;
                _cachedInvokedItemContainerProperty = argsType.GetProperty("InvokedItemContainer");
            }
            
            if (_cachedInvokedItemContainerProperty == null)
            {
                System.Diagnostics.Debug.WriteLine("SettingsWindow: InvokedItemContainer property is null");
                return;
            }
            
            var invokedItem = _cachedInvokedItemContainerProperty.GetValue(args) as NavigationViewItem;
            System.Diagnostics.Debug.WriteLine($"SettingsWindow: InvokedItem: {invokedItem}, Tag: {invokedItem?.Tag}, Tag type: {invokedItem?.Tag?.GetType().Name}");
            
            if (invokedItem?.Tag is not string tag)
            {
                System.Diagnostics.Debug.WriteLine("SettingsWindow: Tag is not a string or invokedItem is null");
                return;
            }
            
            System.Diagnostics.Debug.WriteLine($"SettingsWindow: Tag value: '{tag}'");

            // タグに応じて適切なページにナビゲート
            switch (tag)
            {
                case "General":
                    System.Diagnostics.Debug.WriteLine("SettingsWindow: Navigating to GeneralSettingsPage");
                    NavigateToPage(typeof(GeneralSettingsPage));
                    break;
                case "Appearance":
                    System.Diagnostics.Debug.WriteLine("SettingsWindow: Navigating to AppearanceSettingsPage");
                    NavigateToPage(typeof(AppearanceSettingsPage));
                    break;
                case "About":
                    System.Diagnostics.Debug.WriteLine("SettingsWindow: Navigating to AboutSettingsPage");
                    NavigateToPage(typeof(AboutSettingsPage));
                    break;
                default:
                    System.Diagnostics.Debug.WriteLine($"SettingsWindow: Unknown tag '{tag}', navigating to GeneralSettingsPage");
                    NavigateToPage(typeof(GeneralSettingsPage));
                    break;
            }
        }

        /// <summary>
        /// 指定されたページタイプにナビゲートします
        /// </summary>
        /// <param name="pageType">ナビゲート先のページタイプ</param>
        private void NavigateToPage(Type pageType)
        {
            System.Diagnostics.Debug.WriteLine($"SettingsWindow: NavigateToPage called with type: {pageType.Name}");
            
            try
            {
                if (pageType == typeof(GeneralSettingsPage))
                {
                    System.Diagnostics.Debug.WriteLine("SettingsWindow: Creating GeneralSettingsPage");
                    SettingsContentFrame.Navigate(new GeneralSettingsPage(ViewModel));
                }
                else if (pageType == typeof(AppearanceSettingsPage))
                {
                    System.Diagnostics.Debug.WriteLine("SettingsWindow: Creating AppearanceSettingsPage");
                    SettingsContentFrame.Navigate(new AppearanceSettingsPage(ViewModel));
                }
                else if (pageType == typeof(AboutSettingsPage))
                {
                    System.Diagnostics.Debug.WriteLine("SettingsWindow: Creating AboutSettingsPage");
                    SettingsContentFrame.Navigate(new AboutSettingsPage(ViewModel));
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"SettingsWindow: Unknown page type: {pageType.Name}");
                }
                
                System.Diagnostics.Debug.WriteLine("SettingsWindow: Navigation completed successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SettingsWindow: Error during navigation: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"SettingsWindow: Stack trace: {ex.StackTrace}");
            }
        }
    }
}
