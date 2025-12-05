using System.Windows;
using System.Windows.Controls;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// TabControlの拡張機能を提供するクラス
    /// </summary>
    public static class TabControlExtensions
    {
        #region CanAddTabs アタッチ可能プロパティ

        /// <summary>
        /// CanAddTabs添付プロパティの識別子
        /// </summary>
        public static readonly DependencyProperty CanAddTabsProperty =
            DependencyProperty.RegisterAttached(
                "CanAddTabs",
                typeof(bool),
                typeof(TabControlExtensions),
                new PropertyMetadata(false, OnCanAddTabsChanged));

        /// <summary>
        /// CanAddTabs添付プロパティの値を取得します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <returns>CanAddTabsの値</returns>
        public static bool GetCanAddTabs(DependencyObject obj)
        {
            return (bool)obj.GetValue(CanAddTabsProperty);
        }

        /// <summary>
        /// CanAddTabs添付プロパティの値を設定します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="value">設定する値</param>
        public static void SetCanAddTabs(DependencyObject obj, bool value)
        {
            obj.SetValue(CanAddTabsProperty, value);
        }

        private static void OnCanAddTabsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            // 必要に応じてイベントハンドラーを設定
        }

        #endregion

        #region CanReorderTabs アタッチ可能プロパティ

        /// <summary>
        /// CanReorderTabs添付プロパティの識別子
        /// </summary>
        public static readonly DependencyProperty CanReorderTabsProperty =
            DependencyProperty.RegisterAttached(
                "CanReorderTabs",
                typeof(bool),
                typeof(TabControlExtensions),
                new PropertyMetadata(false));

        /// <summary>
        /// CanReorderTabs添付プロパティの値を取得します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <returns>CanReorderTabsの値</returns>
        public static bool GetCanReorderTabs(DependencyObject obj)
        {
            return (bool)obj.GetValue(CanReorderTabsProperty);
        }

        /// <summary>
        /// CanReorderTabs添付プロパティの値を設定します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="value">設定する値</param>
        public static void SetCanReorderTabs(DependencyObject obj, bool value)
        {
            obj.SetValue(CanReorderTabsProperty, value);
        }

        #endregion

        #region TabAdding イベント

        /// <summary>
        /// TabAddingイベントの識別子
        /// </summary>
        public static readonly RoutedEvent TabAddingEvent =
            EventManager.RegisterRoutedEvent(
                "TabAdding",
                RoutingStrategy.Bubble,
                typeof(TabAddingEventHandler),
                typeof(TabControlExtensions));

        /// <summary>
        /// TabAddingイベントハンドラーを追加します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="handler">イベントハンドラー</param>
        public static void AddTabAddingHandler(DependencyObject obj, TabAddingEventHandler handler)
        {
            if (obj is UIElement element)
            {
                element.AddHandler(TabAddingEvent, handler);
            }
        }

        /// <summary>
        /// TabAddingイベントハンドラーを削除します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="handler">イベントハンドラー</param>
        public static void RemoveTabAddingHandler(DependencyObject obj, TabAddingEventHandler handler)
        {
            if (obj is UIElement element)
            {
                element.RemoveHandler(TabAddingEvent, handler);
            }
        }

        /// <summary>
        /// TabAddingイベントハンドラーを設定するための添付プロパティ（XAMLで使用）
        /// </summary>
        public static readonly DependencyProperty TabAddingProperty =
            DependencyProperty.RegisterAttached(
                "TabAdding",
                typeof(TabAddingEventHandler),
                typeof(TabControlExtensions),
                new PropertyMetadata(null, OnTabAddingChanged));

        /// <summary>
        /// TabAddingイベントハンドラーを取得します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <returns>イベントハンドラー</returns>
        public static TabAddingEventHandler? GetTabAdding(DependencyObject obj)
        {
            return (TabAddingEventHandler?)obj.GetValue(TabAddingProperty);
        }

        /// <summary>
        /// TabAddingイベントハンドラーを設定します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="value">イベントハンドラー</param>
        public static void SetTabAdding(DependencyObject obj, TabAddingEventHandler? value)
        {
            obj.SetValue(TabAddingProperty, value);
        }

        private static void OnTabAddingChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is UIElement element)
            {
                if (e.OldValue is TabAddingEventHandler oldHandler)
                {
                    element.RemoveHandler(TabAddingEvent, oldHandler);
                }
                if (e.NewValue is TabAddingEventHandler newHandler)
                {
                    element.AddHandler(TabAddingEvent, newHandler);
                }
            }
        }

        #endregion

        #region TabClosing イベント

        /// <summary>
        /// TabClosingイベントの識別子
        /// </summary>
        public static readonly RoutedEvent TabClosingEvent =
            EventManager.RegisterRoutedEvent(
                "TabClosing",
                RoutingStrategy.Bubble,
                typeof(TabClosingEventHandler),
                typeof(TabControlExtensions));

        /// <summary>
        /// TabClosingイベントハンドラーを追加します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="handler">イベントハンドラー</param>
        public static void AddTabClosingHandler(DependencyObject obj, TabClosingEventHandler handler)
        {
            if (obj is UIElement element)
            {
                element.AddHandler(TabClosingEvent, handler);
            }
        }

        /// <summary>
        /// TabClosingイベントハンドラーを削除します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="handler">イベントハンドラー</param>
        public static void RemoveTabClosingHandler(DependencyObject obj, TabClosingEventHandler handler)
        {
            if (obj is UIElement element)
            {
                element.RemoveHandler(TabClosingEvent, handler);
            }
        }

        /// <summary>
        /// TabClosingイベントハンドラーを設定するための添付プロパティ（XAMLで使用）
        /// </summary>
        public static readonly DependencyProperty TabClosingProperty =
            DependencyProperty.RegisterAttached(
                "TabClosing",
                typeof(TabClosingEventHandler),
                typeof(TabControlExtensions),
                new PropertyMetadata(null, OnTabClosingChanged));

        /// <summary>
        /// TabClosingイベントハンドラーを取得します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <returns>イベントハンドラー</returns>
        public static TabClosingEventHandler? GetTabClosing(DependencyObject obj)
        {
            return (TabClosingEventHandler?)obj.GetValue(TabClosingProperty);
        }

        /// <summary>
        /// TabClosingイベントハンドラーを設定します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="value">イベントハンドラー</param>
        public static void SetTabClosing(DependencyObject obj, TabClosingEventHandler? value)
        {
            obj.SetValue(TabClosingProperty, value);
        }

        private static void OnTabClosingChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is UIElement element)
            {
                if (e.OldValue is TabClosingEventHandler oldHandler)
                {
                    element.RemoveHandler(TabClosingEvent, oldHandler);
                }
                if (e.NewValue is TabClosingEventHandler newHandler)
                {
                    element.AddHandler(TabClosingEvent, newHandler);
                }
            }
        }

        #endregion

        #region IsClosable アタッチ可能プロパティ（TabItem用）

        /// <summary>
        /// IsClosable添付プロパティの識別子
        /// </summary>
        public static readonly DependencyProperty IsClosableProperty =
            DependencyProperty.RegisterAttached(
                "IsClosable",
                typeof(bool),
                typeof(TabControlExtensions),
                new PropertyMetadata(false));

        /// <summary>
        /// IsClosable添付プロパティの値を取得します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <returns>IsClosableの値</returns>
        public static bool GetIsClosable(DependencyObject obj)
        {
            return (bool)obj.GetValue(IsClosableProperty);
        }

        /// <summary>
        /// IsClosable添付プロパティの値を設定します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="value">設定する値</param>
        public static void SetIsClosable(DependencyObject obj, bool value)
        {
            obj.SetValue(IsClosableProperty, value);
        }

        #endregion
    }

    /// <summary>
    /// TabAddingイベントのイベントハンドラーデリゲート
    /// </summary>
    /// <param name="sender">イベントの送信元</param>
    /// <param name="e">イベント引数</param>
    public delegate void TabAddingEventHandler(object sender, TabAddingEventArgs e);

    /// <summary>
    /// TabClosingイベントのイベントハンドラーデリゲート
    /// </summary>
    /// <param name="sender">イベントの送信元</param>
    /// <param name="e">イベント引数</param>
    public delegate void TabClosingEventHandler(object sender, TabClosingEventArgs e);

    /// <summary>
    /// TabAddingイベントのイベント引数
    /// </summary>
    public class TabAddingEventArgs : RoutedEventArgs
    {
        /// <summary>
        /// イベントをキャンセルするかどうかを取得または設定します
        /// </summary>
        public bool Cancel { get; set; }

        /// <summary>
        /// <see cref="TabAddingEventArgs"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        public TabAddingEventArgs()
            : base(TabControlExtensions.TabAddingEvent)
        {
        }

        /// <summary>
        /// <see cref="TabAddingEventArgs"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="routedEvent">ルーティングイベント</param>
        public TabAddingEventArgs(RoutedEvent routedEvent)
            : base(routedEvent)
        {
        }

        /// <summary>
        /// <see cref="TabAddingEventArgs"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="routedEvent">ルーティングイベント</param>
        /// <param name="source">イベントソース</param>
        public TabAddingEventArgs(RoutedEvent routedEvent, object source)
            : base(routedEvent, source)
        {
        }
    }

    /// <summary>
    /// TabClosingイベントのイベント引数
    /// </summary>
    public class TabClosingEventArgs : RoutedEventArgs
    {
        /// <summary>
        /// 閉じられるTabItemを取得または設定します
        /// </summary>
        public TabItem? TabItem { get; set; }

        /// <summary>
        /// イベントをキャンセルするかどうかを取得または設定します
        /// </summary>
        public bool Cancel { get; set; }

        /// <summary>
        /// <see cref="TabClosingEventArgs"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        public TabClosingEventArgs()
            : base(TabControlExtensions.TabClosingEvent)
        {
        }

        /// <summary>
        /// <see cref="TabClosingEventArgs"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="routedEvent">ルーティングイベント</param>
        public TabClosingEventArgs(RoutedEvent routedEvent)
            : base(routedEvent)
        {
        }

        /// <summary>
        /// <see cref="TabClosingEventArgs"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="routedEvent">ルーティングイベント</param>
        /// <param name="source">イベントソース</param>
        public TabClosingEventArgs(RoutedEvent routedEvent, object source)
            : base(routedEvent, source)
        {
        }
    }
}
