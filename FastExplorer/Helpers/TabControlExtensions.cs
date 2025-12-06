using System.Windows;
using System.Windows.Controls;
using Wpf.Ui.Controls;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// TabControlの拡張機能を提供するクラス
    /// TabAddingとTabClosingのイベントハンドラーをXAMLで設定するためのラッパー
    /// </summary>
    public static class TabControlExtensions
    {
        #region TabAdding イベント

        /// <summary>
        /// TabAddingイベントの識別子（WPFUIのTabControlExtensionsから取得）
        /// </summary>
        public static RoutedEvent TabAddingEvent => Wpf.Ui.Controls.TabControlExtensions.TabAddingEvent;

        /// <summary>
        /// TabAddingイベントハンドラーを追加します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="handler">イベントハンドラー</param>
        public static void AddTabAddingHandler(DependencyObject obj, TabAddingEventHandler handler)
        {
            Wpf.Ui.Controls.TabControlExtensions.AddTabAddingHandler(obj, handler);
        }

        /// <summary>
        /// TabAddingイベントハンドラーを削除します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="handler">イベントハンドラー</param>
        public static void RemoveTabAddingHandler(DependencyObject obj, TabAddingEventHandler handler)
        {
            Wpf.Ui.Controls.TabControlExtensions.RemoveTabAddingHandler(obj, handler);
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
                    RemoveTabAddingHandler(element, oldHandler);
                }
                if (e.NewValue is TabAddingEventHandler newHandler)
                {
                    AddTabAddingHandler(element, newHandler);
                }
            }
        }

        #endregion

        #region TabClosing イベント

        /// <summary>
        /// TabClosingイベントの識別子（WPFUIのTabControlExtensionsから取得）
        /// </summary>
        public static RoutedEvent TabClosingEvent => Wpf.Ui.Controls.TabControlExtensions.TabClosingEvent;

        /// <summary>
        /// TabClosingイベントハンドラーを追加します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="handler">イベントハンドラー</param>
        public static void AddTabClosingHandler(DependencyObject obj, TabClosingEventHandler handler)
        {
            Wpf.Ui.Controls.TabControlExtensions.AddTabClosingHandler(obj, handler);
        }

        /// <summary>
        /// TabClosingイベントハンドラーを削除します
        /// </summary>
        /// <param name="obj">対象のオブジェクト</param>
        /// <param name="handler">イベントハンドラー</param>
        public static void RemoveTabClosingHandler(DependencyObject obj, TabClosingEventHandler handler)
        {
            Wpf.Ui.Controls.TabControlExtensions.RemoveTabClosingHandler(obj, handler);
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
                    RemoveTabClosingHandler(element, oldHandler);
                }
                if (e.NewValue is TabClosingEventHandler newHandler)
                {
                    AddTabClosingHandler(element, newHandler);
                }
            }
        }

        #endregion
    }
}
