using System.ComponentModel;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using SharpVectors.Converters;
using SharpVectors.Renderers.Wpf;
using SharpVectors.Runtime;
using Wpf.Ui.Appearance;

namespace FastExplorer.Controls
{
    /// <summary>
    /// テーマに応じて色を変更できるSVGアイコンコントロール
    /// </summary>
    public class ThemedSvgIcon : SvgViewbox
    {
        private static readonly System.Collections.Generic.List<ThemedSvgIcon> _instances = new();
        /// <summary>
        /// アイコンのブラシを取得または設定します
        /// </summary>
        public static readonly DependencyProperty IconBrushProperty =
            DependencyProperty.Register(
                nameof(IconBrush),
                typeof(Brush),
                typeof(ThemedSvgIcon),
                new PropertyMetadata(null, OnIconBrushChanged));

        /// <summary>
        /// アイコンのブラシを取得または設定します
        /// </summary>
        public Brush IconBrush
        {
            get => (Brush)GetValue(IconBrushProperty);
            set => SetValue(IconBrushProperty, value);
        }

        private bool _brushApplied = false;
        private ApplicationTheme _lastTheme = ApplicationTheme.Light;
        private Color? _lastBrushColor = null;

        /// <summary>
        /// <see cref="ThemedSvgIcon"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        public ThemedSvgIcon()
        {
            _instances.Add(this);
            _lastTheme = ApplicationThemeManager.GetAppTheme();
            
            Loaded += (_, __) => 
            {
                _brushApplied = false;
                _lastBrushColor = null;
                ApplyIconBrushDelayed();
            };
            
            Unloaded += (_, __) =>
            {
                _instances.Remove(this);
            };
            
            LayoutUpdated += (_, __) =>
            {
                // テーマが変更された場合は再適用
                var currentTheme = ApplicationThemeManager.GetAppTheme();
                
                // IconBrushプロパティから色を取得（DynamicResourceが解決された値を使用）
                Color? currentBrushColor = null;
                if (IconBrush is SolidColorBrush scb)
                {
                    currentBrushColor = scb.Color;
                }
                else
                {
                    // IconBrushプロパティがnullの場合は、リソースから直接取得を試みる
                    try
                    {
                        var resource = FindResource("IconBrush");
                        if (resource is SolidColorBrush solidBrush)
                        {
                            currentBrushColor = solidBrush.Color;
                        }
                    }
                    catch
                    {
                        // リソースが見つからない場合は無視
                    }
                }
                
                if (currentTheme != _lastTheme || currentBrushColor != _lastBrushColor)
                {
                    _lastTheme = currentTheme;
                    _lastBrushColor = currentBrushColor;
                    _brushApplied = false;
                    ApplyIconBrushDelayed();
                }
                else if (!_brushApplied && currentBrushColor != null)
                {
                    ApplyIconBrushDelayed();
                }
            };
            
            // Sourceプロパティの変更を監視
            var sourceDescriptor = DependencyPropertyDescriptor.FromProperty(SourceProperty, typeof(SvgViewbox));
            sourceDescriptor?.AddValueChanged(this, (s, e) => 
            {
                _brushApplied = false;
                _lastBrushColor = null;
                ApplyIconBrushDelayed();
            });
        }

        /// <summary>
        /// IconBrushプロパティが変更されたときに呼び出されます
        /// </summary>
        private static void OnIconBrushChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ThemedSvgIcon icon)
            {
                icon._brushApplied = false; // 強制的に再適用
                if (e.NewValue is SolidColorBrush scb)
                {
                    icon._lastBrushColor = scb.Color;
                }
                else
                {
                    icon._lastBrushColor = null;
                }
                icon.ApplyIconBrushDelayed();
            }
        }

        /// <summary>
        /// アイコンブラシを遅延適用します
        /// </summary>
        private void ApplyIconBrushDelayed()
        {
            if (IconBrush == null)
                return;

            // 複数回試行してSVGが読み込まれるのを待つ
            Dispatcher.BeginInvoke(new System.Action(() =>
            {
                ApplyIconBrush();
            }), System.Windows.Threading.DispatcherPriority.Loaded);
            
            // さらに遅延して再試行
            Dispatcher.BeginInvoke(new System.Action(() =>
            {
                ApplyIconBrush();
            }), System.Windows.Threading.DispatcherPriority.Render);
            
            // さらに遅延して再試行（SVGの読み込みが遅い場合に対応）
            Dispatcher.BeginInvoke(new System.Action(() =>
            {
                ApplyIconBrush();
            }), System.Windows.Threading.DispatcherPriority.Background);
            
            // さらに遅延して再試行（確実に適用するため）
            Dispatcher.BeginInvoke(new System.Action(() =>
            {
                ApplyIconBrush();
            }), System.Windows.Threading.DispatcherPriority.SystemIdle);
        }

        /// <summary>
        /// 描画時に呼び出されます
        /// </summary>
        protected override void OnRender(DrawingContext drawingContext)
        {
            base.OnRender(drawingContext);
            
            // 描画後にブラシを適用
            if (!_brushApplied && IconBrush != null)
            {
                Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    ApplyIconBrush();
                }), System.Windows.Threading.DispatcherPriority.Render);
            }
        }

        /// <summary>
        /// アイコンブラシを適用します
        /// </summary>
        private void ApplyIconBrush()
        {
            // IconBrushプロパティから取得（DynamicResourceが解決された値を使用）
            Brush? brushToApply = IconBrush;
            
            // IconBrushプロパティがnullの場合は、リソースから直接取得を試みる
            if (brushToApply == null)
            {
                try
                {
                    var resource = FindResource("IconBrush");
                    if (resource is Brush resourceBrush)
                    {
                        brushToApply = resourceBrush;
                    }
                }
                catch
                {
                    // リソースが見つからない場合は終了
                    return;
                }
            }
            
            if (brushToApply == null)
                return;

            // ChildからDrawingを取得
            if (Child is SvgDrawingCanvas canvas)
            {
                bool applied = false;

                // DrawObjectsプロパティにアクセス（最も確実な方法）
                var drawObjects = canvas.DrawObjects;
                if (drawObjects != null && drawObjects.Count > 0)
                {
                    foreach (var drawing in drawObjects)
                    {
                        ApplyBrushRecursive(drawing, brushToApply);
                        applied = true;
                    }
                }

                // DrawObjectsが空の場合、または適用できなかった場合は、他の方法を試す
                if (!applied)
                {
                    // すべてのプロパティをリフレクションで探索
                    var properties = typeof(SvgDrawingCanvas).GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    foreach (var prop in properties)
                    {
                        if (prop.PropertyType == typeof(Drawing) || prop.PropertyType.IsSubclassOf(typeof(Drawing)))
                        {
                            try
                            {
                                var value = prop.GetValue(canvas);
                                if (value is Drawing drawing)
                                {
                                    ApplyBrushRecursive(drawing, brushToApply);
                                    applied = true;
                                }
                            }
                            catch
                            {
                                // プロパティにアクセスできない場合は無視
                            }
                        }
                    }

                    // フィールドも探索
                    if (!applied)
                    {
                        var fields = typeof(SvgDrawingCanvas).GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                        foreach (var field in fields)
                        {
                            if (field.FieldType == typeof(Drawing) || field.FieldType.IsSubclassOf(typeof(Drawing)))
                            {
                                try
                                {
                                    var value = field.GetValue(canvas);
                                    if (value is Drawing drawing)
                                    {
                                        ApplyBrushRecursive(drawing, brushToApply);
                                        applied = true;
                                    }
                                }
                                catch
                                {
                                    // フィールドにアクセスできない場合は無視
                                }
                            }
                        }
                    }

                    // VisualTreeHelperで子要素を探索
                    if (!applied)
                    {
                        for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(canvas); i++)
                        {
                            var child = System.Windows.Media.VisualTreeHelper.GetChild(canvas, i);
                            if (child is DrawingVisual drawingVisual)
                            {
                                var drawing = drawingVisual.Drawing;
                                if (drawing != null)
                                {
                                    ApplyBrushRecursive(drawing, brushToApply);
                                    applied = true;
                                }
                            }
                        }
                    }
                }

                if (applied)
                {
                    _brushApplied = true;
                    // 再描画を強制
                    canvas.InvalidateVisual();
                    InvalidateVisual();
                }
            }
        }

        /// <summary>
        /// 再帰的にDrawingにブラシを適用します
        /// </summary>
        /// <param name="drawing">適用するDrawing</param>
        /// <param name="brush">適用するブラシ</param>
        private static void ApplyBrushRecursive(Drawing drawing, Brush brush)
        {
            if (drawing == null)
                return;

            if (drawing is DrawingGroup group)
            {
                foreach (var child in group.Children)
                {
                    ApplyBrushRecursive(child, brush);
                }
            }
            else if (drawing is GeometryDrawing geo)
            {
                // 塗りつぶし（Brushがnullでも適用）
                geo.Brush = brush;

                // 線（Penが存在する場合のみ）
                if (geo.Pen != null)
                {
                    geo.Pen.Brush = brush;
                }
            }
            else if (drawing is GlyphRunDrawing glyph)
            {
                // GlyphRunDrawingの場合もForegroundBrushを設定
                if (glyph.ForegroundBrush != null)
                {
                    glyph.ForegroundBrush = brush;
                }
            }
        }

        /// <summary>
        /// すべてのインスタンスにブラシを再適用します
        /// </summary>
        public static void RefreshAllInstances()
        {
            foreach (var instance in _instances.ToArray())
            {
                if (instance != null)
                {
                    instance._brushApplied = false;
                    instance._lastBrushColor = null;
                    instance._lastTheme = ApplicationTheme.Light; // 強制的に再適用させる
                    instance.ApplyIconBrushDelayed();
                }
            }
        }
    }
}

