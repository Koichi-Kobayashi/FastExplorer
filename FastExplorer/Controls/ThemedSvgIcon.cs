using System.ComponentModel;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using SharpVectors.Converters;
using SharpVectors.Renderers.Wpf;
using SharpVectors.Runtime;

namespace FastExplorer.Controls
{
    /// <summary>
    /// テーマに応じて色を変更できるSVGアイコンコントロール
    /// </summary>
    public class ThemedSvgIcon : SvgViewbox
    {
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

        /// <summary>
        /// <see cref="ThemedSvgIcon"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        public ThemedSvgIcon()
        {
            Loaded += (_, __) => 
            {
                _brushApplied = false;
                ApplyIconBrushDelayed();
            };
            
            LayoutUpdated += (_, __) =>
            {
                if (!_brushApplied && IconBrush != null)
                {
                    ApplyIconBrushDelayed();
                }
            };
            
            // Sourceプロパティの変更を監視
            var sourceDescriptor = DependencyPropertyDescriptor.FromProperty(SourceProperty, typeof(SvgViewbox));
            sourceDescriptor?.AddValueChanged(this, (s, e) => 
            {
                _brushApplied = false;
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
            if (IconBrush == null)
                return;

            // ChildからDrawingを取得
            if (Child is SvgDrawingCanvas canvas)
            {
                bool applied = false;

                // DrawObjectsプロパティにアクセス
                var drawObjects = canvas.DrawObjects;
                if (drawObjects != null && drawObjects.Count > 0)
                {
                    foreach (var drawing in drawObjects)
                    {
                        ApplyBrushRecursive(drawing, IconBrush);
                        applied = true;
                    }
                }

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
                                ApplyBrushRecursive(drawing, IconBrush);
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
                                ApplyBrushRecursive(drawing, IconBrush);
                                applied = true;
                            }
                        }
                        catch
                        {
                            // フィールドにアクセスできない場合は無視
                        }
                    }
                }

                // VisualTreeHelperで子要素を探索
                for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(canvas); i++)
                {
                    var child = System.Windows.Media.VisualTreeHelper.GetChild(canvas, i);
                    if (child is DrawingVisual drawingVisual)
                    {
                        var drawing = drawingVisual.Drawing;
                        if (drawing != null)
                        {
                            ApplyBrushRecursive(drawing, IconBrush);
                            applied = true;
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
    }
}

