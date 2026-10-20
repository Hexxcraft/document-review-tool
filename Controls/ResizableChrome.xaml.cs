using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using HexxcorpDocumentReviewUtility.Models;

namespace HexxcorpDocumentReviewUtility.Controls
{
    /// <summary>
    /// Provides the resizable/movable "chrome" (drag header, 8-point resize
    /// handles, remove button) wrapped around a workspace canvas item's
    /// content (an image or a <see cref="PdfBoxControl"/>). The control's
    /// <see cref="DataContext"/> is expected to be a <see cref="CanvasItemModel"/>;
    /// position/size changes are written directly to that model so the
    /// Canvas.Left/Top/Width/Height bindings on the host container update live.
    /// </summary>
    public partial class ResizableChrome : UserControl
    {
        private const double MinItemWidth = 120;
        private const double MinItemHeight = 90;
        private const double MinZoom = 0.25;
        private const double MaxZoom = 4.0;
        private const double ZoomStep = 0.1;

        /// <summary>Matches the fixed Canvas size declared in WorkspaceWindow.xaml's ItemsHost ItemsPanelTemplate. Items are kept fully within these bounds so they can't be dragged/resized out of the workspace's valid area.</summary>
        private const double CanvasWidth = 2600;
        private const double CanvasHeight = 1700;

        private static int _zIndexCounter;

        public static readonly DependencyProperty InnerContentProperty = DependencyProperty.Register(
            nameof(InnerContent), typeof(object), typeof(ResizableChrome), new PropertyMetadata(null, OnInnerContentChanged));

        public static readonly DependencyProperty FileNameProperty = DependencyProperty.Register(
            nameof(FileName), typeof(string), typeof(ResizableChrome));

        public static readonly DependencyProperty ZoomProperty = DependencyProperty.Register(
            nameof(Zoom), typeof(double), typeof(ResizableChrome),
            new FrameworkPropertyMetadata(1.0, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnZoomChanged));

        public object? InnerContent
        {
            get => GetValue(InnerContentProperty);
            set => SetValue(InnerContentProperty, value);
        }

        public string? FileName
        {
            get => (string?)GetValue(FileNameProperty);
            set => SetValue(FileNameProperty, value);
        }

        /// <summary>Content magnification level (1.0 = 100%), clamped to [<see cref="MinZoom"/>, <see cref="MaxZoom"/>].</summary>
        public double Zoom
        {
            get => (double)GetValue(ZoomProperty);
            set => SetValue(ZoomProperty, Math.Clamp(value, MinZoom, MaxZoom));
        }

        /// <summary>Raised when the user clicks the remove ("x") button on this item.</summary>
        public event EventHandler<CanvasItemModel>? RemoveRequested;

        private CanvasItemModel? Model => DataContext as CanvasItemModel;

        public ResizableChrome()
        {
            InitializeComponent();
            PreviewMouseLeftButtonDown += (_, _) => Panel.SetZIndex(this, ++_zIndexCounter);
            ContentHost.SizeChanged += ContentHost_SizeChanged;
        }

        /// <summary>
        /// Applies the zoom level. PDF content (<see cref="PdfBoxControl"/>) manages its own internal
        /// page zoom/scroll, so for that content type the value is simply forwarded to
        /// <see cref="PdfBoxControl.Zoom"/> and this control's own outer scale is left at 1.0. For any
        /// other content (e.g. a plain image) the scale is applied here via the <see cref="ZoomHost"/>
        /// LayoutTransform, which - unlike a RenderTransform - changes the reported layout size so the
        /// wrapping <see cref="ContentScrollViewer"/> grows/shrinks its scrollable extent and keeps
        /// panning available at any zoom level.
        /// </summary>
        /// <summary>
        /// Re-applies the current zoom whenever the inner content is (re)assigned: XAML property-element
        /// order means <see cref="InnerContent"/> can be set after <see cref="Zoom"/> during object
        /// construction, so the very first zoom application might not yet know whether to route to a
        /// <see cref="PdfBoxControl"/>'s own internal zoom or this control's outer scale.
        /// </summary>
        private static void OnInnerContentChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ResizableChrome chrome)
            {
                OnZoomChanged(chrome, new DependencyPropertyChangedEventArgs(ZoomProperty, chrome.Zoom, chrome.Zoom));
            }
        }

        private static void OnZoomChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not ResizableChrome chrome) return;

            double zoom = (double)e.NewValue;
            chrome.ZoomPercentText.Text = $"{zoom * 100:0}%";

            if (chrome.InnerContent is PdfBoxControl pdfBox)
            {
                chrome.ContentScale.ScaleX = 1.0;
                chrome.ContentScale.ScaleY = 1.0;
                pdfBox.Zoom = zoom;
            }
            else
            {
                chrome.ContentScale.ScaleX = zoom;
                chrome.ContentScale.ScaleY = zoom;
            }
        }

        /// <summary>
        /// Keeps the (unscaled, "100%") zoom host size in sync with the content host's actual box size so
        /// the LayoutTransform-based zoom scales relative to the current box dimensions rather than a
        /// stale size from before a resize.
        /// </summary>
        private void ContentHost_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            ZoomHost.Width = e.NewSize.Width;
            ZoomHost.Height = e.NewSize.Height;
        }

        /// <summary>Ctrl+mouse wheel zooms the content; a plain wheel is left unhandled so it still scrolls PDF pages / the canvas as normal.</summary>
        private void ContentHost_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.Control)
            {
                return;
            }

            Zoom += e.Delta > 0 ? ZoomStep : -ZoomStep;
            e.Handled = true;
        }

        private void ZoomInButton_Click(object sender, RoutedEventArgs e) => Zoom += ZoomStep;

        private void ZoomOutButton_Click(object sender, RoutedEventArgs e) => Zoom -= ZoomStep;

        private void ZoomResetButton_Click(object sender, RoutedEventArgs e) => Zoom = 1.0;

        private void ZoomPercentText_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e) => Zoom = 1.0;

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            if (Model is { } model)
            {
                RemoveRequested?.Invoke(this, model);
            }
        }

        private void MoveThumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (Model is not { } model) return;
            model.X = Math.Clamp(model.X + e.HorizontalChange, 0, Math.Max(0, CanvasWidth - model.Width));
            model.Y = Math.Clamp(model.Y + e.VerticalChange, 0, Math.Max(0, CanvasHeight - model.Height));
        }

        // ---- Corner/edge resize handlers -------------------------------------------------

        private void ResizeTopLeft_DragDelta(object sender, DragDeltaEventArgs e)
        {
            ResizeLeft(e.HorizontalChange);
            ResizeTop(e.VerticalChange);
        }

        private void ResizeTopRight_DragDelta(object sender, DragDeltaEventArgs e)
        {
            ResizeRight(e.HorizontalChange);
            ResizeTop(e.VerticalChange);
        }

        private void ResizeBottomLeft_DragDelta(object sender, DragDeltaEventArgs e)
        {
            ResizeLeft(e.HorizontalChange);
            ResizeBottom(e.VerticalChange);
        }

        private void ResizeBottomRight_DragDelta(object sender, DragDeltaEventArgs e)
        {
            ResizeRight(e.HorizontalChange);
            ResizeBottom(e.VerticalChange);
        }

        private void ResizeTop_DragDelta(object sender, DragDeltaEventArgs e) => ResizeTop(e.VerticalChange);
        private void ResizeBottom_DragDelta(object sender, DragDeltaEventArgs e) => ResizeBottom(e.VerticalChange);
        private void ResizeLeft_DragDelta(object sender, DragDeltaEventArgs e) => ResizeLeft(e.HorizontalChange);
        private void ResizeRight_DragDelta(object sender, DragDeltaEventArgs e) => ResizeRight(e.HorizontalChange);

        private void ResizeLeft(double horizontalChange)
        {
            if (Model is not { } model) return;
            double newWidth = Math.Max(MinItemWidth, model.Width - horizontalChange);
            double newX = model.X + model.Width - newWidth;
            if (newX < 0)
            {
                // Don't let the left edge cross the canvas's left border; shrink the
                // requested width instead so the item "scoots back" within bounds.
                newWidth += newX;
                newX = 0;
            }

            model.X = newX;
            model.Width = Math.Max(MinItemWidth, newWidth);
        }

        private void ResizeRight(double horizontalChange)
        {
            if (Model is not { } model) return;
            double maxWidth = Math.Max(MinItemWidth, CanvasWidth - model.X);
            model.Width = Math.Clamp(model.Width + horizontalChange, MinItemWidth, maxWidth);
        }

        private void ResizeTop(double verticalChange)
        {
            if (Model is not { } model) return;
            double newHeight = Math.Max(MinItemHeight, model.Height - verticalChange);
            double newY = model.Y + model.Height - newHeight;
            if (newY < 0)
            {
                newHeight += newY;
                newY = 0;
            }

            model.Y = newY;
            model.Height = Math.Max(MinItemHeight, newHeight);
        }

        private void ResizeBottom(double verticalChange)
        {
            if (Model is not { } model) return;
            double maxHeight = Math.Max(MinItemHeight, CanvasHeight - model.Y);
            model.Height = Math.Clamp(model.Height + verticalChange, MinItemHeight, maxHeight);
        }
    }
}
