using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PDFtoImage;
using SkiaSharp;

namespace HexxcorpDocumentReviewUtility.Controls
{
    /// <summary>
    /// Renders every page of a PDF file into a vertically stacked, scrollable
    /// viewer. Supports smooth scrolling via the mouse wheel, PageUp/PageDown,
    /// and the Up/Down arrow keys.
    /// </summary>
    public partial class PdfBoxControl : UserControl
    {
        /// <summary>Pixel step used for a single arrow-key or mouse-wheel scroll tick.</summary>
        private const double LineScrollAmount = 48;

        /// <summary>Render width (px) used for each page bitmap; height follows the page aspect ratio.</summary>
        private const int RenderWidth = 1100;

        private const double MinZoom = 0.25;
        private const double MaxZoom = 4.0;

        /// <summary>
        /// Magnification applied to the rendered pages via a <see cref="System.Windows.Media.ScaleTransform"/>
        /// LayoutTransform (not a RenderTransform) so the internal <see cref="PdfScrollViewer"/> correctly
        /// recomputes its scrollable extent at any zoom level: zooming in grows the pages so more of the
        /// viewer can be panned into view via scroll/PageUp/PageDown/arrow keys, and zooming out shrinks the
        /// pages so more of the document fits in the same on-screen box at once.
        /// </summary>
        public static readonly DependencyProperty ZoomProperty = DependencyProperty.Register(
            nameof(Zoom), typeof(double), typeof(PdfBoxControl),
            new FrameworkPropertyMetadata(1.0, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnZoomChanged));

        public double Zoom
        {
            get => (double)GetValue(ZoomProperty);
            set => SetValue(ZoomProperty, Math.Clamp(value, MinZoom, MaxZoom));
        }

        private static void OnZoomChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is PdfBoxControl control)
            {
                control.PagesScale.ScaleX = (double)e.NewValue;
                control.PagesScale.ScaleY = (double)e.NewValue;
            }
        }

        public static readonly DependencyProperty FilePathProperty = DependencyProperty.Register(
            nameof(FilePath),
            typeof(string),
            typeof(PdfBoxControl),
            new PropertyMetadata(null, OnFilePathChanged));

        // Registered as 'object' (boxing the byte[]) rather than 'byte[]'
        // directly: WPF's XAML compiler cannot bind an array-typed dependency
        // property from inside a DataTemplate (fails with MC4102 at compile
        // time), so 'object' is used here purely to work around that limitation.
        public static readonly DependencyProperty SourceBytesProperty = DependencyProperty.Register(
            nameof(SourceBytes),
            typeof(object),
            typeof(PdfBoxControl),
            new PropertyMetadata(null, OnSourceBytesChanged));

        /// <summary>
        /// Fallback-only source: a path on disk to read the PDF from if
        /// <see cref="SourceBytes"/> is not set. Not used once the workspace
        /// embeds file content, but kept for any direct/standalone usage.
        /// </summary>
        public string? FilePath
        {
            get => (string?)GetValue(FilePathProperty);
            set => SetValue(FilePathProperty, value);
        }

        /// <summary>
        /// The PDF's raw bytes, embedded in the workspace project file. When
        /// set, this is always preferred over <see cref="FilePath"/> so the
        /// control never needs to touch the disk to render a loaded workspace.
        /// Declared as <see cref="object"/> (not <c>byte[]</c>) because WPF's
        /// XAML compiler cannot bind an array-typed CLR property from inside
        /// a DataTemplate.
        /// </summary>
        public object? SourceBytes
        {
            get => GetValue(SourceBytesProperty);
            set => SetValue(SourceBytesProperty, value);
        }

        public PdfBoxControl()
        {
            InitializeComponent();
        }

        private static void OnFilePathChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is PdfBoxControl control && control.SourceBytes is not byte[] { Length: > 0 } && e.NewValue is string path && !string.IsNullOrWhiteSpace(path))
            {
                _ = control.LoadPdfAsync();
            }
        }

        private static void OnSourceBytesChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is PdfBoxControl control && e.NewValue is byte[] { Length: > 0 })
            {
                _ = control.LoadPdfAsync();
            }
        }

        private async Task LoadPdfAsync()
        {
            PagesPanel.Children.Clear();
            StatusText.Visibility = Visibility.Visible;
            StatusText.Text = "Loading PDF...";

            string displayName = string.IsNullOrWhiteSpace(FilePath) ? "PDF" : Path.GetFileName(FilePath);

            try
            {
                // Prefer embedded bytes (always present for items added or
                // loaded by this version of the app) and only fall back to
                // reading the original path from disk for project files
                // saved before embedding was introduced.
                byte[] bytes = SourceBytes is byte[] { Length: > 0 } embedded
                    ? embedded
                    : await Task.Run(() => File.ReadAllBytes(FilePath!));

                var images = await Task.Run(() => RenderAllPages(bytes));

                PagesPanel.Children.Clear();
                foreach (var bitmapSource in images)
                {
                    var image = new Image
                    {
                        Source = bitmapSource,
                        Stretch = Stretch.Uniform,
                        Margin = new Thickness(0, 0, 0, 4),
                        SnapsToDevicePixels = true
                    };
                    PagesPanel.Children.Add(image);
                }

                StatusText.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException or ArgumentException)
            {
                StatusText.Visibility = Visibility.Visible;
                StatusText.Text = $"Unable to load PDF:\n{displayName}\n({ex.Message})";
            }
        }

        /// <summary>
        /// Renders every page of the given PDF to a frozen, UI-thread-safe
        /// <see cref="BitmapImage"/>. Executed on a background thread.
        /// </summary>
        private static System.Collections.Generic.List<BitmapImage> RenderAllPages(byte[] pdfBytes)
        {
            var result = new System.Collections.Generic.List<BitmapImage>();

            using var stream = new MemoryStream(pdfBytes);
            var options = new PDFtoImage.RenderOptions(Width: RenderWidth, WithAspectRatio: true);

            foreach (SKBitmap page in Conversion.ToImages(stream, leaveOpen: true, password: null, options: options))
            {
                using (page)
                {
                    using SKData data = page.Encode(SKEncodedImageFormat.Png, 90);
                    using var memoryStream = new MemoryStream(data.ToArray());

                    var bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.StreamSource = memoryStream;
                    bitmapImage.EndInit();
                    bitmapImage.Freeze();

                    result.Add(bitmapImage);
                }
            }

            return result;
        }

        private void PdfScrollViewer_MouseEnter(object sender, MouseEventArgs e)
        {
            // Ensure the viewer has keyboard focus so Up/Down/PageUp/PageDown
            // scroll this PDF frame rather than bubbling elsewhere.
            PdfScrollViewer.Focus();
        }

        private void PdfScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            double offset = PdfScrollViewer.VerticalOffset - (e.Delta / 120.0) * LineScrollAmount;
            PdfScrollViewer.ScrollToVerticalOffset(offset);
            e.Handled = true;
        }

        private void PdfScrollViewer_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Up:
                    PdfScrollViewer.ScrollToVerticalOffset(PdfScrollViewer.VerticalOffset - LineScrollAmount);
                    e.Handled = true;
                    break;
                case Key.Down:
                    PdfScrollViewer.ScrollToVerticalOffset(PdfScrollViewer.VerticalOffset + LineScrollAmount);
                    e.Handled = true;
                    break;
                case Key.PageUp:
                    PdfScrollViewer.ScrollToVerticalOffset(PdfScrollViewer.VerticalOffset - PdfScrollViewer.ViewportHeight);
                    e.Handled = true;
                    break;
                case Key.PageDown:
                    PdfScrollViewer.ScrollToVerticalOffset(PdfScrollViewer.VerticalOffset + PdfScrollViewer.ViewportHeight);
                    e.Handled = true;
                    break;
            }
        }
    }
}
