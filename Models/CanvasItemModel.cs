using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Media.Imaging;

namespace HexxcorpDocumentReviewUtility.Models
{
    /// <summary>
    /// Identifies the kind of content a workspace canvas item renders.
    /// </summary>
    public enum WorkspaceItemKind
    {
        Image,
        Pdf
    }

    /// <summary>
    /// Base class for an item placed on the PureRef-style workspace canvas.
    /// Implements <see cref="INotifyPropertyChanged"/> so that the canvas
    /// position/size bindings update live as the user drags or resizes items.
    /// </summary>
    public abstract class CanvasItemModel : INotifyPropertyChanged
    {
        private double _x;
        private double _y;
        private double _width;
        private double _height;
        private double _zoom = 1.0;
        private bool _isHidden;
        private string _filePath = string.Empty;
        private byte[]? _fileBytes;

        /// <summary>Unique identifier used for diffing/removal and project-file round-tripping.</summary>
        public Guid Id { get; init; } = Guid.NewGuid();

        public abstract WorkspaceItemKind Kind { get; }

        public double X
        {
            get => _x;
            set { _x = value; OnPropertyChanged(); }
        }

        public double Y
        {
            get => _y;
            set { _y = value; OnPropertyChanged(); }
        }

        public double Width
        {
            get => _width;
            set { _width = value; OnPropertyChanged(); }
        }

        public double Height
        {
            get => _height;
            set { _height = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Content magnification level (1.0 = 100%), independent of the
        /// frame's <see cref="Width"/>/<see cref="Height"/>. Persisted with
        /// the workspace file so a document's zoom level is restored on load.
        /// </summary>
        public double Zoom
        {
            get => _zoom;
            set { _zoom = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// When true, the item is hidden from the canvas (removed via the "X" chrome
        /// button) but remains in the workspace's item list and the saved .lqw file.
        /// Restored via the green restore icon in the document drawer.
        /// </summary>
        public bool IsHidden
        {
            get => _isHidden;
            set { _isHidden = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Original source path, kept only so the item's display name can be
        /// derived from the original file name. Not used for loading content
        /// once <see cref="FileBytes"/> is populated - the workspace never
        /// re-reads this path from disk, so saved/traded workspace files keep
        /// working even if this path no longer exists on another machine.
        /// </summary>
        public string FilePath
        {
            get => _filePath;
            set { _filePath = value; OnPropertyChanged(); OnPropertyChanged(nameof(DisplayName)); }
        }

        /// <summary>
        /// The raw file content (image or PDF bytes), embedded directly into
        /// the workspace project file on save. This is what makes a saved
        /// .lqw file fully self-contained and portable to another, even
        /// disconnected, workstation that only has the app's .exe.
        /// </summary>
        public byte[]? FileBytes
        {
            get => _fileBytes;
            set { _fileBytes = value; OnPropertyChanged(); }
        }

        /// <summary>File name only, shown in the item's drag header.</summary>
        public string DisplayName => string.IsNullOrWhiteSpace(FilePath) ? "Untitled" : Path.GetFileName(FilePath);

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    /// <summary>A resizable/movable image frame on the workspace canvas.</summary>
    public sealed class ImageItemModel : CanvasItemModel
    {
        public override WorkspaceItemKind Kind => WorkspaceItemKind.Image;

        private BitmapImage? _image;

        /// <summary>Decoded bitmap displayed inside the frame. Not persisted directly; reloaded from <see cref="CanvasItemModel.FilePath"/>.</summary>
        public BitmapImage? Image
        {
            get => _image;
            set { _image = value; OnPropertyChanged(); }
        }
    }

    /// <summary>A resizable/movable scrollable PDF viewer frame on the workspace canvas.</summary>
    public sealed class PdfItemModel : CanvasItemModel
    {
        public override WorkspaceItemKind Kind => WorkspaceItemKind.Pdf;
    }
}
