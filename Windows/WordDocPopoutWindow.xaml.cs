using System;
using System.Windows;

namespace HexxcorpDocumentReviewUtility.Windows
{
    /// <summary>
    /// A simple, maximized host window used as the "Pop Out" destination for
    /// the embedded Word document viewer. It owns its own native docking
    /// slot (<see cref="HostPanel"/>); the real Word window is reparented
    /// directly onto that native handle by <see cref="Services.EmbeddedWordController"/>
    /// - no WPF content is ever moved between windows, so the same live
    /// Word session is reused throughout and no edits are ever lost in the
    /// hand-off.
    /// </summary>
    public partial class WordDocPopoutWindow : Window
    {
        /// <summary>Raised when the user clicks "Pop In"; the owner is responsible for reattaching the document and closing this window.</summary>
        public event EventHandler? PopInRequested;

        /// <summary>The native docking slot that Word's window gets reparented onto while popped out.</summary>
        public Services.Win32PanelHost HostPanel { get; } = new();

        public WordDocPopoutWindow()
        {
            InitializeComponent();
            ContentHost.Children.Add(HostPanel);
        }

        private void PopInButton_Click(object sender, RoutedEventArgs e)
        {
            PopInRequested?.Invoke(this, EventArgs.Empty);
        }
    }
}
