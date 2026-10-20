using System;
using System.IO;
using System.Runtime.InteropServices;

namespace HexxcorpDocumentReviewUtility.Services
{
    /// <summary>
    /// Owns a real, hidden-chrome Microsoft Word instance used to view and
    /// edit a single .docx document with true 1:1 fidelity (tables, images,
    /// fonts, styles, headers/footers, everything - because it genuinely is
    /// Word doing the rendering/editing/saving, not a custom approximation).
    /// </summary>
    /// <remarks>
    /// Word is automated via late-bound COM (<c>dynamic</c>) so the app has
    /// no compile-time dependency on the Office Primary Interop Assemblies -
    /// it simply fails gracefully with <see cref="WordNotAvailableException"/>
    /// if Word isn't installed on the machine. The document's real top-level
    /// window is then reparented (<c>SetParent</c>) into a native container
    /// supplied by the caller (see <see cref="Win32PanelHost"/>), so the
    /// actual Word UI is embedded directly inside the WPF window.
    /// </remarks>
    public sealed class EmbeddedWordController : IDisposable
    {
        private const int WdAlertsNone = 0;
        private const int WdDoNotSaveChanges = 0;
        private const int WdFormatXMLDocument = 12; // .docx

        private dynamic? _wordApp;
        private dynamic? _document;
        private string? _backingFilePath;
        private IntPtr _wordHwnd = IntPtr.Zero;
        private IntPtr _currentContainer = IntPtr.Zero;

        /// <summary>True once a document is currently open in this controller's Word instance.</summary>
        public bool IsOpen => _document is not null;

        /// <summary>Checks (without starting Word) whether Word's COM automation class is registered on this machine.</summary>
        public static bool IsWordInstalled => Type.GetTypeFromProgID("Word.Application") is not null;

        /// <summary>
        /// Starts (if needed) a Word instance, opens the given document bytes
        /// from a private temp copy, and reparents Word's real window into
        /// <paramref name="containerHandle"/>.
        /// </summary>
        public void Open(byte[] docBytes, IntPtr containerHandle)
        {
            Close();

            Type? wordType = Type.GetTypeFromProgID("Word.Application")
                ?? throw new WordNotAvailableException("Microsoft Word does not appear to be installed on this computer.");

            _backingFilePath = Path.Combine(Path.GetTempPath(), $"LQU_WordEdit_{Guid.NewGuid():N}.docx");
            File.WriteAllBytes(_backingFilePath, docBytes);

            try
            {
                _wordApp = Activator.CreateInstance(wordType)!;
                _wordApp.Visible = true;
                _wordApp.DisplayAlerts = WdAlertsNone;

                _document = _wordApp.Documents.Open(
                    FileName: _backingFilePath,
                    ConfirmConversions: false,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: true);

                dynamic docWindow = _document.Windows[1];
                _wordHwnd = new IntPtr((int)docWindow.Hwnd);
            }
            catch (COMException ex)
            {
                Close();
                throw new WordNotAvailableException("Microsoft Word could not be started or did not respond. Make sure Word is installed and not blocked by a policy/add-in.", ex);
            }

            AttachTo(containerHandle);
        }

        /// <summary>
        /// Starts (if needed) a Word instance, creates a brand-new blank
        /// document (saved immediately to a private temp backing file so it
        /// behaves identically to <see cref="Open"/> from then on), and
        /// reparents Word's real window into <paramref name="containerHandle"/>.
        /// </summary>
        public void New(IntPtr containerHandle)
        {
            Close();

            Type? wordType = Type.GetTypeFromProgID("Word.Application")
                ?? throw new WordNotAvailableException("Microsoft Word does not appear to be installed on this computer.");

            _backingFilePath = Path.Combine(Path.GetTempPath(), $"LQU_WordEdit_{Guid.NewGuid():N}.docx");

            try
            {
                _wordApp = Activator.CreateInstance(wordType)!;
                _wordApp.Visible = true;
                _wordApp.DisplayAlerts = WdAlertsNone;

                _document = _wordApp.Documents.Add();
                _document.SaveAs2(FileName: _backingFilePath, FileFormat: WdFormatXMLDocument);

                dynamic docWindow = _document.Windows[1];
                _wordHwnd = new IntPtr((int)docWindow.Hwnd);
            }
            catch (COMException ex)
            {
                Close();
                throw new WordNotAvailableException("Microsoft Word could not be started or did not respond. Make sure Word is installed and not blocked by a policy/add-in.", ex);
            }

            AttachTo(containerHandle);
        }

        /// <summary>Reparents the live Word window into a (possibly different) native container - used for Pop Out/Pop In.</summary>
        public void AttachTo(IntPtr containerHandle)
        {
            if (_wordHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_wordHwnd))
            {
                return;
            }

            _currentContainer = containerHandle;

            int style = NativeMethods.GetWindowLong(_wordHwnd, NativeMethods.GWL_STYLE);
            style &= ~(NativeMethods.WS_POPUP | NativeMethods.WS_CAPTION | NativeMethods.WS_THICKFRAME
                       | NativeMethods.WS_SYSMENU | NativeMethods.WS_MINIMIZEBOX | NativeMethods.WS_MAXIMIZEBOX);
            style |= NativeMethods.WS_CHILD;
            NativeMethods.SetWindowLong(_wordHwnd, NativeMethods.GWL_STYLE, style);

            NativeMethods.SetParent(_wordHwnd, containerHandle);
            NativeMethods.SetWindowPos(_wordHwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_FRAMECHANGED);

            ResizeToFillContainer();
            NativeMethods.ShowWindow(_wordHwnd, NativeMethods.SW_SHOW);
        }

        /// <summary>Call whenever the currently-attached container's size changes so the embedded Word window keeps filling it.</summary>
        public void ResizeToFillContainer()
        {
            if (_wordHwnd == IntPtr.Zero || _currentContainer == IntPtr.Zero)
            {
                return;
            }

            if (NativeMethods.GetClientRect(_currentContainer, out NativeMethods.RECT rect))
            {
                int width = Math.Max(rect.Right - rect.Left, 1);
                int height = Math.Max(rect.Bottom - rect.Top, 1);
                NativeMethods.MoveWindow(_wordHwnd, 0, 0, width, height, true);
            }
        }

        /// <summary>Sets Word's own zoom level (10-500%), using Word's native zoom/scroll engine - no custom transform required.</summary>
        public void SetZoomPercent(int percent)
        {
            if (_wordApp is null)
            {
                return;
            }

            try
            {
                percent = Math.Clamp(percent, 10, 500);
                _wordApp.ActiveWindow.View.Zoom.Percentage = percent;
            }
            catch (COMException)
            {
                // Best-effort only; zoom isn't critical path.
            }
        }

        /// <summary>Returns the document's current zoom percentage, or 100 if unavailable.</summary>
        public int GetZoomPercent()
        {
            try
            {
                return _wordApp is null ? 100 : (int)_wordApp.ActiveWindow.View.Zoom.Percentage;
            }
            catch (COMException)
            {
                return 100;
            }
        }

        /// <summary>Has Word save the live document to its private backing file, then returns the resulting bytes for persistence.</summary>
        public byte[]? SaveAndGetBytes()
        {
            if (_document is null || _backingFilePath is null)
            {
                return null;
            }

            try
            {
                _document.Save();
                return File.ReadAllBytes(_backingFilePath);
            }
            catch (COMException)
            {
                return null;
            }
        }

        /// <summary>Exports the live document to a user-chosen path, then re-targets Word's save path back to the private backing file so the embedded session is unaffected.</summary>
        public void ExportTo(string path)
        {
            if (_document is null || _backingFilePath is null)
            {
                return;
            }

            _document.SaveAs2(FileName: path, FileFormat: WdFormatXMLDocument);
            _document.SaveAs2(FileName: _backingFilePath, FileFormat: WdFormatXMLDocument);
        }

        /// <summary>Closes the document and quits this controller's private Word instance, discarding the temp backing file.</summary>
        public void Close()
        {
            try
            {
                _document?.Close(SaveChanges: WdDoNotSaveChanges);
            }
            catch (COMException)
            {
                // Ignore - best-effort cleanup.
            }

            try
            {
                _wordApp?.Quit();
            }
            catch (COMException)
            {
                // Ignore - best-effort cleanup.
            }

            if (_document is not null)
            {
                Marshal.FinalReleaseComObject(_document);
                _document = null;
            }

            if (_wordApp is not null)
            {
                Marshal.FinalReleaseComObject(_wordApp);
                _wordApp = null;
            }

            _wordHwnd = IntPtr.Zero;
            _currentContainer = IntPtr.Zero;

            if (_backingFilePath is not null)
            {
                try
                {
                    if (File.Exists(_backingFilePath))
                    {
                        File.Delete(_backingFilePath);
                    }
                }
                catch (IOException)
                {
                    // Ignore - temp file, harmless if left behind.
                }
                catch (UnauthorizedAccessException)
                {
                    // Ignore - temp file, harmless if left behind.
                }

                _backingFilePath = null;
            }
        }

        public void Dispose() => Close();
    }

    /// <summary>Thrown when Microsoft Word is not installed/available for embedded document editing.</summary>
    public sealed class WordNotAvailableException : Exception
    {
        public WordNotAvailableException(string message) : base(message)
        {
        }

        public WordNotAvailableException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
