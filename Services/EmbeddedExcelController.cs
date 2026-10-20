using System;
using System.IO;
using System.Runtime.InteropServices;

namespace HexxcorpDocumentReviewUtility.Services
{
    /// <summary>
    /// Owns a real, hidden-chrome Microsoft Excel instance used to view and
    /// edit an Excel workbook (.xlsx/.xls/.xlsm/.csv) with true 1:1 fidelity
    /// (formulas, sheets, formatting, calculations, charts, everything -
    /// because it genuinely is Excel doing the rendering/editing/saving, not
    /// a custom approximation).
    /// </summary>
    /// <remarks>
    /// Excel is automated via late-bound COM (<c>dynamic</c>) so the app has
    /// no compile-time dependency on Office Primary Interop Assemblies -
    /// it simply fails gracefully with <see cref="ExcelNotAvailableException"/>
    /// if Excel isn't installed on the machine. The workbook's real top-level
    /// window is then reparented (<c>SetParent</c>) into a native container
    /// supplied by the caller (see <see cref="Win32PanelHost"/>), so the
    /// actual Excel UI is embedded directly inside the WPF window.
    /// </remarks>
    public sealed class EmbeddedExcelController : IDisposable
    {
        private const int XlOpenXMLWorkbook = 51; // .xlsx
        private const int XlNormal = -4143;
        private const int XlMaximized = -4137;

        private dynamic? _excelApp;
        private dynamic? _workbook;
        private string? _backingFilePath;
        private IntPtr _excelHwnd = IntPtr.Zero;
        private IntPtr _currentContainer = IntPtr.Zero;
        private uint _attachedThreadId = 0;
        private int _lastWidth = -1;
        private int _lastHeight = -1;

        /// <summary>True once a workbook is currently open in this controller's Excel instance.</summary>
        public bool IsOpen => _workbook is not null;

        /// <summary>Checks (without starting Excel) whether Excel's COM automation class is registered on this machine.</summary>
        public static bool IsExcelInstalled => Type.GetTypeFromProgID("Excel.Application") is not null;

        /// <summary>
        /// Starts (if needed) an Excel instance, opens the given spreadsheet bytes
        /// from a private temp copy, and reparents Excel's real window into
        /// <paramref name="containerHandle"/>.
        /// </summary>
        public void Open(byte[] sheetBytes, IntPtr containerHandle, string fileName = "")
        {
            Close();

            Type? excelType = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new ExcelNotAvailableException("Microsoft Excel does not appear to be installed on this computer.");

            string ext = Path.GetExtension(fileName);
            if (string.IsNullOrWhiteSpace(ext) || ext.ToLowerInvariant() is not (".xlsx" or ".xls" or ".xlsm" or ".csv"))
            {
                ext = ".xlsx";
            }

            _backingFilePath = Path.Combine(Path.GetTempPath(), $"LQU_ExcelEdit_{Guid.NewGuid():N}{ext}");
            File.WriteAllBytes(_backingFilePath, sheetBytes);

            try
            {
                _excelApp = Activator.CreateInstance(excelType)!;
                _excelApp.Visible = true;
                _excelApp.DisplayAlerts = false;

                _workbook = _excelApp.Workbooks.Open(_backingFilePath);
                _excelHwnd = new IntPtr((int)_excelApp.Hwnd);
                EnsureWindowStates();
            }
            catch (COMException ex)
            {
                Close();
                throw new ExcelNotAvailableException("Microsoft Excel could not be started or did not respond. Make sure Excel is installed and not blocked by a policy/add-in.", ex);
            }

            AttachTo(containerHandle);
        }

        /// <summary>
        /// Starts (if needed) an Excel instance, creates a brand-new blank
        /// workbook (saved immediately to a private temp backing file so it
        /// behaves identically to <see cref="Open"/> from then on), and
        /// reparents Excel's real window into <paramref name="containerHandle"/>.
        /// </summary>
        public void New(IntPtr containerHandle)
        {
            Close();

            Type? excelType = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new ExcelNotAvailableException("Microsoft Excel does not appear to be installed on this computer.");

            _backingFilePath = Path.Combine(Path.GetTempPath(), $"LQU_ExcelEdit_{Guid.NewGuid():N}.xlsx");

            try
            {
                _excelApp = Activator.CreateInstance(excelType)!;
                _excelApp.Visible = true;
                _excelApp.DisplayAlerts = false;

                _workbook = _excelApp.Workbooks.Add();
                _workbook.SaveAs(_backingFilePath, XlOpenXMLWorkbook);

                _excelHwnd = new IntPtr((int)_excelApp.Hwnd);
                EnsureWindowStates();
            }
            catch (COMException ex)
            {
                Close();
                throw new ExcelNotAvailableException("Microsoft Excel could not be started or did not respond. Make sure Excel is installed and not blocked by a policy/add-in.", ex);
            }

            AttachTo(containerHandle);
        }

        /// <summary>Reparents the live Excel window into a (possibly different) native container - used for Pop Out/Pop In.</summary>
        public void AttachTo(IntPtr containerHandle)
        {
            if (_excelHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_excelHwnd))
            {
                return;
            }

            _currentContainer = containerHandle;
            _lastWidth = -1;
            _lastHeight = -1;

            uint excelThread = NativeMethods.GetWindowThreadProcessId(_excelHwnd, out _);
            uint currentThread = NativeMethods.GetCurrentThreadId();
            if (excelThread != 0 && excelThread != currentThread && excelThread != _attachedThreadId)
            {
                if (_attachedThreadId != 0)
                {
                    NativeMethods.AttachThreadInput(currentThread, _attachedThreadId, false);
                }
                NativeMethods.AttachThreadInput(currentThread, excelThread, true);
                _attachedThreadId = excelThread;
            }

            NativeMethods.SetParent(_excelHwnd, containerHandle);
            EnsureWindowStates();
            NativeMethods.SetWindowPos(_excelHwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_FRAMECHANGED);

            ResizeToFillContainer();
            NativeMethods.ShowWindow(_excelHwnd, NativeMethods.SW_SHOW);
        }

        /// <summary>Call whenever the currently-attached container's size changes so the embedded Excel window keeps filling it.</summary>
        public void ResizeToFillContainer()
        {
            if (_excelHwnd == IntPtr.Zero || _currentContainer == IntPtr.Zero)
            {
                return;
            }

            if (NativeMethods.GetClientRect(_currentContainer, out NativeMethods.RECT rect))
            {
                int width = Math.Max(rect.Right - rect.Left, 1);
                int height = Math.Max(rect.Bottom - rect.Top, 1);

                if (width == _lastWidth && height == _lastHeight)
                {
                    return;
                }

                _lastWidth = width;
                _lastHeight = height;

                NativeMethods.MoveWindow(_excelHwnd, 0, 0, width, height, true);
            }
        }

        /// <summary>
        /// Enforces xlNormal on Excel's top-level application window (so Win32
        /// sizing can expand to arbitrary widths and heights without monitor caps)
        /// while keeping the workbook's active window xlMaximized (so spreadsheet
        /// cells and sheets fill 100% of Excel's client viewport without floating borders).
        /// </summary>
        private void EnsureWindowStates()
        {
            if (_excelApp is null)
            {
                return;
            }

            try
            {
                if ((int)_excelApp.WindowState != XlNormal)
                {
                    _excelApp.WindowState = XlNormal;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (_workbook is not null)
                {
                    dynamic? windows = _workbook.Windows;
                    if (windows is not null && (int)windows.Count > 0)
                    {
                        dynamic? win = windows[1];
                        if (win is not null && (int)win.WindowState != XlMaximized)
                        {
                            win.WindowState = XlMaximized;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                dynamic? activeWin = _excelApp.ActiveWindow;
                if (activeWin is not null && (int)activeWin.WindowState != XlMaximized)
                {
                    activeWin.WindowState = XlMaximized;
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>Sets Excel's own zoom level (10-400%), using Excel's native zoom engine.</summary>
        public void SetZoomPercent(int percent)
        {
            if (_excelApp is null)
            {
                return;
            }

            try
            {
                percent = Math.Clamp(percent, 10, 400);
                _excelApp.ActiveWindow.Zoom = percent;
            }
            catch (COMException)
            {
                // Best-effort only; zoom isn't critical path.
            }
        }

        /// <summary>Returns the workbook's current zoom percentage, or 100 if unavailable.</summary>
        public int GetZoomPercent()
        {
            try
            {
                return _excelApp is null ? 100 : (int)_excelApp.ActiveWindow.Zoom;
            }
            catch (COMException)
            {
                return 100;
            }
        }

        /// <summary>Has Excel save the live workbook to its private backing file, then returns the resulting bytes for persistence.</summary>
        public byte[]? SaveAndGetBytes()
        {
            if (_workbook is null || _backingFilePath is null)
            {
                return null;
            }

            try
            {
                _workbook.Save();
                return File.ReadAllBytes(_backingFilePath);
            }
            catch (COMException)
            {
                return null;
            }
        }

        /// <summary>Exports the live workbook to a user-chosen path.</summary>
        public void ExportTo(string path)
        {
            if (_workbook is null || _backingFilePath is null)
            {
                return;
            }

            _workbook.Save();
            File.Copy(_backingFilePath, path, true);
        }

        /// <summary>Closes the workbook and quits this controller's private Excel instance, discarding the temp backing file.</summary>
        public void Close()
        {
            if (_attachedThreadId != 0)
            {
                uint currentThread = NativeMethods.GetCurrentThreadId();
                NativeMethods.AttachThreadInput(currentThread, _attachedThreadId, false);
                _attachedThreadId = 0;
            }

            _lastWidth = -1;
            _lastHeight = -1;

            try
            {
                _workbook?.Close(SaveChanges: false);
            }
            catch (COMException)
            {
                // Ignore - best-effort cleanup.
            }

            try
            {
                _excelApp?.Quit();
            }
            catch (COMException)
            {
                // Ignore - best-effort cleanup.
            }

            if (_workbook is not null)
            {
                Marshal.FinalReleaseComObject(_workbook);
                _workbook = null;
            }

            if (_excelApp is not null)
            {
                Marshal.FinalReleaseComObject(_excelApp);
                _excelApp = null;
            }

            _excelHwnd = IntPtr.Zero;
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

    /// <summary>Thrown when Microsoft Excel is not installed/available for embedded spreadsheet editing.</summary>
    public sealed class ExcelNotAvailableException : Exception
    {
        public ExcelNotAvailableException(string message) : base(message)
        {
        }

        public ExcelNotAvailableException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
