using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace HexxcorpDocumentReviewUtility.Services
{
    /// <summary>
    /// A minimal native Win32 child window hosted inside the WPF visual tree.
    /// This exists purely as a stable, resizable "docking slot" - the real
    /// Microsoft Word application window gets reparented into its handle
    /// (via <see cref="EmbeddedWordController"/>) so that Word itself renders
    /// and edits the document with true 1:1 fidelity, instead of a custom
    /// from-scratch document renderer.
    /// </summary>
    public sealed class Win32PanelHost : HwndHost
    {
        private const int WS_CHILD = 0x40000000;
        private const int WS_CLIPCHILDREN = 0x02000000;
        private const int WS_VISIBLE = 0x10000000;
        private const int WS_EX_CONTROLPARENT = 0x00010000;

        private IntPtr _hwnd = IntPtr.Zero;

        /// <summary>Raised whenever this panel's on-screen bounds change, so an attached native window can be resized to fill it.</summary>
        public event EventHandler? SizeOrPositionChanged;

        /// <summary>The native handle of this container; reparent other native windows onto this to dock them inside the WPF layout.</summary>
        public new IntPtr Handle => _hwnd;

        protected override HandleRef BuildWindowCore(HandleRef hwndParent)
        {
            _hwnd = NativeMethods.CreateWindowEx(
                WS_EX_CONTROLPARENT,
                "static",
                string.Empty,
                WS_CHILD | WS_CLIPCHILDREN | WS_VISIBLE,
                0,
                0,
                Math.Max((int)ActualWidth, 1),
                Math.Max((int)ActualHeight, 1),
                hwndParent.Handle,
                IntPtr.Zero,
                IntPtr.Zero,
                IntPtr.Zero);

            return new HandleRef(this, _hwnd);
        }

        protected override void DestroyWindowCore(HandleRef hwnd)
        {
            NativeMethods.DestroyWindow(hwnd.Handle);
            _hwnd = IntPtr.Zero;
        }

        protected override void OnWindowPositionChanged(Rect rcBoundingBox)
        {
            base.OnWindowPositionChanged(rcBoundingBox);
            SizeOrPositionChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    /// <summary>Shared Win32 P/Invoke declarations for native window embedding (used by <see cref="Win32PanelHost"/> and <see cref="EmbeddedWordController"/>).</summary>
    internal static class NativeMethods
    {
        public const int GWL_STYLE = -16;
        public const int WS_CHILD = 0x40000000;
        public const int WS_POPUP = unchecked((int)0x80000000);
        public const int WS_CAPTION = 0x00C00000;
        public const int WS_THICKFRAME = 0x00040000;
        public const int WS_SYSMENU = 0x00080000;
        public const int WS_MINIMIZEBOX = 0x00020000;
        public const int WS_MAXIMIZEBOX = 0x00010000;
        public const int WS_MAXIMIZE = 0x01000000;
        public const int WS_CLIPCHILDREN = 0x02000000;
        public const int WS_CLIPSIBLINGS = 0x04000000;
        public const uint SWP_NOZORDER = 0x0004;
        public const uint SWP_NOACTIVATE = 0x0010;
        public const uint SWP_FRAMECHANGED = 0x0020;
        public const int SW_SHOW = 5;
        public const int SW_HIDE = 0;

        public const uint RDW_INVALIDATE = 0x0001;
        public const uint RDW_INTERNALPAINT = 0x0002;
        public const uint RDW_ERASE = 0x0004;
        public const uint RDW_ALLCHILDREN = 0x0080;
        public const uint RDW_UPDATENOW = 0x0100;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr CreateWindowEx(
            int dwExStyle,
            string lpClassName,
            string lpWindowName,
            int dwStyle,
            int x,
            int y,
            int nWidth,
            int nHeight,
            IntPtr hWndParent,
            IntPtr hMenu,
            IntPtr hInstance,
            IntPtr lpParam);

        [DllImport("user32.dll")]
        public static extern bool DestroyWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, bool bErase);

        [DllImport("user32.dll")]
        public static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, string? lpszClass, string? lpszWindow);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    }
}
