using System;
using System.IO;
using System.Runtime.InteropServices;

namespace HexxcorpDocumentReviewUtility.Services
{
    /// <summary>
    /// Resolves well-known Windows shell folders that have no corresponding
    /// <see cref="Environment.SpecialFolder"/> value (notably Downloads).
    /// </summary>
    internal static class KnownFolders
    {
        private static readonly Guid DownloadsFolderId = new("374DE290-123F-4565-9164-39C4925E467B");

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        private static extern int SHGetKnownFolderPath(ref Guid id, uint flags, IntPtr token, out IntPtr path);

        /// <summary>
        /// Returns the current user's Downloads folder path, honoring any
        /// redirection the user has configured. Falls back to the default
        /// "%USERPROFILE%\Downloads" location if the shell API is unavailable
        /// or fails for any reason.
        /// </summary>
        public static string GetDownloadsFolderPath()
        {
            try
            {
                Guid id = DownloadsFolderId;
                if (SHGetKnownFolderPath(ref id, 0, IntPtr.Zero, out IntPtr pathPtr) == 0)
                {
                    try
                    {
                        string? path = Marshal.PtrToStringUni(pathPtr);
                        if (!string.IsNullOrEmpty(path))
                        {
                            return path;
                        }
                    }
                    finally
                    {
                        Marshal.FreeCoTaskMem(pathPtr);
                    }
                }
            }
            catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException)
            {
                // Fall through to the default-path fallback below.
            }

            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        }
    }
}
