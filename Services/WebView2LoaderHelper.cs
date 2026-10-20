using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Web.WebView2.Core;

namespace HexxcorpDocumentReviewUtility.Services;

/// <summary>
/// Ensures WebView2Loader.dll is located, extracted, and loaded properly across all deployment modes:
/// single-file executable, unpacked directory, loose files, or bare .exe on the Windows Desktop.
/// </summary>
public static class WebView2LoaderHelper
{
    private static bool _initialized;
    private static readonly object _syncLock = new();

    public static void EnsureLoaderInitialized()
    {
        if (_initialized) return;

        lock (_syncLock)
        {
            if (_initialized) return;

            try
            {
                string? loaderFolder = ResolveAndMaterializeLoader();
                if (!string.IsNullOrEmpty(loaderFolder))
                {
                    try
                    {
                        CoreWebView2Environment.SetLoaderDllFolderPath(loaderFolder);
                    }
                    catch
                    {
                        // Ignore if already set or loaded
                    }

                    string fullDllPath = Path.Combine(loaderFolder, "WebView2Loader.dll");
                    if (File.Exists(fullDllPath))
                    {
                        try
                        {
                            NativeLibrary.Load(fullDllPath);
                        }
                        catch
                        {
                        }

                        try
                        {
                            var webView2Assembly = typeof(CoreWebView2Environment).Assembly;
                            NativeLibrary.SetDllImportResolver(webView2Assembly, (libraryName, assembly, searchPath) =>
                            {
                                if (string.Equals(libraryName, "WebView2Loader.dll", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(libraryName, "WebView2Loader", StringComparison.OrdinalIgnoreCase))
                                {
                                    if (File.Exists(fullDllPath))
                                    {
                                        return NativeLibrary.Load(fullDllPath);
                                    }
                                }
                                return IntPtr.Zero;
                            });
                        }
                        catch
                        {
                        }
                    }
                }
            }
            catch
            {
            }

            _initialized = true;
        }
    }

    private static string? ResolveAndMaterializeLoader()
    {
        // 1. Check next to current executing process / base directory
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        if (File.Exists(Path.Combine(baseDir, "WebView2Loader.dll")))
        {
            return baseDir;
        }
        string subPath = Path.Combine(baseDir, "runtimes", "win-x64", "native");
        if (File.Exists(Path.Combine(subPath, "WebView2Loader.dll")))
        {
            return subPath;
        }

        // 2. Check .NET single-file bundle extraction folder in Temp\.net\HexxcorpDocumentReviewUtility
        try
        {
            string tempNet = Path.Combine(Path.GetTempPath(), ".net", "HexxcorpDocumentReviewUtility");
            if (Directory.Exists(tempNet))
            {
                var latestBundle = Directory.EnumerateDirectories(tempNet)
                    .OrderByDescending(Directory.GetLastWriteTimeUtc)
                    .FirstOrDefault();
                if (latestBundle != null)
                {
                    if (File.Exists(Path.Combine(latestBundle, "WebView2Loader.dll")))
                    {
                        return latestBundle;
                    }
                    string bundleSub = Path.Combine(latestBundle, "runtimes", "win-x64", "native");
                    if (File.Exists(Path.Combine(bundleSub, "WebView2Loader.dll")))
                    {
                        return bundleSub;
                    }
                }
            }
        }
        catch
        {
        }

        // 3. Fallback: extract the embedded WebView2Loader.dll resource to %LOCALAPPDATA%\HexxcorpDocumentReviewUtility\runtimes\win-x64\native\
        try
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string targetDir = Path.Combine(localAppData, "HexxcorpDocumentReviewUtility", "runtimes", "win-x64", "native");
            Directory.CreateDirectory(targetDir);
            string targetFile = Path.Combine(targetDir, "WebView2Loader.dll");

            var assembly = typeof(WebView2LoaderHelper).Assembly;
            string resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(n => n.EndsWith("WebView2Loader.dll", StringComparison.OrdinalIgnoreCase)) ?? string.Empty;

            if (!string.IsNullOrEmpty(resourceName))
            {
                using var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream != null)
                {
                    if (!File.Exists(targetFile) || new FileInfo(targetFile).Length != stream.Length)
                    {
                        using var fs = new FileStream(targetFile, FileMode.Create, FileAccess.Write, FileShare.None);
                        stream.CopyTo(fs);
                    }
                    return targetDir;
                }
            }

            if (File.Exists(targetFile))
            {
                return targetDir;
            }
        }
        catch
        {
        }

        return null;
    }
}
