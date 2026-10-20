using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;

namespace HexxcorpDocumentReviewUtility.Services
{
    /// <summary>
    /// Adds a simple tab strip to an embedded WebView2 browser. The XAML-defined WebView2 is tab 1;
    /// pop-ups / target=_blank links open as new tabs (keeping window.opener intact) instead of separate windows.
    /// </summary>
    public sealed class BrowserTabManager
    {
        private sealed class Tab
        {
            public WebView2 View = null!;
            public Border Header = null!;
            public TextBlock Title = null!;
        }

        private readonly Panel _host;
        private readonly Panel _strip;
        private readonly Func<CoreWebView2Environment?> _env;
        private readonly string _homeUrl;
        private readonly List<Tab> _tabs = new();
        private Tab? _active;

        public event Action<string>? UrlChanged;
        public event Action? NavStateChanged;
        public event Action<string>? NavigationStarted;
        public event Action? NavigationFinished;
        public event Action<string>? DownloadCompleted;

        public BrowserTabManager(Panel host, Panel strip, WebView2 firstView, Func<CoreWebView2Environment?> env, string homeUrl)
        {
            _host = host;
            _strip = strip;
            _env = env;
            _homeUrl = homeUrl;
            var first = new Tab { View = firstView };
            _tabs.Add(first);
            _active = first;
            Rebuild();
        }

        public WebView2 Active => (_active ?? _tabs[0]).View;

        /// <summary>Call once the first tab's CoreWebView2 is ready.</summary>
        public void AttachFirst() => Hook(_tabs[0]);

        private void Hook(Tab tab)
        {
            var core = tab.View.CoreWebView2;
            core.Settings.IsGeneralAutofillEnabled = true;
            core.Settings.IsPasswordAutosaveEnabled = true;
            core.SourceChanged += (_, _) =>
            {
                if (tab == _active && tab.View.Source != null) UrlChanged?.Invoke(tab.View.Source.ToString());
            };
            core.DocumentTitleChanged += (_, _) => SetTitle(tab);
            core.NavigationStarting += (_, e) => { if (tab == _active) NavigationStarted?.Invoke(e.Uri); };
            core.NavigationCompleted += (_, _) =>
            {
                if (tab == _active) { NavigationFinished?.Invoke(); }
                SetTitle(tab);
                NavStateChanged?.Invoke();
            };
            core.HistoryChanged += (_, _) => NavStateChanged?.Invoke();
            core.WindowCloseRequested += (_, _) => Close(tab);
            core.DownloadStarting += (_, e) =>
            {
                var op = e.DownloadOperation;
                op.StateChanged += (s, _) =>
                {
                    if (op.State == CoreWebView2DownloadState.Completed) DownloadCompleted?.Invoke(op.ResultFilePath);
                };
            };
            core.NewWindowRequested += async (_, e) =>
            {
                var deferral = e.GetDeferral();
                try
                {
                    var created = await CreateTabAsync();
                    if (created != null)
                    {
                        e.NewWindow = created.View.CoreWebView2;
                        e.Handled = true;
                        Select(created);
                    }
                }
                catch (Exception)
                {
                    e.Handled = false;
                }
                finally
                {
                    deferral.Complete();
                }
            };
            SetTitle(tab);
        }

        private async Task<Tab?> CreateTabAsync()
        {
            var env = _env();
            if (env == null) return null;
            var view = new WebView2 { Visibility = Visibility.Collapsed };
            _host.Children.Add(view);
            var tab = new Tab { View = view };
            _tabs.Add(tab);
            await view.EnsureCoreWebView2Async(env);
            Hook(tab);
            Rebuild();
            return tab;
        }

        /// <summary>Opens a blank tab and navigates it to the given URL (or the home page).</summary>
        public async Task NewTabAsync(string? url = null)
        {
            var tab = await CreateTabAsync();
            if (tab == null) return;
            Select(tab);
            tab.View.CoreWebView2.Navigate(string.IsNullOrWhiteSpace(url) ? _homeUrl : url);
        }

        private void Select(Tab tab)
        {
            _active = tab;
            foreach (var t in _tabs) t.View.Visibility = t == tab ? Visibility.Visible : Visibility.Collapsed;
            Rebuild();
            if (tab.View.Source != null) UrlChanged?.Invoke(tab.View.Source.ToString());
            NavStateChanged?.Invoke();
        }

        private void Close(Tab tab)
        {
            if (_tabs.Count <= 1 || !_tabs.Contains(tab)) return;
            int idx = _tabs.IndexOf(tab);
            _tabs.Remove(tab);
            _host.Children.Remove(tab.View);
            try { tab.View.Dispose(); } catch (Exception) { }
            if (_active == tab) Select(_tabs[Math.Min(idx, _tabs.Count - 1)]);
            else Rebuild();
        }

        private void SetTitle(Tab tab)
        {
            string t = tab.View.CoreWebView2?.DocumentTitle ?? string.Empty;
            if (string.IsNullOrWhiteSpace(t)) t = "New tab";
            tab.Title ??= new TextBlock();
            tab.Title.Text = t.Length > 22 ? t[..22] + "…" : t;
        }

        private void Rebuild()
        {
            _strip.Children.Clear();
            foreach (var tab in _tabs)
            {
                var t = tab;
                bool isActive = t == _active;
                t.Title ??= new TextBlock { Text = "Tab" };
                t.Title.Foreground = isActive ? Brushes.Black : Brushes.White;
                t.Title.FontSize = 11;
                t.Title.VerticalAlignment = VerticalAlignment.Center;

                var row = new StackPanel { Orientation = Orientation.Horizontal };
                if (t.Title.Parent is Panel oldParent) oldParent.Children.Remove(t.Title);
                row.Children.Add(t.Title);
                if (_tabs.Count > 1)
                {
                    var close = new Button
                    {
                        Content = "✕",
                        FontSize = 9,
                        Padding = new Thickness(4, 0, 4, 0),
                        Margin = new Thickness(6, 0, 0, 0),
                        Background = Brushes.Transparent,
                        BorderThickness = new Thickness(0),
                        Foreground = isActive ? Brushes.Black : Brushes.White,
                        Cursor = System.Windows.Input.Cursors.Hand,
                        ToolTip = "Close tab"
                    };
                    close.Click += (_, _) => Close(t);
                    row.Children.Add(close);
                }

                t.Header = new Border
                {
                    Child = row,
                    Padding = new Thickness(8, 3, 6, 3),
                    Margin = new Thickness(0, 0, 3, 0),
                    CornerRadius = new CornerRadius(3, 3, 0, 0),
                    Background = isActive ? new SolidColorBrush(Color.FromRgb(0xFF, 0xC2, 0x0E)) : new SolidColorBrush(Color.FromRgb(0x1B, 0x4A, 0x8F)),
                    Cursor = System.Windows.Input.Cursors.Hand,
                    MaxWidth = 190
                };
                t.Header.MouseLeftButtonDown += (_, _) => Select(t);
                t.Header.MouseDown += (_, e) => { if (e.ChangedButton == System.Windows.Input.MouseButton.Middle) Close(t); };
                _strip.Children.Add(t.Header);
            }

            var plus = new Button
            {
                Content = "+",
                FontWeight = FontWeights.Bold,
                Padding = new Thickness(8, 0, 8, 0),
                Background = Brushes.Transparent,
                Foreground = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xFF, 0xC2, 0x0E)),
                BorderThickness = new Thickness(1),
                Cursor = System.Windows.Input.Cursors.Hand,
                ToolTip = "New tab"
            };
            plus.Click += async (_, _) => await NewTabAsync();
            _strip.Children.Add(plus);
        }

        public void DisposeAll()
        {
            foreach (var t in _tabs)
            {
                try { t.View.Dispose(); } catch (Exception) { }
            }
        }
    }
}
