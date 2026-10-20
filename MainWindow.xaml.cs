using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using System.Windows.Threading;
using HexxcorpDocumentReviewUtility.Models;
using HexxcorpDocumentReviewUtility.Services;
using HexxcorpDocumentReviewUtility.ViewModels;
using HexxcorpDocumentReviewUtility.Windows;

namespace HexxcorpDocumentReviewUtility;

/// <summary>
/// Main dashboard window for the Hexxcorp Document Review Utility. Hosts the
/// live clock / 30-day header, the interactive date calculator, the
/// persisted "Current PSP Snapshot Date" field, and the entry points to the
/// Workspace canvas and the editable review-criteria reference modal.
/// </summary>
public partial class MainWindow : Window
{
    // Matches an optional leading '-' followed by one or more digits, used to
    // validate the "+/- N days" offset text box as the user types.
    private static readonly Regex SignedIntegerPattern = new(@"^-?\d*$", RegexOptions.Compiled);

    // WM_SYSCOMMAND / SC_* constants used to detect minimize/restore/maximize
    // requests on this window before Windows applies its native "hide owned
    // windows when the owner is minimized" behavior (see DetachSecondaryWindows).
    private const int WM_SYSCOMMAND = 0x0112;
    private const int ScCommandMask = 0xFFF0;
    private const int ScMinimize = 0xF020;
    private const int ScMaximize = 0xF030;
    private const int ScRestore = 0xF120;

    private readonly MainViewModel _viewModel = new();
    private readonly DispatcherTimer _clockTimer;
    private bool _isLoadingConfig;
    private readonly DispatcherTimer _trackerSaveTimer;
    private double _trackerDrawerWidth = 480;

    private void TrackerToggleButton_Click(object sender, RoutedEventArgs e)
    {
        bool opening = TrackerDrawer.Visibility != Visibility.Visible;
        if (opening)
        {
            TrackerDrawer.Visibility = Visibility.Visible;
            TrackerSplitter.Visibility = Visibility.Visible;
            TrackerColumn.Width = new GridLength(_trackerDrawerWidth);
            MinHeight = Math.Max(MinHeight, 300);
            if (WindowState == WindowState.Normal)
            {
                Width += _trackerDrawerWidth + 6;
                Height = Math.Max(Height, 640);
            }
            TrackerToggleButton.Content = "Hide Dates & Work History <";
        }
        else
        {
            _trackerDrawerWidth = TrackerColumn.ActualWidth > 0 ? TrackerColumn.ActualWidth : _trackerDrawerWidth;
            TrackerDrawer.Visibility = Visibility.Collapsed;
            TrackerSplitter.Visibility = Visibility.Collapsed;
            TrackerColumn.Width = new GridLength(0);
            if (WindowState == WindowState.Normal)
            {
                Width = Math.Max(MinWidth, Width - _trackerDrawerWidth - 6);
            }
            TrackerToggleButton.Content = "Dates & Work History >";
        }
    }

    private void TrackerSplitter_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
    {
        // Dragging left widens the drawer (and the window) so the left panel keeps its size.
        double target = Math.Max(300, TrackerColumn.ActualWidth - e.HorizontalChange);
        double delta = target - TrackerColumn.ActualWidth;
        TrackerColumn.Width = new GridLength(target);
        _trackerDrawerWidth = target;
        if (WindowState == WindowState.Normal)
        {
            Left -= 0;
            Width += delta;
        }
    }

    // The workspace window is a singleton: it is created once and reused via
    // Show()/Hide() for the lifetime of the application (see
    // WorkspaceWindow's own Closing handler for why it can never truly Close()
    // until real application shutdown).
    private WorkspaceWindow? _workspaceWindow;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;

        LoadConfig();
        MainTrackerPanel.DataContext = _viewModel.Tracker;

        // Debounced persistence of the tracker into config.json.
        _trackerSaveTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _trackerSaveTimer.Tick += (_, _) => { _trackerSaveTimer.Stop(); SaveConfig(); };
        _viewModel.Tracker.PropertyChanged += (_, _) =>
        {
            if (_isLoadingConfig) return;
            _trackerSaveTimer.Stop();
            _trackerSaveTimer.Start();
        };
        // Persist to config.json whenever the calculator or snapshot date changes.
        _viewModel.PropertyChanged += ViewModel_PropertyChanged;

        // Keep Workspace/Minimums windows visible and at whatever
        // state (e.g. maximized) they were in, even when this root
        // window is minimized - see the WM_SYSCOMMAND hook and
        // MainWindow_StateChanged below.
        StateChanged += MainWindow_StateChanged;

        _clockTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1)
        };
        _clockTimer.Tick += (_, _) => _viewModel.CurrentTimeText = DateTime.Now.ToString("h:mm:ss tt", CultureInfo.InvariantCulture);
        _clockTimer.Start();
    }

    /// <summary>
    /// Hooks this window's native message pump so we can detach owned
    /// windows the instant a minimize command is issued - before Windows
    /// auto-hides them - rather than reacting after the fact.
    /// </summary>
    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        if (PresentationSource.FromVisual(this) is HwndSource source)
        {
            source.AddHook(WindowProcHook);
        }
    }

    private IntPtr WindowProcHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WM_SYSCOMMAND)
        {
            int command = wParam.ToInt32() & ScCommandMask;
            if (command == ScMinimize)
            {
                DetachSecondaryWindows();
            }
            else if (command is ScMaximize or ScRestore)
            {
                ReattachSecondaryWindows();
            }
        }

        return IntPtr.Zero;
    }

    /// <summary>
    /// Fallback for any minimize/restore path that does not go through
    /// WM_SYSCOMMAND (e.g. "Show Desktop" / Win+D). Idempotent - safe to
    /// call even if the WM_SYSCOMMAND hook already handled it.
    /// </summary>
    private void MainWindow_StateChanged(object? sender, EventArgs e)
    {
        if (WindowState == WindowState.Minimized)
        {
            DetachSecondaryWindows();
        }
        else
        {
            ReattachSecondaryWindows();
        }
    }

    /// <summary>
    /// Breaks the native Win32 "owned window" link for every window this
    /// dashboard opened (Workspace, Review Criteria, etc.) so
    /// Windows does not automatically hide/minimize them along with this
    /// root window, and re-shows any that were already hidden by the time
    /// this runs. Windows otherwise hides owned windows whenever their
    /// owner is minimized - this keeps them visible (and in whatever state,
    /// e.g. maximized, the user left them in) regardless of what happens to
    /// the main window.
    /// </summary>
    private void DetachSecondaryWindows()
    {
        foreach (Window window in Application.Current.Windows)
        {
            if (ReferenceEquals(window, this) || !ReferenceEquals(window.Owner, this))
            {
                continue;
            }

            new WindowInteropHelper(window).Owner = IntPtr.Zero;

            if (!window.IsVisible)
            {
                window.Show();
            }
        }
    }

    /// <summary>Restores the normal owner link once this window is no longer minimized.</summary>
    private void ReattachSecondaryWindows()
    {
        IntPtr mainHandle = new WindowInteropHelper(this).Handle;
        foreach (Window window in Application.Current.Windows)
        {
            if (ReferenceEquals(window, this) || !ReferenceEquals(window.Owner, this))
            {
                continue;
            }

            var helper = new WindowInteropHelper(window);
            if (helper.Owner == IntPtr.Zero)
            {
                helper.Owner = mainHandle;
            }
        }
    }

    /// <summary>Loads persisted values from config.json into the view model.</summary>
    private void LoadConfig()
    {
        _isLoadingConfig = true;
        try
        {
            AppConfig config = ConfigService.Load();

            _viewModel.CurrentPspSnapshotDate = TryParseDate(config.CurrentPspSnapshotDate);
            _viewModel.TargetDate = TryParseDate(config.LastTargetDate) ?? DateTime.Today;
            _viewModel.DayOffset = config.LastDayOffset != 0 ? config.LastDayOffset : 30;

            if (config.Tracker != null)
            {
                _viewModel.Tracker.Load(config.Tracker);
            }
            // config's PSP date is authoritative (it is also shared with the main date picker).
            if (_viewModel.CurrentPspSnapshotDate.HasValue)
            {
                _viewModel.Tracker.PspSnapshot = TrackerDates.Show(_viewModel.CurrentPspSnapshotDate);
            }
            }
        catch (Exception ex)
        {
            // Config loading must never prevent the app from starting.
            MessageBox.Show(
                $"Could not load saved settings; defaults were used.\n\n{ex.Message}",
                "Hexxcorp Document Review Utility",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
        finally
        {
            _isLoadingConfig = false;
        }
    }

    private static DateTime? TryParseDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return DateTime.TryParseExact(
            value,
            MainViewModel.DateStorageFormat,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out DateTime parsed)
            ? parsed
            : null;
    }

    private void ViewModel_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (_isLoadingConfig)
        {
            return;
        }

        if (e.PropertyName is nameof(MainViewModel.CurrentPspSnapshotDate)
            or nameof(MainViewModel.TargetDate)
            or nameof(MainViewModel.DayOffset))
        {
            SaveConfig();
        }
    }

    private void SaveConfig()
    {
        try
        {
            var config = new AppConfig
            {
                CurrentPspSnapshotDate = _viewModel.CurrentPspSnapshotDate?.ToString(MainViewModel.DateStorageFormat, CultureInfo.InvariantCulture),
                LastTargetDate = _viewModel.TargetDate?.ToString(MainViewModel.DateStorageFormat, CultureInfo.InvariantCulture),
                LastDayOffset = _viewModel.DayOffset,
                Tracker = _viewModel.Tracker.HasAnyData ? _viewModel.Tracker.ToData() : null
            };

            ConfigService.Save(config);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Could not save settings to config.json.\n\n{ex.Message}",
                "Hexxcorp Document Review Utility",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    /// <summary>Flips the header's "30 Days" reference between "past" (default, subtracts) and "future" (adds).</summary>
    private void DateDirectionToggleButton_Click(object sender, RoutedEventArgs e) => _viewModel.LookingBackward = !_viewModel.LookingBackward;

    private void WorkspaceButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_workspaceWindow == null)
            {
                _workspaceWindow = new WorkspaceWindow
                {
                    MainViewModel = _viewModel
                };
            }
            else
            {
                _workspaceWindow.MainViewModel = _viewModel;
                _workspaceWindow.SyncPspFromViewModel();
            }

            _workspaceWindow.Show();
            _workspaceWindow.Activate();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Could not open the Workspace window.\n\n{DescribeException(ex)}",
                "Hexxcorp Document Review Utility",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void MinimumsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var minimumsWindow = new MinimumsWindow
            {
                Owner = this
            };
            minimumsWindow.ShowDialog();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Could not open the Review Criteria reference.\n\n{DescribeException(ex)}",
                "Hexxcorp Document Review Utility",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private static string DescribeException(Exception ex)
    {
        var messages = new List<string>();
        for (var current = ex; current != null; current = current.InnerException)
        {
            messages.Add(current.Message);
        }
        return string.Join("\n\n", messages.Distinct());
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        _clockTimer.Stop();
        SaveConfig();

        // Real application shutdown: allow the (otherwise persistent)
        // Workspace window to actually close rather than just hide.
        if (_workspaceWindow != null)
        {
            _workspaceWindow.AllowRealClose();
            _workspaceWindow.Close();
        }
    }
}