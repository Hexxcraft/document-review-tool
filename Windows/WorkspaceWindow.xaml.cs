using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using HexxcorpDocumentReviewUtility.Controls;
using HexxcorpDocumentReviewUtility.Models;
using HexxcorpDocumentReviewUtility.Services;
using HexxcorpDocumentReviewUtility.ViewModels;
using Microsoft.Win32;

namespace HexxcorpDocumentReviewUtility.Windows
{
    /// <summary>
    /// A modeless, simplified PureRef-style reference canvas. Supports
    /// drag-and-drop ingestion of images and PDFs (plus one-click ingestion
    /// of recent Downloads-folder files), free resizing/repositioning/zooming
    /// of each item, manual project save/load, a 30-minute autosave backup,
    /// a top interactive date calculator with driver/contractor header
    /// fields, and a bottom "quick date reference" bar (30/60/90-day and
    /// 1/3/5/10-year rolling milestones plus a "Backdate" box) that can be
    /// toggled between "past" (default: subtracts time, oldest date first)
    /// and "future" (adds time, youngest date first) modes. All displayed
    /// dates use a numeric-only format (MM/dd/yyyy) with no month or day names.
    /// </summary>
    public partial class WorkspaceWindow : Window
    {
        private const double DefaultImageWidth = 320;
        private const double DefaultImageHeight = 240;
        private const double DefaultPdfWidth = 420;
        private const double DefaultPdfHeight = 540;
        private static readonly TimeSpan AutosaveInterval = TimeSpan.FromMinutes(30);
        private static readonly TimeSpan QuickDateRefreshInterval = TimeSpan.FromMinutes(1);
        private static readonly TimeSpan RecentDownloadWindow = TimeSpan.FromMinutes(5);
        private const string QuickDateFormat = "MM/dd/yyyy";
        private static readonly Regex SignedIntegerPattern = new(@"^-?\d*$", RegexOptions.Compiled);

        private readonly ObservableCollection<CanvasItemModel> _items = new();
        private readonly ObservableCollection<DrawerEntry> _drawerItems = new();
        private readonly DispatcherTimer _autosaveTimer;
        private readonly DispatcherTimer _quickDateTimer;
        private int _dropCounter;

        /// <summary>The calendar date the quick-date bar was last computed for, so we only touch the UI when the day actually rolls over.</summary>
        private DateTime _quickDatesComputedFor = DateTime.MinValue;

        /// <summary>True when the bottom quick-date bar is in "past" mode (default, subtracts time, oldest-first); false is "future" mode (adds time, youngest-first).</summary>
        private bool _quickDatesLookingBackward = true;

        /// <summary>True while the window is performing a confirmed close; prevents re-entrant Closing prompts.</summary>
        private bool _allowHide;

        /// <summary>Display/default-save file name for the embedded Word document loaded in the Notes panel, empty when none is loaded.</summary>
        private string _wordDocFileName = string.Empty;

        /// <summary>Owns the hidden-chrome Microsoft Word instance used to view/edit the currently loaded document (true 1:1 fidelity - Word itself does the rendering), if any.</summary>
        private EmbeddedWordController? _wordController;

        /// <summary>The native docking slot (inside <c>WordDocHostContainer</c>) that Word's window is reparented onto while docked in this window.</summary>
        private Win32PanelHost? _wordHostPanel;

        /// <summary>Set while the Word document viewer is popped out into its own window.</summary>
        private WordDocPopoutWindow? _wordDocPopout;

        private const int WordDocZoomStepPercent = 10;

        /// <summary>Display/default-save file name for the embedded Excel workbook loaded in the Sheet panel, empty when none is loaded.</summary>
        private string _excelDocFileName = string.Empty;

        /// <summary>Owns the hidden-chrome Microsoft Excel instance used to view/edit the currently loaded workbook (true 1:1 fidelity - Excel itself does the rendering), if any.</summary>
        private EmbeddedExcelController? _excelController;

        /// <summary>The native docking slot (inside <c>ExcelSheetHostContainer</c>) that Excel's window is reparented onto while docked in this window.</summary>
        private Win32PanelHost? _excelHostPanel;

        /// <summary>Set while the Excel workbook viewer is popped out into its own window.</summary>
        private ExcelSheetPopoutWindow? _excelSheetPopout;

        private const int ExcelZoomStepPercent = 10;

        private MainViewModel? _mainViewModel;
        private bool _isSyncingPsp;

        /// <summary>Reference to MainWindow's ViewModel to keep the PSP Snapshot Date in sync across both windows.</summary>
        public MainViewModel? MainViewModel
        {
            get => _mainViewModel;
            set
            {
                if (_mainViewModel != null)
                {
                    _mainViewModel.PropertyChanged -= MainViewModel_PropertyChanged;
                }
                _mainViewModel = value;
                if (_mainViewModel != null)
                {
                    _mainViewModel.PropertyChanged += MainViewModel_PropertyChanged;
                    SyncPspFromViewModel();
                    AttachTracker();
                }
            }
        }

        // ======================= Dates & Employment History tracker sync =======================

        private bool _syncingTracker;

        private IEnumerable<(TextBox Box, Func<TrackerModel, string> Get, Action<TrackerModel, string> Set)> TrackerBoxMap()
        {
            yield return (DrugTestRanDateBox, t => t.DtDate, (t, v) => t.DtDate = v);
            yield return (MvrLastReviewDateBox, t => t.MvrRan, (t, v) => t.MvrRan = v);
            yield return (FingerprintDateBox, t => t.FpRan, (t, v) => t.FpRan = v);
            yield return (PhysicalNextDateBox, t => t.PhysicalExp, (t, v) => t.PhysicalExp = v);
            yield return (CertificationDateBox, t => t.Certification, (t, v) => t.Certification = v);
            yield return (LimitedConsentDateBox, t => t.LimitedConsent, (t, v) => t.LimitedConsent = v);
            yield return (FullConsentDateBox, t => t.FullConsent, (t, v) => t.FullConsent = v);
            yield return (CdlisDacDateBox, t => t.CdlisDac, (t, v) => t.CdlisDac = v);
            yield return (PspSnapshotDateBox, t => t.PspSnapshot, (t, v) => t.PspSnapshot = v);
            yield return (PspRanDateBox, t => t.PspRan, (t, v) => t.PspRan = v);
        }

        private void AttachTracker()
        {
            if (_mainViewModel == null) return;
            var tracker = _mainViewModel.Tracker;
            WorkspaceFileService.TrackerProvider = () => _mainViewModel?.Tracker.ToData();
            tracker.PropertyChanged -= Tracker_PropertyChanged;
            tracker.PropertyChanged += Tracker_PropertyChanged;
            TrackerPanelHost.DataContext = tracker;
            TrackerPanelHost.CompactMode = true;
            PullTrackerIntoBoxes();
        }

        private void Tracker_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e) => PullTrackerIntoBoxes();

        /// <summary>Model to boxes: fills File Dates boxes from the tracker (only with non-empty tracker values).</summary>
        private void PullTrackerIntoBoxes()
        {
            if (_syncingTracker || _mainViewModel == null) return;
            var t = _mainViewModel.Tracker;
            _syncingTracker = true;
            try
            {
                foreach (var (box, get, _) in TrackerBoxMap())
                {
                    string v = get(t);
                    if (!string.IsNullOrEmpty(v) && box.Text != v) box.Text = v;
                }
                string latest = t.LatestJobEndText;
                if (!string.IsNullOrEmpty(latest) && LatestJobEvDateBox.Text != latest) LatestJobEvDateBox.Text = latest;
            }
            finally { _syncingTracker = false; }
            UpdateExpirationHints();
        }

        /// <summary>Boxes to model: copies complete (or emptied) File Dates boxes into the tracker.</summary>
        private void PushBoxesToTracker()
        {
            if (_syncingTracker || _mainViewModel == null) return;
            var t = _mainViewModel.Tracker;
            _syncingTracker = true;
            try
            {
                foreach (var (box, get, set) in TrackerBoxMap())
                {
                    string text = box.Text ?? string.Empty;
                    if ((text.Length == 0 || text.Length == 10) && get(t) != text) set(t, text);
                }
            }
            finally { _syncingTracker = false; }
            UpdateExpirationHints();
        }

        private void UpdateExpirationHints()
        {
            if (_mainViewModel == null) return;
            var t = _mainViewModel.Tracker;
            string Recorded(string value) => "Entered: " + (string.IsNullOrEmpty(value) ? "—" : value);
            DrugTestExpHint.Text = Recorded(t.DtDate);
            MvrRanExpHint.Text = Recorded(t.MvrRan);
            FingerprintExpHint.Text = Recorded(t.FpRan);
            PspExpHint.Text = Recorded(t.PspRan);
            PhysicalLookbackHint.Text = Recorded(t.PhysicalExp);
            LatestJobEvExpHint.Text = Recorded(LatestJobEvDateBox.Text);
        }

        private void MainViewModel_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(MainViewModel.CurrentPspSnapshotDate))
            {
                SyncPspFromViewModel();
            }
        }

        public void SyncPspFromViewModel()
        {
            if (_isSyncingPsp) return;
            _isSyncingPsp = true;
            try
            {
                string text = _mainViewModel?.CurrentPspSnapshotDate?.ToString(FileDateFormat, CultureInfo.InvariantCulture) ?? string.Empty;
                if (PspSnapshotDateBox.Text != text)
                {
                    PspSnapshotDateBox.Text = text;
                }
            }
            finally
            {
                _isSyncingPsp = false;
            }
            RevalidateAllDateFields();
        }

        public WorkspaceWindow()
        {
            InitializeComponent();
            foreach (TextBox dateBox in new[]
            {
                CertificationDateBox, LimitedConsentDateBox, FullConsentDateBox, DrugTestRanDateBox, DrugTestReceivedDateBox,
                CdlExpirationDateBox, MvrCdlExpirationDateBox, MvrLastReviewDateBox, MvrNextReviewDateBox,
                PhysicalLastDateBox, PhysicalNextDateBox, FingerprintDateBox, HazmatExpirationDateBox,
                LatestJobEvDateBox, SevenDayLogDateBox, StatementOfLeaseDateBox, StatementOfLeaseTrailerDateBox, DayQualifiedDateBox, PspSnapshotDateBox, PspRanDateBox, CdlisDacDateBox
            })
            {
                dateBox.MaxLength = 10;
            }

            Services.CopilotContextBuilder.TrackerContextProvider = () => _mainViewModel?.Tracker.ToPlainText();
            TrackerModel.AsOf = () => Services.CopilotContextBuilder.EffectiveToday;
            ItemsHost.ItemsSource = _items;
            DrawerItemsControl.ItemsSource = _drawerItems;
            _items.CollectionChanged += (_, _) => RefreshDrawerList();

            _wordHostPanel = new Win32PanelHost();
            _wordHostPanel.SizeOrPositionChanged += (_, _) => _wordController?.ResizeToFillContainer();
            WordDocHostContainer.Children.Add(_wordHostPanel);

            _excelHostPanel = new Win32PanelHost();
            _excelHostPanel.SizeOrPositionChanged += (_, _) => _excelController?.ResizeToFillContainer();
            ExcelSheetHostContainer.Children.Add(_excelHostPanel);

            _autosaveTimer = new DispatcherTimer { Interval = AutosaveInterval };
            _autosaveTimer.Tick += AutosaveTimer_Tick;
            _autosaveTimer.Start();

            AutosaveStatusText.Text = "Autosave: every 30 min";

            _quickDateTimer = new DispatcherTimer { Interval = QuickDateRefreshInterval };
            _quickDateTimer.Tick += (_, _) => RefreshQuickDateBar();
            _quickDateTimer.Start();
            RefreshQuickDateBar();


            RevalidateAllDateFields();
        }

        // ======================= Quick date reference bar =======================

        /// <summary>
        /// Recomputes the bottom quick-reference bar (30/60/90-day and
        /// 1/3/5/10-year milestones, plus the "Backdate" box) from the
        /// current system date. Cheap to call on a timer tick; only touches
        /// the UI when the calendar date has actually rolled over since the
        /// last computation (the toggle button forces a recompute regardless).
        /// </summary>
        private void RefreshQuickDateBar(bool forceRecompute = false)
        {
            DateTime today = DateTime.Today;
            if (!forceRecompute && today == _quickDatesComputedFor)
            {
                return;
            }

            _quickDatesComputedFor = today;

            // Each milestone's label plus its "future" (add) and "past" (subtract) date functions.
            (string Label, Func<DateTime, DateTime> Future, Func<DateTime, DateTime> Past)[] futureOrder =
            {
                ("30 DAYS", d => d.AddDays(30), d => d.AddDays(-30)),
                ("60 DAYS", d => d.AddDays(60), d => d.AddDays(-60)),
                ("90 DAYS", d => d.AddDays(90), d => d.AddDays(-90)),
                ("1 YEAR", d => d.AddYears(1), d => d.AddYears(-1)),
                ("3 YEARS", d => d.AddYears(3), d => d.AddYears(-3)),
                ("5 YEARS", d => d.AddYears(5), d => d.AddYears(-5)),
                ("10 YEARS", d => d.AddYears(10), d => d.AddYears(-10))
            };

            // "Past" mode subtracts instead of adds, and reverses the display order so the oldest date is first; "future" mode keeps the natural youngest-first order.
            IEnumerable<(string Label, Func<DateTime, DateTime> Compute)> milestones = _quickDatesLookingBackward
                ? futureOrder.Reverse().Select(m => (m.Label, m.Past))
                : futureOrder.Select(m => (m.Label, m.Future));

            TextBlock[] labels = { Milestone1Label, Milestone2Label, Milestone3Label, Milestone4Label, Milestone5Label, Milestone6Label, Milestone7Label };
            TextBlock[] values = { Milestone1Text, Milestone2Text, Milestone3Text, Milestone4Text, Milestone5Text, Milestone6Text, Milestone7Text };

            int i = 0;
            foreach ((string label, Func<DateTime, DateTime> compute) in milestones)
            {
                labels[i].Text = label;
                values[i].Text = compute(today).ToString(QuickDateFormat, CultureInfo.InvariantCulture);
                i++;
            }

            // "Backdate": yesterday, except when today is Monday, in which
            // case it's the previous Friday (so it never lands on a weekend).
            DateTime backdate = ComputeBackdate(today);
            BackdateText.Text = backdate.ToString(QuickDateFormat, CultureInfo.InvariantCulture);
        }

        /// <summary>Flips the quick-date bar between "past" (default, subtract/oldest-first) and "future" (add/youngest-first) modes.</summary>
        private void DateDirectionToggleButton_Click(object sender, RoutedEventArgs e)
        {
            _quickDatesLookingBackward = !_quickDatesLookingBackward;
            DateDirectionToggleText.Text = _quickDatesLookingBackward ? "PAST" : "FUTURE";
            RefreshQuickDateBar(forceRecompute: true);
        }

        // ======================= Driver/contractor header fields =======================

        /// <summary>Only allows digits through (used by the SSN, FID, and Phone# boxes, which apply their own formatting separately).</summary>
        private void DigitsOnlyTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !e.Text.All(char.IsDigit);
        }

        /// <summary>Formats the SSN box as the user types: ###-##-####.</summary>
        private void SsnTextBox_TextChanged(object sender, TextChangedEventArgs e) =>
            ReformatDigitsBox(SsnTextBox, FormatSsn);

        /// <summary>Formats the phone box as the user types: (###) ###-####.</summary>
        private void PhoneTextBox_TextChanged(object sender, TextChangedEventArgs e) =>
            ReformatDigitsBox(PhoneTextBox, FormatPhone);

        /// <summary>Formats the FID box as the user types: XX-XXXXXXX.</summary>
        private void FidTextBox_TextChanged(object sender, TextChangedEventArgs e) =>
            ReformatDigitsBox(FidTextBox, FormatFid);

        /// <summary>
        /// Re-applies a digit-formatting function to a text box's content
        /// without re-entering TextChanged recursively, and keeps the caret
        /// at the end (acceptable for these short, fixed-pattern fields).
        /// </summary>
        private static void ReformatDigitsBox(TextBox textBox, Func<string, string> formatter)
        {
            string digitsOnly = new(textBox.Text.Where(char.IsDigit).ToArray());
            string formatted = formatter(digitsOnly);
            if (formatted == textBox.Text)
            {
                return;
            }

            textBox.Text = formatted;
            textBox.CaretIndex = textBox.Text.Length;
        }

        private static string FormatSsn(string digits)
        {
            digits = digits.Length > 9 ? digits[..9] : digits;
            return digits.Length switch
            {
                <= 3 => digits,
                <= 5 => $"{digits[..3]}-{digits[3..]}",
                _ => $"{digits[..3]}-{digits[3..5]}-{digits[5..]}"
            };
        }

        private static string FormatPhone(string digits)
        {
            digits = digits.Length > 10 ? digits[..10] : digits;
            return digits.Length switch
            {
                <= 3 => digits,
                <= 6 => $"({digits[..3]}) {digits[3..]}",
                _ => $"({digits[..3]}) {digits[3..6]}-{digits[6..]}"
            };
        }

        private static string FormatFid(string digits)
        {
            digits = digits.Length > 9 ? digits[..9] : digits;
            return digits.Length switch
            {
                <= 2 => digits,
                _ => $"{digits[..2]}-{digits[2..]}"
            };
        }

        /// <summary>Shows/hides the collapsible File Details row and flips the toggle button's label.</summary>
        private void FileDetailsToggleButton_Click(object sender, RoutedEventArgs e)
        {
            bool expanding = FileDetailsPanel.Visibility != Visibility.Visible;
            FileDetailsPanel.Visibility = expanding ? Visibility.Visible : Visibility.Collapsed;
            FileDetailsToggleButton.Content = expanding ? "Hide File Details" : "View/Edit File Details";
        }

        /// <summary>Snapshots the current header text boxes into a serializable record.</summary>
        private WorkspaceHeaderInfo BuildHeaderInfo() => new()
        {
            TruckNumber = TruckNumberTextBox.Text,
            TrailerNumber = TrailerNumberTextBox.Text,
            ContractorName = ContractorNameTextBox.Text,
            DriverName = DriverNameTextBox.Text,
            Ssn = SsnTextBox.Text,
            Fid = FidTextBox.Text,
            Phone = PhoneTextBox.Text
        };

        /// <summary>Restores the header text boxes from a loaded (or freshly cleared) record.</summary>
        private void ApplyHeaderInfo(WorkspaceHeaderInfo header)
        {
            TruckNumberTextBox.Text = header.TruckNumber;
            TrailerNumberTextBox.Text = header.TrailerNumber;
            ContractorNameTextBox.Text = header.ContractorName;
            DriverNameTextBox.Text = header.DriverName;
            SsnTextBox.Text = header.Ssn;
            FidTextBox.Text = header.Fid;
            PhoneTextBox.Text = header.Phone;
        }

        // ======================= File Dates rules-engine panel =======================

        private const string FileDateFormat = "MM/dd/yyyy";

        /// <summary>Every red-flag reason found by the most recent <see cref="RevalidateAllDateFields"/> pass, shown by the "Conflict List" button.</summary>
        private readonly List<string> _dateFieldConflicts = new();

        /// <summary>Shows/hides the collapsible File Dates panel and flips the toggle button's label.</summary>
        private void FileDatesToggleButton_Click(object sender, RoutedEventArgs e)
        {
            bool expanding = FileDatesPanel.Visibility != Visibility.Visible;
            FileDatesPanel.Height = Math.Clamp(_fileDatesPanelHeight, 70, Math.Max(70, ActualHeight - 220));
            FileDatesPanel.Visibility = expanding ? Visibility.Visible : Visibility.Collapsed;
            FileDatesToggleButton.Content = expanding ? "Hide File Dates" : "File Dates";
        }

        private double _fileDatesPanelHeight = 270;

        /// <summary>Lets the user drag the top edge of the File Dates panel to resize it vertically (it scrolls when squeezed small).</summary>
        private void FileDatesPanelSplitter_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            double maxHeight = Math.Max(70, ActualHeight - 220);
            double newHeight = Math.Clamp(FileDatesPanel.ActualHeight - e.VerticalChange, 70, maxHeight);
            FileDatesPanel.Height = newHeight;
            _fileDatesPanelHeight = newHeight;
        }

        /// <summary>Progressively masks digits-only input into MM/dd/yyyy as the user types (any manually-typed "/" is stripped by <see cref="ReformatDigitsBox"/> before this runs).</summary>
        private static string FormatDateMmDdYy(string digits)
        {
            digits = digits.Length > 8 ? digits[..8] : digits;
            return digits.Length switch
            {
                <= 2 => digits,
                <= 4 => $"{digits[..2]}/{digits[2..]}",
                _ => $"{digits[..2]}/{digits[2..4]}/{digits[4..]}"
            };
        }

        /// <summary>Shared handler for every File Dates box except MVR Last Review (which also auto-populates MVR Next Review below).</summary>
        private bool _syncingAuditDate;

        private void AuditDateBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_syncingAuditDate) return;
            ReformatDigitsBox(AuditDateBox, FormatDateMmDdYy);
            _syncingAuditDate = true;
            try { if (DayQualifiedDateBox.Text != AuditDateBox.Text) DayQualifiedDateBox.Text = AuditDateBox.Text; }
            finally { _syncingAuditDate = false; }
            RevalidateAllDateFields();
        }

        private void DateFieldTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ReformatDigitsBox((TextBox)sender, FormatDateMmDdYy);
            if (ReferenceEquals(sender, DayQualifiedDateBox) && !_syncingAuditDate)
            {
                _syncingAuditDate = true;
                try { if (AuditDateBox.Text != DayQualifiedDateBox.Text) AuditDateBox.Text = DayQualifiedDateBox.Text; }
                finally { _syncingAuditDate = false; }
            }
            RevalidateAllDateFields();
        }

        /// <summary>Reformats the MVR review date and refreshes date readability states.</summary>
        private void MvrLastReviewDateBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ReformatDigitsBox(MvrLastReviewDateBox, FormatDateMmDdYy);
            RevalidateAllDateFields();
        }

        private void FingerprintDateBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ReformatDigitsBox(FingerprintDateBox, FormatDateMmDdYy);
            RevalidateAllDateFields();
        }

        private void PspSnapshotDateBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ReformatDigitsBox(PspSnapshotDateBox, FormatDateMmDdYy);

            if (!_isSyncingPsp)
            {
                _isSyncingPsp = true;
                try
                {
                    if (TryParseFileDate(PspSnapshotDateBox.Text, out DateTime dt))
                    {
                        if (_mainViewModel != null && _mainViewModel.CurrentPspSnapshotDate != dt)
                        {
                            _mainViewModel.CurrentPspSnapshotDate = dt;
                        }
                    }
                    else if (string.IsNullOrWhiteSpace(PspSnapshotDateBox.Text))
                    {
                        if (_mainViewModel != null && _mainViewModel.CurrentPspSnapshotDate != null)
                        {
                            _mainViewModel.CurrentPspSnapshotDate = null;
                        }
                    }
                }
                finally
                {
                    _isSyncingPsp = false;
                }
            }

            RevalidateAllDateFields();
        }

        /// <summary>Opens a simple popup listing every active red-flag reason found by the last validation pass.</summary>
        private void ConflictListButton_Click(object sender, RoutedEventArgs e)
        {
            string body = _dateFieldConflicts.Count == 0
                ? "No conflicts - every entered File Dates field currently passes validation."
                : string.Join("\n\n", _dateFieldConflicts.Select((c, i) => $"{i + 1}. {c}"));

            MessageBox.Show(this, body, "Conflict List", MessageBoxButton.OK,
                _dateFieldConflicts.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private static bool TryParseFileDate(string text, out DateTime date) =>
            DateTime.TryParseExact(text, FileDateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);

        private static DateTime? ParseFileDateOrNull(string text) =>
            TryParseFileDate(text, out DateTime date) ? date : null;

        /// <summary>"Yesterday", except when today is Monday, in which case the previous Friday (so it never lands on a weekend) - shared by the Quick Date bar and the File Dates rules below.</summary>
        private static DateTime ComputeBackdate(DateTime today) =>
            today.DayOfWeek == DayOfWeek.Monday ? today.AddDays(-3) : today.AddDays(-1);

        /// <summary>Paints a date box Neutral (blank), Green (readable date), or Red (invalid date input).</summary>
        private void ApplyFieldState(TextBox box, string label, DateTime? value, bool isRed, string reason)
        {
            if (value is null && !isRed)
            {
                box.ClearValue(TextBox.BackgroundProperty);
                box.ClearValue(TextBox.ForegroundProperty);
                box.ToolTip = null;
                return;
            }

            box.Background = (System.Windows.Media.Brush)FindResource(isRed ? "DateFieldRedBrush" : "DateFieldGreenBrush");
            box.Foreground = (System.Windows.Media.Brush)FindResource("WhiteBrush");
            box.ToolTip = isRed ? reason : "Readable date only; review criteria are evaluated separately.";

            if (isRed)
            {
                _dateFieldConflicts.Add($"{label}: {reason}");
            }
        }

        /// <summary>
        /// Re-evaluates every File Dates field against the qualification rules
        /// spec, coloring each box Neutral/Green/Red and rebuilding the
        /// Conflict List. Only fields that contain a value are evaluated;
        /// blank fields are always left Neutral.
        /// </summary>
        private void RevalidateAllDateFields()
        {
            PushBoxesToTracker();
            DateTime? auditDate = _mode == QualMode.PostReviewAudit ? ParseFileDateOrNull(DayQualifiedDateBox.Text) : null;
            CopilotContextBuilder.AuditDate = auditDate;
            _mainViewModel?.Tracker.Refresh();
            DateTime today = auditDate ?? DateTime.Today;
            FileDatesTodayLabel.Text = auditDate.HasValue ? "AUDIT AS OF:" : "TODAY:";
            FileDatesCurrentDateText.Text = today.ToString(FileDateFormat, CultureInfo.InvariantCulture);
            _dateFieldConflicts.Clear();

            (TextBox Box, string Label)[] fields =
            {
                (CertificationDateBox, "Certification"),
                (LimitedConsentDateBox, "Limited Consent"),
                (FullConsentDateBox, "Full Consent"),
                (DrugTestRanDateBox, "Drug Test Ran"),
                (DrugTestReceivedDateBox, "Drug Test Received"),
                (CdlExpirationDateBox, "CDL Expiration"),
                (MvrCdlExpirationDateBox, "MVR CDL Expiration"),
                (HazmatExpirationDateBox, "Hazmat Expiration"),
                (FingerprintDateBox, "Fingerprint Date"),
                (MvrLastReviewDateBox, "MVR Last Review"),
                (MvrNextReviewDateBox, "MVR Next Review"),
                (PhysicalLastDateBox, "Physical Last"),
                (PhysicalNextDateBox, "Physical Next"),
                (LatestJobEvDateBox, "Latest Job Verification"),
                (SevenDayLogDateBox, "7-Day Log"),
                (StatementOfLeaseDateBox, "Statement of Lease"),
                (StatementOfLeaseTrailerDateBox, "Trailer Statement of Lease"),
                (DayQualifiedDateBox, "Day Qualified"),
                (PspSnapshotDateBox, "PSP Snapshot"),
                (PspRanDateBox, "PSP Run Date"),
                (CdlisDacDateBox, "CDLIS / DAC Date")
            };

            foreach ((TextBox box, string label) in fields)
            {
                if (string.IsNullOrWhiteSpace(box.Text))
                {
                    ApplyFieldState(box, label, null, false, string.Empty);
                    continue;
                }

                DateTime? value = ParseFileDateOrNull(box.Text);
                bool invalid = value is null;
                ApplyFieldState(box, label, value, invalid, invalid ? "Enter a valid date in MM/dd/yyyy format." : string.Empty);
            }
        }
        /// <summary>Snapshots the current File Dates boxes into a serializable record.</summary>
        private WorkspaceDateFieldsInfo BuildDateFieldsInfo() => new()
        {
            Certification = ParseFileDateOrNull(CertificationDateBox.Text),
            LimitedConsent = ParseFileDateOrNull(LimitedConsentDateBox.Text),
            FullConsent = ParseFileDateOrNull(FullConsentDateBox.Text),
            DrugTestRan = ParseFileDateOrNull(DrugTestRanDateBox.Text),
            DrugTestReceived = ParseFileDateOrNull(DrugTestReceivedDateBox.Text),
            CdlExpiration = ParseFileDateOrNull(CdlExpirationDateBox.Text),
            MvrCdlExpiration = ParseFileDateOrNull(MvrCdlExpirationDateBox.Text),
            HazmatExpiration = ParseFileDateOrNull(HazmatExpirationDateBox.Text),
            FingerprintDate = ParseFileDateOrNull(FingerprintDateBox.Text),
            MvrLastReview = ParseFileDateOrNull(MvrLastReviewDateBox.Text),
            MvrNextReview = ParseFileDateOrNull(MvrNextReviewDateBox.Text),
            PhysicalLast = ParseFileDateOrNull(PhysicalLastDateBox.Text),
            PhysicalNext = ParseFileDateOrNull(PhysicalNextDateBox.Text),
            LatestJobEv = ParseFileDateOrNull(LatestJobEvDateBox.Text),
            SevenDayLog = ParseFileDateOrNull(SevenDayLogDateBox.Text),
            StatementOfLease = ParseFileDateOrNull(StatementOfLeaseDateBox.Text),
            StatementOfLeaseTrailer = ParseFileDateOrNull(StatementOfLeaseTrailerDateBox.Text),
            DayQualified = ParseFileDateOrNull(DayQualifiedDateBox.Text),
            PspSnapshotDate = ParseFileDateOrNull(PspSnapshotDateBox.Text)
        };

        /// <summary>Restores the File Dates boxes from a loaded (or freshly cleared) record, then revalidates.</summary>
        private void ApplyDateFieldsInfo(WorkspaceDateFieldsInfo fields)
        {
            static string Fmt(DateTime? d) => d is { } v ? v.ToString(FileDateFormat, CultureInfo.InvariantCulture) : string.Empty;

            CertificationDateBox.Text = Fmt(fields.Certification);
            LimitedConsentDateBox.Text = Fmt(fields.LimitedConsent);
            FullConsentDateBox.Text = Fmt(fields.FullConsent);
            DrugTestRanDateBox.Text = Fmt(fields.DrugTestRan);
            DrugTestReceivedDateBox.Text = Fmt(fields.DrugTestReceived);
            CdlExpirationDateBox.Text = Fmt(fields.CdlExpiration);
            MvrCdlExpirationDateBox.Text = Fmt(fields.MvrCdlExpiration);
            HazmatExpirationDateBox.Text = Fmt(fields.HazmatExpiration);
            FingerprintDateBox.Text = Fmt(fields.FingerprintDate);
            MvrLastReviewDateBox.Text = Fmt(fields.MvrLastReview);
            MvrNextReviewDateBox.Text = Fmt(fields.MvrNextReview);
            PhysicalLastDateBox.Text = Fmt(fields.PhysicalLast);
            PhysicalNextDateBox.Text = Fmt(fields.PhysicalNext);
            LatestJobEvDateBox.Text = Fmt(fields.LatestJobEv);
            SevenDayLogDateBox.Text = Fmt(fields.SevenDayLog);
            StatementOfLeaseDateBox.Text = Fmt(fields.StatementOfLease);
            StatementOfLeaseTrailerDateBox.Text = Fmt(fields.StatementOfLeaseTrailer);
            DayQualifiedDateBox.Text = Fmt(fields.DayQualified);

            if (fields.PspSnapshotDate != null)
            {
                PspSnapshotDateBox.Text = Fmt(fields.PspSnapshotDate);
            }
            else if (_mainViewModel?.CurrentPspSnapshotDate != null)
            {
                PspSnapshotDateBox.Text = _mainViewModel.CurrentPspSnapshotDate.Value.ToString(FileDateFormat, CultureInfo.InvariantCulture);
            }
            else
            {
                PspSnapshotDateBox.Text = string.Empty;
            }

            RevalidateAllDateFields();
        }

        // ======================= Embedded Word document viewer (real Microsoft Word, embedded) =======================

        /// <summary>Shared by both the centered "empty state" button and the top-bar button once a document is loaded.</summary>
        private void LoadWordDocButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Load Word Document",
                Filter = "Word Document (*.docx)|*.docx|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                byte[] bytes = File.ReadAllBytes(dialog.FileName);
                LoadWordDocSession(bytes, Path.GetFileName(dialog.FileName));
                RefreshDrawerList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or WordNotAvailableException)
            {
                MessageBox.Show(this, $"Could not open the Word document:\n{ex.Message}", "Load Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveWordDocButton_Click(object sender, RoutedEventArgs e)
        {
            if (_wordController is not { IsOpen: true })
            {
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "Save Word Document",
                Filter = "Word Document (*.docx)|*.docx",
                DefaultExt = ".docx",
                FileName = string.IsNullOrWhiteSpace(_wordDocFileName) ? "Document.docx" : _wordDocFileName
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                _wordController.ExportTo(dialog.FileName);
                _wordDocFileName = Path.GetFileName(dialog.FileName);
                MessageBox.Show(this, "Word document saved successfully.", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
                RefreshDrawerList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or COMException)
            {
                MessageBox.Show(this, $"Could not save the Word document:\n{ex.Message}", "Save Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>Pops the Word document viewer out into its own maximized window (or re-activates it if already popped out). Word's real window is simply reparented onto the popout window's native docking slot - the same live session continues uninterrupted.</summary>
        private void WordDocPopOutButton_Click(object sender, RoutedEventArgs e)
        {
            if (_wordController is not { IsOpen: true })
            {
                return;
            }

            if (_wordDocPopout is not null)
            {
                _wordDocPopout.Activate();
                return;
            }

            _wordDocPopout = new WordDocPopoutWindow { Owner = this };
            _wordDocPopout.PopInRequested += WordDocPopout_PopInRequested;
            _wordDocPopout.Closed += WordDocPopout_Closed;
            _wordDocPopout.HostPanel.SizeOrPositionChanged += (_, _) => _wordController?.ResizeToFillContainer();
            _wordDocPopout.Show();

            // Must attach after Show() so the popout's native container handle actually exists.
            _wordController.AttachTo(_wordDocPopout.HostPanel.Handle);
        }

        private void WordDocPopout_PopInRequested(object? sender, EventArgs e)
        {
            _wordDocPopout?.Close();
        }

        private void WordDocPopout_Closed(object? sender, EventArgs e)
        {
            if (_wordDocPopout is null)
            {
                return;
            }

            _wordDocPopout.PopInRequested -= WordDocPopout_PopInRequested;
            _wordDocPopout.Closed -= WordDocPopout_Closed;
            _wordDocPopout = null;

            if (_wordController is { IsOpen: true } && _wordHostPanel is not null)
            {
                _wordController.AttachTo(_wordHostPanel.Handle);
            }
        }

        private void WordDocZoomInButton_Click(object sender, RoutedEventArgs e) => AdjustWordDocZoom(WordDocZoomStepPercent);

        private void WordDocZoomOutButton_Click(object sender, RoutedEventArgs e) => AdjustWordDocZoom(-WordDocZoomStepPercent);

        private void AdjustWordDocZoom(int deltaPercent)
        {
            if (_wordController is not { IsOpen: true })
            {
                return;
            }

            _wordController.SetZoomPercent(_wordController.GetZoomPercent() + deltaPercent);
        }

        /// <summary>True once a Word document has been loaded into this session (center button hidden, top bar/viewer shown).</summary>
        private bool HasWordDocLoaded => WordDocLoadedPanel.Visibility == Visibility.Visible;

        /// <summary>Snapshots the embedded Word document (if any) for persistence with the workspace file - has the live Word session save itself, then reads back the resulting bytes.</summary>
        private byte[]? BuildWordDocBytes() => _wordController?.SaveAndGetBytes();

        /// <summary>Opens the given .docx bytes in a real, embedded Microsoft Word instance and shows it in the viewer.</summary>
        private void LoadWordDocSession(byte[] bytes, string fileName)
        {
            if (!EmbeddedWordController.IsWordInstalled)
            {
                throw new WordNotAvailableException("Microsoft Word does not appear to be installed on this computer. Word must be installed to view/edit Word documents.");
            }

            _wordDocPopout?.Close();
            _wordController?.Dispose();
            _wordController = new EmbeddedWordController();
            _wordController.Open(bytes, _wordHostPanel!.Handle);

            _wordDocFileName = fileName;
            WordDocEmptyStatePanel.Visibility = Visibility.Collapsed;
            WordDocLoadedPanel.Visibility = Visibility.Visible;
        }

        /// <summary>Creates a brand-new, blank document in a real, embedded Microsoft Word instance and shows it in the viewer.</summary>
        private void NewWordDocButton_Click(object sender, RoutedEventArgs e)
        {
            if (HasWordDocLoaded)
            {
                var confirm = MessageBox.Show(
                    this,
                    "Starting a new document will close the one currently open here. Save it first if you want to keep it. Continue?",
                    "New Word Document",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning,
                    MessageBoxResult.No);

                if (confirm != MessageBoxResult.Yes)
                {
                    return;
                }
            }

            try
            {
                if (!EmbeddedWordController.IsWordInstalled)
                {
                    throw new WordNotAvailableException("Microsoft Word does not appear to be installed on this computer. Word must be installed to view/edit Word documents.");
                }

                _wordDocPopout?.Close();
                _wordController?.Dispose();
                _wordController = new EmbeddedWordController();
                _wordController.New(_wordHostPanel!.Handle);

                _wordDocFileName = "New Document.docx";
                WordDocEmptyStatePanel.Visibility = Visibility.Collapsed;
                WordDocLoadedPanel.Visibility = Visibility.Visible;
                RefreshDrawerList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or WordNotAvailableException or COMException)
            {
                MessageBox.Show(this, $"Could not create a new Word document:\n{ex.Message}", "New Document Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>Restores (or clears) the embedded Word document from a loaded/reset workspace file.</summary>
        private void ApplyWordDoc(byte[]? bytes, string fileName)
        {
            if (bytes is { Length: > 0 })
            {
                try
                {
                    LoadWordDocSession(bytes, fileName);
                    RefreshDrawerList();
                    return;
                }
                catch (WordNotAvailableException ex)
                {
                    MessageBox.Show(this, $"This workspace contains a Word document, but it could not be reopened:\n{ex.Message}", "Word Not Available", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }

            _wordDocPopout?.Close();
            _wordController?.Dispose();
            _wordController = null;
            _wordDocFileName = string.Empty;
            WordDocLoadedPanel.Visibility = Visibility.Collapsed;
            WordDocEmptyStatePanel.Visibility = Visibility.Visible;
            RefreshDrawerList();
        }

        // ======================= Embedded Excel sheet viewer (real Microsoft Excel, embedded) =======================

        /// <summary>Shared by both the centered "empty state" button and the top-bar button once a sheet is loaded.</summary>
        private void LoadExcelSheetButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Load Excel Spreadsheet",
                Filter = "Excel Files (*.xlsx;*.xls;*.xlsm;*.csv)|*.xlsx;*.xls;*.xlsm;*.csv|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                byte[] bytes = File.ReadAllBytes(dialog.FileName);
                LoadExcelSheetSession(bytes, Path.GetFileName(dialog.FileName));
                RefreshDrawerList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ExcelNotAvailableException)
            {
                MessageBox.Show(this, $"Could not open the Excel spreadsheet:\n{ex.Message}", "Load Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveExcelSheetButton_Click(object sender, RoutedEventArgs e)
        {
            if (_excelController is not { IsOpen: true })
            {
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "Save Excel Spreadsheet",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 Workbook (*.xls)|*.xls",
                DefaultExt = ".xlsx",
                FileName = string.IsNullOrWhiteSpace(_excelDocFileName) ? "Workbook.xlsx" : _excelDocFileName
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                _excelController.ExportTo(dialog.FileName);
                _excelDocFileName = Path.GetFileName(dialog.FileName);
                MessageBox.Show(this, "Excel spreadsheet saved successfully.", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
                RefreshDrawerList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or COMException)
            {
                MessageBox.Show(this, $"Could not save the Excel spreadsheet:\n{ex.Message}", "Save Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>Pops the Excel sheet viewer out into its own maximized window (or re-activates it if already popped out). Excel's real window is reparented onto the popout window's native docking slot - the same live session continues uninterrupted.</summary>
        private void ExcelPopOutButton_Click(object sender, RoutedEventArgs e)
        {
            if (_excelController is not { IsOpen: true })
            {
                return;
            }

            if (_excelSheetPopout is not null)
            {
                _excelSheetPopout.Activate();
                return;
            }

            _excelSheetPopout = new ExcelSheetPopoutWindow { Owner = this };
            _excelSheetPopout.PopInRequested += ExcelSheetPopout_PopInRequested;
            _excelSheetPopout.Closed += ExcelSheetPopout_Closed;
            _excelSheetPopout.HostPanel.SizeOrPositionChanged += (_, _) => _excelController?.ResizeToFillContainer();
            _excelSheetPopout.Show();

            _excelController.AttachTo(_excelSheetPopout.HostPanel.Handle);
        }

        private void ExcelSheetPopout_PopInRequested(object? sender, EventArgs e)
        {
            _excelSheetPopout?.Close();
        }

        private void ExcelSheetPopout_Closed(object? sender, EventArgs e)
        {
            if (_excelSheetPopout is null)
            {
                return;
            }

            _excelSheetPopout.PopInRequested -= ExcelSheetPopout_PopInRequested;
            _excelSheetPopout.Closed -= ExcelSheetPopout_Closed;
            _excelSheetPopout = null;

            if (_excelController is { IsOpen: true } && _excelHostPanel is not null)
            {
                _excelController.AttachTo(_excelHostPanel.Handle);
            }
        }

        private void ExcelZoomInButton_Click(object sender, RoutedEventArgs e) => AdjustExcelZoom(ExcelZoomStepPercent);

        private void ExcelZoomOutButton_Click(object sender, RoutedEventArgs e) => AdjustExcelZoom(-ExcelZoomStepPercent);

        private void AdjustExcelZoom(int deltaPercent)
        {
            if (_excelController is not { IsOpen: true })
            {
                return;
            }

            _excelController.SetZoomPercent(_excelController.GetZoomPercent() + deltaPercent);
        }

        /// <summary>True once an Excel spreadsheet has been loaded into this session (center button hidden, top bar/viewer shown).</summary>
        private bool HasExcelLoaded => ExcelLoadedPanel.Visibility == Visibility.Visible;

        /// <summary>Snapshots the embedded Excel workbook (if any) for persistence with the workspace file - has the live Excel session save itself, then reads back the resulting bytes.</summary>
        private byte[]? BuildExcelDocBytes() => _excelController?.SaveAndGetBytes();

        /// <summary>Opens the given spreadsheet bytes in a real, embedded Microsoft Excel instance and shows it in the viewer.</summary>
        private void LoadExcelSheetSession(byte[] bytes, string fileName)
        {
            if (!EmbeddedExcelController.IsExcelInstalled)
            {
                throw new ExcelNotAvailableException("Microsoft Excel does not appear to be installed on this computer. Excel must be installed to view/edit spreadsheets.");
            }

            _excelSheetPopout?.Close();
            _excelController?.Dispose();
            _excelController = new EmbeddedExcelController();
            _excelController.Open(bytes, _excelHostPanel!.Handle, fileName);

            _excelDocFileName = fileName;
            ExcelEmptyStatePanel.Visibility = Visibility.Collapsed;
            ExcelLoadedPanel.Visibility = Visibility.Visible;
        }

        /// <summary>Creates a brand-new, blank workbook in a real, embedded Microsoft Excel instance and shows it in the viewer.</summary>
        private void NewExcelSheetButton_Click(object sender, RoutedEventArgs e)
        {
            if (HasExcelLoaded)
            {
                var confirm = MessageBox.Show(
                    this,
                    "Starting a new spreadsheet will close the one currently open here. Save it first if you want to keep it. Continue?",
                    "New Excel Spreadsheet",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning,
                    MessageBoxResult.No);

                if (confirm != MessageBoxResult.Yes)
                {
                    return;
                }
            }

            try
            {
                if (!EmbeddedExcelController.IsExcelInstalled)
                {
                    throw new ExcelNotAvailableException("Microsoft Excel does not appear to be installed on this computer. Excel must be installed to view/edit spreadsheets.");
                }

                _excelSheetPopout?.Close();
                _excelController?.Dispose();
                _excelController = new EmbeddedExcelController();
                _excelController.New(_excelHostPanel!.Handle);

                _excelDocFileName = "New Spreadsheet.xlsx";
                ExcelEmptyStatePanel.Visibility = Visibility.Collapsed;
                ExcelLoadedPanel.Visibility = Visibility.Visible;
                RefreshDrawerList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ExcelNotAvailableException or COMException)
            {
                MessageBox.Show(this, $"Could not create a new Excel spreadsheet:\n{ex.Message}", "New Spreadsheet Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>Restores (or clears) the embedded Excel workbook from a loaded/reset workspace file.</summary>
        private void ApplyExcelDoc(byte[]? bytes, string fileName)
        {
            if (bytes is { Length: > 0 })
            {
                try
                {
                    LoadExcelSheetSession(bytes, fileName);
                    RefreshDrawerList();
                    return;
                }
                catch (ExcelNotAvailableException ex)
                {
                    MessageBox.Show(this, $"This workspace contains an Excel spreadsheet, but it could not be reopened:\n{ex.Message}", "Excel Not Available", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }

            _excelSheetPopout?.Close();
            _excelController?.Dispose();
            _excelController = null;
            _excelDocFileName = string.Empty;
            ExcelLoadedPanel.Visibility = Visibility.Collapsed;
            ExcelEmptyStatePanel.Visibility = Visibility.Visible;
            RefreshDrawerList();
        }

        // ======================= Document drawer =======================

        /// <summary>A single row in the document drawer - either a canvas item (image/PDF) or the synthetic Word/Excel document entry.</summary>
        private sealed class DrawerEntry : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler? PropertyChanged;
            private bool _isSelected;
            public bool IsSelected
            {
                get => _isSelected;
                set
                {
                    if (_isSelected == value) return;
                    _isSelected = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                }
            }

            /// <summary>Stable identity that survives list rebuilds so selection is preserved.</summary>
            public object Key => (object?)Source ?? (IsWordDoc ? "word" : IsExcelSheet ? "excel" : "tracker");
            public string Icon { get; init; } = string.Empty;
            public string DisplayName { get; init; } = string.Empty;
            public CanvasItemModel? Source { get; init; }
            public bool IsWordDoc { get; init; }
            public bool IsExcelSheet { get; init; }

            /// <summary>True for the pinned, undeletable Dates &amp; Employment History entry.</summary>
            public bool IsTracker { get; init; }
            public bool CanModify => !IsTracker;
            /// <summary>True when the entry's canvas item was hidden via its "X" chrome button; shows the drawer's green restore icon.</summary>
            public bool IsHidden { get; init; }
        }

        private double _storedFilesDrawerWidth = 240;

        /// <summary>Shows/hides the document drawer and (re)populates its list when opening.</summary>
        private void ViewStoredFilesButton_Click(object sender, RoutedEventArgs e)
        {
            bool opening = DrawerHost.Visibility != Visibility.Visible;
            DrawerHost.Visibility = opening ? Visibility.Visible : Visibility.Collapsed;
            DrawerColumn.Width = new GridLength(opening ? _storedFilesDrawerWidth : 0);
            StoredFilesToggleButton.Content = opening ? "Hide Stored Files" : "Stored Files";
            if (opening)
            {
                RefreshDrawerList();
            }
        }

        /// <summary>Lets the user drag-resize the View Stored Files drawer horizontally via its gold splitter handle.</summary>
        private void DrawerSplitter_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            double newWidth = Math.Clamp(DrawerColumn.Width.Value + e.HorizontalChange, 180, 800);
            DrawerColumn.Width = new GridLength(newWidth);
            _storedFilesDrawerWidth = newWidth;
        }

        /// <summary>Rebuilds the drawer's list from the current canvas items plus the Word document (if loaded).</summary>
        private void RefreshDrawerList()
        {
            _drawerItems.Clear();
            _drawerItems.Add(new DrawerEntry
            {
                Icon = "\U0001F9EE",
                DisplayName = "Dates & Employment History (driver-reported, unverified)",
                IsTracker = true
            });
            foreach (CanvasItemModel item in _items)
            {
                _drawerItems.Add(new DrawerEntry
                {
                    Icon = item.Kind == WorkspaceItemKind.Pdf ? "\U0001F4C4" : "\U0001F5BC",
                    DisplayName = item.DisplayName,
                    Source = item,
                    IsHidden = item.IsHidden
                });
            }

            if (HasWordDocLoaded)
            {
                _drawerItems.Add(new DrawerEntry
                {
                    Icon = "\U0001F4DD",
                    DisplayName = string.IsNullOrWhiteSpace(_wordDocFileName) ? "Document.docx" : _wordDocFileName,
                    IsWordDoc = true
                });
            }

            if (HasExcelLoaded)
            {
                _drawerItems.Add(new DrawerEntry
                {
                    Icon = "\U0001F4CA",
                    DisplayName = string.IsNullOrWhiteSpace(_excelDocFileName) ? "Workbook.xlsx" : _excelDocFileName,
                    IsExcelSheet = true
                });
            }

            var live = _drawerItems.Select(d => d.Key).ToHashSet();
            _selectedKeys.RemoveWhere(k => !live.Contains(k));
            ApplySelectionToEntries();
        }

        // ---- Multi-select ----

        private readonly HashSet<object> _selectedKeys = new();
        private object? _selectionAnchor;

        private void ApplySelectionToEntries()
        {
            foreach (var d in _drawerItems)
            {
                d.IsSelected = !d.IsTracker && _selectedKeys.Contains(d.Key);
            }
            int n = _selectedKeys.Count;
            DrawerSelectionText.Text = n == 0 ? "No files selected" : n == 1 ? "1 file selected" : $"{n} files selected";
        }

        private void DrawerRow_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if ((sender as FrameworkElement)?.Tag is not DrawerEntry entry || entry.IsTracker)
            {
                return;
            }

            DrawerPanel.Focus();
            var mods = Keyboard.Modifiers;
            if ((mods & ModifierKeys.Shift) != 0 && _selectionAnchor is not null)
            {
                int a = _drawerItems.ToList().FindIndex(d => d.Key.Equals(_selectionAnchor));
                int b = _drawerItems.IndexOf(entry);
                if (a >= 0 && b >= 0)
                {
                    if ((mods & ModifierKeys.Control) == 0) _selectedKeys.Clear();
                    for (int i = Math.Min(a, b); i <= Math.Max(a, b); i++)
                    {
                        if (!_drawerItems[i].IsTracker) _selectedKeys.Add(_drawerItems[i].Key);
                    }
                }
            }
            else if ((mods & ModifierKeys.Control) != 0)
            {
                if (!_selectedKeys.Remove(entry.Key)) _selectedKeys.Add(entry.Key);
                _selectionAnchor = entry.Key;
            }
            else if (e.ClickCount == 1)
            {
                // Plain click on an already-selected row in a multi-selection keeps it so double-click actions still work on the group.
                if (!(entry.IsSelected && _selectedKeys.Count > 1))
                {
                    _selectedKeys.Clear();
                    _selectedKeys.Add(entry.Key);
                }
                _selectionAnchor = entry.Key;
            }
            ApplySelectionToEntries();
        }

        private void DrawerSelectAll_Click(object sender, RoutedEventArgs e) => SelectAllDrawerEntries();
        private void DrawerPanel_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            for (var d = e.OriginalSource as DependencyObject; d != null && !ReferenceEquals(d, DrawerPanel); d = d is System.Windows.Media.Visual or System.Windows.Media.Media3D.Visual3D ? System.Windows.Media.VisualTreeHelper.GetParent(d) : LogicalTreeHelper.GetParent(d))
            {
                if (d is System.Windows.Controls.Primitives.ButtonBase || (d is Border { Tag: DrawerEntry }))
                {
                    return;
                }
            }

            if (_selectedKeys.Count > 0)
            {
                _selectedKeys.Clear();
                _selectionAnchor = null;
                ApplySelectionToEntries();
            }
        }

        private void DrawerSelectNone_Click(object sender, RoutedEventArgs e)
        {
            _selectedKeys.Clear();
            ApplySelectionToEntries();
        }

        private void SelectAllDrawerEntries()
        {
            _selectedKeys.Clear();
            foreach (var d in _drawerItems.Where(d => !d.IsTracker)) _selectedKeys.Add(d.Key);
            ApplySelectionToEntries();
        }

        /// <summary>The entries an action should apply to: the whole selection if the clicked entry is part of a multi-selection, otherwise just that entry.</summary>
        private List<DrawerEntry> GetActionTargets(DrawerEntry clicked)
        {
            if (clicked.IsSelected && _selectedKeys.Count > 1)
            {
                return _drawerItems.Where(d => d.IsSelected && !d.IsTracker).ToList();
            }
            return new List<DrawerEntry> { clicked };
        }

        private void DrawerPanel_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.OriginalSource is TextBox) return;
            bool ctrl = (Keyboard.Modifiers & ModifierKeys.Control) != 0;
            if (ctrl && e.Key == Key.A) { SelectAllDrawerEntries(); e.Handled = true; }
            else if (ctrl && e.Key == Key.S) { DrawerSaveSelectedButton_Click(this, new RoutedEventArgs()); e.Handled = true; }
            else if (ctrl && e.Key == Key.O) { AddFileButton_Click(this, new RoutedEventArgs()); e.Handled = true; }
            else if (e.Key == Key.Escape) { DrawerSelectNone_Click(this, new RoutedEventArgs()); e.Handled = true; }
            else if (e.Key == Key.Delete)
            {
                var targets = _drawerItems.Where(d => d.IsSelected && !d.IsTracker).ToList();
                if (targets.Count > 0) { RemoveDrawerEntries(targets); e.Handled = true; }
            }
        }

        private void DrawerPanel_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
            {
                AddFilesToCanvas(files);
            }
            e.Handled = true;
        }

        private static readonly string[] SupportedExtensions = { ".png", ".jpg", ".jpeg", ".pdf" };

        private void AddFileButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Add Files to Workspace",
                Multiselect = true,
                CheckFileExists = true,
                Filter = "Supported files (*.pdf;*.png;*.jpg;*.jpeg)|*.pdf;*.png;*.jpg;*.jpeg|PDF documents (*.pdf)|*.pdf|Images (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg",
                InitialDirectory = KnownFolders.GetDownloadsFolderPath()
            };
            if (dialog.ShowDialog(this) == true)
            {
                AddFilesToCanvas(dialog.FileNames);
            }
        }

        /// <summary>Adds files (staggered so they don't stack exactly), skipping unsupported types, and selects the newly added ones.</summary>
        private void AddFilesToCanvas(IEnumerable<string> files)
        {
            var failures = new List<string>();
            int offset = 0;
            var before = new HashSet<CanvasItemModel>(_items);
            foreach (string file in files)
            {
                if (!SupportedExtensions.Contains(Path.GetExtension(file).ToLowerInvariant()))
                {
                    failures.Add($"{Path.GetFileName(file)}: unsupported file type");
                    continue;
                }
                try
                {
                    AddFileToCanvas(file, new Point(40 + offset, 40 + offset));
                    offset += 24;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
                {
                    failures.Add($"{Path.GetFileName(file)}: {ex.Message}");
                }
            }

            _selectedKeys.Clear();
            foreach (var item in _items.Where(i => !before.Contains(i))) _selectedKeys.Add(item);
            RefreshDrawerList();

            if (failures.Count > 0)
            {
                MessageBox.Show(this, "The following file(s) could not be added:\n\n" + string.Join("\n", failures),
                    "Unable to Add File", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>Double-click an icon in the drawer to save that single file to disk.</summary>
        private void DrawerIcon_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount != 2 || (sender as FrameworkElement)?.Tag is not DrawerEntry entry)
            {
                return;
            }

            if (entry.IsTracker)
            {
                OpenTrackerDrawerButton_Click(sender, e);
                return;
            }

            SaveDrawerEntries(GetActionTargets(entry));
        }

        private static byte[]? GetEntryBytes(DrawerEntry entry, WorkspaceWindow w) => entry.IsWordDoc
            ? w.BuildWordDocBytes()
            : entry.IsExcelSheet ? w.BuildExcelDocBytes() : entry.Source?.FileBytes;

        /// <summary>One entry: Save dialog. Several: choose a folder once; existing names get a numeric suffix instead of being overwritten.</summary>
        private void SaveDrawerEntries(List<DrawerEntry> entries)
        {
            entries = entries.Where(en => !en.IsTracker).ToList();
            if (entries.Count == 0) return;
            if (entries.Count == 1)
            {
                SaveDrawerEntry(entries[0]);
                return;
            }

            var folderDialog = new OpenFolderDialog { Title = $"Choose a folder to save {entries.Count} files into" };
            if (folderDialog.ShowDialog(this) != true) return;

            int saved = 0;
            var errors = new List<string>();
            foreach (var entry in entries)
            {
                try
                {
                    byte[]? bytes = GetEntryBytes(entry, this);
                    if (bytes is null) continue;
                    string name = string.Join("_", entry.DisplayName.Split(Path.GetInvalidFileNameChars()));
                    string dest = Path.Combine(folderDialog.FolderName, name);
                    for (int i = 2; File.Exists(dest); i++)
                    {
                        dest = Path.Combine(folderDialog.FolderName, $"{Path.GetFileNameWithoutExtension(name)} ({i}){Path.GetExtension(name)}");
                    }
                    File.WriteAllBytes(dest, bytes);
                    saved++;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    errors.Add($"{entry.DisplayName}: {ex.Message}");
                }
            }

            string summary = $"Saved {saved} of {entries.Count} file(s).";
            if (errors.Count > 0) summary += "\n\nErrors:\n" + string.Join("\n", errors);
            MessageBox.Show(this, summary, "Save File", MessageBoxButton.OK,
                errors.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private void DrawerSaveSelectedButton_Click(object sender, RoutedEventArgs e)
        {
            var targets = _drawerItems.Where(d => d.IsSelected && !d.IsTracker).ToList();
            if (targets.Count == 0)
            {
                MessageBox.Show(this, "Select one or more files in the drawer first (click, Ctrl+click or Shift+click).", "Save File", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            SaveDrawerEntries(targets);
        }

        /// <summary>Double-click a name in the drawer to rename that file; the change propagates live to the canvas header.</summary>
        private void DrawerName_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount != 2 || (sender as FrameworkElement)?.Tag is not DrawerEntry entry || entry.IsTracker)
            {
                return;
            }

            string extension = Path.GetExtension(entry.DisplayName);
            string baseName = Path.GetFileNameWithoutExtension(entry.DisplayName);

            var prompt = new TextInputWindow("Rename File", "Enter a new file name:", baseName) { Owner = this };
            if (prompt.ShowDialog() != true)
            {
                return;
            }

            string newBaseName = string.Join("_", prompt.ResultText.Split(Path.GetInvalidFileNameChars()));
            string newName = newBaseName + extension;

            if (entry.IsWordDoc)
            {
                _wordDocFileName = newName;
            }
            else if (entry.IsExcelSheet)
            {
                _excelDocFileName = newName;
            }
            else if (entry.Source is not null)
            {
                entry.Source.FilePath = newName; // DisplayName/canvas header both update live from this.
            }

            RefreshDrawerList();
        }

        /// <summary>Paperclip icon in the drawer: attaches this specific file directly into Assistant chat.</summary>
        private async void DrawerAttach_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if ((sender as FrameworkElement)?.Tag is not DrawerEntry entry)
            {
                return;
            }

            e.Handled = true;
            foreach (var target in GetActionTargets(entry))
            {
                await AttachFileToCopilotAsync(target.DisplayName);
            }
        }

        /// <summary>Green restore icon in the drawer: shows a hidden ("X"-closed) item on the canvas again without re-adding it (it was never removed from the workspace).</summary>
        private void DrawerRestore_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if ((sender as FrameworkElement)?.Tag is not DrawerEntry entry || entry.Source is null)
            {
                return;
            }

            e.Handled = true;
            foreach (var target in GetActionTargets(entry))
            {
                if (target.Source is not null) target.Source.IsHidden = false;
            }
            RefreshDrawerList();
        }

        /// <summary>Trash icon in the drawer: removes the file from both the workspace and the render canvas.</summary>
        private void DrawerDelete_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if ((sender as FrameworkElement)?.Tag is not DrawerEntry entry)
            {
                return;
            }

            e.Handled = true;
            RemoveDrawerEntries(GetActionTargets(entry));
        }

        private void RemoveDrawerEntries(List<DrawerEntry> targets)
        {
            targets = targets.Where(t => !t.IsTracker).ToList();
            if (targets.Count == 0) return;

            string prompt = targets.Count == 1
                ? $"Remove \"{targets[0].DisplayName}\" from this workspace?"
                : $"Remove these {targets.Count} files from this workspace?\n\n" + string.Join("\n", targets.Take(8).Select(t => "• " + t.DisplayName)) + (targets.Count > 8 ? "\n…" : "");
            var result = MessageBox.Show(this, prompt, "Confirm Remove", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            foreach (var entry in targets)
            {
                if (entry.IsWordDoc)
                {
                    ApplyWordDoc(null, string.Empty);
                }
                else if (entry.IsExcelSheet)
                {
                    ApplyExcelDoc(null, string.Empty);
                }
                else if (entry.Source is not null)
                {
                    _items.Remove(entry.Source);
                }
            }

            RefreshDrawerList();
        }

        /// <summary>"Save All" at the top of the drawer: prompts once for a destination folder, then writes every file using its current display name.</summary>
        private void DrawerSaveAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (_drawerItems.Count == 0)
            {
                MessageBox.Show(this, "There are no files in this workspace yet.", "Save All", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var folderDialog = new OpenFolderDialog { Title = "Choose a folder to save all workspace files into" };
            if (folderDialog.ShowDialog(this) != true)
            {
                return;
            }

            int saved = 0;
            var errors = new List<string>();
            foreach (DrawerEntry entry in _drawerItems)
            {
                try
                {
                    string destination = Path.Combine(folderDialog.FolderName, entry.DisplayName);
                    byte[]? bytes = entry.IsWordDoc
                        ? BuildWordDocBytes()
                        : entry.IsExcelSheet
                            ? BuildExcelDocBytes()
                            : entry.Source?.FileBytes;
                    if (bytes is null)
                    {
                        continue;
                    }

                    File.WriteAllBytes(destination, bytes);
                    saved++;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    errors.Add($"{entry.DisplayName}: {ex.Message}");
                }
            }

            string summary = $"Saved {saved} of {_drawerItems.Count} file(s).";
            if (errors.Count > 0)
            {
                summary += "\n\nErrors:\n" + string.Join("\n", errors);
            }

            MessageBox.Show(this, summary, "Save All", MessageBoxButton.OK,
                errors.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        /// <summary>Saves a single drawer entry's bytes to disk via a Save dialog defaulted to its current display name.</summary>
        private void SaveDrawerEntry(DrawerEntry entry)
        {
            byte[]? bytes = entry.IsWordDoc
                ? BuildWordDocBytes()
                : entry.IsExcelSheet
                    ? BuildExcelDocBytes()
                    : entry.Source?.FileBytes;
            if (bytes is null)
            {
                MessageBox.Show(this, "This file has no content to save.", "Save", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string extension = Path.GetExtension(entry.DisplayName);
            var dialog = new SaveFileDialog
            {
                Title = $"Save {entry.DisplayName}",
                FileName = entry.DisplayName,
                DefaultExt = extension,
                Filter = extension switch
                {
                    ".pdf" => "PDF Document (*.pdf)|*.pdf",
                    ".docx" => "Word Document (*.docx)|*.docx",
                    ".xlsx" or ".xls" => "Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 Workbook (*.xls)|*.xls",
                    _ => "Image (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg"
                }
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                File.WriteAllBytes(dialog.FileName, bytes);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                MessageBox.Show(this, $"Could not save this file:\n{ex.Message}", "Save Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ======================= Drag & drop ingestion =======================

        private void Canvas_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void Canvas_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            if (e.Data.GetData(DataFormats.FileDrop) is not string[] files)
            {
                return;
            }

            Point dropPoint = e.GetPosition(ItemsHost);
            var failures = new System.Collections.Generic.List<string>();

            foreach (string file in files)
            {
                try
                {
                    AddFileToCanvas(file, dropPoint);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
                {
                    failures.Add($"{Path.GetFileName(file)}: {ex.Message}");
                }
            }

            if (failures.Count > 0)
            {
                MessageBox.Show(
                    this,
                    "The following file(s) could not be loaded:\n\n" + string.Join("\n", failures),
                    "Unable to Load File",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Scans the Windows Downloads folder for PNG/JPG/PDF files that
        /// finished downloading within the last 5 minutes and ingests each
        /// one onto the canvas exactly as if it had been dropped by the
        /// cursor. Anything older than the 5-minute window is left alone.
        /// </summary>
        private void AddRecentDownloadsButton_Click(object sender, RoutedEventArgs e)
        {
            string downloadsFolder = KnownFolders.GetDownloadsFolderPath();
            if (!Directory.Exists(downloadsFolder))
            {
                MessageBox.Show(this, "Could not locate your Downloads folder.", "Add Recent Downloads", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DateTime cutoff = DateTime.Now - RecentDownloadWindow;
            List<FileInfo> candidates;
            try
            {
                candidates = Directory.EnumerateFiles(downloadsFolder)
                    .Where(path => Path.GetExtension(path).ToLowerInvariant() is ".png" or ".jpg" or ".jpeg" or ".pdf")
                    .Select(path => new FileInfo(path))
                    // A file is "downloaded within the last 5 minutes" if either timestamp
                    // falls in the window - covers browsers that touch LastWriteTime on
                    // completion as well as ones that only set CreationTime.
                    .Where(info => info.CreationTime >= cutoff || info.LastWriteTime >= cutoff)
                    .OrderBy(info => info.CreationTime)
                    .ToList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                MessageBox.Show(this, $"Could not read the Downloads folder:\n{ex.Message}", "Add Recent Downloads", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (candidates.Count == 0)
            {
                MessageBox.Show(this, "No PNG, JPG, or PDF files found in Downloads from the last 5 minutes.", "Add Recent Downloads", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dropPoint = new Point(40, 40);
            var failures = new List<string>();
            foreach (var file in candidates)
            {
                try
                {
                    AddFileToCanvas(file.FullName, dropPoint);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
                {
                    failures.Add($"{file.Name}: {ex.Message}");
                }
            }

            if (failures.Count > 0)
            {
                MessageBox.Show(
                    this,
                    "The following recent download(s) could not be loaded:\n\n" + string.Join("\n", failures),
                    "Unable to Load File",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }

            int loaded = candidates.Count - failures.Count;
            if (loaded > 0)
            {
                MessageBox.Show(this, $"Added {loaded} recent download(s) to the workspace.", "Add Recent Downloads", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private readonly HashSet<string> _autoAddedDownloads = new(StringComparer.OrdinalIgnoreCase);

        private void AddFileToCanvas(string filePath, Point dropPoint)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();

            if (extension is not (".png" or ".jpg" or ".jpeg" or ".pdf"))
            {
                throw new NotSupportedException("Only PNG, JPG, and PDF files are supported.");
            }

            // Read the file into memory once and embed those bytes on the
            // item itself. From this point on the item never needs the
            // original file on disk again - this is what lets a saved
            // workspace file be fully self-contained and portable.
            byte[] bytes = File.ReadAllBytes(filePath);

            // Cascade subsequent items so multi-file drops don't stack exactly on top of one another.
            double offset = _dropCounter * 24;
            _dropCounter++;
            double x = Math.Max(0, dropPoint.X + offset);
            double y = Math.Max(0, dropPoint.Y + offset);

            switch (extension)
            {
                case ".png":
                case ".jpg":
                case ".jpeg":
                    _items.Add(new ImageItemModel
                    {
                        FilePath = filePath,
                        FileBytes = bytes,
                        X = x,
                        Y = y,
                        Width = DefaultImageWidth,
                        Height = DefaultImageHeight,
                        Image = LoadBitmapImageFromBytes(bytes)
                    });
                    break;

                case ".pdf":
                    _items.Add(new PdfItemModel
                    {
                        FilePath = filePath,
                        FileBytes = bytes,
                        X = x,
                        Y = y,
                        Width = DefaultPdfWidth,
                        Height = DefaultPdfHeight
                    });
                    break;
            }
        }

        /// <summary>Builds a frozen, UI-thread-safe bitmap from in-memory bytes (no disk access).</summary>
        private static BitmapImage LoadBitmapImageFromBytes(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.StreamSource = stream;
            bitmap.EndInit();
            bitmap.Freeze();
            return bitmap;
        }

        private void Chrome_RemoveRequested(object? sender, CanvasItemModel item)
        {
            item.IsHidden = true;
            RefreshDrawerList();
        }

        // ======================= Toolbar actions =======================

        private void SaveLayoutButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Title = "Save Workspace Layout",
                Filter = "Hexxcorp Workspace (*.lqw)|*.lqw",
                DefaultExt = WorkspaceFileService.ProjectFileExtension,
                FileName = $"Workspace_{DateTime.Now:yyyyMMdd_HHmm}"
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                WorkspaceFileService.Save(dialog.FileName, _items, string.Empty, BuildHeaderInfo(), BuildDateFieldsInfo(), BuildWordDocBytes(), _wordDocFileName, BuildExcelDocBytes(), _excelDocFileName);
                MessageBox.Show(this, "Workspace layout saved successfully.", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                MessageBox.Show(this, $"Could not save the workspace file:\n{ex.Message}", "Save Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadLayoutButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Load Workspace Layout",
                Filter = "Hexxcorp Workspace (*.lqw)|*.lqw|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                var project = WorkspaceFileService.Load(dialog.FileName);
                _items.Clear();
                ApplyHeaderInfo(project.Header);
                ApplyMode(project.Mode);
                ApplyDateFieldsInfo(project.DateFields ?? new WorkspaceDateFieldsInfo());
                if (_mainViewModel != null)
                {
                    if (project.Tracker != null) _mainViewModel.Tracker.Load(project.Tracker);
                    else _mainViewModel.Tracker.Clear();
                    PushBoxesToTracker();
                }
                ApplyWordDoc(project.WordDocBytes, project.WordDocFileName);
                ApplyExcelDoc(project.ExcelDocBytes, project.ExcelDocFileName);
                var missingFiles = new System.Collections.Generic.List<string>();

                foreach (var record in project.Items)
                {
                    try
                    {
                        // Prefer the embedded bytes (present for files saved by
                        // this version). Fall back to reading the original path
                        // from disk only for older project files saved before
                        // embedding was introduced - if that path is also gone,
                        // this item is reported as missing below.
                        byte[] bytes = record.Data is { Length: > 0 }
                            ? record.Data
                            : File.ReadAllBytes(record.FilePath);

                        CanvasItemModel item = record.Kind == WorkspaceItemKind.Image
                            ? new ImageItemModel { Image = LoadBitmapImageFromBytes(bytes) }
                            : new PdfItemModel();

                        item.FilePath = record.FilePath;
                        item.FileBytes = bytes;
                        item.X = record.X;
                        item.Y = record.Y;
                        item.Width = record.Width;
                        item.Height = record.Height;
                        item.Zoom = record.Zoom <= 0 ? 1.0 : record.Zoom;
                        item.IsHidden = record.IsHidden;
                        _items.Add(item);
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
                    {
                        missingFiles.Add($"{Path.GetFileName(record.FilePath)}: {ex.Message}");
                    }
                }

                if (missingFiles.Count > 0)
                {
                    MessageBox.Show(
                        this,
                        "Some referenced files could not be restored:\n\n" + string.Join("\n", missingFiles),
                        "Partial Load",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
            {
                MessageBox.Show(this, $"Could not load the workspace file:\n{ex.Message}", "Load Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearCanvasButton_Click(object sender, RoutedEventArgs e)
        {
            if (_items.Count == 0)
            {
                return;
            }

            var result = MessageBox.Show(
                this,
                "Clear all items from the canvas? Unsaved work will be lost.",
                "Clear Canvas",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (result == MessageBoxResult.Yes)
            {
                _items.Clear();
            }
        }

        /// <summary>Default width (in device-independent pixels) the View/Edit Doc drawer opens to the first time it's used.</summary>
        private double _wordDocDrawerWidth = 380;

        /// <summary>Default width (in device-independent pixels) the View/Edit Sheet drawer opens to the first time it's used.</summary>
        private double _excelDrawerWidth = 480;

        /// <summary>Default width (in device-independent pixels) the View Browser drawer opens to the first time it's used.</summary>
        private double _browserDrawerWidth = 480;

        /// <summary>Default width (in device-independent pixels) the View Assistant drawer opens to the first time it's used.</summary>
        private double _copilotDrawerWidth = 520;

        private bool _isBrowserInitialized;
        private Microsoft.Web.WebView2.Core.CoreWebView2Environment? _browserEnv;
        private BrowserTabManager? _browserTabs;
        private Microsoft.Web.WebView2.Wpf.WebView2 ActiveBrowser => _browserTabs?.Active ?? BrowserWebView;
        private QualMode _mode = QualMode.DriverFile;

        // ======================= Review Mode (Driver File / Full Unit / Post-Review Audit) =======================

        private void ModeToggleButton_Click(object sender, RoutedEventArgs e)
        {
            ApplyMode((QualMode)(((int)_mode + 1) % 3));
        }

        /// <summary>Switches review mode: updates labels, grays out/enables lease boxes, and tells the prompt builder.</summary>
        private void ApplyMode(QualMode mode)
        {
            _mode = mode;
            CopilotContextBuilder.Mode = mode;
            bool full = mode == QualMode.FullUnit;

            ModeToggleButton.Content = mode switch
            {
                QualMode.FullUnit => "Mode: Full Unit",
                QualMode.PostReviewAudit => "Mode: Post-Review Audit",
                _ => "Mode: Driver File"
            };
            SolLabelText.Text = full ? "SOL - TRUCK" : "STATEMENT OF LEASE (N/A)";
            SolTrailerPanel.Visibility = full ? Visibility.Visible : Visibility.Collapsed;
            StatementOfLeaseDateBox.IsEnabled = full;
            StatementOfLeaseTrailerDateBox.IsEnabled = full;
            StatementOfLeaseDateBox.Opacity = full ? 1.0 : 0.4;
            SolLabelText.Opacity = full ? 1.0 : 0.4;
            AuditDatePanel.Visibility = mode == QualMode.PostReviewAudit ? Visibility.Visible : Visibility.Collapsed;
            if (mode == QualMode.PostReviewAudit)
            {
                AuditDateBox.Text = DayQualifiedDateBox.Text;
                if (AuditDateBox.Text.Length == 0) AuditDateBox.Focus();
            }
            RevalidateAllDateFields();
        }

        private bool _isInitializingBrowser;
        private bool _isCopilotInitialized;
        private bool _isInitializingCopilot;
        private bool _isCopilotLoading;
        private string? _pendingAttachFilePath;
        private CancellationTokenSource? _autoApplyCts;

        private void UpdateDrawerMargins()
        {
            bool trackerOpen = TrackerDrawerHost.Visibility == Visibility.Visible;
            bool copilotOpen = CopilotDrawerHost.Visibility == Visibility.Visible;
            bool browserOpen = BrowserDrawerHost.Visibility == Visibility.Visible;
            bool excelOpen = ExcelDrawerHost.Visibility == Visibility.Visible;
            bool docOpen = WordDocDrawerHost.Visibility == Visibility.Visible;

            // Word is rightmost: right margin is always 8; left is 4 if an open drawer is to its left, else 0.
            double docLeft = (trackerOpen || copilotOpen || browserOpen || excelOpen) ? 4 : 0;
            NotesPanel.Margin = new Thickness(docLeft, 8, 8, 8);

            // Excel is second from right: right is 4 if Word is open, else 8; left is 4 if Browser/Copilot/Tracker is open, else 0.
            double excelRight = docOpen ? 4 : 8;
            double excelLeft = (trackerOpen || copilotOpen || browserOpen) ? 4 : 0;
            ExcelPanel.Margin = new Thickness(excelLeft, 8, excelRight, 8);

            // Browser is third from right: right is 4 if Excel/Doc is open, else 8; left is 4 if Copilot/Tracker is open, else 0.
            double browserRight = (excelOpen || docOpen) ? 4 : 8;
            double browserLeft = (trackerOpen || copilotOpen) ? 4 : 0;
            BrowserPanel.Margin = new Thickness(browserLeft, 8, browserRight, 8);

            // Copilot: left is 4 if the Tracker is open to its left, else 0; right is 4 if any drawer to its right is open, else 8.
            double copilotRight = (browserOpen || excelOpen || docOpen) ? 4 : 8;
            CopilotPanel.Margin = new Thickness(trackerOpen ? 4 : 0, 8, copilotRight, 8);

            // Tracker is leftmost of the shelf drawers.
            double trackerRight = (copilotOpen || browserOpen || excelOpen || docOpen) ? 4 : 8;
            TrackerPanelBorder.Margin = new Thickness(0, 8, trackerRight, 8);
        }

        private double _trackerDrawerWidth = 520;

        /// <summary>Opens or closes the Dates &amp; Work History tracker drawer on the shelf.</summary>
        private void TrackerToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (TrackerDrawerHost.Visibility != Visibility.Visible)
            {
                TrackerDrawerHost.Width = _trackerDrawerWidth;
                TrackerDrawerHost.Visibility = Visibility.Visible;
                TrackerToggleButton.Content = "Hide Work History";
            }
            else
            {
                if (TrackerDrawerHost.Width > 0) _trackerDrawerWidth = TrackerDrawerHost.Width;
                TrackerDrawerHost.Width = 0;
                TrackerDrawerHost.Visibility = Visibility.Collapsed;
                TrackerToggleButton.Content = "Work History";
            }
            UpdateDrawerMargins();
        }

        /// <summary>Opens the tracker drawer if it is closed (used by the Sheet drawer's third button).</summary>
        private void OpenTrackerDrawerButton_Click(object sender, RoutedEventArgs e)
        {
            if (TrackerDrawerHost.Visibility != Visibility.Visible) TrackerToggleButton_Click(sender, e);
        }

        private void TrackerDrawerSplitter_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            double others = new[] { CopilotDrawerHost, BrowserDrawerHost, ExcelDrawerHost, WordDocDrawerHost }
                .Where(h => h.Visibility == Visibility.Visible).Sum(h => h.Width);
            double max = Math.Max(300, ActualWidth - others - 250);
            double w = Math.Clamp(TrackerDrawerHost.Width - e.HorizontalChange, 300, max);
            TrackerDrawerHost.Width = w;
            _trackerDrawerWidth = w;
        }

        /// <summary>Opens or closes the View/Edit Doc drawer by snapping its host width, remembering the last drag-resized width for next time.</summary>
        private void NotesToggleButton_Click(object sender, RoutedEventArgs e)
        {
            bool opening = WordDocDrawerHost.Visibility != Visibility.Visible;
            if (opening)
            {
                WordDocDrawerHost.Width = _wordDocDrawerWidth;
                WordDocDrawerHost.Visibility = Visibility.Visible;
                NotesToggleButton.Content = "Hide Doc";
                UpdateDrawerMargins();
            }
            else
            {
                if (WordDocDrawerHost.Width > 0)
                {
                    _wordDocDrawerWidth = WordDocDrawerHost.Width;
                }

                WordDocDrawerHost.Width = 0;
                WordDocDrawerHost.Visibility = Visibility.Collapsed;
                NotesToggleButton.Content = "View/Edit Doc";
                UpdateDrawerMargins();
            }
        }

        /// <summary>Opens or closes the View/Edit Sheet drawer by snapping its host width, remembering the last drag-resized width for next time.</summary>
        private void SheetToggleButton_Click(object sender, RoutedEventArgs e)
        {
            bool opening = ExcelDrawerHost.Visibility != Visibility.Visible;
            if (opening)
            {
                ExcelDrawerHost.Width = _excelDrawerWidth;
                ExcelDrawerHost.Visibility = Visibility.Visible;
                SheetToggleButton.Content = "Hide Sheet";
                UpdateDrawerMargins();
            }
            else
            {
                if (ExcelDrawerHost.Width > 0)
                {
                    _excelDrawerWidth = ExcelDrawerHost.Width;
                }

                ExcelDrawerHost.Width = 0;
                ExcelDrawerHost.Visibility = Visibility.Collapsed;
                SheetToggleButton.Content = "View/Edit Sheet";
                UpdateDrawerMargins();
            }
        }

        /// <summary>Opens or closes the View Browser drawer by snapping its host width, remembering the last drag-resized width for next time.</summary>
        private async void BrowserToggleButton_Click(object sender, RoutedEventArgs e)
        {
            bool opening = BrowserDrawerHost.Visibility != Visibility.Visible;
            if (opening)
            {
                BrowserDrawerHost.Width = _browserDrawerWidth;
                BrowserDrawerHost.Visibility = Visibility.Visible;
                BrowserToggleButton.Content = "Hide Browser";
                UpdateDrawerMargins();
                await EnsureBrowserInitializedAsync();
            }
            else
            {
                if (BrowserDrawerHost.Width > 0)
                {
                    _browserDrawerWidth = BrowserDrawerHost.Width;
                }

                BrowserDrawerHost.Width = 0;
                BrowserDrawerHost.Visibility = Visibility.Collapsed;
                BrowserToggleButton.Content = "Browser";
                UpdateDrawerMargins();
            }
        }

        /// <summary>Opens or closes the View Assistant drawer by snapping its host width, remembering the last drag-resized width for next time.</summary>
        private async void CopilotToggleButton_Click(object sender, RoutedEventArgs e)
        {
            bool opening = CopilotDrawerHost.Visibility != Visibility.Visible;
            if (opening)
            {
                CopilotDrawerHost.Width = _copilotDrawerWidth;
                CopilotDrawerHost.Visibility = Visibility.Visible;
                CopilotToggleButton.Content = "Hide Assistant";
                UpdateDrawerMargins();
                await EnsureCopilotInitializedAsync();
            }
            else
            {
                if (CopilotDrawerHost.Width > 0)
                {
                    _copilotDrawerWidth = CopilotDrawerHost.Width;
                }

                CopilotDrawerHost.Width = 0;
                CopilotDrawerHost.Visibility = Visibility.Collapsed;
                CopilotToggleButton.Content = "Assistant";
                UpdateDrawerMargins();
            }
        }

        /// <summary>Lets the user freely drag-resize the open View/Edit Doc drawer via its gold splitter handle.</summary>
        private void WordDocDrawerSplitter_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            double currentCopilotWidth = CopilotDrawerHost.Visibility == Visibility.Visible ? CopilotDrawerHost.Width : 0;
            double currentBrowserWidth = BrowserDrawerHost.Visibility == Visibility.Visible ? BrowserDrawerHost.Width : 0;
            double currentExcelWidth = ExcelDrawerHost.Visibility == Visibility.Visible ? ExcelDrawerHost.Width : 0;
            double maxDocWidth = Math.Max(260, ActualWidth - currentCopilotWidth - currentBrowserWidth - currentExcelWidth - 250);
            double newWidth = Math.Clamp(WordDocDrawerHost.Width - e.HorizontalChange, 260, maxDocWidth);
            WordDocDrawerHost.Width = newWidth;
            _wordDocDrawerWidth = newWidth;
            WordDocDrawerHost.UpdateLayout();
            _wordController?.ResizeToFillContainer();
        }

        /// <summary>Lets the user freely drag-resize the open View/Edit Sheet drawer via its gold splitter handle.</summary>
        private void ExcelDrawerSplitter_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            double currentDocWidth = WordDocDrawerHost.Visibility == Visibility.Visible ? WordDocDrawerHost.Width : 0;
            double currentCopilotWidth = CopilotDrawerHost.Visibility == Visibility.Visible ? CopilotDrawerHost.Width : 0;
            double currentBrowserWidth = BrowserDrawerHost.Visibility == Visibility.Visible ? BrowserDrawerHost.Width : 0;
            double maxExcelWidth = Math.Max(260, ActualWidth - currentDocWidth - currentCopilotWidth - currentBrowserWidth - 250);
            double newWidth = Math.Clamp(ExcelDrawerHost.Width - e.HorizontalChange, 260, maxExcelWidth);
            ExcelDrawerHost.Width = newWidth;
            _excelDrawerWidth = newWidth;
            ExcelDrawerHost.UpdateLayout();
            _excelController?.ResizeToFillContainer();
        }

        /// <summary>Lets the user freely drag-resize the open View Browser drawer via its gold splitter handle.</summary>
        private void BrowserDrawerSplitter_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            double currentDocWidth = WordDocDrawerHost.Visibility == Visibility.Visible ? WordDocDrawerHost.Width : 0;
            double currentExcelWidth = ExcelDrawerHost.Visibility == Visibility.Visible ? ExcelDrawerHost.Width : 0;
            double currentCopilotWidth = CopilotDrawerHost.Visibility == Visibility.Visible ? CopilotDrawerHost.Width : 0;
            double maxBrowserWidth = Math.Max(260, ActualWidth - currentDocWidth - currentExcelWidth - currentCopilotWidth - 250);
            double newWidth = Math.Clamp(BrowserDrawerHost.Width - e.HorizontalChange, 260, maxBrowserWidth);
            BrowserDrawerHost.Width = newWidth;
            _browserDrawerWidth = newWidth;
        }

        /// <summary>Lets the user freely drag-resize the open View Co-Pilot drawer via its gold splitter handle.</summary>
        private void CopilotDrawerSplitter_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            double currentDocWidth = WordDocDrawerHost.Visibility == Visibility.Visible ? WordDocDrawerHost.Width : 0;
            double currentExcelWidth = ExcelDrawerHost.Visibility == Visibility.Visible ? ExcelDrawerHost.Width : 0;
            double currentBrowserWidth = BrowserDrawerHost.Visibility == Visibility.Visible ? BrowserDrawerHost.Width : 0;
            double maxCopilotWidth = Math.Max(280, ActualWidth - currentDocWidth - currentExcelWidth - currentBrowserWidth - 250);
            double newWidth = Math.Clamp(CopilotDrawerHost.Width - e.HorizontalChange, 280, maxCopilotWidth);
            CopilotDrawerHost.Width = newWidth;
            _copilotDrawerWidth = newWidth;
        }

        // ======================= Embedded WebView2 Browser =======================

        private async Task EnsureBrowserInitializedAsync()
        {
            if (_isBrowserInitialized || _isInitializingBrowser)
            {
                return;
            }

            _isInitializingBrowser = true;
            try
            {
                Services.WebView2LoaderHelper.EnsureLoaderInitialized();

                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string userDataFolder = System.IO.Path.Combine(localAppData, "HexxcorpDocumentReviewUtility", "WebView2Data");
                System.IO.Directory.CreateDirectory(userDataFolder);

                var options = new Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions
                {
                    AllowSingleSignOnUsingOSPrimaryAccount = false
                };

                var environment = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, userDataFolder, options);
                _browserEnv = environment;
                await BrowserWebView.EnsureCoreWebView2Async(environment);

                _isBrowserInitialized = true;

                _browserTabs = new BrowserTabManager(BrowserTabHost, BrowserTabStrip, BrowserWebView, () => _browserEnv, "about:blank");
                _browserTabs.UrlChanged += u => BrowserUrlTextBox.Text = u;
                _browserTabs.NavigationStarted += u =>
                {
                    BrowserStatusText.Text = $"Loading {u}...";
                    BrowserStatusText.Visibility = Visibility.Visible;
                };
                _browserTabs.NavigationFinished += () => BrowserStatusText.Visibility = Visibility.Collapsed;
                _browserTabs.NavStateChanged += UpdateBrowserNavButtons;
                _browserTabs.AttachFirst();

                NavigateToUrl("about:blank");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    $"Could not initialize the embedded web browser.\n\n{ex.Message}\n\nPlease ensure Microsoft Edge or Microsoft Edge WebView2 Runtime is installed.",
                    "Web Browser",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            finally
            {
                _isInitializingBrowser = false;
            }
        }

        private void UpdateBrowserNavButtons()
        {
            if (!_isBrowserInitialized || BrowserWebView.CoreWebView2 == null)
            {
                BrowserBackButton.IsEnabled = false;
                BrowserForwardButton.IsEnabled = false;
                return;
            }

            var av = _browserTabs?.Active ?? BrowserWebView;
            BrowserBackButton.IsEnabled = av.CanGoBack;
            BrowserForwardButton.IsEnabled = av.CanGoForward;
        }

        private void BrowserBackButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isBrowserInitialized && ActiveBrowser.CanGoBack)
            {
                ActiveBrowser.GoBack();
            }
        }

        private void BrowserForwardButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isBrowserInitialized && ActiveBrowser.CanGoForward)
            {
                ActiveBrowser.GoForward();
            }
        }

        private void BrowserRefreshButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isBrowserInitialized)
            {
                ActiveBrowser.Reload();
            }
        }

        private void BrowserHomeButton_Click(object sender, RoutedEventArgs e)
        {
            NavigateToUrl("about:blank");
        }

        private void BrowserGoButton_Click(object sender, RoutedEventArgs e)
        {
            NavigateToUrl(BrowserUrlTextBox.Text);
        }

        private void BrowserUrlTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                NavigateToUrl(BrowserUrlTextBox.Text);
                e.Handled = true;
            }
        }

        private async void NavigateToUrl(string? input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return;
            }

            if (!_isBrowserInitialized)
            {
                await EnsureBrowserInitializedAsync();
                if (!_isBrowserInitialized)
                {
                    return;
                }
            }

            input = input.Trim();
            if (!input.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                !input.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                if (input.Contains('.') && !input.Contains(' '))
                {
                    input = "https://" + input;
                }
                else
                {
                    input = "https://www.google.com/search?q=" + Uri.EscapeDataString(input);
                }
            }

            try
            {
                BrowserUrlTextBox.Text = input;
                ActiveBrowser.CoreWebView2?.Navigate(input);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Could not navigate to:\n{input}\n\n{ex.Message}", "Browser Navigation", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ======================= Embedded Microsoft 365 Copilot Work Chat =======================

        private const string CopilotWorkUrl = "https://copilot.microsoft.com/";

        private async Task EnsureCopilotInitializedAsync()
        {
            if (_isCopilotInitialized || _isInitializingCopilot)
            {
                return;
            }

            _isInitializingCopilot = true;
            try
            {
                Services.WebView2LoaderHelper.EnsureLoaderInitialized();

                CopilotStatusText.Text = "Initializing Assistant...";
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string userDataFolder = System.IO.Path.Combine(localAppData, "HexxcorpDocumentReviewUtility", "WebView2Data");
                System.IO.Directory.CreateDirectory(userDataFolder);

                var options = new Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions
                {
                    AllowSingleSignOnUsingOSPrimaryAccount = false
                };

                var environment = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, userDataFolder, options);
                await CopilotWebView.EnsureCoreWebView2Async(environment);

                _isCopilotInitialized = true;

                // Setup CDP file chooser interceptor so programmatic attachments bypass native dialogs
                try
                {
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("Page.enable", "{}");
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("DOM.enable", "{}");
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                        "Page.setInterceptFileChooserDialog",
                        "{\"enabled\":true}");

                    var fileChooserReceiver = CopilotWebView.CoreWebView2.GetDevToolsProtocolEventReceiver("Page.fileChooserOpened");
                    fileChooserReceiver.DevToolsProtocolEventReceived += OnCopilotFileChooserOpened;
                }
                catch
                {
                }

                CopilotWebView.CoreWebView2.NavigationStarting += (_, _) =>
                {
                    _isCopilotLoading = true;
                    CopilotStatusText.Text = "Connecting...";
                };

                CopilotWebView.CoreWebView2.NavigationCompleted += async (_, e) =>
                {
                    _isCopilotLoading = false;
                    CopilotStatusText.Text = e.IsSuccess ? "Assistant Ready" : "Navigation Failed";

                    if (e.IsSuccess)
                    {
                        try
                        {
                            await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("Page.enable", "{}");
                            await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("DOM.enable", "{}");
                            await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                                "Page.setInterceptFileChooserDialog",
                                "{\"enabled\":true}");
                        }
                        catch
                        {
                        }
                    }
                };

                CopilotWebView.CoreWebView2.Navigate(CopilotWorkUrl);
            }
            catch (Exception ex)
            {
                CopilotStatusText.Text = "Failed to load Assistant";
                MessageBox.Show(
                    this,
                    $"Could not initialize Microsoft Copilot Work Chat.\n\n{ex.Message}\n\nPlease ensure Microsoft Edge or Microsoft Edge WebView2 Runtime is installed.",
                    "Qualification Assistant",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            finally
            {
                _isInitializingCopilot = false;
            }
        }

        private async void OnCopilotFileChooserOpened(object? sender, Microsoft.Web.WebView2.Core.CoreWebView2DevToolsProtocolEventReceivedEventArgs e)
        {
            string? targetPath = _pendingAttachFilePath;
            if (string.IsNullOrWhiteSpace(targetPath) || !File.Exists(targetPath))
            {
                return;
            }

            try
            {
                string handleParams = JsonSerializer.Serialize(new { action = "accept", files = new[] { targetPath } });
                await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("Page.handleFileChooser", handleParams);
            }
            catch
            {
            }

            try
            {
                string handleParams = JsonSerializer.Serialize(new { action = "accept", files = new[] { targetPath } });
                await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("Page.handleFileChooserDialog", handleParams);
            }
            catch
            {
            }

            try
            {
                using var doc = JsonDocument.Parse(e.ParameterObjectAsJson);
                if (doc.RootElement.TryGetProperty("backendNodeId", out var bId))
                {
                    string domParams = JsonSerializer.Serialize(new { files = new[] { targetPath }, backendNodeId = bId.GetInt32() });
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("DOM.setFileInputFiles", domParams);
                }
            }
            catch
            {
            }
        }

        /// <summary>
        /// Locates or extracts a stored workspace file (PDF, image, Word doc, or Excel workbook) to a physical file on disk
        /// so it can be attached to Assistant chat, pasted via Windows Clipboard, or opened in File Explorer.
        /// </summary>
        public (string FilePath, string FileName)? GetOrExtractStoredFilePath(string? preferredFileName = null)
        {
            CanvasItemModel? match = null;
            if (!string.IsNullOrWhiteSpace(preferredFileName))
            {
                match = _items.FirstOrDefault(i => string.Equals(System.IO.Path.GetFileName(i.FilePath), preferredFileName, StringComparison.OrdinalIgnoreCase))
                     ?? _items.FirstOrDefault(i => string.Equals(i.DisplayName, preferredFileName, StringComparison.OrdinalIgnoreCase));
            }

            // Prefer PDF files if no specific file was requested
            match ??= _items.FirstOrDefault(i => i.Kind == WorkspaceItemKind.Pdf);
            match ??= _items.FirstOrDefault();

            if (match != null)
            {
                string fileName = match.DisplayName;
                if (string.IsNullOrWhiteSpace(fileName))
                {
                    fileName = match.Kind == WorkspaceItemKind.Pdf ? "qualification_file.pdf" : "qualification_image.png";
                }

                if (!string.IsNullOrWhiteSpace(match.FilePath) && File.Exists(match.FilePath))
                {
                    return (match.FilePath, fileName);
                }

                if (match.FileBytes != null && match.FileBytes.Length > 0)
                {
                    try
                    {
                        string tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "HexxcorpDocumentReviewFiles");
                        System.IO.Directory.CreateDirectory(tempDir);
                        string safeName = string.Join("_", fileName.Split(System.IO.Path.GetInvalidFileNameChars()));
                        string targetPath = System.IO.Path.Combine(tempDir, safeName);
                        File.WriteAllBytes(targetPath, match.FileBytes);
                        return (targetPath, fileName);
                    }
                    catch
                    {
                    }
                }
            }

            if (HasWordDocLoaded && (string.IsNullOrWhiteSpace(preferredFileName) || string.Equals(preferredFileName, _wordDocFileName, StringComparison.OrdinalIgnoreCase)))
            {
                string fileName = string.IsNullOrWhiteSpace(_wordDocFileName) ? "Document.docx" : _wordDocFileName;
                byte[]? bytes = BuildWordDocBytes();
                if (bytes != null && bytes.Length > 0)
                {
                    try
                    {
                        string tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "HexxcorpDocumentReviewFiles");
                        System.IO.Directory.CreateDirectory(tempDir);
                        string targetPath = System.IO.Path.Combine(tempDir, fileName);
                        File.WriteAllBytes(targetPath, bytes);
                        return (targetPath, fileName);
                    }
                    catch
                    {
                    }
                }
            }

            if (HasExcelLoaded && (string.IsNullOrWhiteSpace(preferredFileName) || string.Equals(preferredFileName, _excelDocFileName, StringComparison.OrdinalIgnoreCase)))
            {
                string fileName = string.IsNullOrWhiteSpace(_excelDocFileName) ? "Workbook.xlsx" : _excelDocFileName;
                byte[]? bytes = BuildExcelDocBytes();
                if (bytes != null && bytes.Length > 0)
                {
                    try
                    {
                        string tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "HexxcorpDocumentReviewFiles");
                        System.IO.Directory.CreateDirectory(tempDir);
                        string targetPath = System.IO.Path.Combine(tempDir, fileName);
                        File.WriteAllBytes(targetPath, bytes);
                        return (targetPath, fileName);
                    }
                    catch
                    {
                    }
                }
            }

            return null;
        }

        public List<(string FilePath, string FileName)> GetAllStoredFilePaths()
        {
            var list = new List<(string FilePath, string FileName)>();
            foreach (var item in _items)
            {
                var resolved = GetOrExtractStoredFilePath(item.DisplayName);
                if (resolved != null && !list.Any(x => x.FilePath == resolved.Value.FilePath))
                {
                    list.Add(resolved.Value);
                }
            }
            if (HasWordDocLoaded)
            {
                var resolved = GetOrExtractStoredFilePath(_wordDocFileName);
                if (resolved != null && !list.Any(x => x.FilePath == resolved.Value.FilePath))
                {
                    list.Add(resolved.Value);
                }
            }
            if (HasExcelLoaded)
            {
                var resolved = GetOrExtractStoredFilePath(_excelDocFileName);
                if (resolved != null && !list.Any(x => x.FilePath == resolved.Value.FilePath))
                {
                    list.Add(resolved.Value);
                }
            }
            return list;
        }

        /// <summary>
        /// Attaches a stored qualification document into Microsoft Copilot Work Chat via
        /// Chrome DevTools Protocol automation and Windows Clipboard integration.
        /// </summary>
        public async Task<bool> AttachFileToCopilotAsync(string? preferredFileName = null)
        {
            var fileInfo = GetOrExtractStoredFilePath(preferredFileName);
            if (fileInfo == null)
            {
                MessageBox.Show(
                    this,
                    "No stored files (PDF, image, Word, or Excel) were found in the workspace to attach.\n\nPlease drag and drop a qualification PDF into the workspace first.",
                    "Attach File",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return false;
            }

            string filePath = fileInfo.Value.FilePath;
            string fileName = fileInfo.Value.FileName;
            _pendingAttachFilePath = filePath;

            // 1. Populate the Windows Clipboard with the file drop list and text path so user can press Ctrl+V anytime
            try
            {
                var dropList = new System.Collections.Specialized.StringCollection { filePath };
                var dataObj = new DataObject();
                dataObj.SetFileDropList(dropList);
                dataObj.SetText(filePath);
                Clipboard.SetDataObject(dataObj, true);
            }
            catch
            {
            }

            // 2. Open Assistant drawer if not open
            if (!_isCopilotInitialized || CopilotWebView.CoreWebView2 == null)
            {
                await EnsureCopilotInitializedAsync();
            }

            if (CopilotDrawerHost.Visibility != Visibility.Visible)
            {
                CopilotDrawerHost.Width = _copilotDrawerWidth;
                CopilotDrawerHost.Visibility = Visibility.Visible;
                CopilotToggleButton.Content = "Hide Assistant";
                UpdateDrawerMargins();
            }

            CopilotStatusText.Text = $"Attaching {fileName}...";

            // If navigating, wait briefly
            int waited = 0;
            while (_isCopilotLoading && waited < 8000)
            {
                await Task.Delay(400);
                waited += 400;
            }

            if (CopilotWebView.CoreWebView2 == null)
            {
                CopilotStatusText.Text = $"{fileName} copied! (Press Ctrl+V in chat)";
                return true;
            }

            try
            {
                // Ensure page, DOM, and file chooser interception are active
                await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("Page.enable", "{}");
                await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("DOM.enable", "{}");
                await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                    "Page.setInterceptFileChooserDialog",
                    "{\"enabled\":true}");

                // Helper script that finds and clicks either the 'Upload images and files' menu item (if already open)
                // or opens the '+' / 'Add content' menu
                string menuInteractScript = @"
(function() {
    function queryAllDeep(selector, root) {
        root = root || document;
        let results = [];
        try {
            const matches = root.querySelectorAll(selector);
            for (let i = 0; i < matches.length; i++) results.push(matches[i]);
        } catch (_) {}
        try {
            const all = root.querySelectorAll('*');
            for (let i = 0; i < all.length; i++) {
                if (all[i].shadowRoot) results = results.concat(queryAllDeep(selector, all[i].shadowRoot));
            }
        } catch (_) {}
        try {
            const iframes = root.querySelectorAll('iframe');
            for (let i = 0; i < iframes.length; i++) {
                try {
                    const doc = iframes[i].contentDocument || iframes[i].contentWindow?.document;
                    if (doc) results = results.concat(queryAllDeep(selector, doc));
                } catch (_) {}
            }
        } catch (_) {}
        return results;
    }

    function isVisible(el) {
        if (!el) return false;
        try {
            const rect = el.getBoundingClientRect();
            if (rect.width <= 0 || rect.height <= 0) return false;
            const style = window.getComputedStyle(el);
            if (style.visibility === 'hidden' || style.display === 'none' || style.opacity === '0') return false;
            return true;
        } catch (_) { return false; }
    }

    function findUploadMenuItem() {
        const candidates = queryAllDeep('[role=""menuitem""], [role=""button""], button, li, a, div, span');
        for (let i = 0; i < candidates.length; i++) {
            const el = candidates[i];
            if (!isVisible(el)) continue;
            const txt = (el.innerText || el.textContent || '').trim().toLowerCase();
            if (txt.includes('attach cloud files') && txt.includes('upload images and files') && el.children.length > 2) {
                continue;
            }
            if (txt.includes('upload images and files') ||
                txt.includes('upload from this device') ||
                txt.includes('upload from computer') ||
                txt.includes('upload images') ||
                txt === 'upload images and files' ||
                (txt.includes('upload') && !txt.includes('cloud'))) {
                return el;
            }
        }
        return null;
    }

    function clickElement(el) {
        let target = el;
        while (target && target !== document.body &&
               target.getAttribute('role') !== 'menuitem' &&
               target.tagName !== 'BUTTON' &&
               target.tagName !== 'A') {
            if (target.parentElement && target.parentElement.getAttribute('role') === 'menuitem') {
                target = target.parentElement;
                break;
            }
            target = target.parentElement;
        }
        if (!target || target === document.body) target = el;

        target.focus();
        const opts = { bubbles: true, cancelable: true, view: window };
        target.dispatchEvent(new MouseEvent('pointerdown', opts));
        target.dispatchEvent(new MouseEvent('mousedown', opts));
        target.dispatchEvent(new MouseEvent('pointerup', opts));
        target.dispatchEvent(new MouseEvent('mouseup', opts));
        target.dispatchEvent(new MouseEvent('click', opts));
        if (typeof target.click === 'function') target.click();
        return true;
    }

    // 1. If 'Upload images and files' is ALREADY open on screen, click it directly!
    const existing = findUploadMenuItem();
    if (existing) {
        clickElement(existing);
        return JSON.stringify({ action: 'CLICKED_EXISTING_UPLOAD_MENU' });
    }

    // 2. Otherwise find and click the '+' / 'Add content' button
    const attachSelectors = [
        'button[aria-label*=""Add content"" i]',
        'button[aria-label*=""Attach"" i]',
        'button[aria-label*=""Upload"" i]',
        'button[aria-label*=""Add"" i]',
        'button[title*=""Add content"" i]',
        'button[title*=""Attach"" i]',
        'button[title*=""Upload"" i]',
        'button[data-testid*=""attach"" i]',
        'button[data-testid*=""upload"" i]'
    ];
    for (let s = 0; s < attachSelectors.length; s++) {
        const btns = queryAllDeep(attachSelectors[s]);
        for (let i = 0; i < btns.length; i++) {
            if (isVisible(btns[i])) {
                clickElement(btns[i]);
                return JSON.stringify({ action: 'CLICKED_PLUS_BUTTON' });
            }
        }
    }

    const footerButtons = queryAllDeep('footer button, div[role=""region""] button, [data-content=""chat-input""] button');
    for (let i = 0; i < footerButtons.length; i++) {
        const btn = footerButtons[i];
        if (isVisible(btn) && btn.querySelector('svg')) {
            const txt = (btn.getAttribute('aria-label') || btn.innerText || '').toLowerCase();
            if (!txt.includes('send') && !txt.includes('submit') && !txt.includes('voice') && !txt.includes('mic')) {
                clickElement(btn);
                return JSON.stringify({ action: 'CLICKED_GENERIC_PLUS_BUTTON' });
            }
        }
    }

    return JSON.stringify({ action: 'NONE_FOUND' });
})();";

                string attachRes = await CopilotWebView.CoreWebView2.ExecuteScriptAsync(menuInteractScript);

                if (attachRes.Contains("CLICKED_PLUS_BUTTON") || attachRes.Contains("CLICKED_GENERIC_PLUS_BUTTON"))
                {
                    // Allow popup menu to animate open
                    await Task.Delay(300);

                    string clickUploadAfterOpenScript = @"
(function() {
    function queryAllDeep(selector, root) {
        root = root || document;
        let results = [];
        try {
            const matches = root.querySelectorAll(selector);
            for (let i = 0; i < matches.length; i++) results.push(matches[i]);
        } catch (_) {}
        try {
            const all = root.querySelectorAll('*');
            for (let i = 0; i < all.length; i++) {
                if (all[i].shadowRoot) results = results.concat(queryAllDeep(selector, all[i].shadowRoot));
            }
        } catch (_) {}
        try {
            const iframes = root.querySelectorAll('iframe');
            for (let i = 0; i < iframes.length; i++) {
                try {
                    const doc = iframes[i].contentDocument || iframes[i].contentWindow?.document;
                    if (doc) results = results.concat(queryAllDeep(selector, doc));
                } catch (_) {}
            }
        } catch (_) {}
        return results;
    }

    function isVisible(el) {
        if (!el) return false;
        try {
            const rect = el.getBoundingClientRect();
            if (rect.width <= 0 || rect.height <= 0) return false;
            const style = window.getComputedStyle(el);
            if (style.visibility === 'hidden' || style.display === 'none' || style.opacity === '0') return false;
            return true;
        } catch (_) { return false; }
    }

    const candidates = queryAllDeep('[role=""menuitem""], [role=""button""], button, li, a, div, span');
    for (let i = 0; i < candidates.length; i++) {
        const el = candidates[i];
        if (!isVisible(el)) continue;
        const txt = (el.innerText || el.textContent || '').trim().toLowerCase();
        if (txt.includes('attach cloud files') && txt.includes('upload images and files') && el.children.length > 2) {
            continue;
        }
        if (txt.includes('upload images and files') ||
            txt.includes('upload from this device') ||
            txt.includes('upload from computer') ||
            txt.includes('upload images') ||
            txt === 'upload images and files' ||
            (txt.includes('upload') && !txt.includes('cloud'))) {
            
            let target = el;
            while (target && target !== document.body &&
                   target.getAttribute('role') !== 'menuitem' &&
                   target.tagName !== 'BUTTON' &&
                   target.tagName !== 'A') {
                if (target.parentElement && target.parentElement.getAttribute('role') === 'menuitem') {
                    target = target.parentElement;
                    break;
                }
                target = target.parentElement;
            }
            if (!target || target === document.body) target = el;

            target.focus();
            const opts = { bubbles: true, cancelable: true, view: window };
            target.dispatchEvent(new MouseEvent('pointerdown', opts));
            target.dispatchEvent(new MouseEvent('mousedown', opts));
            target.dispatchEvent(new MouseEvent('pointerup', opts));
            target.dispatchEvent(new MouseEvent('mouseup', opts));
            target.dispatchEvent(new MouseEvent('click', opts));
            if (typeof target.click === 'function') target.click();
            return JSON.stringify({ clicked: true, text: txt });
        }
    }
    return JSON.stringify({ clicked: false });
})();";
                    await CopilotWebView.CoreWebView2.ExecuteScriptAsync(clickUploadAfterOpenScript);
                }

                await Task.Delay(350);

                // Check for any <input type="file"> in DOM or shadow roots
                string checkInputsScript = @"
(function() {
    function queryAllDeep(selector, root) {
        root = root || document;
        let results = [];
        try {
            const matches = root.querySelectorAll(selector);
            for (let i = 0; i < matches.length; i++) results.push(matches[i]);
        } catch (_) {}
        try {
            const all = root.querySelectorAll('*');
            for (let i = 0; i < all.length; i++) {
                if (all[i].shadowRoot) results = results.concat(queryAllDeep(selector, all[i].shadowRoot));
            }
        } catch (_) {}
        try {
            const iframes = root.querySelectorAll('iframe');
            for (let i = 0; i < iframes.length; i++) {
                try {
                    const doc = iframes[i].contentDocument || iframes[i].contentWindow?.document;
                    if (doc) results = results.concat(queryAllDeep(selector, doc));
                } catch (_) {}
            }
        } catch (_) {}
        return results;
    }

    const inputs = queryAllDeep('input[type=""file""]');
    if (inputs && inputs.length > 0) {
        window.__documentReviewFileInput = inputs[inputs.length - 1];
        return JSON.stringify({ found: true, count: inputs.length });
    }
    return JSON.stringify({ found: false });
})();";

                string checkInputsRes = await CopilotWebView.CoreWebView2.ExecuteScriptAsync(checkInputsScript);
                if (checkInputsRes.Contains("\"found\":true"))
                {
                    try
                    {
                        string evalRes = await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                            "Runtime.evaluate",
                            "{\"expression\":\"window.__documentReviewFileInput\",\"returnByValue\":false}");

                        using var doc = JsonDocument.Parse(evalRes);
                        if (doc.RootElement.TryGetProperty("result", out var resObj) &&
                            resObj.TryGetProperty("objectId", out var objIdProp))
                        {
                            string objectId = objIdProp.GetString()!;
                            string setFileParams = JsonSerializer.Serialize(new
                            {
                                files = new[] { filePath },
                                objectId = objectId
                            });

                            await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("DOM.setFileInputFiles", setFileParams);

                            await CopilotWebView.CoreWebView2.ExecuteScriptAsync(@"
(function() {
    if (window.__documentReviewFileInput) {
        window.__documentReviewFileInput.dispatchEvent(new Event('input', { bubbles: true }));
        window.__documentReviewFileInput.dispatchEvent(new Event('change', { bubbles: true }));
    }
})();");
                        }
                    }
                    catch
                    {
                    }
                }

                // Focus the chat input box and dispatch Ctrl+V via CDP keyboard events
                string focusInputScript = @"
(function() {
    function queryAllDeep(selector, root) {
        root = root || document;
        let results = [];
        try {
            const matches = root.querySelectorAll(selector);
            for (let i = 0; i < matches.length; i++) results.push(matches[i]);
        } catch (_) {}
        try {
            const all = root.querySelectorAll('*');
            for (let i = 0; i < all.length; i++) {
                if (all[i].shadowRoot) results = results.concat(queryAllDeep(selector, all[i].shadowRoot));
            }
        } catch (_) {}
        try {
            const iframes = root.querySelectorAll('iframe');
            for (let i = 0; i < iframes.length; i++) {
                try {
                    const doc = iframes[i].contentDocument || iframes[i].contentWindow?.document;
                    if (doc) results = results.concat(queryAllDeep(selector, doc));
                } catch (_) {}
            }
        } catch (_) {}
        return results;
    }
    const inputs = queryAllDeep('div[contenteditable=""true""], div[role=""textbox""], textarea');
    for (let i = 0; i < inputs.length; i++) {
        try {
            inputs[i].focus();
            return true;
        } catch (_) {}
    }
    return false;
})();";
                await CopilotWebView.CoreWebView2.ExecuteScriptAsync(focusInputScript);

                try
                {
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                        "Input.dispatchKeyEvent",
                        "{\"type\":\"rawKeyDown\",\"windowsVirtualKeyCode\":17,\"modifiers\":2}");
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                        "Input.dispatchKeyEvent",
                        "{\"type\":\"keyDown\",\"windowsVirtualKeyCode\":86,\"modifiers\":2,\"text\":\"\\u0016\"}");
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                        "Input.dispatchKeyEvent",
                        "{\"type\":\"keyUp\",\"windowsVirtualKeyCode\":86,\"modifiers\":2}");
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                        "Input.dispatchKeyEvent",
                        "{\"type\":\"keyUp\",\"windowsVirtualKeyCode\":17,\"modifiers\":0}");
                }
                catch
                {
                }

                CopilotStatusText.Text = $"Attached {fileName}! (Ready in chat)";
                return true;
            }
            catch
            {
                CopilotStatusText.Text = $"{fileName} copied! (Paste with Ctrl+V into chat)";
                return true;
            }
        }

        private async void CopilotAuditButton_Click(object sender, RoutedEventArgs e)
        {
            var storedFiles = GetStoredFileNames();
            if (storedFiles.Count > 0)
            {
                await AttachFileToCopilotAsync();
                await Task.Delay(400);
            }

            string prompt = CopilotContextBuilder.BuildAuditPrompt(
                BuildHeaderInfo(),
                BuildDateFieldsInfo(),
                _dateFieldConflicts,
                storedFiles);

            await SendPromptToCopilotAsync(prompt, "Auditing file and filling fields...");
            _ = MonitorAndAutoApplyCopilotUpdatesAsync();
        }

        private async void CopilotSummarizeButton_Click(object sender, RoutedEventArgs e)
        {
            var storedFiles = GetStoredFileNames();
            if (storedFiles.Count > 0)
            {
                await AttachFileToCopilotAsync();
                await Task.Delay(400);
            }

            string prompt = CopilotContextBuilder.BuildSummarizePrompt(
                BuildHeaderInfo(),
                BuildDateFieldsInfo(),
                _dateFieldConflicts,
                storedFiles);

            await SendPromptToCopilotAsync(prompt, "Generating driver summarization...");
        }

        private async void CopilotDraftRejectEmailButton_Click(object sender, RoutedEventArgs e)
        {
            var storedFiles = GetStoredFileNames();
            if (storedFiles.Count > 0)
            {
                await AttachFileToCopilotAsync();
                await Task.Delay(400);
            }

            string prompt = CopilotContextBuilder.BuildDraftRejectEmailPrompt(
                BuildHeaderInfo(),
                BuildDateFieldsInfo(),
                _dateFieldConflicts,
                storedFiles);

            await SendPromptToCopilotAsync(prompt, "Drafting concise rejection email...");
        }

        private async void CopilotGatherDatesButton_Click(object sender, RoutedEventArgs e)
        {
            var storedFiles = GetStoredFileNames();
            if (storedFiles.Count > 0)
            {
                await AttachFileToCopilotAsync();
                await Task.Delay(400);
            }

            string prompt = CopilotContextBuilder.BuildGatherDatesPrompt(
                BuildHeaderInfo(),
                BuildDateFieldsInfo(),
                _dateFieldConflicts,
                storedFiles);

            await SendPromptToCopilotAsync(prompt, "Gathering all qualification dates...");
            _ = MonitorAndAutoApplyCopilotUpdatesAsync();
        }

        private void CopilotReloadButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isCopilotInitialized && CopilotWebView.CoreWebView2 != null)
            {
                CopilotStatusText.Text = "Reloading...";
                CopilotWebView.Reload();
            }
        }

        private async Task SendPromptToCopilotAsync(string prompt, string statusMsg)
        {
            // Always copy to clipboard as primary/fallback so the user has it ready
            try
            {
                Clipboard.SetText(prompt);
            }
            catch
            {
            }

            if (!_isCopilotInitialized || CopilotWebView.CoreWebView2 == null)
            {
                await EnsureCopilotInitializedAsync();
            }

            if (CopilotDrawerHost.Visibility != Visibility.Visible)
            {
                CopilotDrawerHost.Width = _copilotDrawerWidth;
                CopilotDrawerHost.Visibility = Visibility.Visible;
                CopilotToggleButton.Content = "Hide Assistant";
                UpdateDrawerMargins();
            }

            CopilotStatusText.Text = statusMsg;

            // If the chat is actively navigating, give it up to a few seconds to finish loading
            int waited = 0;
            while (_isCopilotLoading && waited < 8000)
            {
                await Task.Delay(400);
                waited += 400;
            }

            if (CopilotWebView.CoreWebView2 == null)
            {
                CopilotStatusText.Text = "Prompt copied to clipboard! (Paste with Ctrl+V into chat)";
                return;
            }

            string jsEscapedPrompt = JsonSerializer.Serialize(prompt);

            // Phase 1: Deep search across light DOM, shadow roots, and child frames for the chat input.
            // Uses bounding rectangle checks (never offsetParent, which returns null for position:fixed).
            // Focuses the element and types via execCommand('insertText') / native property setter.
            string insertScript = @"
(function() {
    const textToSend = " + jsEscapedPrompt + @";

    function isVisible(el) {
        if (!el) return false;
        try {
            const rect = el.getBoundingClientRect();
            if (rect.width <= 0 || rect.height <= 0) return false;
            const style = window.getComputedStyle(el);
            if (style.visibility === 'hidden' || style.display === 'none' || style.opacity === '0') return false;
            return true;
        } catch (_) {
            return false;
        }
    }

    function queryAllDeep(selector, root) {
        root = root || document;
        let results = [];
        try {
            const matches = root.querySelectorAll(selector);
            for (let i = 0; i < matches.length; i++) results.push(matches[i]);
        } catch (_) {}

        try {
            const all = root.querySelectorAll('*');
            for (let i = 0; i < all.length; i++) {
                if (all[i].shadowRoot) {
                    results = results.concat(queryAllDeep(selector, all[i].shadowRoot));
                }
            }
        } catch (_) {}

        try {
            const iframes = root.querySelectorAll('iframe');
            for (let i = 0; i < iframes.length; i++) {
                try {
                    const doc = iframes[i].contentDocument || iframes[i].contentWindow?.document;
                    if (doc) results = results.concat(queryAllDeep(selector, doc));
                } catch (_) {}
            }
        } catch (_) {}

        return results;
    }

    function findInput() {
        const selectors = [
            'textarea[placeholder*=""Copilot"" i]',
            'textarea[placeholder*=""message"" i]',
            'textarea[placeholder*=""Ask"" i]',
            'div[contenteditable=""true""][aria-label*=""Copilot"" i]',
            'div[contenteditable=""true""][aria-label*=""message"" i]',
            'div[contenteditable=""true""][aria-label*=""Ask"" i]',
            'div[role=""textbox""][aria-label*=""Copilot"" i]',
            'div[role=""textbox""][aria-label*=""message"" i]',
            'div[role=""textbox""][aria-label*=""Ask"" i]',
            'div[contenteditable=""true""]',
            'div[role=""textbox""]',
            'textarea#userInput',
            'textarea[data-id=""chat-input-textarea""]',
            'textarea'
        ];

        for (let s = 0; s < selectors.length; s++) {
            const elements = queryAllDeep(selectors[s]);
            for (let i = 0; i < elements.length; i++) {
                if (isVisible(elements[i])) {
                    return elements[i];
                }
            }
        }

        // Broad fallback: any element mentioning Copilot or message
        const all = queryAllDeep('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            const p = el.getAttribute('placeholder') || '';
            const a = el.getAttribute('aria-label') || '';
            if (/copilot|message|ask/i.test(p) || /copilot|message|ask/i.test(a)) {
                if (el.tagName === 'TEXTAREA' || el.getAttribute('contenteditable') === 'true' || el.getAttribute('role') === 'textbox') {
                    if (isVisible(el)) return el;
                }
            }
        }

        return null;
    }

    const input = findInput();
    if (!input) {
        return JSON.stringify({ status: 'INPUT_NOT_FOUND' });
    }

    try { input.focus(); } catch (_) {}
    try { input.click(); } catch (_) {}

    const isTextarea = input.tagName.toLowerCase() === 'textarea' || input.tagName.toLowerCase() === 'input';

    if (isTextarea) {
        try {
            const proto = Object.getPrototypeOf(input);
            const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set
                || Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set
                || Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
            if (setter) {
                setter.call(input, textToSend);
            } else {
                input.value = textToSend;
            }
        } catch (_) {
            input.value = textToSend;
        }
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    } else {
        const doc = input.ownerDocument || document;
        try {
            const sel = doc.getSelection();
            if (sel) {
                const range = doc.createRange();
                range.selectNodeContents(input);
                sel.removeAllRanges();
                sel.addRange(range);
            }
            doc.execCommand('selectAll', false, null);
            doc.execCommand('delete', false, null);
        } catch (_) {}

        try {
            const lines = textToSend.split('\n');
            for (let i = 0; i < lines.length; i++) {
                if (lines[i].length > 0) {
                    doc.execCommand('insertText', false, lines[i]);
                }
                if (i < lines.length - 1) {
                    const brOk = doc.execCommand('insertLineBreak');
                    if (!brOk) doc.execCommand('insertHTML', false, '<br>');
                }
            }
        } catch (_) {}

        const curText = (input.innerText || input.textContent || '').trim();
        if (!curText) {
            try {
                input.textContent = textToSend;
            } catch (_) {}
        }

        try {
            input.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: textToSend }));
        } catch (_) {}
    }

    const finalLength = (input.value || input.innerText || input.textContent || '').trim().length;

    return JSON.stringify({
        status: 'INPUT_FOCUSED',
        inputTag: input.tagName,
        contentLength: finalLength
    });
})();";

            try
            {
                bool foundInput = false;
                int contentLength = 0;

                // Attempt input detection up to 3 times in case the React component is mounting
                for (int attempt = 0; attempt < 3; attempt++)
                {
                    string rawInsertResult = await CopilotWebView.CoreWebView2.ExecuteScriptAsync(insertScript);
                    try
                    {
                        string unescaped = JsonSerializer.Deserialize<string>(rawInsertResult) ?? rawInsertResult;
                        using var doc = JsonDocument.Parse(unescaped);
                        if (doc.RootElement.TryGetProperty("status", out var st) && st.GetString() == "INPUT_FOCUSED")
                        {
                            foundInput = true;
                            if (doc.RootElement.TryGetProperty("contentLength", out var cl))
                            {
                                contentLength = cl.GetInt32();
                            }
                            break;
                        }
                    }
                    catch
                    {
                        if (rawInsertResult.Contains("INPUT_FOCUSED"))
                        {
                            foundInput = true;
                            break;
                        }
                    }

                    await Task.Delay(500);
                }

                if (!foundInput)
                {
                    CopilotStatusText.Text = "Prompt copied to clipboard! (Paste with Ctrl+V into chat)";
                    return;
                }

                // If JS text insertion registered 0 characters, type via Chrome DevTools Protocol Input.insertText
                if (contentLength == 0)
                {
                    try
                    {
                        string cdpTextParams = JsonSerializer.Serialize(new { text = prompt });
                        await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync("Input.insertText", cdpTextParams);
                    }
                    catch
                    {
                    }
                }

                // Allow 300ms for React/Lexical state updates to display the Send button (replaces the mic icon)
                await Task.Delay(300);

                // Phase 2: Find and click the Send button, or dispatch Enter
                string clickSendScript = @"
(function() {
    function isVisible(el) {
        if (!el) return false;
        try {
            const rect = el.getBoundingClientRect();
            if (rect.width <= 0 || rect.height <= 0) return false;
            const style = window.getComputedStyle(el);
            if (style.visibility === 'hidden' || style.display === 'none' || style.opacity === '0') return false;
            return true;
        } catch (_) {
            return false;
        }
    }

    function queryAllDeep(selector, root) {
        root = root || document;
        let results = [];
        try {
            const matches = root.querySelectorAll(selector);
            for (let i = 0; i < matches.length; i++) results.push(matches[i]);
        } catch (_) {}
        try {
            const all = root.querySelectorAll('*');
            for (let i = 0; i < all.length; i++) {
                if (all[i].shadowRoot) results = results.concat(queryAllDeep(selector, all[i].shadowRoot));
            }
        } catch (_) {}
        try {
            const iframes = root.querySelectorAll('iframe');
            for (let i = 0; i < iframes.length; i++) {
                try {
                    const doc = iframes[i].contentDocument || iframes[i].contentWindow?.document;
                    if (doc) results = results.concat(queryAllDeep(selector, doc));
                } catch (_) {}
            }
        } catch (_) {}
        return results;
    }

    const sendSelectors = [
        'button[aria-label*=""Send"" i]',
        'button[aria-label*=""Submit"" i]',
        'button[title*=""Send"" i]',
        'button[title*=""Submit"" i]',
        'button[data-testid*=""send"" i]',
        'button[data-testid*=""submit"" i]',
        'button[aria-label*=""Prompt"" i]',
        'button.send-button'
    ];

    for (let s = 0; s < sendSelectors.length; s++) {
        const btns = queryAllDeep(sendSelectors[s]);
        for (let i = 0; i < btns.length; i++) {
            const btn = btns[i];
            if (isVisible(btn) && !btn.disabled && btn.getAttribute('aria-disabled') !== 'true') {
                btn.focus();
                btn.click();
                btn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
                btn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
                btn.dispatchEvent(new MouseEvent('click', { bubbles: true }));
                return JSON.stringify({ status: 'BUTTON_CLICKED' });
            }
        }
    }

    // Fallback: Dispatch Enter key on active element
    const active = document.activeElement;
    if (active) {
        const enterEvent = new KeyboardEvent('keydown', {
            bubbles: true,
            cancelable: true,
            key: 'Enter',
            code: 'Enter',
            keyCode: 13,
            which: 13
        });
        active.dispatchEvent(enterEvent);
        return JSON.stringify({ status: 'ENTER_DISPATCHED' });
    }

    return JSON.stringify({ status: 'BUTTON_NOT_FOUND' });
})();";

                string rawClickResult = await CopilotWebView.CoreWebView2.ExecuteScriptAsync(clickSendScript);
                bool sent = rawClickResult.Contains("BUTTON_CLICKED") || rawClickResult.Contains("ENTER_DISPATCHED");

                // Dispatch native Enter key event via Chrome DevTools Protocol to ensure submission
                try
                {
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                        "Input.dispatchKeyEvent",
                        "{\"type\":\"rawKeyDown\",\"windowsVirtualKeyCode\":13,\"nativeVirtualKeyCode\":13,\"unmodifiedText\":\"\\r\",\"text\":\"\\r\"}");
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                        "Input.dispatchKeyEvent",
                        "{\"type\":\"keyUp\",\"windowsVirtualKeyCode\":13,\"nativeVirtualKeyCode\":13,\"unmodifiedText\":\"\\r\",\"text\":\"\\r\"}");
                    sent = true;
                }
                catch
                {
                }

                // Brief pause then follow up with a secondary send-click pass if the input still has content
                await Task.Delay(400);
                try
                {
                    await CopilotWebView.CoreWebView2.ExecuteScriptAsync(clickSendScript);
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                        "Input.dispatchKeyEvent",
                        "{\"type\":\"rawKeyDown\",\"windowsVirtualKeyCode\":13,\"nativeVirtualKeyCode\":13,\"unmodifiedText\":\"\\r\",\"text\":\"\\r\"}");
                    await CopilotWebView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                        "Input.dispatchKeyEvent",
                        "{\"type\":\"keyUp\",\"windowsVirtualKeyCode\":13,\"nativeVirtualKeyCode\":13,\"unmodifiedText\":\"\\r\",\"text\":\"\\r\"}");
                }
                catch
                {
                }

                CopilotStatusText.Text = sent ? "Prompt sent to Assistant!" : "Prompt entered! (Click Send or press Enter)";
            }
            catch
            {
                CopilotStatusText.Text = "Prompt copied to clipboard! (Paste with Ctrl+V into chat)";
            }
        }

        /// <summary>
        /// Reads the latest response or update block from the Copilot chat DOM across light DOM, shadow DOM, and iframes,
        /// strictly ignoring user messages, prompt echos, and unrendered template text.
        /// </summary>
        private async Task<string> ScrapeLatestCopilotResponseAsync(string openTag = "[APP_UPDATE]", string closeTag = "[/APP_UPDATE]")
        {
            if (!_isCopilotInitialized || CopilotWebView.CoreWebView2 == null)
            {
                return string.Empty;
            }

            string script = @"
(function() {
    function queryAllDeep(selector, root) {
        root = root || document;
        let results = [];
        try {
            const matches = root.querySelectorAll(selector);
            for (let i = 0; i < matches.length; i++) results.push(matches[i]);
        } catch (_) {}
        try {
            const all = root.querySelectorAll('*');
            for (let i = 0; i < all.length; i++) {
                if (all[i].shadowRoot) results = results.concat(queryAllDeep(selector, all[i].shadowRoot));
            }
        } catch (_) {}
        try {
            const iframes = root.querySelectorAll('iframe');
            for (let i = 0; i < iframes.length; i++) {
                try {
                    const doc = iframes[i].contentDocument || iframes[i].contentWindow?.document;
                    if (doc) results = results.concat(queryAllDeep(selector, doc));
                } catch (_) {}
            }
        } catch (_) {}
        return results;
    }

    function isUserElement(el) {
        if (!el) return false;
        let cur = el;
        while (cur && cur !== document.body) {
            const role = (cur.getAttribute('role') || '').toLowerCase();
            const author = (cur.getAttribute('data-author') || cur.getAttribute('data-message-author-role') || '').toLowerCase();
            const testId = (cur.getAttribute('data-testid') || '').toLowerCase();
            const className = (typeof cur.className === 'string' ? cur.className : '').toLowerCase();
            const tag = cur.tagName ? cur.tagName.toLowerCase() : '';

            if (tag === 'textarea' || cur.getAttribute('contenteditable') === 'true') return true;
            if (author === 'user' || role === 'user' || testId.includes('user-message') || className.includes('user-message')) return true;
            if (className.includes('chat-input') || testId.includes('chat-input')) return true;
            cur = cur.parentElement || (cur.getRootNode && cur.getRootNode().host);
        }
        return false;
    }

    function isPromptOrTemplate(text) {
        if (!text) return true;
        const lower = text.toLowerCase();
        if (lower.includes('document review - findings and data extraction') ||
            lower.includes('=== app update block') ||
            lower.includes('extract dates and history') ||
            lower.includes('output everything in the strict [app_update]') ||
            lower.includes('=== app update schema ===') ||
            lower.includes('<driver name>') ||
            lower.includes('line 1 name') ||
            lower.includes('john doe') ||
            lower.includes('doe trucking') ||
            lower.includes('<mm/dd/yy>') || lower.includes('<mm/dd/yyyy>')) {
            return true;
        }
        return false;
    }

    const assistantSelectors = [
        '[data-content=""ai-message""]',
        'div[data-author=""copilot""]',
        'div[data-author=""assistant""]',
        'div[data-message-author-role=""assistant""]',
        'div.response-message',
        'div.ac-container',
        'div[role=""region""]'
    ];

    // First: Look through explicit assistant message elements in reverse order
    for (let s = 0; s < assistantSelectors.length; s++) {
        const nodes = queryAllDeep(assistantSelectors[s]);
        if (nodes && nodes.length > 0) {
            for (let i = nodes.length - 1; i >= 0; i--) {
                const el = nodes[i];
                if (isUserElement(el)) continue;
                const txt = el.innerText || el.textContent || '';
                if (txt.includes('[APP_UPDATE]') && txt.includes('[/APP_UPDATE]')) {
                    if (!isPromptOrTemplate(txt)) {
                        return txt;
                    }
                }
            }
        }
    }

    // Second: Fallback to all DOM elements in reverse, skipping user elements & prompt templates
    const all = queryAllDeep('*');
    for (let i = all.length - 1; i >= 0; i--) {
        const el = all[i];
        if (isUserElement(el)) continue;
        const txt = el.innerText || el.textContent || '';
        if (txt.includes('[APP_UPDATE]') && txt.includes('[/APP_UPDATE]')) {
            if (!isPromptOrTemplate(txt)) {
                return txt;
            }
        }
    }

    return '';
})();";

            script = script.Replace("[APP_UPDATE]", openTag).Replace("[/APP_UPDATE]", closeTag);
            string scrapedJson = await CopilotWebView.CoreWebView2.ExecuteScriptAsync(script);
            try
            {
                return JsonSerializer.Deserialize<string>(scrapedJson) ?? string.Empty;
            }
            catch
            {
                return scrapedJson ?? string.Empty;
            }
        }

        /// <summary>
        /// Monitors Copilot output in the background. As soon as a structured [APP_UPDATE] block is detected,
        /// automatically copies it to the clipboard and applies it to the app without requiring manual clicks.
        /// </summary>
        private async Task MonitorAndAutoApplyCopilotUpdatesAsync()
        {
            _autoApplyCts?.Cancel();
            _autoApplyCts = new CancellationTokenSource();
            var token = _autoApplyCts.Token;

            try
            {
                // Wait 4 seconds for Copilot to begin streaming/generating its reply
                await Task.Delay(4000, token);

                for (int i = 0; i < 40; i++)
                {
                    if (token.IsCancellationRequested) return;

                    if (_isCopilotInitialized && CopilotWebView.CoreWebView2 != null)
                    {
                        string response = await ScrapeLatestCopilotResponseAsync();
                        if (!string.IsNullOrWhiteSpace(response) &&
                            response.Contains("[APP_UPDATE]") &&
                            response.Contains("[/APP_UPDATE]"))
                        {
                            var updates = CopilotContextBuilder.ParseSuggestedUpdates(response);
                            if (updates.TotalCount > 0)
                            {
                                try
                                {
                                    Clipboard.SetText(response);
                                }
                                catch
                                {
                                }

                                ApplySuggestedFieldsAndDates(updates);
                                CopilotStatusText.Text = $"⚡ Auto-applied {updates.TotalCount} field(s) & date(s)!";
                                return;
                            }
                        }
                    }

                    await Task.Delay(1500, token);
                }
            }
            catch (TaskCanceledException)
            {
            }
            catch
            {
            }
        }

        private async void CopilotApplyDatesButton_Click(object sender, RoutedEventArgs e)
        {
            string responseText = string.Empty;
            if (_isCopilotInitialized && CopilotWebView.CoreWebView2 != null)
            {
                CopilotStatusText.Text = "Reading Assistant output...";
                responseText = await ScrapeLatestCopilotResponseAsync();
            }

            var updates = CopilotContextBuilder.ParseSuggestedUpdates(responseText);
            if (updates.TotalCount == 0 && Clipboard.ContainsText())
            {
                try
                {
                    string clip = Clipboard.GetText();
                    updates = CopilotContextBuilder.ParseSuggestedUpdates(clip);
                }
                catch
                {
                }
            }

            if (updates.TotalCount == 0)
            {
                CopilotStatusText.Text = "No date updates detected";
                MessageBox.Show(
                    this,
                    "No structured updates were found in the Assistant's response or clipboard.\n\nClick 'Audit File & Report Issues' or 'Gather All Dates', or copy the [APP_UPDATE] text block to the clipboard and click this button again.",
                    "No Updates Found",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            // Sync clipboard automatically so updates flow out cleanly
            if (!string.IsNullOrWhiteSpace(responseText))
            {
                try
                {
                    Clipboard.SetText(responseText);
                }
                catch
                {
                }
            }

            ApplySuggestedFieldsAndDates(updates);
            CopilotStatusText.Text = $"Applied {updates.TotalCount} field(s) & date(s) to file!";
        }

        private void ApplySuggestedFieldsAndDates(CopilotExtractedData data)
        {
            // Apply header fields to top bar if present
            bool updatedHeader = false;
            foreach (var (k, v) in data.HeaderFields)
            {
                if (string.IsNullOrWhiteSpace(v)) continue;
                switch (k.ToLowerInvariant())
                {
                    case "drivername": DriverNameTextBox.Text = v; updatedHeader = true; break;
                    case "contractorname": ContractorNameTextBox.Text = v; updatedHeader = true; break;
                    case "trucknumber": TruckNumberTextBox.Text = v; updatedHeader = true; break;
                    case "trailernumber": TrailerNumberTextBox.Text = v; updatedHeader = true; break;
                    case "ssn": SsnTextBox.Text = v; updatedHeader = true; break;
                    case "fid": FidTextBox.Text = v; updatedHeader = true; break;
                    case "phone": PhoneTextBox.Text = v; updatedHeader = true; break;
                }
            }

            if (updatedHeader && FileDetailsPanel.Visibility != Visibility.Visible)
            {
                FileDetailsPanel.Visibility = Visibility.Visible;
                FileDetailsToggleButton.Content = "Hide File Details";
            }

            // Apply file dates to date bar
            if (data.Dates.Count > 0)
            {
                ApplySuggestedDatesToBoxes(data.Dates);
            }

            // Apply driver-reported employment history and tracker-only dates to the shared tracker
            if (_mainViewModel != null)
            {
                var t = _mainViewModel.Tracker;
                if (data.Dates.TryGetValue("PspRan", out var pr)) t.PspRan = pr;
                if (data.Dates.TryGetValue("CdlisDac", out var cd)) t.CdlisDac = cd;
                if (data.Employment.Count > 0)
                {
                    t.Rows.Clear();
                    foreach (var r in data.Employment)
                        t.Rows.Add(new EmploymentRow { Description = r.Description, IsOtr = r.IsOtr, Start = r.Start, End = r.End });
                    while (t.Rows.Count < 12) t.Rows.Add(new EmploymentRow());
                }
                PullTrackerIntoBoxes();
            }
        }

        private void ApplySuggestedDatesToBoxes(Dictionary<string, string> dates)
        {
            foreach (var (k, v) in dates)
            {
                if (string.IsNullOrWhiteSpace(v)) continue;
                switch (k.ToLowerInvariant())
                {
                    case "certification": CertificationDateBox.Text = v; break;
                    case "limitedconsent": LimitedConsentDateBox.Text = v; break;
                    case "fullconsent": FullConsentDateBox.Text = v; break;
                    case "drugtestran": DrugTestRanDateBox.Text = v; break;
                    case "drugtestreceived": DrugTestReceivedDateBox.Text = v; break;
                    case "cdlexpiration": CdlExpirationDateBox.Text = v; break;
                    case "mvrcdlexpiration": MvrCdlExpirationDateBox.Text = v; break;
                    case "mvrlastreview": MvrLastReviewDateBox.Text = v; break;
                    case "mvrnextreview": MvrNextReviewDateBox.Text = v; break;
                    case "physicallast": PhysicalLastDateBox.Text = v; break;
                    case "physicalnext": PhysicalNextDateBox.Text = v; break;
                    case "fingerprintdate": FingerprintDateBox.Text = v; break;
                    case "hazmatexpiration": HazmatExpirationDateBox.Text = v; break;
                    case "latestjobev": LatestJobEvDateBox.Text = v; break;
                    case "sevendaylog": SevenDayLogDateBox.Text = v; break;
                    case "statementoflease": if (_mode == QualMode.FullUnit) StatementOfLeaseDateBox.Text = v; break;
                    case "statementofleasetrailer": if (_mode == QualMode.FullUnit) StatementOfLeaseTrailerDateBox.Text = v; break;
                    case "dayqualified": DayQualifiedDateBox.Text = v; break;
                    case "pspsnapshotdate":
                        PspSnapshotDateBox.Text = v;
                        if (MainViewModel != null && DateTime.TryParse(v, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                        {
                            MainViewModel.CurrentPspSnapshotDate = dt;
                        }
                        break;
                }
            }

            if (FileDatesPanel.Visibility != Visibility.Visible)
            {
                FileDatesPanel.Visibility = Visibility.Visible;
                FileDatesToggleButton.Content = "Hide File Dates";
            }

            RevalidateAllDateFields();
        }

        private List<string> GetStoredFileNames()
        {
            var list = new List<string>();
            foreach (var item in _items)
            {
                if (!string.IsNullOrWhiteSpace(item.FilePath))
                {
                    list.Add(System.IO.Path.GetFileName(item.FilePath));
                }
            }
            if (HasWordDocLoaded && !string.IsNullOrWhiteSpace(_wordDocFileName))
            {
                list.Add(_wordDocFileName);
            }
            if (HasExcelLoaded && !string.IsNullOrWhiteSpace(_excelDocFileName))
            {
                list.Add(_excelDocFileName);
            }
            return list;
        }

        /// <summary>
        /// Reverts the workspace window back to the main dashboard without
        /// touching the loaded items - they remain in memory for next time.
        /// </summary>
        private void HideWorkspaceButton_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }

        // ======================= Window close warning =======================

        private void WorkspaceWindow_Closing(object? sender, CancelEventArgs e)
        {
            if (_allowHide)
            {
                // Programmatic shutdown (application exiting); allow it through unchanged.
                return;
            }

            // Every native close attempt (title bar X, Alt+F4) is intercepted:
            // we never truly destroy this window instance so it can be reused
            // the next time the user clicks "Workspace" on the main dashboard.
            e.Cancel = true;

            var result = MessageBox.Show(
                this,
                "Are you sure you want to close the workspace? Save first or your work will be lost.",
                "Close Workspace",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (result == MessageBoxResult.Yes)
            {
                ResetWorkspaceState();
                Hide();
            }
        }

        /// <summary>Clears everything (canvas, fields, tracker, docs, drawers, modes) back to a fresh workspace.</summary>
        private void ResetWorkspaceState()
        {
            {
                _items.Clear();
                _autoAddedDownloads.Clear();
                ApplyHeaderInfo(new WorkspaceHeaderInfo());
                ApplyDateFieldsInfo(new WorkspaceDateFieldsInfo());
                ApplyMode(QualMode.DriverFile);
                _mainViewModel?.Tracker.Clear();
                ApplyWordDoc(null, string.Empty);
                ApplyExcelDoc(null, string.Empty);
                if (CopilotDrawerHost.Visibility == Visibility.Visible)
                {
                    CopilotDrawerHost.Width = 0;
                    CopilotDrawerHost.Visibility = Visibility.Collapsed;
                    CopilotToggleButton.Content = "Assistant";
                }
                if (BrowserDrawerHost.Visibility == Visibility.Visible)
                {
                    BrowserDrawerHost.Width = 0;
                    BrowserDrawerHost.Visibility = Visibility.Collapsed;
                    BrowserToggleButton.Content = "Browser";
                }
                if (ExcelDrawerHost.Visibility == Visibility.Visible)
                {
                    ExcelDrawerHost.Width = 0;
                    ExcelDrawerHost.Visibility = Visibility.Collapsed;
                    SheetToggleButton.Content = "View/Edit Sheet";
                }
                if (WordDocDrawerHost.Visibility == Visibility.Visible)
                {
                    WordDocDrawerHost.Width = 0;
                    WordDocDrawerHost.Visibility = Visibility.Collapsed;
                    NotesToggleButton.Content = "View/Edit Doc";
                }
                UpdateDrawerMargins();
            }
        }

        private void NewWorkspaceButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                this,
                "Start a new workspace? This clears everything. Save first or your work will be lost.",
                "New Workspace",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);
            if (result == MessageBoxResult.Yes)
            {
                ResetWorkspaceState();
            }
        }

        private void AutosaveTimer_Tick(object? sender, EventArgs e)
        {
            string? path = WorkspaceFileService.CreateAutosaveBackup(_items, string.Empty, BuildHeaderInfo(), BuildDateFieldsInfo(), BuildWordDocBytes(), _wordDocFileName, BuildExcelDocBytes(), _excelDocFileName);
            AutosaveStatusText.Text = path is null
                ? "Autosave: every 30 min"
                : $"Last autosave: {DateTime.Now:h:mm tt}";
        }

        /// <summary>Called by the host application during real shutdown so the Closing handler doesn't re-prompt.</summary>
        public void AllowRealClose()
        {
            _allowHide = true;
            _autosaveTimer.Stop();
            _quickDateTimer.Stop();
            _wordDocPopout?.Close();
            _wordController?.Dispose();
            _wordController = null;
            _excelSheetPopout?.Close();
            _excelController?.Dispose();
            _excelController = null;
            if (_browserTabs != null) _browserTabs.DisposeAll(); else BrowserWebView?.Dispose();
            CopilotWebView?.Dispose();
        }
    }
}
