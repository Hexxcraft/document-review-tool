using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace HexxcorpDocumentReviewUtility.ViewModels
{
    /// <summary>
    /// Backing view model for <see cref="Windows.MainWindow"/> (the main
    /// dashboard). Exposes the live clock, the interactive date calculator,
    /// and the persisted "Current PSP Snapshot Date" field via data binding.
    /// </summary>
    public class MainViewModel : INotifyPropertyChanged
    {
        public const string DateStorageFormat = "yyyy-MM-dd";

        /// <summary>Numeric-only display format used everywhere a date is shown to the user (no month/day names).</summary>
        public const string DateDisplayFormat = "MM/dd/yyyy";

        private bool _syncingTrackerPsp;

        /// <summary>The single shared Dates &amp; Employment History tracker used by the main drawer, the workspace and Copilot.</summary>
        public HexxcorpDocumentReviewUtility.Models.TrackerModel Tracker { get; } = new();

        public MainViewModel()
        {
            Tracker.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(Models.TrackerModel.CalcTargetDate)) OnPropertyChanged(nameof(TargetDate));
                else if (e.PropertyName == nameof(Models.TrackerModel.CalcDayOffset)) OnPropertyChanged(nameof(DayOffset));
                if (e.PropertyName != nameof(Models.TrackerModel.PspSnapshot) || _syncingTrackerPsp) return;
                _syncingTrackerPsp = true;
                try
                {
                    DateTime? parsed = Models.TrackerDates.TryParse(Tracker.PspSnapshot, out DateTime d) ? d : null;
                    if (CurrentPspSnapshotDate != parsed) CurrentPspSnapshotDate = parsed;
                }
                finally { _syncingTrackerPsp = false; }
            };
        }

        private string _currentTimeText = DateTime.Now.ToString("h:mm:ss tt", CultureInfo.InvariantCulture);
        private DateTime? _currentPspSnapshotDate;
        private bool _lookingBackward = true;

        /// <summary>Live system clock text, refreshed once per second.</summary>
        public string CurrentTimeText
        {
            get => _currentTimeText;
            set { _currentTimeText = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// True when the header's 30-day reference is in "past" mode
        /// (default - subtracts 30 days); false is "future" mode (adds 30
        /// days). Flipped by the header's Past/Future toggle button.
        /// </summary>
        public bool LookingBackward
        {
            get => _lookingBackward;
            set { _lookingBackward = value; OnPropertyChanged(); OnPropertyChanged(nameof(ThirtyDaysFromTodayText)); OnPropertyChanged(nameof(DateDirectionLabel)); }
        }

        /// <summary>Button caption for the header Past/Future toggle.</summary>
        public string DateDirectionLabel => LookingBackward ? "PAST" : "FUTURE";

        /// <summary>"30 Days From/Before Today" label, recomputed from the current system date and the Past/Future toggle.</summary>
        public string ThirtyDaysFromTodayText => LookingBackward
            ? $"30 Days Before Today: {DateTime.Today.AddDays(-30).ToString(DateDisplayFormat, CultureInfo.InvariantCulture)}"
            : $"30 Days From Today: {DateTime.Today.AddDays(30).ToString(DateDisplayFormat, CultureInfo.InvariantCulture)}";

        /// <summary>Target date selected in the interactive date calculator.</summary>
        public DateTime? TargetDate
        {
            get => Tracker.CalcTargetDate;
            set { Tracker.CalcTargetDate = value; }
        }

        /// <summary>Signed number of days to add to <see cref="TargetDate"/>.</summary>
        public int DayOffset
        {
            get => Tracker.CalcDayOffset;
            set { Tracker.CalcDayOffset = value; }
        }

        /// <summary>Live-computed result of TargetDate + DayOffset.</summary>
        public string ResultDateText => TargetDate.HasValue
            ? TargetDate.Value.AddDays(DayOffset).ToString(DateDisplayFormat, CultureInfo.InvariantCulture)
            : "Select a target date";

        /// <summary>
        /// The editable, auto-persisted "Current PSP Snapshot Date" field.
        /// </summary>
        public DateTime? CurrentPspSnapshotDate
        {
            get => _currentPspSnapshotDate;
            set
            {
                _currentPspSnapshotDate = value;
                OnPropertyChanged();
                if (!_syncingTrackerPsp)
                {
                    _syncingTrackerPsp = true;
                    try { Tracker.PspSnapshot = Models.TrackerDates.Show(value); }
                    finally { _syncingTrackerPsp = false; }
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
