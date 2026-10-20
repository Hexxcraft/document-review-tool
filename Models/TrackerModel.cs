using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

namespace HexxcorpDocumentReviewUtility.Models
{
    /// <summary>Serializable snapshot of the Dates &amp; Employment History tracker (config.json and .lqw).</summary>
    public class TrackerData
    {
        public string DtDate { get; set; } = string.Empty;
        public string MvrRan { get; set; } = string.Empty;
        public string FpRan { get; set; } = string.Empty;
        public string PspSnapshot { get; set; } = string.Empty;
        public string PspRan { get; set; } = string.Empty;
        public string PhysicalExp { get; set; } = string.Empty;
        public string Certification { get; set; } = string.Empty;
        public string CdlisDac { get; set; } = string.Empty;
        public string LimitedConsent { get; set; } = string.Empty;
        public string FullConsent { get; set; } = string.Empty;
        public string Program { get; set; } = "All History";
        public List<TrackerRowData> Rows { get; set; } = new();
    }

    public class TrackerRowData
    {
        public string Description { get; set; } = string.Empty;
        public bool IsOtr { get; set; }
        public string Start { get; set; } = string.Empty;
        public string End { get; set; } = string.Empty;
    }

    /// <summary>Shared date helpers for the tracker (strict MM/dd/yyyy output, lenient input).</summary>
    public static class TrackerDates
    {
        public const string Format = "MM/dd/yyyy";

        private static readonly string[] Accepted =
        {
            "MM/dd/yyyy", "M/d/yyyy", "MM/d/yyyy", "M/dd/yyyy",
            "MM/dd/yy", "M/d/yy", "MM-dd-yyyy", "M-d-yyyy", "MM-dd-yy", "M-d-yy",
            "yyyy-MM-dd", "yyyy/MM/dd", "MMddyyyy", "MMddyy", "M/d/yyyy h:mm:ss tt", "M/d/yyyy H:mm:ss"
        };

        public static bool TryParse(string? text, out DateTime date)
        {
            date = default;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string t = text.Trim();
            if (DateTime.TryParseExact(t, Accepted, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) return true;
            return DateTime.TryParse(t, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
        }

        public static string Normalize(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            return TryParse(text, out DateTime d) ? d.ToString(Format, CultureInfo.InvariantCulture) : text.Trim();
        }

        public static string Show(DateTime? d) => d.HasValue ? d.Value.ToString(Format, CultureInfo.InvariantCulture) : string.Empty;

        /// <summary>Whole calendar months between two dates (same semantics as Excel DATEDIF "m").</summary>
        public static int WholeMonths(DateTime start, DateTime end)
        {
            if (end < start) return 0;
            int months = (end.Year - start.Year) * 12 + end.Month - start.Month;
            if (end.Day < start.Day) months--;
            return Math.Max(0, months);
        }

        /// <summary>
        /// Months covered by the union of the given intervals, so overlapping jobs are only counted once.
        /// Overlapping or back-to-back intervals are merged before whole months are counted.
        /// </summary>
        public static int UnionMonths(IEnumerable<(DateTime Start, DateTime End)> intervals)
        {
            var list = intervals.Where(i => i.End >= i.Start).OrderBy(i => i.Start).ToList();
            if (list.Count == 0) return 0;
            int total = 0;
            DateTime curStart = list[0].Start, curEnd = list[0].End;
            for (int i = 1; i < list.Count; i++)
            {
                if (list[i].Start <= curEnd.AddDays(1))
                {
                    if (list[i].End > curEnd) curEnd = list[i].End;
                }
                else
                {
                    total += WholeMonths(curStart, curEnd);
                    curStart = list[i].Start;
                    curEnd = list[i].End;
                }
            }
            total += WholeMonths(curStart, curEnd);
            return total;
        }
    }

    /// <summary>One employment line in the tracker.</summary>
    public class EmploymentRow : INotifyPropertyChanged
    {
        private string _description = string.Empty;
        private bool _isOtr;
        private string _start = string.Empty;
        private string _end = string.Empty;

        public string Description { get => _description; set { _description = value ?? string.Empty; On(); } }
        public bool IsOtr { get => _isOtr; set { _isOtr = value; On(); On(nameof(OtrMonths)); } }
        public string Start { get => _start; set { _start = TrackerDates.Normalize(value); On(); On(nameof(Months)); On(nameof(OtrMonths)); } }
        public string End { get => _end; set { _end = TrackerDates.Normalize(value); On(); On(nameof(Months)); On(nameof(OtrMonths)); } }

        public DateTime? StartDate => TrackerDates.TryParse(_start, out DateTime d) ? d : null;
        public DateTime? EndDate => TrackerDates.TryParse(_end, out DateTime d) ? d : null;

        /// <summary>This line's own whole months (not de-duplicated against other jobs).</summary>
        public int Months => StartDate.HasValue && EndDate.HasValue ? TrackerDates.WholeMonths(StartDate.Value, EndDate.Value) : 0;
        public int OtrMonths => _isOtr ? Months : 0;

        public bool IsBlank => string.IsNullOrWhiteSpace(_description) && string.IsNullOrWhiteSpace(_start) && string.IsNullOrWhiteSpace(_end);

        public event PropertyChangedEventHandler? PropertyChanged;
        private void On([CallerMemberName] string? n = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    /// <summary>
    /// The single shared Dates &amp; Employment History tracker. The main-window drawer, the workspace drawer and
    /// the File Dates boxes all read and write this one instance.
    /// </summary>
    public class TrackerModel : INotifyPropertyChanged
    {
        public const int DefaultRowCount = 12;

        private string _dtDate = "", _mvrRan = "", _fpRan = "", _pspSnapshot = "", _pspRan = "", _physicalExp = "";
        private string _certification = "", _cdlisDac = "", _limitedConsent = "", _fullConsent = "";

        public TrackerModel()
        {
            Rows.CollectionChanged += (_, e) =>
            {
                if (e.NewItems != null) foreach (EmploymentRow r in e.NewItems) r.PropertyChanged += Row_PropertyChanged;
                if (e.OldItems != null) foreach (EmploymentRow r in e.OldItems) r.PropertyChanged -= Row_PropertyChanged;
                RaiseTotals();
            };
            for (int i = 0; i < DefaultRowCount; i++) Rows.Add(new EmploymentRow());
        }

        public ObservableCollection<EmploymentRow> Rows { get; } = new();

        private DateTime? _calcTargetDate = DateTime.Today;
        private int _calcDayOffset = 30;

        /// <summary>Interactive date calculator: target date (shared by the main window and workspace).</summary>
        public DateTime? CalcTargetDate
        {
            get => _calcTargetDate;
            set { _calcTargetDate = value; Raise(nameof(CalcTargetDate)); Raise(nameof(CalcResultText)); }
        }

        /// <summary>Interactive date calculator: signed day offset.</summary>
        public int CalcDayOffset
        {
            get => _calcDayOffset;
            set { _calcDayOffset = value; Raise(nameof(CalcDayOffset)); Raise(nameof(CalcResultText)); }
        }

        public string CalcResultText => _calcTargetDate.HasValue
            ? _calcTargetDate.Value.AddDays(_calcDayOffset).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)
            : "Select a target date";

        public string DtDate { get => _dtDate; set => Set(ref _dtDate, value); }
        public string MvrRan { get => _mvrRan; set => Set(ref _mvrRan, value); }
        public string FpRan { get => _fpRan; set => Set(ref _fpRan, value); }
        public string PspSnapshot { get => _pspSnapshot; set => Set(ref _pspSnapshot, value); }
        public string PspRan { get => _pspRan; set => Set(ref _pspRan, value); }
        public string PhysicalExp { get => _physicalExp; set => Set(ref _physicalExp, value); }
        public string Certification { get => _certification; set => Set(ref _certification, value); }
        public string CdlisDac { get => _cdlisDac; set => Set(ref _cdlisDac, value); }
        public string LimitedConsent { get => _limitedConsent; set => Set(ref _limitedConsent, value); }
        public string FullConsent { get => _fullConsent; set => Set(ref _fullConsent, value); }

        /// <summary>The date used to cap ongoing history entries during a retrospective review.</summary>
        public static Func<DateTime> AsOf { get; set; } = () => DateTime.Today;

        public DateTime WindowStart => DateTime.MinValue;

        /// <summary>All entered intervals are included; overlapping dates count once.</summary>
        private IEnumerable<(DateTime, DateTime)> Intervals(bool otrOnly)
        {
            DateTime asOf = AsOf();
            return Rows.Where(r => r.StartDate.HasValue && r.EndDate.HasValue && (!otrOnly || r.IsOtr))
                .Select(r => (Start: r.StartDate!.Value, End: r.EndDate!.Value))
                .Where(i => i.Start <= asOf)
                .Select(i => (i.Start, i.End > asOf ? asOf : i.End));
        }

        public bool IsOutsideWindow(DateTime jobEnd) => false;

        /// <summary>Re-raises every computed total (call when the as-of date changes).</summary>
        public void Refresh() => RaiseTotals();

        /// <summary>Months of entered work history, with overlapping jobs counted once.</summary>
        public int TotalMonths => TrackerDates.UnionMonths(Intervals(false));

        /// <summary>OTR months across entered history, with overlapping jobs counted once.</summary>
        public int OtrMonths => TrackerDates.UnionMonths(Intervals(true));

        /// <summary>Simple sum of every in-window line (what the old paper sheet did) - shown for comparison.</summary>
        public int RawSumMonths => Intervals(false).Sum(i => TrackerDates.WholeMonths(i.Item1, i.Item2));

        /// <summary>OTR months inside the selected program's window.</summary>
        public int OtrLast36Months => OtrMonths;
        public string TotalText => Describe(TotalMonths);
        public string OtrText => Describe(OtrMonths);
        public static readonly string[] Programs = { "All History" };

        private string _program = "All History";
        public string Program
        {
            get => _program;
            set
            {
                string v = "All History";
                if (_program == v) return;
                _program = v;
                Raise(nameof(Program));
                RaiseTotals();
            }
        }

        public int RequiredOtrMonths => 0;
        public int WindowMonths => 0;

        /// <summary>OTR months (overlap-safe) across all entered history.</summary>
        public int OtrInWindow => OtrMonths;

        public string WindowLabel => "All available history";
        public string OtrLast36Text => Describe(OtrInWindow);
        public bool MeetsOtrMinimum => true;
        public bool BelowOtrMinimum => false;

        public string OverlapNote
        {
            get
            {
                int diff = RawSumMonths - TotalMonths;
                return diff > 0 ? $"Overlapping jobs removed: {diff} duplicate month(s) not counted." : "No overlapping jobs detected.";
            }
        }

        /// <summary>Latest employment end date across all lines.</summary>
        public DateTime? LatestJobEnd
        {
            get
            {
                DateTime? best = null;
                foreach (var r in Rows)
                    if (r.EndDate is DateTime e && (best == null || e > best)) best = e;
                return best;
            }
        }

        public string LatestJobEndText => TrackerDates.Show(LatestJobEnd);

        private static string Describe(int months) => $"{months} mo ({months / 12} yr {months % 12} mo)";

        private void Row_PropertyChanged(object? s, PropertyChangedEventArgs e) => RaiseTotals();

        private void RaiseTotals()
        {
            foreach (string n in new[] { nameof(TotalMonths), nameof(OtrMonths), nameof(RawSumMonths), nameof(OtrLast36Months), nameof(TotalText),
                nameof(OtrText), nameof(OtrLast36Text), nameof(MeetsOtrMinimum), nameof(BelowOtrMinimum), nameof(OtrInWindow), nameof(WindowLabel), nameof(RequiredOtrMonths), nameof(WindowMonths), nameof(OverlapNote), nameof(LatestJobEnd), nameof(LatestJobEndText) })
            {
                Raise(n);
            }
        }

        private void Set(ref string field, string? value, string? alsoRaise = null, [CallerMemberName] string? name = null)
        {
            string v = TrackerDates.Normalize(value);
            if (field == v) return;
            field = v;
            Raise(name);
            if (alsoRaise != null) Raise(alsoRaise);
        }

        public TrackerData ToData()
        {
            var d = new TrackerData
            {
                DtDate = DtDate, MvrRan = MvrRan, FpRan = FpRan, PspSnapshot = PspSnapshot, PspRan = PspRan, PhysicalExp = PhysicalExp,
                Certification = Certification, CdlisDac = CdlisDac, LimitedConsent = LimitedConsent, FullConsent = FullConsent,
                                Program = Program
                            };
            foreach (var r in Rows.Where(r => !r.IsBlank))
            {
                d.Rows.Add(new TrackerRowData { Description = r.Description, IsOtr = r.IsOtr, Start = r.Start, End = r.End });
            }
            return d;
        }

        public void Load(TrackerData? d)
        {
            d ??= new TrackerData();
            DtDate = d.DtDate; MvrRan = d.MvrRan; FpRan = d.FpRan; PspSnapshot = d.PspSnapshot; PspRan = d.PspRan; PhysicalExp = d.PhysicalExp;
            Certification = d.Certification; CdlisDac = d.CdlisDac; LimitedConsent = d.LimitedConsent; FullConsent = d.FullConsent;
            Program = d.Program;
            Rows.Clear();
            foreach (var r in d.Rows)
            {
                Rows.Add(new EmploymentRow { Description = r.Description, IsOtr = r.IsOtr, Start = r.Start, End = r.End });
            }
            while (Rows.Count < DefaultRowCount) Rows.Add(new EmploymentRow());
        }

        public void Clear() => Load(new TrackerData());

        /// <summary>
        /// Pastes tab/comma-separated clipboard text into the rows starting at <paramref name="startIndex"/>.
        /// Accepts "Description, OTR?, Start, End" or "Description, Start, End"; an Excel header row is skipped.
        /// </summary>
        public int PasteRows(string text, int startIndex)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            int count = 0;
            int idx = Math.Max(0, startIndex);
            foreach (string rawLine in text.Replace("\r", "").Split('\n'))
            {
                if (string.IsNullOrWhiteSpace(rawLine)) continue;
                string[] c = rawLine.Contains('\t') ? rawLine.Split('\t') : rawLine.Split(',');
                c = c.Select(x => x.Trim().Trim('"')).ToArray();
                if (c.Length > 0 && c[0].StartsWith("Description", StringComparison.OrdinalIgnoreCase)) continue;

                string desc = c.Length > 0 ? c[0] : "";
                bool otr = false;
                string start = "", end = "";
                bool hasOtrColumn = c.Length >= 4 && !TrackerDates.TryParse(c[1], out _);
                if (hasOtrColumn)
                {
                    otr = c[1].StartsWith("y", StringComparison.OrdinalIgnoreCase) || c[1].Equals("true", StringComparison.OrdinalIgnoreCase) || c[1] == "1";
                    start = c[2]; end = c[3];
                }
                else if (c.Length >= 3)
                {
                    start = c[1]; end = c[2];
                }
                else if (c.Length == 2)
                {
                    start = c[1];
                }

                while (Rows.Count <= idx) Rows.Add(new EmploymentRow());
                var row = Rows[idx++];
                row.Description = desc; row.IsOtr = otr; row.Start = start; row.End = end;
                count++;
            }
            return count;
        }

        /// <summary>Plain-text rendering used for Copilot context, the stored-files drawer entry and export.</summary>
        public string ToPlainText()
        {
            var sb = new StringBuilder();
            string V(string s) => string.IsNullOrWhiteSpace(s) ? "(blank)" : s;
            sb.AppendLine("USER-ENTERED DATES (not verified against documents)");
            sb.AppendLine($"DT Date: {V(DtDate)}");
            sb.AppendLine($"MVR Ran: {V(MvrRan)}");
            sb.AppendLine($"Fingerprint Date: {V(FpRan)}");
            sb.AppendLine($"PSP Snapshot: {V(PspSnapshot)}");
            sb.AppendLine($"PSP Ran: {V(PspRan)}");
            sb.AppendLine($"Medical Expiration: {V(PhysicalExp)}");
            sb.AppendLine($"Certification Date: {V(Certification)}");
            sb.AppendLine($"CDLIS / DAC: {V(CdlisDac)}");
            sb.AppendLine($"Confirmation IA / Limited Consent: {V(LimitedConsent)}");
            sb.AppendLine($"Clearinghouse / Full Consent: {V(FullConsent)}");
            sb.AppendLine();
            sb.AppendLine("DRIVER-REPORTED EMPLOYMENT HISTORY (Description | OTR? | Start | End | Months)");
            int n = 0;
            foreach (var r in Rows.Where(r => !r.IsBlank))
            {
                n++;
                sb.AppendLine($"{n}. {V(r.Description)} | {(r.IsOtr ? "OTR" : "Non-OTR")} | {V(r.Start)} | {V(r.End)} | {r.Months} mo");
            }
            if (n == 0) sb.AppendLine("(no employment lines entered)");
            sb.AppendLine();
            sb.AppendLine($"Total work experience (overlaps counted once): {TotalText}");
            sb.AppendLine($"Total OTR experience (overlaps counted once): {OtrText}");
            sb.AppendLine("History scope: all entered records; no OTR minimum is configured.");
            sb.AppendLine($"OTR history total: {OtrLast36Text}");
            sb.AppendLine($"Latest job end date: {V(LatestJobEndText)}");
            sb.AppendLine(OverlapNote);
            return sb.ToString();
        }

        public bool HasAnyData => ToData() is var d &&
            (d.Rows.Count > 0 || !string.IsNullOrWhiteSpace(d.DtDate + d.MvrRan + d.FpRan + d.PspSnapshot + d.PspRan + d.PhysicalExp + d.Certification + d.CdlisDac + d.LimitedConsent + d.FullConsent));

        public event PropertyChangedEventHandler? PropertyChanged;
        private void Raise(string? n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }
}
