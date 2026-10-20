using System;

namespace HexxcorpDocumentReviewUtility.Models
{
    /// <summary>
    /// Persisted application configuration. Serialized as JSON to a
    /// <c>config.json</c> file located alongside the executable.
    /// </summary>
    public class AppConfig
    {
        /// <summary>
        /// The user-entered "Current PSP Snapshot Date" value, stored in a
        /// culture-invariant round-trippable format (yyyy-MM-dd). Null when
        /// the user has never set a value.
        /// </summary>
        public string? CurrentPspSnapshotDate { get; set; }

        /// <summary>
        /// Last target date selected in the interactive date calculator, so the
        /// tool re-opens with the operator's most recent working date.
        /// </summary>
        public string? LastTargetDate { get; set; }

        /// <summary>
        /// Last "+/- N days" offset used in the interactive date calculator.
        /// </summary>
        public int LastDayOffset { get; set; }

        /// <summary>The Dates &amp; Employment History tracker, so it survives restarts.</summary>
        public TrackerData? Tracker { get; set; }
    }
}
