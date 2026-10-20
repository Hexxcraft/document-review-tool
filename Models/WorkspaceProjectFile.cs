using System;
using System.Collections.Generic;

namespace HexxcorpDocumentReviewUtility.Models
{
    /// <summary>Which kind of file is being reviewed; changes required fields and prompts.</summary>
    public enum QualMode
    {
        DriverFile = 0,
        FullUnit = 1,
        PostReviewAudit = 2
    }

    /// <summary>
    /// Serializable representation of one workspace canvas item, used only for
    /// saving/loading project files. The source image/PDF bytes are embedded
    /// directly in <see cref="Data"/> (base64-encoded by the JSON serializer)
    /// so a saved .lqw file is fully self-contained and can be opened on any
    /// machine that has the app, with no access to the original source file.
    /// </summary>
    public class WorkspaceItemRecord
    {
        public WorkspaceItemKind Kind { get; set; }

        /// <summary>
        /// Original source path, kept only to recover a friendly display name.
        /// Also used as a fallback content source when loading older project
        /// files saved before embedding was added (<see cref="Data"/> is null).
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>Embedded raw file bytes (image or PDF). Null only for project files saved by older app versions.</summary>
        public byte[]? Data { get; set; }

        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }

        /// <summary>Per-item content zoom level (1.0 = 100%). Defaults to 1.0 so older project files without this field still load at normal size.</summary>
        public double Zoom { get; set; } = 1.0;

        /// <summary>Added in schema Version 8: true when the item was hidden from the canvas via its "X" chrome button (still present in the file, restorable from the document drawer).</summary>
        public bool IsHidden { get; set; }
    }

    /// <summary>
    /// Driver/contractor identifying details entered in the workspace's top
    /// header bar. Persisted with the workspace project file so they travel
    /// with it; every field defaults to empty so older (or hand-edited) .lqw
    /// files without this section still load without error.
    /// </summary>
    public class WorkspaceHeaderInfo
    {
        public string DriverName { get; set; } = string.Empty;
        public string ContractorName { get; set; } = string.Empty;
        public string Ssn { get; set; } = string.Empty;
        public string Fid { get; set; } = string.Empty;
        public string Phone { get; set; } = string.Empty;

        /// <summary>Added in schema Version 10: equipment tracking numbers.</summary>
        public string TruckNumber { get; set; } = string.Empty;
        public string TrailerNumber { get; set; } = string.Empty;
    }

    /// <summary>
    /// Values for the collapsible "File Dates" panel's evaluation fields.
    /// Every property is a nullable <see cref="DateTime"/> that defaults to
    /// null so older (or hand-edited) .lqw files without this section still
    /// load without error, and so blank fields stay completely unflagged.
    /// </summary>
    public class WorkspaceDateFieldsInfo
    {
        public DateTime? Certification { get; set; }
        public DateTime? LimitedConsent { get; set; }
        public DateTime? FullConsent { get; set; }
        public DateTime? DrugTestRan { get; set; }
        public DateTime? DrugTestReceived { get; set; }
        public DateTime? CdlExpiration { get; set; }
        public DateTime? MvrCdlExpiration { get; set; }
        public DateTime? HazmatExpiration { get; set; }

        /// <summary>Added in schema Version 6, splitting the old combined Hazmat/FP Expiration field into separate Hazmat Endorsement Expiration and Fingerprint Date rules.</summary>
        public DateTime? FingerprintDate { get; set; }
        public DateTime? MvrLastReview { get; set; }
        public DateTime? MvrNextReview { get; set; }
        public DateTime? PhysicalLast { get; set; }
        public DateTime? PhysicalNext { get; set; }
        public DateTime? LatestJobEv { get; set; }
        public DateTime? SevenDayLog { get; set; }
        public DateTime? StatementOfLease { get; set; }

        /// <summary>Added in schema Version 13: Full Unit mode splits the lease into Truck (StatementOfLease) and Trailer.</summary>
        public DateTime? StatementOfLeaseTrailer { get; set; }
        public DateTime? DayQualified { get; set; }

        /// <summary>Added in schema Version 9, synchronizes with MainWindow's Current PSP Snapshot Date.</summary>
        public DateTime? PspSnapshotDate { get; set; }
    }

    /// <summary>
    /// Root object persisted to a workspace project file (*.lqw) and to
    /// timestamped autosave backups.
    /// </summary>
    public class WorkspaceProjectFile
    {
        /// <summary>Schema version, allows future format changes without breaking older saves. Version 3 added per-item zoom and the driver/contractor header fields. Version 4 added the File Dates panel values. Version 5 added the embedded Word document. Version 6 split Hazmat/FP Expiration into Hazmat Endorsement Expiration and a new Fingerprint Date field. Version 7 added the auto-populated FP Exp field. Version 8 added per-item IsHidden state. Version 9 added PspSnapshotDate sync and removed FP Exp field. Version 10 added TruckNumber and TrailerNumber and collapsible file details. Version 11 added embedded Excel spreadsheet.</summary>
        public int Version { get; set; } = 13;

        /// <summary>Added in schema Version 13: the review mode (Driver File, Full Unit, Post-Review Audit).</summary>
        public QualMode Mode { get; set; } = QualMode.DriverFile;

        public List<WorkspaceItemRecord> Items { get; set; } = new();

        /// <summary>Free-form notes text entered via the workspace's Notes panel.</summary>
        public string Notes { get; set; } = string.Empty;

        /// <summary>Driver/contractor header fields entered in the top bar.</summary>
        public WorkspaceHeaderInfo Header { get; set; } = new();

        /// <summary>Values entered into the collapsible "File Dates" rules-engine panel.</summary>
        public WorkspaceDateFieldsInfo DateFields { get; set; } = new();

        /// <summary>
        /// Raw bytes of the .docx loaded into the Notes panel's embedded Word
        /// viewer, if any. Null/empty when no document has been loaded, so
        /// the saved workspace stays fully self-contained and portable.
        /// </summary>
        public byte[]? WordDocBytes { get; set; }

        /// <summary>Display/default-save file name for the embedded Word document (e.g. "Report.docx").</summary>
        public string WordDocFileName { get; set; } = string.Empty;

        /// <summary>
        /// Raw bytes of the spreadsheet loaded into the Sheet panel's embedded
        /// Excel viewer, if any. Null/empty when no sheet has been loaded, so
        /// the saved workspace stays fully self-contained and portable.
        /// </summary>
        public byte[]? ExcelDocBytes { get; set; }

        /// <summary>Display/default-save file name for the embedded Excel spreadsheet (e.g. "Data.xlsx").</summary>
        public string ExcelDocFileName { get; set; } = string.Empty;

        /// <summary>Added in schema Version 12: the Dates &amp; Employment History tracker (driver-reported, unverified).</summary>
        public TrackerData? Tracker { get; set; }
    }
}
