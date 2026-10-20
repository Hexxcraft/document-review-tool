using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using HexxcorpDocumentReviewUtility.Models;

namespace HexxcorpDocumentReviewUtility.Services
{
    /// <summary>
    /// Handles saving/loading workspace project files (manual saves) and
    /// managing the automated autosave backup rotation.
    /// </summary>
    public static class WorkspaceFileService
    {
        public const string ProjectFileExtension = ".lqw";

        private static readonly JsonSerializerOptions SerializerOptions = new()
        {
            WriteIndented = true
        };

        /// <summary>Directory where autosave backups are written, created on first use.</summary>
        public static string BackupDirectory
        {
            get
            {
                string dir = Path.Combine(AppContext.BaseDirectory, "WorkspaceBackups");
                Directory.CreateDirectory(dir);
                return dir;
            }
        }

        /// <summary>Maximum number of autosave backups retained; older ones are pruned.</summary>
        private const int MaxAutosaveBackups = 20;

        /// <summary>Supplies the live tracker snapshot to embed in every saved/autosaved project (set by the workspace window).</summary>
        public static Func<TrackerData?>? TrackerProvider { get; set; }

        /// <summary>
        /// Builds a serializable project file from the live canvas item collection.
        /// </summary>
        public static WorkspaceProjectFile BuildProjectFile(
            IEnumerable<CanvasItemModel> items,
            string notes = "",
            WorkspaceHeaderInfo? header = null,
            WorkspaceDateFieldsInfo? dateFields = null,
            byte[]? wordDocBytes = null,
            string wordDocFileName = "",
            byte[]? excelDocBytes = null,
            string excelDocFileName = "")
        {
            var file = new WorkspaceProjectFile
            {
                Notes = notes,
                Mode = CopilotContextBuilder.Mode,
                Header = header ?? new WorkspaceHeaderInfo(),
                DateFields = dateFields ?? new WorkspaceDateFieldsInfo(),
                WordDocBytes = wordDocBytes,
                WordDocFileName = wordDocFileName,
                ExcelDocBytes = excelDocBytes,
                ExcelDocFileName = excelDocFileName,
                Tracker = TrackerProvider?.Invoke()
            };
            foreach (var item in items)
            {
                file.Items.Add(new WorkspaceItemRecord
                {
                    Kind = item.Kind,
                    FilePath = item.FilePath,
                    Data = item.FileBytes,
                    X = item.X,
                    Y = item.Y,
                    Width = item.Width,
                    Height = item.Height,
                    Zoom = item.Zoom,
                    IsHidden = item.IsHidden
                });
            }
            return file;
        }

        /// <summary>Saves a project file to an explicit path chosen by the user.</summary>
        public static void Save(
            string path,
            IEnumerable<CanvasItemModel> items,
            string notes = "",
            WorkspaceHeaderInfo? header = null,
            WorkspaceDateFieldsInfo? dateFields = null,
            byte[]? wordDocBytes = null,
            string wordDocFileName = "",
            byte[]? excelDocBytes = null,
            string excelDocFileName = "")
        {
            var file = BuildProjectFile(items, notes, header, dateFields, wordDocBytes, wordDocFileName, excelDocBytes, excelDocFileName);
            string json = JsonSerializer.Serialize(file, SerializerOptions);
            File.WriteAllText(path, json);
        }

        /// <summary>Loads a project file, throwing on malformed/unreadable content so the caller can report it.</summary>
        public static WorkspaceProjectFile Load(string path)
        {
            string json = File.ReadAllText(path);
            var file = JsonSerializer.Deserialize<WorkspaceProjectFile>(json, SerializerOptions);
            return file ?? new WorkspaceProjectFile();
        }

        /// <summary>
        /// Writes a timestamped autosave backup and prunes old backups beyond
        /// <see cref="MaxAutosaveBackups"/> to avoid unbounded disk growth.
        /// Returns the path written, or null if there was nothing to save or
        /// the write failed.
        /// </summary>
        public static string? CreateAutosaveBackup(
            IEnumerable<CanvasItemModel> items,
            string notes = "",
            WorkspaceHeaderInfo? header = null,
            WorkspaceDateFieldsInfo? dateFields = null,
            byte[]? wordDocBytes = null,
            string wordDocFileName = "",
            byte[]? excelDocBytes = null,
            string excelDocFileName = "")
        {
            try
            {
                var itemList = items.ToList();
                bool hasHeaderData = header is { } h && (!string.IsNullOrWhiteSpace(h.DriverName)
                    || !string.IsNullOrWhiteSpace(h.ContractorName) || !string.IsNullOrWhiteSpace(h.Ssn)
                    || !string.IsNullOrWhiteSpace(h.Fid) || !string.IsNullOrWhiteSpace(h.Phone)
                    || !string.IsNullOrWhiteSpace(h.TruckNumber) || !string.IsNullOrWhiteSpace(h.TrailerNumber));
                bool hasWordDoc = wordDocBytes is { Length: > 0 };
                bool hasExcelDoc = excelDocBytes is { Length: > 0 };

                if (itemList.Count == 0 && string.IsNullOrWhiteSpace(notes) && !hasHeaderData && !hasWordDoc && !hasExcelDoc)
                {
                    return null;
                }

                string fileName = $"autosave_{DateTime.Now:yyyyMMdd_HHmmss}{ProjectFileExtension}";
                string fullPath = Path.Combine(BackupDirectory, fileName);
                Save(fullPath, itemList, notes, header, dateFields, wordDocBytes, wordDocFileName, excelDocBytes, excelDocFileName);

                PruneOldBackups();
                return fullPath;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
            {
                return null;
            }
        }

        private static void PruneOldBackups()
        {
            var backups = Directory.GetFiles(BackupDirectory, $"autosave_*{ProjectFileExtension}")
                .Select(p => new FileInfo(p))
                .OrderByDescending(f => f.CreationTimeUtc)
                .ToList();

            foreach (var stale in backups.Skip(MaxAutosaveBackups))
            {
                try { stale.Delete(); } catch (IOException) { /* best-effort cleanup */ }
            }
        }
    }
}
