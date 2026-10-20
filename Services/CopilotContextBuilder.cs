using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using HexxcorpDocumentReviewUtility.Models;

namespace HexxcorpDocumentReviewUtility.Services
{
    /// <summary>
    /// Holds structured updates extracted from Copilot's response: driver header fields,
    /// file qualification dates, and (optionally) driver employment history rows.
    /// </summary>
    public class CopilotExtractedData
    {
        public Dictionary<string, string> Dates { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> HeaderFields { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<TrackerRowData> Employment { get; } = new();
        public int TotalCount => Dates.Count + HeaderFields.Count + Employment.Count;
    }

    /// <summary>
    /// Builds document-review prompts from loaded documents, app data, and editable review criteria,
    /// then parses supported suggestions back into app updates.
    /// </summary>
    public static class CopilotContextBuilder
    {
        private const string DateFormat = "MM/dd/yyyy";

        /// <summary>Supplies the tracker's plain-text rendering (set by the workspace window).</summary>
        public static Func<string?>? TrackerContextProvider { get; set; }

        /// <summary>Current review mode, set by the Workspace mode toggle; the prompt builders read it.</summary>
        public static QualMode Mode { get; set; } = QualMode.DriverFile;

        /// <summary>Post-Review Audit "as of" date (the Qualified Date). Null = not set.</summary>
        public static DateTime? AuditDate { get; set; }

        private static bool AuditActive => Mode == QualMode.PostReviewAudit && AuditDate.HasValue;

        /// <summary>The date treated as "today": the Qualified Date in Post-Review Audit mode, otherwise the real date.</summary>
        public static DateTime EffectiveToday => AuditActive ? AuditDate!.Value : DateTime.Today;

        // ======================= Prompt builders =======================

        /// <summary>
        /// Audit &amp; auto-fill in one pass: terse findings for the reviewer, followed by an [APP_UPDATE]
        /// block that fills the header fields, file dates and employment history.
        /// </summary>
        public static string BuildAuditPrompt(
            WorkspaceHeaderInfo header,
            WorkspaceDateFieldsInfo dates,
            IReadOnlyList<string> conflicts,
            IReadOnlyList<string> fileNames)
        {
            var sb = new StringBuilder();
            sb.AppendLine("DOCUMENT REVIEW - FINDINGS AND DATA EXTRACTION");
            sb.AppendLine("Review the attached documents against the applicable DOT reference and the user-configured review criteria below. Then extract supported details for the app. This tool is modular: apply only criteria relevant to this document set and industry.");
            sb.AppendLine();
            AppendGroundRules(sb, allowDocNames: true);
            AppendChecklist(sb);
            sb.AppendLine("=== PART A - FINDINGS (strict format; as few words as possible) ===");
            sb.AppendLine("STATUS: READY / BLOCKED / PENDING DOCS - plus a reason of 8 words or fewer.");
            sb.AppendLine("FAIL: one bullet per confirmed rule failure, 15 words or fewer: item - what the file shows (date/value) - rule missed (with day count).");
            sb.AppendLine("CHECK: one bullet per irregularity, mismatch, or item needing human judgment, 15 words or fewer (include Tracker-vs-document differences).");
            sb.AppendLine("MISSING: required documents, signatures, dates, and blank required fields, as one comma-separated line, naming the source for each blank.");
            sb.AppendLine("PASS: one line of item names only.");
            sb.AppendLine("Omit any heading that has nothing under it (except STATUS). No introduction, recap, explanations, tables, or advice. If an item is not required or cannot be judged, leave it out entirely.");
            sb.AppendLine("CLOSING NOTE (once, at the very END of your whole reply, AFTER the [/APP_UPDATE] tag; never after individual findings): write 2-3 short plain sentences in a natural voice. Tell the user these findings are a first pass and are best confirmed before anything is sent out, and that for the most accurate results they should research or double-check ONE issue at a time. Then ask whether they would like to drill down into any single item, and say that for that item you will give the page number, where in the file the discrepancy appears, and a fresh double-check of the finding. Do not put this note inside the findings or the [APP_UPDATE] block.");
            sb.AppendLine();
            sb.AppendLine("=== PART B - AUTO-FILL (after the findings) ===");
            AppendExtractionRules(sb, includeHeader: true);
            AppendDriverContext(sb, header, dates, conflicts, fileNames);
            AppendUpdateSchemaInstructions(sb, includeHeader: true);
            return sb.ToString();
        }

        /// <summary>Driver summarization: short summary first, then a compact deep-dive.</summary>
        public static string BuildSummarizePrompt(
            WorkspaceHeaderInfo header,
            WorkspaceDateFieldsInfo dates,
            IReadOnlyList<string> conflicts,
            IReadOnlyList<string> fileNames)
        {
            var sb = new StringBuilder();
            sb.AppendLine("DOCUMENT REVIEW SUMMARY");
            sb.AppendLine("Summarize the attached documents using only evidence in the files, app data, DOT reference where applicable, and user-configured criteria.");
            sb.AppendLine();
            AppendGroundRules(sb, allowDocNames: true);
            AppendChecklist(sb);
            sb.AppendLine("=== OUTPUT ===");
            sb.AppendLine("PART 1 - SUMMARY: 3 sentences maximum. Identify the document set, overall review status (ready / blocked / pending information), and main strength or concern.");
            sb.AppendLine("PART 2 - DETAIL: use these headings when applicable, with 1-3 short bullets under each (about 300 words total). Describe evidence, criteria outcome, and differences between sources.");
            sb.AppendLine("1. Identity and Key Details");
            sb.AppendLine("2. Dates and Validity");
            sb.AppendLine("3. History and Supporting Records");
            sb.AppendLine("4. Required Forms and Signatures");
            sb.AppendLine("5. Configured Criteria");
            sb.AppendLine("6. Open Questions and Next Actions");
            sb.AppendLine("Skip any heading with nothing to report. No tables, no repeating the checklist.");
            sb.AppendLine();
            AppendDriverContext(sb, header, dates, conflicts, fileNames);
            return sb.ToString();
        }

        /// <summary>Bare list of every problem point, ready to paste into an email body.</summary>
        public static string BuildDraftRejectEmailPrompt(
            WorkspaceHeaderInfo header,
            WorkspaceDateFieldsInfo dates,
            IReadOnlyList<string> conflicts,
            IReadOnlyList<string> fileNames)
        {
            var sb = new StringBuilder();
            sb.AppendLine("DRAFT DOCUMENT REVIEW FINDINGS");
            sb.AppendLine("Review the attached documents using the applicable DOT reference and user-configured criteria. List supported issues and questions as concise notes that can be adapted for a review message.");
            sb.AppendLine("USE THE WHOLE CONVERSATION: this chat may already contain an earlier audit and a back-and-forth where the user and you investigated individual issues. Build the list from the CURRENT state of that discussion. Drop any item that was found to be NOT a problem, keep every confirmed item, ADD any new issue found during the discussion, and use corrected facts (dates, pages, names) from the discussion over earlier claims. If the user said an item is fine or should be ignored, leave it out.");
            sb.AppendLine();
            AppendGroundRules(sb, allowDocNames: false);
            AppendChecklist(sb);
            sb.AppendLine("=== OUTPUT - THE LIST ONLY ===");
            sb.AppendLine("- Write each bullet with the document or field first, then the supported issue, then a concise suggested follow-up. Use actual names and dates from the files; never invent examples or facts.");
            sb.AppendLine("- Output only a bullet list. One bullet per confirmed issue, missing item, or question, 20 words or fewer.");
            sb.AppendLine("- NO subject line, NO greeting, NO driver or contractor name, NO closing, NO introduction, NO summary, NO advice.");
            sb.AppendLine("- NO file names, NO PDF references, NO citations, NO footnotes, NO numbered source markers, and NO links of any kind.");
            sb.AppendLine("- Include only items you can support from the file. Skip anything that passed or is not required.");
            sb.AppendLine("- If there are no supported issues, output exactly: No issues found.");
            sb.AppendLine();
            AppendDriverContext(sb, header, dates, conflicts, fileNames);
            return sb.ToString();
        }

        /// <summary>Backward-compatible alias.</summary>
        public static string BuildDraftEmailPrompt(
            WorkspaceHeaderInfo header,
            WorkspaceDateFieldsInfo dates,
            IReadOnlyList<string> conflicts,
            IReadOnlyList<string> fileNames) => BuildDraftRejectEmailPrompt(header, dates, conflicts, fileNames);

        /// <summary>Backward-compatible alias (audit and minimums check are one prompt).</summary>
        public static string BuildCheckMinimumsPrompt(
            WorkspaceHeaderInfo header,
            WorkspaceDateFieldsInfo dates,
            IReadOnlyList<string> conflicts,
            IReadOnlyList<string> fileNames) => BuildAuditPrompt(header, dates, conflicts, fileNames);

        /// <summary>Dates and employment history only (header fields are left alone).</summary>
        public static string BuildGatherDatesPrompt(
            WorkspaceHeaderInfo header,
            WorkspaceDateFieldsInfo dates,
            IReadOnlyList<string> conflicts,
            IReadOnlyList<string> fileNames)
        {
            var sb = new StringBuilder();
            sb.AppendLine("EXTRACT DATES AND HISTORY");
            sb.AppendLine("Read the attached documents and extract relevant dates and history for the app. Do not audit or make eligibility decisions.");
            sb.AppendLine();
            AppendGroundRules(sb, allowDocNames: true);
            AppendExtractionRules(sb, includeHeader: false);
            AppendDriverContext(sb, header, dates, conflicts, fileNames);
            AppendUpdateSchemaInstructions(sb, includeHeader: false);
            return sb.ToString();
        }

        // ======================= Shared prompt sections =======================

        private static void AppendGroundRules(StringBuilder sb, bool allowDocNames)
        {
            string today = EffectiveToday.ToString(DateFormat, CultureInfo.InvariantCulture);
            sb.AppendLine("=== RULES (READ FIRST) ===");
            sb.AppendLine("1. Evidence only. Use the attached documents and the app data below. Never guess or fill a gap from memory. If something is not in the file, say NOT IN FILE or leave it out. If you are not sure, label it CHECK instead of stating it as fact.");
            sb.AppendLine("2. Source ranking. Documents outrank app data. The Tracker is driver-reported and UNVERIFIED: treat it as the driver's claim, compare it with the documents, and flag differences.");
            sb.AppendLine("3. EV notes. If the notes/comments section at the bottom of an Employment Verification corrects or contradicts anything above it, the bottom note is the most accurate; use it.");
            sb.AppendLine("4. Dates are MM/dd/yyyy. TODAY is " + today + (AuditActive ? " (the Qualified Date; see rule 10 - ignore the real calendar date)" : "") + ". File completion = the Day Qualified date in the app data if entered, otherwise TODAY. Count calendar days and show the arithmetic in a few words.");
            sb.AppendLine("5. Re-read the relevant document before calling anything missing or expired.");
            sb.AppendLine("6. Do not treat a document or field as required unless an applicable DOT reference or user-configured criterion says it is required. Ignore criteria marked not applicable.");
            sb.AppendLine("7. Never print a full SSN in your findings (last four digits only).");
            sb.AppendLine(allowDocNames
                ? "8. Plain text only. Keep document references to a short name; no links, footnotes, or citation markers."
                : "8. Plain text only. No document names, links, footnotes, or citation markers.");
            sb.AppendLine("9. DUPLICATE DOCUMENTS. If the same document appears more than once, treat the most recently created/dated copy as the truth. Briefly note (one CHECK line) that an older similar copy exists, and whether it is an extra source or conflicts with the current one. If one copy would cause an issue and the other would not, say so and tell the user to delete or archive the unneeded or conflicting copy and not submit it to the Assistant again. Do not audit every copy separately - no double work.");
            AppendModeRules(sb);
            sb.AppendLine();
        }

        /// <summary>Mode-specific context for document set and retrospective reviews.</summary>
        private static void AppendModeRules(StringBuilder sb)
        {
            switch (Mode)
            {
                case QualMode.FullUnit:
                    sb.AppendLine("10. MODE: COMPLETE RECORD REVIEW. Review all supplied documents and equipment details. Require a lease, inspection, or other record only when the user-configured criteria or applicable law requires it.");
                    sb.AppendLine("11. HANDWRITTEN OR UNCLEAR RECORDS: flag only what is legible. If an important detail cannot be read, ask for human verification instead of guessing.");
                    break;
                case QualMode.PostReviewAudit:
                    if (AuditActive)
                        sb.AppendLine("AUDIT DATE: " + AuditDate!.Value.ToString(DateFormat, CultureInfo.InvariantCulture) + " is the review reference date. Evaluate each document as it existed on that date; a document created later was not available then and should be identified as a timing question.");
                    sb.AppendLine("10. MODE: RETROSPECTIVE REVIEW. Separate legal-reference findings from user-configured criteria and identify which source supports each finding. Do not assume requirements beyond the selected criteria.");
                    break;
                default:
                    sb.AppendLine("10. MODE: DOCUMENT SET REVIEW. Assess the supplied materials against applicable DOT references and user-configured criteria only. Do not infer missing requirements.");
                    break;
            }
        }

        private static void AppendChecklist(StringBuilder sb)
        {
            sb.AppendLine("=== REVIEW CRITERIA ===");
            sb.AppendLine("Use only the applicable reference and configured criteria below. User-configured requirements must be supported by their stated source and applicability. Do not invent, extend, or treat an empty template as a rule.");
            MinimumsData criteria = MinimumsDataService.Load();
            foreach (MinimumsSection section in criteria.Sections)
            {
                sb.AppendLine("[" + section.Title + "]");
                foreach (string line in section.Lines)
                {
                    if (!string.IsNullOrWhiteSpace(line)) sb.AppendLine("- " + line);
                }
            }
            sb.AppendLine();
        }

        private static void AppendExtractionRules(StringBuilder sb, bool includeHeader)
        {
            if (includeHeader)
            {
                sb.AppendLine("APP FIELDS (only when supported by the documents): these fields reflect this app's current record model. Fill applicable values only; omit fields that do not apply to the document set.");
            }
            sb.AppendLine("DATE FIELDS (MM/dd/yyyy, 4-digit year, only if printed in the documents):");
            sb.AppendLine("- Certification: the date on the document titled Certification. The same date usually also appears on the PSP Consent, Clearinghouse consent, and Disclosures & Consents. If these disagree: keep the app's current Certification date when it matches one of them; otherwise re-examine and choose the most valid date (the Certification document's date, backed by the most consent forms, not in the future). If you changed or could not resolve it, note it in one CHECK line.");
            sb.AppendLine("- LimitedConsent: ALWAYS the date the Operating History (OH) was signed - from the OH alone, never from another form. FullConsent: ALWAYS the date on the final Clearinghouse results page - never from its corresponding consent form or any other document.");
            sb.AppendLine("- DrugTestRan / DrugTestReceived: only if a document in the file shows them; never a finding when absent");
            sb.AppendLine("- CdlExpiration: license expiration; MvrCdlExpiration: CDL expiration shown on the MVR");
            sb.AppendLine("- MvrLastReview: the date the MVR was pulled");
            sb.AppendLine("- PhysicalLast / PhysicalNext: medical exam date and medical certificate expiration as shown on the MVR/CDLIS (or the med cert if one is in the file)");
            sb.AppendLine("- HazmatExpiration: Hazmat endorsement expiration; FingerprintDate: fingerprint receipt date (omit when a valid Hazmat endorsement is held)");
            sb.AppendLine("- LatestJobEv: the most recent employer's quit/end date per its EV (after applying any bottom-note correction)");
            sb.AppendLine("- SevenDayLog, DayQualified (and in Full Unit mode StatementOfLease for the truck and StatementOfLeaseTrailer for the trailer): use the qualification date printed in the file for all of them; omit them if none is printed. Never use today's date.");
            sb.AppendLine("- PspSnapshotDate: PSP snapshot date; PspRan: date the PSP was run; CdlisDac: CDLIS/DAC check date");
            sb.AppendLine("HISTORY (EmploymentHistory array, when applicable): include every supported record in date order. Do not apply an unstated look-back window or exclude older records.");
            sb.AppendLine("- Description: organization, role, or activity in a few words. Otr: use the JSON boolean true only when the source explicitly identifies over-the-road work; otherwise false.");
            sb.AppendLine("- Start and End: full MM/dd/yyyy dates from the best available source. If only a month is given, use the 1st for Start and the last day of that month for End. For ongoing records, use today's date (or the audit reference date in retrospective mode). Omit unsupported values.");
            sb.AppendLine("EXTRACTION RULES: omit any field you cannot find. Never output placeholder, sample, or guessed values. 10/03/2025 is October 3rd; if a date is ambiguous or unreadable, omit it.");
            sb.AppendLine();
        }

        private static void AppendDriverContext(
            StringBuilder sb,
            WorkspaceHeaderInfo header,
            WorkspaceDateFieldsInfo dates,
            IReadOnlyList<string> conflicts,
            IReadOnlyList<string> fileNames)
        {
            sb.AppendLine("=== CURRENT FILE DETAILS (app data; may be blank or unverified) ===");
            sb.AppendLine($"Driver: {Val(header.DriverName)} | Contractor: {Val(header.ContractorName)} | Truck: {Val(header.TruckNumber)} | Trailer: {Val(header.TrailerNumber)}");
            sb.AppendLine($"SSN: {MaskSsn(header.Ssn)} | FID: {Val(header.Fid)} | Phone: {Val(header.Phone)}");
            sb.AppendLine(AuditActive
                ? $"Today (as-of Qualified Date for this audit): {EffectiveToday.ToString(DateFormat, CultureInfo.InvariantCulture)}"
                : $"Today: {DateTime.Today.ToString(DateFormat, CultureInfo.InvariantCulture)}");
            sb.AppendLine($"Certification: {Fmt(dates.Certification)} | Limited Consent: {Fmt(dates.LimitedConsent)} | Full Consent: {Fmt(dates.FullConsent)}");
            sb.AppendLine($"Drug Test Ran: {Fmt(dates.DrugTestRan)} | Drug Test Received: {Fmt(dates.DrugTestReceived)}");
            sb.AppendLine($"CDL Exp: {Fmt(dates.CdlExpiration)} | MVR CDL Exp: {Fmt(dates.MvrCdlExpiration)} | MVR Pulled: {Fmt(dates.MvrLastReview)}");
            sb.AppendLine($"Physical Last: {Fmt(dates.PhysicalLast)} | Physical Next: {Fmt(dates.PhysicalNext)}");
            sb.AppendLine($"Fingerprint: {Fmt(dates.FingerprintDate)} | Hazmat Exp: {Fmt(dates.HazmatExpiration)} | Latest Job EV: {Fmt(dates.LatestJobEv)}");
            sb.AppendLine(Mode == QualMode.FullUnit
                ? $"7-Day Log: {Fmt(dates.SevenDayLog)} | Statement of Lease (Truck): {Fmt(dates.StatementOfLease)} | Statement of Lease (Trailer): {Fmt(dates.StatementOfLeaseTrailer)} | Day Qualified: {Fmt(dates.DayQualified)}"
                : $"7-Day Log: {Fmt(dates.SevenDayLog)} | Day Qualified: {Fmt(dates.DayQualified)}");
            sb.AppendLine("Review mode: " + Mode);
            sb.AppendLine($"PSP Snapshot: {Fmt(dates.PspSnapshotDate)}");
            sb.AppendLine();

            string? tracker = null;
            try { tracker = TrackerContextProvider?.Invoke(); } catch (Exception) { }
            sb.AppendLine("=== DATES & EMPLOYMENT HISTORY TRACKER (DRIVER-REPORTED, NOT VERIFIED) ===");
            sb.AppendLine(string.IsNullOrWhiteSpace(tracker) ? "(tracker not available)" : tracker.TrimEnd());
            sb.AppendLine();

            if (conflicts != null && conflicts.Count > 0)
            {
                sb.AppendLine("=== APP DATE-RULE FLAGS (informational; confirm against the documents) ===");
                foreach (string c in conflicts)
                {
                    sb.AppendLine($"- {c}");
                }
                sb.AppendLine();
            }

            sb.AppendLine("=== ATTACHED / LOADED WORKSPACE DOCUMENTS ===");
            if (fileNames != null && fileNames.Count > 0)
            {
                foreach (string fn in fileNames)
                {
                    sb.AppendLine($"- {fn} (ATTACHED TO THIS CHAT SESSION)");
                }
                sb.AppendLine("Read every page of every attached document.");
            }
            else
            {
                sb.AppendLine("- No documents are loaded in the workspace. Say so in one line and limit your answer to the app data.");
            }
            sb.AppendLine();
        }

        private static void AppendUpdateSchemaInstructions(StringBuilder sb, bool includeHeader)
        {
            sb.AppendLine("=== APP UPDATE BLOCK (the app reads this to fill itself in) ===");
            sb.AppendLine("Finish with ONE block: the tag [APP_UPDATE], a single JSON object, then [/APP_UPDATE]. Include only fields with real values from the documents; delete every line you have no value for. Template (angle-bracket parts are placeholders to replace or delete):");
            sb.AppendLine("[APP_UPDATE]");
            sb.AppendLine("{");
            if (includeHeader)
            {
                sb.AppendLine("  \"DriverName\": \"<driver name>\",");
                sb.AppendLine("  \"ContractorName\": \"<contractor/fleet>\",");
                sb.AppendLine("  \"TruckNumber\": \"<truck #>\",");
                sb.AppendLine("  \"TrailerNumber\": \"<trailer #>\",");
                sb.AppendLine("  \"Ssn\": \"<ssn>\",");
                sb.AppendLine("  \"Fid\": \"<fid>\",");
                sb.AppendLine("  \"Phone\": \"<phone>\",");
            }
            sb.AppendLine("  \"Certification\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"LimitedConsent\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"FullConsent\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"DrugTestRan\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"DrugTestReceived\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"CdlExpiration\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"MvrCdlExpiration\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"MvrLastReview\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"PhysicalLast\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"PhysicalNext\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"FingerprintDate\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"HazmatExpiration\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"LatestJobEv\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"SevenDayLog\": \"<MM/dd/yyyy>\",");
            if (Mode == QualMode.FullUnit)
            {
                sb.AppendLine("  \"StatementOfLease\": \"<MM/dd/yyyy>\",");
                sb.AppendLine("  \"StatementOfLeaseTrailer\": \"<MM/dd/yyyy>\",");
            }
            sb.AppendLine("  \"DayQualified\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"PspSnapshotDate\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"PspRan\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"CdlisDac\": \"<MM/dd/yyyy>\",");
            sb.AppendLine("  \"EmploymentHistory\": [");
            sb.AppendLine("    { \"Description\": \"<employer - role>\", \"Otr\": true, \"Start\": \"<MM/dd/yyyy>\", \"End\": \"<MM/dd/yyyy>\" }");
            sb.AppendLine("  ]");
            sb.AppendLine("}");
            sb.AppendLine("[/APP_UPDATE]");
            sb.AppendLine("Valid JSON only (double quotes, no comments, no trailing commas). Do not wrap the block in anything else and output it only once, at the very end.");
            if (!includeHeader)
            {
                sb.AppendLine("Before the block, write at most one line: \"NOT FOUND: <field names>\" for fields you could not find. Nothing else.");
            }
        }

        // ======================= Response parsing =======================

        /// <summary>
        /// Identifies whether a JSON snippet is a prompt template or dummy sample rather than real extracted data.
        /// </summary>
        public static bool IsDummyOrPromptJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return true;
            string[] markers =
            {
                "<driver name>", "<MM/dd/yy>", "<MM/dd/yyyy>", "<truck #>", "<trailer #>", "<ssn>", "<fid>", "<phone>",
                "<contractor/fleet>", "<employer - role>", "John Doe", "Doe Trucking", "555-555-5555", "123456789", "123-45-6789"
            };
            foreach (string m in markers)
            {
                if (json.Contains(m, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static readonly string[] StandardDateFormats = new[]
        {
            "MM/dd/yyyy", "M/d/yyyy", "MM/d/yyyy", "M/dd/yyyy",
            "MM-dd-yyyy", "M-d-yyyy",
            "yyyy-MM-dd", "yyyy/MM/dd",
            "MM/dd/yy", "M/d/yy", "MM/d/yy", "M/dd/yy",
            "MM-dd-yy", "M-d-yy"
        };

        /// <summary>
        /// Parses driver header fields, dates, and employment rows from Copilot's response text,
        /// supporting [APP_UPDATE] blocks, markdown code blocks, or raw JSON. Scans in reverse so the
        /// assistant's final answer wins over prompt echoes or earlier turns.
        /// </summary>
        public static CopilotExtractedData ParseSuggestedUpdates(string responseText)
        {
            var data = new CopilotExtractedData();
            if (string.IsNullOrWhiteSpace(responseText))
            {
                return data;
            }

            string jsonToParse = string.Empty;

            // 1. Explicit [APP_UPDATE] ... [/APP_UPDATE] blocks, newest first.
            var matches = Regex.Matches(responseText, @"\[APP_UPDATE\]\s*(\{.*?\})\s*\[/APP_UPDATE\]", RegexOptions.Singleline);
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                string candidate = matches[i].Groups[1].Value;
                if (!IsDummyOrPromptJson(candidate) && IsValidJson(candidate))
                {
                    jsonToParse = candidate;
                    break;
                }
            }

            // 2. Markdown json blocks with known keys.
            if (string.IsNullOrWhiteSpace(jsonToParse))
            {
                var mdMatches = Regex.Matches(responseText, @"```(?:json)?\s*(\{[\s\S]*?(?:Certification|DrugTest|StatementOfLease|DayQualified|SevenDayLog|MvrLastReview|PspSnapshotDate|DriverName|ContractorName|TruckNumber|EmploymentHistory)[\s\S]*?\})\s*```", RegexOptions.IgnoreCase);
                for (int i = mdMatches.Count - 1; i >= 0; i--)
                {
                    string candidate = mdMatches[i].Groups[1].Value;
                    if (!IsDummyOrPromptJson(candidate) && IsValidJson(candidate))
                    {
                        jsonToParse = candidate;
                        break;
                    }
                }
            }

            // 3. Any JSON object containing key fields.
            if (string.IsNullOrWhiteSpace(jsonToParse))
            {
                var generalMatches = Regex.Matches(responseText, @"\{[\s\S]*?(?:""DriverName""|""Certification""|""StatementOfLease""|""DayQualified"")[\s\S]*?\}", RegexOptions.IgnoreCase);
                for (int i = generalMatches.Count - 1; i >= 0; i--)
                {
                    string candidate = generalMatches[i].Value;
                    if (!IsDummyOrPromptJson(candidate) && IsValidJson(candidate))
                    {
                        jsonToParse = candidate;
                        break;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(jsonToParse))
            {
                return data;
            }

            try
            {
                using var doc = JsonDocument.Parse(jsonToParse);
                foreach (var prop in doc.RootElement.EnumerateObject())
                {
                    string key = prop.Name;
                    if (string.Equals(key, "EmploymentHistory", StringComparison.OrdinalIgnoreCase) && prop.Value.ValueKind == JsonValueKind.Array)
                    {
                        ReadEmployment(prop.Value, data);
                        continue;
                    }

                    if (prop.Value.ValueKind != JsonValueKind.String) continue;
                    string val = prop.Value.GetString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(val)) continue;

                    if (IsHeaderField(key))
                    {
                        data.HeaderFields[key] = val;
                    }
                    else if (NormalizeDateString(val, out string normalized))
                    {
                        data.Dates[key] = normalized;
                    }
                }
            }
            catch (JsonException)
            {
            }

            return data;
        }

        private static bool IsValidJson(string json)
        {
            try
            {
                using var _ = JsonDocument.Parse(json);
                return true;
            }
            catch (JsonException)
            {
                return false;
            }
        }

        private static void ReadEmployment(JsonElement array, CopilotExtractedData data)
        {
            foreach (var el in array.EnumerateArray())
            {
                if (el.ValueKind != JsonValueKind.Object) continue;
                string desc = GetAny(el, "Description", "Employer", "Company", "Name", "Job");
                string start = GetAny(el, "Start", "StartDate", "From", "Hired", "HireDate");
                string end = GetAny(el, "End", "EndDate", "To", "Quit", "QuitDate", "Terminated");
                bool otr = false;
                foreach (var prop in el.EnumerateObject())
                {
                    if (!prop.Name.Equals("Otr", StringComparison.OrdinalIgnoreCase) && !prop.Name.Equals("IsOtr", StringComparison.OrdinalIgnoreCase)) continue;
                    var o = prop.Value;
                    otr = o.ValueKind == JsonValueKind.True ||
                          (o.ValueKind == JsonValueKind.Number && o.TryGetInt32(out int n) && n != 0) ||
                          (o.ValueKind == JsonValueKind.String && Regex.IsMatch(o.GetString() ?? "", @"^\s*(t|y|1|otr)", RegexOptions.IgnoreCase));
                    break;
                }

                if (!otr && Regex.IsMatch(desc, @"\b(OTR|over[- ]the[- ]road)\b", RegexOptions.IgnoreCase)) otr = true;

                if (!NormalizeEmploymentDate(start, false, out string s)) s = string.Empty;
                if (!NormalizeEmploymentDate(end, true, out string e)) e = string.Empty;
                if (string.IsNullOrWhiteSpace(desc) && s.Length == 0 && e.Length == 0) continue;

                data.Employment.Add(new TrackerRowData { Description = desc, IsOtr = otr, Start = s, End = e });
            }
        }

        /// <summary>Case-insensitive property lookup with aliases; numbers and booleans are returned as text.</summary>
        private static string GetAny(JsonElement el, params string[] names)
        {
            foreach (string name in names)
            {
                foreach (var prop in el.EnumerateObject())
                {
                    if (!prop.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) continue;
                    string v = prop.Value.ValueKind == JsonValueKind.String ? prop.Value.GetString()?.Trim() ?? string.Empty : string.Empty;
                    if (v.Length > 0) return v;
                }
            }

            return string.Empty;
        }

        /// <summary>Accepts full dates, MM/yyyy (1st / last day of the month) and "present"-style words for an end date.</summary>
        private static bool NormalizeEmploymentDate(string input, bool isEnd, out string formatted)
        {
            formatted = string.Empty;
            if (string.IsNullOrWhiteSpace(input)) return false;
            string s = input.Trim();

            if (isEnd && Regex.IsMatch(s, @"^(present|current|currently|ongoing|now|today|still)", RegexOptions.IgnoreCase))
            {
                formatted = EffectiveToday.ToString(DateFormat, CultureInfo.InvariantCulture);
                return true;
            }

            var my = Regex.Match(s, @"^(\d{1,2})[/\-](\d{4})$");
            if (my.Success && int.Parse(my.Groups[1].Value) is >= 1 and <= 12)
            {
                int month = int.Parse(my.Groups[1].Value);
                int year = int.Parse(my.Groups[2].Value);
                var d = isEnd ? new DateTime(year, month, DateTime.DaysInMonth(year, month)) : new DateTime(year, month, 1);
                formatted = d.ToString(DateFormat, CultureInfo.InvariantCulture);
                return true;
            }

            return NormalizeDateString(s, out formatted);
        }

        private static string GetString(JsonElement el, string name) =>
            el.TryGetProperty(name, out var p) && p.ValueKind == JsonValueKind.String ? p.GetString()?.Trim() ?? string.Empty : string.Empty;
        private static bool IsHeaderField(string key)
        {
            return string.Equals(key, "DriverName", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(key, "ContractorName", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(key, "TruckNumber", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(key, "TrailerNumber", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(key, "Ssn", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(key, "Fid", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(key, "Phone", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Backward-compatible method returning only the extracted dates dictionary.</summary>
        public static Dictionary<string, string> ParseSuggestedDates(string responseText) =>
            ParseSuggestedUpdates(responseText).Dates;

        private static bool NormalizeDateString(string input, out string formatted)
        {
            formatted = string.Empty;
            if (string.IsNullOrWhiteSpace(input)) return false;

            string trimmed = input.Trim();

            if (DateTime.TryParseExact(trimmed, StandardDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime exactDt))
            {
                formatted = exactDt.ToString(DateFormat, CultureInfo.InvariantCulture);
                return true;
            }

            string digits = Regex.Replace(trimmed, @"[^\d]", "");
            if (digits.Length == 6)
            {
                if (DateTime.TryParseExact(digits, "MMddyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt))
                {
                    formatted = dt.ToString(DateFormat, CultureInfo.InvariantCulture);
                    return true;
                }
            }
            else if (digits.Length == 8)
            {
                if (DateTime.TryParseExact(digits, "MMddyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt) ||
                    DateTime.TryParseExact(digits, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                {
                    formatted = dt.ToString(DateFormat, CultureInfo.InvariantCulture);
                    return true;
                }
            }

            if (DateTime.TryParse(trimmed, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime generalDt))
            {
                formatted = generalDt.ToString(DateFormat, CultureInfo.InvariantCulture);
                return true;
            }

            return false;
        }

        private static string Fmt(DateTime? dt) =>
            dt.HasValue ? dt.Value.ToString(DateFormat, CultureInfo.InvariantCulture) : "(not entered)";

        private static string Val(string? s) =>
            string.IsNullOrWhiteSpace(s) ? "(none)" : s.Trim();

        private static string MaskSsn(string? ssn)
        {
            if (string.IsNullOrWhiteSpace(ssn)) return "(none)";
            string clean = Regex.Replace(ssn, @"[^\d]", "");
            return clean.Length >= 4 ? $"***-**-{clean[^4..]}" : "***-**-****";
        }
    }
}

