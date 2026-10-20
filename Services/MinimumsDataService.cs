using System;
using System.IO;
using System.Text.Json;
using HexxcorpDocumentReviewUtility.Models;

namespace HexxcorpDocumentReviewUtility.Services
{
    /// <summary>
    /// Loads and saves the DOT reference and user-configured review criteria
    /// (<c>MinimumsData.json</c>, stored in the application's base directory).
    /// If the file is missing (first run) it is seeded with the built-in
    /// default content; if it is missing or corrupt at any later point, the
    /// built-in defaults are used in-memory so the reference window never
    /// fails to display.
    /// </summary>
    public static class MinimumsDataService
    {
        private static readonly string DataFilePath =
            Path.Combine(AppContext.BaseDirectory, "MinimumsData.json");

        private static readonly JsonSerializerOptions SerializerOptions = new()
        {
            WriteIndented = true
        };

        /// <summary>
        /// Loads the reference content from disk, seeding the file with
        /// built-in defaults on first run, and falling back to the built-in
        /// defaults (without overwriting the file) if it is corrupt/unreadable.
        /// </summary>
        public static MinimumsData Load()
        {
            try
            {
                if (!File.Exists(DataFilePath))
                {
                    var defaults = BuildDefaultData();
                    Save(defaults);
                    return defaults;
                }

                string json = File.ReadAllText(DataFilePath);
                var data = JsonSerializer.Deserialize<MinimumsData>(json, SerializerOptions);
                return data ?? BuildDefaultData();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
            {
                return BuildDefaultData();
            }
        }

        /// <summary>Persists the reference content to disk. Returns false (without throwing) on a locked/read-only file.</summary>
        public static bool Save(MinimumsData data)
        {
            try
            {
                string json = JsonSerializer.Serialize(data, SerializerOptions);
                File.WriteAllText(DataFilePath, json);
                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return false;
            }
        }

        /// <summary>
        /// The built-in DOT starter reference and editable criteria templates, used to seed
        /// <c>MinimumsData.json</c> on first run and as a safe fallback.
        /// </summary>
        public static MinimumsData BuildDefaultData()
        {
            static MinimumsSection S(string title, params string[] lines)
            {
                var section = new MinimumsSection { Title = title };
                section.Lines.AddRange(lines);
                return section;
            }

            return new MinimumsData
            {
                Version = 1,
                Sections =
                {
                    S("DOT Reference (Starter)",
                        "Starter checklist only, not legal advice or a complete compliance program. Verify current federal and state regulations, scope, exceptions, and authoritative sources before use.",
                        "D1 License: verify a valid, unexpired commercial license with the required class and endorsements, applicable restrictions addressed, current status, and consistency with CDLIS/source records.",
                        "D2 Age: interstate commercial driving generally requires age 21 or older. Verify the applicable operating scope and current exceptions.",
                        "D3 Application and history: verify required application information and employment history for the applicable period; identify unexplained gaps for review.",
                        "D4 Motor-vehicle records: obtain and review records for applicable licensing states and regulatory periods; document sources, retrieval dates, and discrepancies.",
                        "D5 Drug and Alcohol Clearinghouse: when the role and current rules require it, obtain required consent and verify the official query result and eligibility status.",
                        "D6 Safety qualification: review relevant disqualifying offenses, suspensions, and safety records against current applicable rules; do not infer an outcome from incomplete evidence.",
                        "D7 Identity and records: compare required identifying information across relevant records; report only supported mismatches and mask sensitive identifiers.",
                        "For each finding, name the applicable rule/source and verify effective date and applicability. Escalate uncertain legal interpretations for qualified review."),
                    S("Organization Requirements (Editable)",
                        "Requirement: [enter one organization-specific requirement or remove this line]",
                        "Source or policy reference: [enter document, section, and effective date]",
                        "Applies to: [enter roles, locations, document types, or conditions]",
                        "Evidence to check: [enter the record or field that demonstrates compliance]"),
                    S("Industry Requirements (Editable)",
                        "Requirement: [enter one industry-specific requirement or remove this line]",
                        "Source or policy reference: [enter authoritative source and effective date]",
                        "Applies to: [enter scope and exceptions]",
                        "Evidence to check: [enter the records needed for review]"),
                    S("Review Preferences (Editable)",
                        "Required document types: [enter or leave blank]",
                        "Required fields and signatures: [enter or leave blank]",
                        "Date windows and validity periods: [enter or leave blank]",
                        "Escalation and exception handling: [enter or leave blank]",
                        "Blank items are not requirements. Confirm configured criteria are current before each review.")
                }
            };
        }
    }
}

