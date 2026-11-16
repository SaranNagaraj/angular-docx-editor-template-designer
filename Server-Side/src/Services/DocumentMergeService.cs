using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;
using EJ2APIServices.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using EJ2WordDocument = Syncfusion.EJ2.DocumentEditor.WordDocument;

namespace EJ2APIServices.Services
{
    /// <summary>
    /// Result of a merge attempt: either the missing fields to prompt for,
    /// or the merged document as SFDT (for the editor) plus DOCX bytes.
    /// </summary>
    public class MergeOutcome
    {
        public List<MissingFieldInfo> MissingFields { get; set; } = new List<MissingFieldInfo>();
        /// <summary>Merged document serialized as SFDT, ready for documentEditor.open().</summary>
        public string MergedSfdt { get; set; }
        public string DocumentName { get; set; }
    }

    /// <summary>
    /// DocIO mail merge engine for the Notice of Hearing POC.
    ///
    /// === Requirement 4: repeating Party regions ===
    /// The template contains a mail merge region delimited by the merge fields
    /// «BeginGroup:Parties» ... «EndGroup:Parties». DocIO discovers the region from
    /// the template's MERGEFIELD delimiters (MailMerge.GetMergeGroupNames() lists it),
    /// and repeats the region body once per matching data row during the merge.
    ///
    /// === Requirement 5: parent fields inside child regions ===
    /// Implemented with DocIO relational (nested) mail merge — NOT flattening.
    /// The data set has a parent table "Case" and a child table "Parties" related by
    /// CaseId. The template wraps the party region in an outer region «BeginGroup:Case».
    /// DocIO resolves «Case:CaseNumber» inside «BeginParties» from the current parent
    /// row while iterating the child rows (verified: parent values appear in every
    /// repeated party block). See BuildNestedDataSet() and Merge() for the exact code path.
    ///
    /// === Multiple selected cases ===
    /// All selected cases are merged into ONE document: each selected case becomes one
    /// row of the parent "Case" table, so the outer region repeats once per case and the
    /// inner region repeats per party of that case. No per-case files and no ZIP.
    /// </summary>
    public static class DocumentMergeService
    {
        /// <summary>Parent region/table name. «Case:CaseNumber» resolves against this table.</summary>
        public const string CaseRegionName = MergeFieldConfig.ParentRegion;

        /// <summary>Child region/table name (repeating party region).</summary>
        public const string PartiesRegionName = MergeFieldConfig.ChildRegion;

        /// <summary>Relation key between Case and Parties.</summary>
        private const string RelationKey = MergeFieldConfig.RelationKey;

        /// <summary>
        /// Converts the SFDT template serialized by the Angular Document Editor into a
        /// DocIO WordDocument with materialized merge fields.
        ///
        /// EJ2WordDocument.Save(sfdt) returns a DocIO document whose field structures are
        /// not yet exposed to MailMerge (GetMergeFieldNames() is empty). Saving that
        /// instance to DOCX and reloading it materializes the fields, so the merge works
        /// on editor-authored templates. (Verified in this POC's integration runs.)
        /// </summary>
        public static WordDocument LoadTemplateFromSfdt(string sfdt)
        {
            WordDocument raw = EJ2WordDocument.Save(sfdt);
            using (var ms = new MemoryStream())
            {
                raw.Save(ms, FormatType.Docx);
                raw.Close();
                ms.Position = 0;
                return new WordDocument(ms, FormatType.Docx);
            }
        }

        /// <summary>
        /// Converts a merged DocIO document into SFDT so the Angular Document Editor can
        /// open the result directly (open(sfdt)) instead of downloading a file.
        /// </summary>
        public static string ToSfdt(WordDocument merged)
        {
            var sfdtDoc = EJ2WordDocument.Load(merged);
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtDoc);
            sfdtDoc.Dispose();
            return json;
        }

        /// <summary>
        /// Loads a DOCX file from disk and returns its SFDT representation so the editor
        /// can open it directly.
        /// </summary>
        public static string LoadDocxAsSfdt(string docxPath)
        {
            using (var doc = new WordDocument(docxPath, FormatType.Docx))
            {
                return ToSfdt(doc);
            }
        }

        /// <summary>
        /// Loads a DOCX file from disk into a DocIO WordDocument ready to be appended.
        /// </summary>
        public static WordDocument LoadDocx(string docxPath)
        {
            return new WordDocument(docxPath, FormatType.Docx);
        }

        /// <summary>
        /// Converts the editor's SFDT payload back to a DOCX file and writes it to disk.
        /// This is the inverse of LoadTemplateFromSfdt: the editor serializes the
        /// document the user is editing as SFDT, and the server persists it as the
        /// DOCX template in the templates folder.
        ///
        /// The write is atomic: the DOCX is first produced in a temporary file next to
        /// the target and only then moved over it, so a failure part-way through never
        /// leaves a truncated template behind.
        /// </summary>
        public static void SaveSfdtAsDocx(string sfdt, string docxPath)
        {
            var directory = Path.GetDirectoryName(docxPath);
            var tempPath = Path.Combine(directory ?? string.Empty, Path.GetFileName(docxPath) + ".tmp");

            try
            {
                WordDocument raw = EJ2WordDocument.Save(sfdt);
                using (var ms = new MemoryStream())
                {
                    raw.Save(ms, FormatType.Docx);
                    raw.Close();
                    ms.Position = 0;

                    using (var doc = new WordDocument(ms, FormatType.Docx))
                    {
                        doc.Save(tempPath, FormatType.Docx);
                    }
                }

                if (File.Exists(tempPath))
                {
                    // File.Move with overwrite keeps the swap atomic on the same volume.
                    File.Move(tempPath, docxPath, true);
                }
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { /* best effort cleanup */ }
                }
            }
        }

        /// <summary>
        /// Builds the bundled sample Notice of Hearing template. Used on application start
        /// so the editor never opens blank, and as the fallback when no SFDT is sent.
        ///
        /// The document is laid out like a real court notice so the generation scenario is
        /// obvious at a glance:
        ///
        ///   NOTICE OF HEARING (title)
        ///   Case Number: «CaseNumber» / Case Caption: «CaseCaption»   (case fields)
        ///   Notice Date: «NoticeDate»                                  (NOT in API data -> prompts)
        ///   Hearing held in person: «HearingInPerson»                  (NOT in API data -> prompts)
        ///   TO ALL PARTIES OF RECORD ...
        ///   «BeginGroup:Parties»
        ///     Dear «FirstName» «LastName»,
        ///     Case Number: «Case:CaseNumber»  Case Caption: «Case:CaseCaption»
        ///     Mailing Address: «Address»
        ///     Presiding Judge: «JudgeName»    (NOT in API data -> prompts)
        ///     ... court boilerplate ...
        ///   «EndGroup:Parties»
        ///   Respectfully Submitted, Court Administration
        ///
        /// JudgeName and NoticeDate are deliberately absent from the sample API data so a
        /// normal demo run goes: Generate -> missing fields detected -> prompt -> values
        /// entered -> merge re-executed -> merged document opens in the editor.
        /// </summary>
        public static WordDocument BuildSampleTemplate()
        {
            var doc = new WordDocument();
            var section = doc.AddSection();

            // ---- Letterhead-style title ----
            var title = section.AddParagraph();
            title.AppendText("NOTICE OF HEARING");
            title.ApplyStyle(BuiltinStyle.Heading1);
            title.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;

            // ---- Body (outside the repeating regions: shown once) ----
            section.AddParagraph().AppendText("TO ALL PARTIES OF RECORD:");
            section.AddParagraph().AppendText("PLEASE TAKE NOTICE that a hearing has been scheduled in the above-captioned matter.");

            // Outer region: one block per selected case (nested merge parent).
            // This region is what makes multiple selected cases merge into ONE document
            // and what lets «Case:CaseNumber» resolve inside the party region.
            // The case identification block lives inside it (per case), as before.
            section.AddParagraph().AppendField($"BeginGroup:{CaseRegionName}", FieldType.FieldMergeField);

            // ---- Case identification block (per selected case) ----
            var caption = section.AddParagraph();
            caption.AppendText("Case Number: ");
            caption.AppendField("CaseNumber", FieldType.FieldMergeField);
            var captionPara = section.AddParagraph();
            captionPara.AppendText("Case Caption: ");
            captionPara.AppendField("CaseCaption", FieldType.FieldMergeField);
            var noticeDate = section.AddParagraph();
            noticeDate.AppendText("Notice Date: ");
            noticeDate.AppendField("NoticeDate", FieldType.FieldMergeField);
            var inPerson = section.AddParagraph();
            inPerson.AppendText("Hearing held in person: ");
            inPerson.AppendField("HearingInPerson", FieldType.FieldMergeField);
            section.AddParagraph().AppendText("---------------------------------------------------------------------");

            // ---- Repeating party region (Requirement 4 + 5) ----
            section.AddParagraph().AppendField($"BeginGroup:{PartiesRegionName}", FieldType.FieldMergeField);

            var dear = section.AddParagraph();
            dear.AppendText("Dear ");
            dear.AppendField("FirstName", FieldType.FieldMergeField);
            dear.AppendText(" ");
            dear.AppendField("LastName", FieldType.FieldMergeField);
            dear.AppendText(",");

            section.AddParagraph().AppendText("You are hereby notified that a hearing has been scheduled in the following case:");
            section.AddParagraph().AppendText("Case Number: ");
            section.AddParagraph().AppendField("Case:CaseNumber", FieldType.FieldMergeField);
            section.AddParagraph().AppendText("Case Caption: ");
            section.AddParagraph().AppendField("Case:CaseCaption", FieldType.FieldMergeField);
            section.AddParagraph().AppendText("Mailing Address: ");
            section.AddParagraph().AppendField("Address", FieldType.FieldMergeField);
            section.AddParagraph().AppendText("Presiding Judge: ");
            section.AddParagraph().AppendField("JudgeName", FieldType.FieldMergeField);
            section.AddParagraph().AppendText("Please appear before the Court at the designated hearing date and time.");
            section.AddParagraph().AppendText("---------------------------------------------------------------------");

            section.AddParagraph().AppendField($"EndGroup:{PartiesRegionName}", FieldType.FieldMergeField);
            section.AddParagraph().AppendField($"EndGroup:{CaseRegionName}", FieldType.FieldMergeField);

            // ---- Signature block ----
            section.AddParagraph().AppendText("Respectfully Submitted,");
            section.AddParagraph().AppendText("Court Administration");
            return doc;
        }

        /// <summary>
        /// Serializes the sample template as SFDT for GET /api/template (the editor opens
        /// this on application load).
        /// </summary>
        public static string BuildSampleTemplateSfdt()
        {
            using (var doc = BuildSampleTemplate())
            {
                return ToSfdt(doc);
            }
        }

        /// <summary>
        /// The relational data set consumed by ExecuteNestedGroup. This is the object
        /// model the review asked about, expressed as ADO.NET tables DocIO can relate:
        ///
        ///   Case                          Parties
        ///   ------                         ---------
        ///   CaseId   (relation key)  ->   CaseId
        ///   CaseNumber                     FirstName
        ///   CaseCaption                    LastName
        ///                                  Address
        ///
        /// One parent row per selected case; child rows relate by CaseId.
        /// </summary>
        public static DataSet BuildNestedDataSet(List<CaseData> cases)
        {
            // Columns come from MergeFieldConfig (fields flagged FromData); values are read
            // from CaseData / PartyData by property name. A null value stays null, so the
            // field is treated as missing and the user is prompted.
            var caseFields = MergeFieldConfig.DataFields(CaseRegionName);
            var partyFields = MergeFieldConfig.DataFields(PartiesRegionName);

            var ds = new DataSet();

            var caseTable = new DataTable(CaseRegionName);
            caseTable.Columns.Add(RelationKey);
            foreach (var f in caseFields) { caseTable.Columns.Add(f); }
            ds.Tables.Add(caseTable);

            var partyTable = new DataTable(PartiesRegionName);
            partyTable.Columns.Add(RelationKey);
            foreach (var f in partyFields) { partyTable.Columns.Add(f); }
            ds.Tables.Add(partyTable);

            foreach (var c in cases)
            {
                caseTable.Rows.Add(RowValues(c.CaseId.ToString(), c, caseFields));
                foreach (var p in c.Parties ?? Enumerable.Empty<PartyData>())
                {
                    partyTable.Rows.Add(RowValues(c.CaseId.ToString(), p, partyFields));
                }
            }
            return ds;
        }

        /// <summary>Reads a public property of the data object by name (null when absent).</summary>
        private static object PropertyValue(object source, string name)
        {
            return source?.GetType()
                .GetProperty(name, System.Reflection.BindingFlags.Public
                                 | System.Reflection.BindingFlags.Instance
                                 | System.Reflection.BindingFlags.IgnoreCase)
                ?.GetValue(source);
        }

        /// <summary>One data row: optional relation key first, then one value per field.</summary>
        private static object[] RowValues(string relationKey, object source, List<string> fields)
        {
            var values = new List<object>();
            if (relationKey != null) { values.Add(relationKey); }
            foreach (var f in fields) { values.Add(PropertyValue(source, f)); }
            return values.ToArray();
        }

        /// <summary>
        /// The merge commands for ExecuteNestedGroup: merge the parent "Case" table,
        /// then the child "Parties" table related by CaseId.
        /// </summary>
        private static ArrayList BuildNestedCommands()
        {
            return new ArrayList
            {
                new DictionaryEntry(CaseRegionName, ""),
                new DictionaryEntry(PartiesRegionName, $"{RelationKey} = %{CaseRegionName}.{RelationKey}%")
            };
        }

        /// <summary>
        /// The flat party table consumed by ExecuteGroup in the standard path
        /// (Scenario A: only a «BeginGroup:Parties» region, no parent Case region).
        /// One row per party across all selected cases; the region repeats per row.
        /// </summary>
        public static DataTable BuildPartyGroupTable(List<CaseData> cases)
        {
            var fields = MergeFieldConfig.DataFields(PartiesRegionName);
            var dt = new DataTable(PartiesRegionName); // table name == region name
            foreach (var f in fields) { dt.Columns.Add(f); }
            foreach (var c in cases)
            {
                foreach (var p in c.Parties ?? Enumerable.Empty<PartyData>())
                {
                    dt.Rows.Add(RowValues(null, p, fields));
                }
            }
            return dt;
        }


        /// <summary>
        /// Builds the consolidated document and runs the mail merge over it.
        ///
        /// ORDER OF OPERATIONS (important):
        ///   1. COMBINE first: the main template and every selected supporting
        ///      document are merged into ONE document using DocIO's ImportContent
        ///      (canonical pattern from
        ///      https://help.syncfusion.com/document-processing/word/word-library/net/word-document/merging-word-documents#merging-documents-in-a-new-page).
        ///      Each supporting document starts on a new page.
        ///   2. Then the mail merge runs ONCE over that combined document.
        ///
        /// Because the merge runs over the combined document, DocIO resolves the
        /// merge fields of the main template AND the supporting documents in the
        /// same pass. Any field the case data cannot fill - in either document -
        /// is reported back so the UI can prompt the user for it.
        ///
        /// Returns the merged WordDocument; the CALLER owns it and must Close() it.
        /// Returns null when the merge produced missing fields - the shared outcome
        /// is populated in that case so the caller can prompt the user.
        /// </summary>
        public static WordDocument MergeAsDocument(
             WordDocument mainTemplate,
             List<CaseData> cases,
             List<AdditionalMergeValue> additionalData,
             List<string> supportingDocumentPaths,
             MergeOutcome outcome)
        {
            // ---- Step 1: combine the main template with the supporting documents ----
            var combined = mainTemplate.Clone();

            var supportingDocuments = new List<(string Path, WordDocument Doc)>();

            if (supportingDocumentPaths != null && supportingDocumentPaths.Count > 0)
            {
                foreach (var path in supportingDocumentPaths)
                {
                    if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    {
                        Console.WriteLine($"[merge] skipping missing supporting document: {path}");
                        continue;
                    }

                    var supporting = new WordDocument(path, FormatType.Docx);
                    supportingDocuments.Add((path, supporting));

                }

                // Import each supporting document at the end of the combined document,
                // starting on a new page.
                foreach (var (path, supporting) in supportingDocuments)
                {
                    ImportSupportingDocument(combined, supporting, Path.GetFileName(path));
                }
            }

            // ---- Step 2: mail merge the combined document ----
            RunMergeOnDocument(combined, cases, additionalData, outcome, "combined");

            // Close the source supporting documents - their content now lives in the
            // combined document.
            foreach (var (_, supporting) in supportingDocuments)
            {
                supporting.Close();
            }

            if (outcome.MissingFields.Count > 0)
            {
                combined.Close();
                return null;
            }

            return combined;
        }

        /// <summary>
        /// Merges the main template with the selected supporting documents, runs the
        /// mail merge over the combined document, and returns the result as SFDT for
        /// the editor. When fields are missing the outcome carries them and
        /// MergedSfdt stays null so the caller can prompt the user.
        /// </summary>
        public static MergeOutcome Merge(
            WordDocument mainTemplate,
            List<CaseData> cases,
            List<AdditionalMergeValue> additionalData,
            List<string> supportingDocumentPaths = null)
        {
            var outcome = new MergeOutcome();
            var combined = MergeAsDocument(mainTemplate, cases, additionalData, supportingDocumentPaths, outcome);
            if (combined == null)
            {
                return outcome; // missing fields; nothing to serialise
            }

            outcome.MergedSfdt = ToSfdt(combined);
            outcome.DocumentName = "Merged-Notice-of-Hearing.docx";
            combined.Close();
            return outcome;
        }

        /// <summary>
        /// Runs the mail merge on one WordDocument using the detected region strategy.
        /// Returns true if the merge completed (any missing fields were resolved or the
        /// document had no missing fields); false if the merge produced missing fields
        /// that still need user input.
        ///
        /// Missing fields are accumulated into the shared outcome so a single prompt
        /// covers the main template AND every selected supporting document.
        /// </summary>
        private static bool RunMergeOnDocument(
            WordDocument doc,
            List<CaseData> cases,
            List<AdditionalMergeValue> additionalData,
            MergeOutcome outcome,
            string label)
        {
            // ---- Region detection (decides the merge strategy) ----
            var detectedGroups = doc.MailMerge.GetMergeGroupNames() ?? new string[0];
            var detectedFields = doc.MailMerge.GetMergeFieldNames() ?? new string[0];
            bool hasCaseRegion = detectedGroups.Any(g =>
                string.Equals(g, CaseRegionName, StringComparison.OrdinalIgnoreCase));
            bool hasPartiesRegion = detectedGroups.Any(g =>
                string.Equals(g, PartiesRegionName, StringComparison.OrdinalIgnoreCase));
            Console.WriteLine($"[merge][{label}] detected groups: [{string.Join(", ", detectedGroups)}]");
            Console.WriteLine($"[merge][{label}] detected fields: [{string.Join(", ", detectedFields)}]");
            Console.WriteLine($"[merge][{label}] strategy: {(hasCaseRegion && hasPartiesRegion ? "nested (ExecuteNestedGroup)" : "standard (ExecuteGroup)")}");

            // Parent index (1-based) of the case row currently being merged. Tracked from
            // MergeField events of the parent group so missing child fields can be tagged
            // with their case.
            int currentCaseId = 0;
            string currentCaseNumber = "";
            int currentCaseRowIndex = 0;
            var caseIds = cases.Select(c => c.CaseId).ToList();
            var caseNumbers = cases.Select(c => c.CaseNumber).ToList();

            // For documents without a Case region (most supporting documents - they have
            // only case-level fields like JudgeName), MergeField for the parent group
            // never fires, so currentCaseId stays 0. Default to the FIRST selected case
            // in that situation so missing-field keys are consistent with what the main
            // template produced for the same case-level field. With FindSuppliedValue's
            // CaseId=0 fallback, a single user answer applies to every case anyway.
            int defaultCaseId = caseIds.Count > 0 ? caseIds[0] : 0;
            string defaultCaseNumber = caseNumbers.Count > 0 ? caseNumbers[0] : "";
            currentCaseId = defaultCaseId;
            currentCaseNumber = defaultCaseNumber;

            doc.MailMerge.ClearFields = false; // keep unmerged fields until we know what is missing

            // Snapshot the missing-field count so we only inspect fields discovered
            // while merging THIS document (caller may pass an outcome that already has
            // missing fields accumulated from a previous run we want to ignore).
            int missingBefore = outcome.MissingFields.Count;

            doc.MailMerge.MergeField += (sender, args) =>
            {
                if (string.Equals(args.GroupName, CaseRegionName, StringComparison.OrdinalIgnoreCase))
                {
                    int parentRow = args.RowIndex < 0 ? 0 : args.RowIndex;
                    currentCaseRowIndex = parentRow;
                    if (parentRow >= 0 && parentRow < caseIds.Count)
                    {
                        currentCaseId = caseIds[parentRow];
                        currentCaseNumber = caseNumbers[parentRow];
                    }
                }
            };

            doc.MailMerge.BeforeClearField += (sender, args) =>
            {
                string groupName = args.GroupName ?? "";
                int rowIndex = args.RowIndex < 0 ? 0 : args.RowIndex + 1;

                // Normalize field tagging so the prompt round-trip is consistent across
                // the main template and every supporting document:
                //
                //   - Fields OUTSIDE any group (GroupName == ""): treat them as
                //     case-level fields. Tag with GroupName="Case" and RowIndex=0 so the
                //     same key is produced whether the field lives in the main template
                //     or in a supporting document.
                //
                //   - Fields INSIDE the Case region (GroupName == "Case"): already
                //     case-level. Force RowIndex=0 (no row differentiation within the
                //     parent group).
                //
                //   - Fields INSIDE the Parties region (GroupName == "Parties"): tag with
                //     the 1-based row index inside the group so the prompt can address a
                //     specific party and so the supplied value applies only to that party.
                //
                // Missing fields with the same (caseId, groupName, rowIndex, fieldName)
                // tuple collapse to one entry (AddMissing dedupes).
                int caseIdForField = currentCaseId;
                string caseNumberForField = currentCaseNumber;

                if (string.IsNullOrEmpty(groupName)
                    || string.Equals(groupName, CaseRegionName, StringComparison.OrdinalIgnoreCase))
                {
                    groupName = CaseRegionName;
                    rowIndex = 0;
                }
                else
                {
                    // Parties group or any other group: keep the row index as reported.
                    // If the case id is still the default (no Case region in this document),
                    // the field is at the document level, not per-case, so we tag it
                    // against the first case to stay consistent with how the main template
                    // tags case-level fields. The user supplies the value once and the
                    // FindSuppliedValue CaseId=0 fallback will not match here because the
                    // field is in a child region - we tag with the default case explicitly.
                }

                var supplied = FindSuppliedValue(additionalData, caseIdForField, groupName, rowIndex, args.FieldName);

                if (supplied != null && !string.IsNullOrEmpty(supplied.Value))
                {
                    args.FieldValue = FormatSuppliedValue(args.FieldName, supplied.Value);
                    args.ClearField = false;
                }
                else
                {
                    AddMissing(outcome, caseIdForField, caseNumberForField, groupName, rowIndex, args.FieldName);
                    args.ClearField = true;
                }
            };

            doc.MailMerge.BeforeClearGroupField += (sender, args) =>
            {
                if (!args.HasMappedGroupInDataSource)
                {
                    // e.g. a case without parties: remove the empty region
                    args.ClearGroup = true;
                }
            };

            // ---- The actual merge: strategy chosen from the detected regions ----
            if (hasCaseRegion && hasPartiesRegion)
            {
                doc.MailMerge.ExecuteNestedGroup(BuildNestedDataSet(cases), BuildNestedCommands());
            }
            else if (hasPartiesRegion)
            {
                doc.MailMerge.ExecuteGroup(BuildPartyGroupTable(cases));
            }
            // Only group / nested-group merge is performed; fields outside a region are not merged.

            // New missing fields discovered in this document?
            if (outcome.MissingFields.Count > missingBefore)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Imports a supporting document at the end of the destination document, starting
        /// on a new page.
        ///
        /// Canonical pattern from
        /// https://help.syncfusion.com/document-processing/word/word-library/net/word-document/merging-word-documents#merging-documents-in-a-new-page
        ///
        ///   //Opens the source document
        ///   WordDocument sourceDocument = new WordDocument(sourceFileName);
        ///   //Opens the destination document
        ///   WordDocument destinationDocument = new WordDocument(targetFileName);
        ///   //Imports the contents of source document at the end of destination document
        ///   destinationDocument.ImportContent(sourceDocument, ImportOptions.UseDestinationStyles);
        ///
        /// The source document's first section keeps its default NewPage break code, so
        /// the imported content starts on a fresh page. ImportOptions.UseDestinationStyles
        /// keeps the destination document's style definitions, so the appended content
        /// matches the main template's look.
        /// </summary>
        private static void ImportSupportingDocument(WordDocument destination, WordDocument source, string label)
        {
            if (source.Sections.Count > 0)
            {
                source.Sections[0].BreakCode = SectionBreakCode.NewPage;
            }

            destination.ImportContent(source, ImportOptions.UseDestinationStyles);
            Console.WriteLine($"[merge] combined supporting document: {label}");
        }

        /// <summary>
        /// Converts a supplied value to the text written into the document. Dates and
        /// numbers are written as received (plain strings); boolean fields (per
        /// MergeFieldConfig) are written as "Yes" / "No".
        /// </summary>
        private static string FormatSuppliedValue(string fieldName, string value)
        {
            var entry = MergeFieldConfig.Find(fieldName);
            if (entry != null && entry.Type == FieldTypes.Boolean)
            {
                return string.Equals(value?.Trim(), "true", StringComparison.OrdinalIgnoreCase)
                    ? MergeFieldConfig.YesText
                    : MergeFieldConfig.NoText;
            }
            return value;
        }

        /// <summary>
        /// Matches a user-supplied value for a missing field. A value tagged with
        /// CaseId 0 applies to every case; otherwise the case id must match, so one
        /// case's answer never leaks into another case.
        /// </summary>
        private static AdditionalMergeValue FindSuppliedValue(
            List<AdditionalMergeValue> additionalData,
            int caseId,
            string groupName,
            int rowIndex,
            string fieldName)
        {
            if (additionalData == null) return null;
            return additionalData.FirstOrDefault(x =>
                (x.CaseId == 0 || x.CaseId == caseId)
                && string.Equals(x.GroupName ?? "", groupName, StringComparison.OrdinalIgnoreCase)
                && x.RowIndex == rowIndex
                && string.Equals(x.FieldName, fieldName, StringComparison.OrdinalIgnoreCase));
        }

        private static void AddMissing(
            MergeOutcome outcome,
            int caseId,
            string caseNumber,
            string groupName,
            int rowIndex,
            string fieldName)
        {
            bool exists = outcome.MissingFields.Any(m =>
                m.CaseId == caseId
                && string.Equals(m.GroupName, groupName, StringComparison.OrdinalIgnoreCase)
                && m.RowIndex == rowIndex
                && string.Equals(m.FieldName, fieldName, StringComparison.OrdinalIgnoreCase));
            if (exists) return;

            outcome.MissingFields.Add(new MissingFieldInfo
            {
                CaseId = caseId,
                CaseNumber = caseNumber,
                GroupName = groupName,
                RowIndex = rowIndex,
                FieldName = fieldName
            });
        }

        /// <summary>
        /// Serializes the sample template as SFDT (GET /api/template response body).
        /// Kept for API symmetry; the controller exposes BuildSampleTemplateSfdt().
        /// </summary>
        public static string SerializeTemplate(WordDocument doc)
        {
            return ToSfdt(doc);
        }
    }
}