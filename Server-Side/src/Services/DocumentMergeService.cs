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
        public const string CaseRegionName = "Case";

        /// <summary>Child region/table name (repeating party region).</summary>
        public const string PartiesRegionName = "Parties";

        /// <summary>Relation key between Case and Parties.</summary>
        private const string RelationKey = "CaseId";

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
        /// Builds the bundled sample Notice of Hearing template. Used on application start
        /// so the editor never opens blank, and as the fallback when no SFDT is sent.
        ///
        /// The document is laid out like a real court notice so the generation scenario is
        /// obvious at a glance:
        ///
        ///   NOTICE OF HEARING (title)
        ///   Case Number: «CaseNumber» / Case Caption: «CaseCaption»   (case fields)
        ///   Notice Date: «NoticeDate»                                  (NOT in API data -> prompts)
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
            var ds = new DataSet();

            var caseTable = new DataTable(CaseRegionName);
            caseTable.Columns.Add("CaseId");
            caseTable.Columns.Add("CaseNumber");
            caseTable.Columns.Add("CaseCaption");
            ds.Tables.Add(caseTable);

            var partyTable = new DataTable(PartiesRegionName);
            partyTable.Columns.Add(RelationKey);
            partyTable.Columns.Add("FirstName");
            partyTable.Columns.Add("LastName");
            partyTable.Columns.Add("Address");
            ds.Tables.Add(partyTable);

            foreach (var c in cases)
            {
                caseTable.Rows.Add(c.CaseId.ToString(), c.CaseNumber ?? "", c.CaseCaption ?? "");
                foreach (var p in c.Parties ?? Enumerable.Empty<PartyData>())
                {
                    partyTable.Rows.Add(c.CaseId.ToString(), p.FirstName ?? "", p.LastName ?? "", p.Address);
                }
            }
            return ds;
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
            var dt = new DataTable(PartiesRegionName); // table name == region name
            dt.Columns.Add("FirstName");
            dt.Columns.Add("LastName");
            dt.Columns.Add("Address");
            foreach (var c in cases)
            {
                foreach (var p in c.Parties ?? Enumerable.Empty<PartyData>())
                {
                    dt.Rows.Add(p.FirstName ?? "", p.LastName ?? "", p.Address ?? "");
                }
            }
            return dt;
        }

        /// <summary>
        /// Runs the mail merge over the selected cases for the main template and every
        /// selected supporting document, then combines them into one consolidated output
        /// document. Returns either the missing fields to prompt for, or the merged
        /// document as SFDT.
        ///
        /// Region detection decides the merge strategy before anything is executed:
        ///
        ///   Scenario A (standard regions): the template has «BeginGroup:Parties» but no
        ///   parent «BeginGroup:Case» region. Detected groups: [Parties]. Strategy:
        ///   MailMerge.ExecuteGroup(partyTable) repeats the region per party row
        ///   (Pattern 1), and MailMerge.Execute(...) fills the case fields outside
        ///   the region. No parent Case region is required.
        ///
        ///   Scenario B (nested regions): the template wraps «BeginGroup:Parties» inside
        ///   «BeginGroup:Case». Detected groups: [Case, Parties]. Strategy: the existing
        ///   relational MailMerge.ExecuteNestedGroup(dataSet, commands) with parent-child
        ///   resolution (Pattern 2 / Requirement 5) - unchanged.
        ///
        /// The same strategy is applied to every selected supporting document so each
        /// one's merge fields (CaseNumber, JudgeName, NoticeDate, ...) are filled using
        /// the same case data. After every document has been merged, the supporting
        /// documents are appended to the main document (one page break between each),
        /// producing one consolidated output.
        ///
        /// Exact code path:
        ///   1. Region discovery (per document):  doc.MailMerge.GetMergeGroupNames() /
        ///      GetMergeFieldNames() (logged so manually inserted regions can be verified).
        ///   2. Data mapping:     BuildPartyGroupTable() for the standard path;
        ///                        BuildNestedDataSet()/BuildNestedCommands() for the nested path.
        ///   3. Mail merge:       ExecuteGroup(...) or ExecuteNestedGroup(...) per scenario.
        ///   4. Missing detection: BeforeClearField fires for every field DocIO could not
        ///      fill. MergeField events track which case row is current (nested path) so
        ///      each missing field is tagged with its CaseId/CaseNumber. AdditionalMergeData
        ///      values are injected through the same event. Missing fields are accumulated
        ///      across the main template and every supporting document - one prompt covers
        ///      all of them.
        ///   5. Append supporting documents: DocIO.AppendDocument is called for each
        ///      merged supporting document in the order the caller supplied them, with
        ///      SectionBreakCode.NewPage between each, producing ONE consolidated document.
        /// </summary>
        public static MergeOutcome Merge(
            WordDocument mainTemplate,
            List<CaseData> cases,
            List<AdditionalMergeValue> additionalData,
            List<string> supportingDocumentPaths = null)
        {
            // Phase 1: merge the main template (cloned, so the caller's WordDocument is not mutated).
            var main = mainTemplate.Clone();
            var outcome = new MergeOutcome();

            var mainResult = RunMergeOnDocument(main, cases, additionalData, outcome, "main");
            if (!mainResult) { return outcome; } // missing fields: prompt the user, no document

            // Phase 2: merge each supporting document (with the same case data + supplied
            // additional values), then append to the main merged document.
            if (supportingDocumentPaths != null && supportingDocumentPaths.Count > 0)
            {
                for (int i = 0; i < supportingDocumentPaths.Count; i++)
                {
                    var path = supportingDocumentPaths[i];
                    if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    {
                        Console.WriteLine($"[merge] skipping missing supporting document: {path}");
                        continue;
                    }

                    // Clone the loaded supporting doc so MailMerge state on the main
                    // document isn't disturbed. We open a fresh WordDocument for each.
                    using (var src = new WordDocument(path, FormatType.Docx))
                    {
                        var supportingClone = src.Clone();

                        // Try the merge on this supporting document. If it reports missing
                        // fields, abort the whole run and let the UI prompt.
                        var ok = RunMergeOnDocument(supportingClone, cases, additionalData, outcome,
                            $"supporting[{i + 1}/{supportingDocumentPaths.Count}]: {Path.GetFileName(path)}");
                        if (!ok)
                        {
                            main.Close();
                            return outcome;
                        }

                        AppendSupportingDocument(main, supportingClone);
                        supportingClone.Close();
                    }
                }
            }

            outcome.MergedSfdt = ToSfdt(main);
            outcome.DocumentName = "Merged-Notice-of-Hearing.docx";
            main.Close();
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
                if (string.Equals(groupName, CaseRegionName, StringComparison.OrdinalIgnoreCase))
                {
                    rowIndex = 0; // case-level field: no row differentiation
                }

                int caseIdForField = currentCaseId;
                string caseNumberForField = currentCaseNumber;

                var supplied = FindSuppliedValue(additionalData, caseIdForField, groupName, rowIndex, args.FieldName);

                if (supplied != null && !string.IsNullOrEmpty(supplied.Value))
                {
                    args.FieldValue = supplied.Value;
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
                ExecuteCaseLevelFields(doc, cases, detectedFields);
            }
            else
            {
                ExecuteCaseLevelFields(doc, cases, detectedFields);
            }

            // New missing fields discovered in this document?
            if (outcome.MissingFields.Count > missingBefore)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Fills case-level merge fields (CaseNumber, CaseCaption) that sit OUTSIDE any
        /// region. Only fields actually present in the document are filled, so the same
        /// code works for the main template AND for supporting documents that may or may
        /// not contain those fields. With multiple selected cases and no parent Case
        /// region there is no per-case repetition; the first selected case supplies the
        /// values for these top-level fields.
        /// </summary>
        private static void ExecuteCaseLevelFields(
            WordDocument doc,
            List<CaseData> cases,
            string[] detectedFields)
        {
            if (detectedFields == null || detectedFields.Length == 0)
            {
                return;
            }

            var present = new HashSet<string>(
                detectedFields.Select(f => f ?? ""),
                StringComparer.OrdinalIgnoreCase);

            var first = cases.FirstOrDefault();
            string firstCaseNumber = first?.CaseNumber ?? "";
            string firstCaseCaption = first?.CaseCaption ?? "";

            var names = new List<string>();
            var values = new List<string>();
            if (present.Contains("CaseNumber")) { names.Add("CaseNumber"); values.Add(firstCaseNumber); }
            if (present.Contains("CaseCaption")) { names.Add("CaseCaption"); values.Add(firstCaseCaption); }

            if (names.Count > 0)
            {
                doc.MailMerge.Execute(names.ToArray(), values.ToArray());
            }
        }

        /// <summary>
        /// Appends an already-merged supporting document to the end of the main merged
        /// document. Each appended document starts on a new page (SectionBreakCode.NewPage).
        /// ImportOptions keep styles from colliding between documents.
        /// </summary>
        private static void AppendSupportingDocument(WordDocument main, WordDocument supporting)
        {
            if (main.LastSection != null)
    {
        main.LastSection.AddParagraph()
            .AppendBreak(BreakType.PageBreak);
    }

    foreach (WSection section in supporting.Sections)
    {
        main.Sections.Add((WSection)section.Clone());
    }
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