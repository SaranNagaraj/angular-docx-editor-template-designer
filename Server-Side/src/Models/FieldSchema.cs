using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace EJ2APIServices.Models
{
    /// <summary>
    /// Type tag for a merge field. The client uses this to render the appropriate
    /// input control in the missing-field dialog (date picker, numeric text box,
    /// checkbox, plain text). All values are returned to the server as strings;
    /// the type only drives UI rendering and validation.
    /// </summary>
    public static class FieldTypes
    {
        public const string String = "string";
        public const string Date = "date";
        public const string Number = "number";
        public const string Boolean = "boolean";
    }

    /// <summary>
    /// One merge field entry in the schema.
    /// </summary>
    public class FieldSchemaEntry
    {
        [JsonPropertyName("type")]
        public string Type { get; set; } = FieldTypes.String;

        [JsonPropertyName("label")]
        public string Label { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        /// <summary>number fields only: accept whole numbers only (e.g. a count of days).</summary>
        [JsonPropertyName("integer")]
        public bool Integer { get; set; }

        /// <summary>number fields only: optional lower bound.</summary>
        [JsonPropertyName("min")]
        public double? Min { get; set; }

        /// <summary>number fields only: optional upper bound.</summary>
        [JsonPropertyName("max")]
        public double? Max { get; set; }

        /// <summary>
        /// True when the value comes from the case data (a property with the same name on
        /// CaseData / PartyData). False means the value is NOT in the data, so the user is
        /// prompted for it. Server-only: not sent to the client.
        /// </summary>
        [JsonIgnore]
        public bool FromData { get; set; }
    }

    /// <summary>
    /// Merge field schema served by GET /api/schema: region (group) name -> field name -> entry.
    /// The Angular Merge Field Explorer AND the missing-field dialog are built entirely
    /// from this response, so the client has no hardcoded field or region names.
    /// </summary>
    public class FieldSchema : Dictionary<string, Dictionary<string, FieldSchemaEntry>>
    {
    }

    /// <summary>
    /// ============================================================================
    ///  THE ONE PLACE FOR MERGE FIELD DEFINITIONS.
    ///  When your merge fields change, edit this class (and your data models) only.
    /// ============================================================================
    ///
    /// It defines:
    ///   1. the region (group) names and the key that relates child rows to parent rows,
    ///   2. every merge field: type, label, validation rules, and whether its value
    ///      comes from the data (FromData) or is prompted for.
    ///
    /// Everything else derives from it:
    ///   - GET /api/schema        -> Merge Fields panel, dialog control, label, validation
    ///   - DocumentMergeService   -> region names, the nested data tables' columns,
    ///                               and how supplied values are written (Yes / No).
    ///
    /// How to change things:
    ///   - New field the user is prompted for : add one entry below (FromData = false).
    ///   - New field filled from the data     : add the property to CaseData / PartyData
    ///                                          and one entry below with FromData = true
    ///                                          (the property name must equal the field name).
    ///   - Change a field's type or rules     : edit its entry.
    ///   - Different region names             : change ParentRegion / ChildRegion and the
    ///                                          keys of the two groups.
    /// Also update your templates (the merge fields inside the regions) to match.
    /// </summary>
    public static class MergeFieldConfig
    {
        /// <summary>Parent region name (repeats once per selected case).</summary>
        public const string ParentRegion = "Case";

        /// <summary>Child region name (repeats once per party inside a case).</summary>
        public const string ChildRegion = "Parties";

        /// <summary>Column that relates child rows to their parent row (not a merge field).</summary>
        public const string RelationKey = "CaseId";

        public const string YesText = "Yes";
        public const string NoText = "No";

        /// <summary>The full schema, built once.</summary>
        public static readonly FieldSchema Schema = Build();

        /// <summary>
        /// Builds the schema: for each region, the list of merge fields it can contain.
        /// This is the ONLY list of merge fields in the application; add, remove or
        /// change fields here.
        ///
        /// Each entry (field name = dictionary key, must match the MERGEFIELD name used in
        /// the templates) has:
        ///   Type      - drives the prompt dialog control and validation
        ///               (string / date / number / boolean).
        ///   Label     - friendly display text only. The UI shows it instead of the raw
        ///               field name: in the missing-value prompt ("Enter Presiding Judge
        ///               (case ...)" rather than "Enter JudgeName"). It is never used for
        ///               matching or merging, so it is safe to change freely. If omitted,
        ///               the raw field name is shown.
        ///   FromData  - true: the value is read from CaseData / PartyData (a property
        ///               with the same name), so no prompt is needed.
        ///               false (default): not in the data, so the user is prompted.
        ///   Integer / Min / Max - optional number rules (whole numbers, bounds).
        /// </summary>
        private static FieldSchema Build()
        {
            return new FieldSchema
            {
                [ParentRegion] = new Dictionary<string, FieldSchemaEntry>
                {
                    // ---- filled from the case data (properties on CaseData) ----
                    ["CaseNumber"]          = new FieldSchemaEntry { Type = FieldTypes.String,  Label = "Case Number", FromData = true },
                    ["CaseCaption"]         = new FieldSchemaEntry { Type = FieldTypes.String,  Label = "Case Caption", FromData = true },
                    // ---- not in the data: the user is prompted when a template uses them ----
                    // Date    -> date picker, the chosen date is sent as a plain string
                    // Number  -> numeric box; NoticeDays accepts whole numbers, minimum 0
                    // Boolean -> checkbox; merged into the document as "Yes" / "No"
                    ["NoticeDate"]          = new FieldSchemaEntry { Type = FieldTypes.Date,    Label = "Notice Date" },
                    ["HearingDate"]         = new FieldSchemaEntry { Type = FieldTypes.Date,    Label = "Hearing Date" },
                    ["JudgeName"]           = new FieldSchemaEntry { Type = FieldTypes.String,  Label = "Presiding Judge" },
                    ["CourtroomNumber"]     = new FieldSchemaEntry { Type = FieldTypes.String,  Label = "Courtroom Number" },
                    ["NoticeDays"]          = new FieldSchemaEntry { Type = FieldTypes.Number,  Label = "Notice Days", Integer = true, Min = 0 },
                    ["HearingInPerson"]     = new FieldSchemaEntry { Type = FieldTypes.Boolean, Label = "Hearing Held In Person" },
                    ["PhotoIdRequired"]     = new FieldSchemaEntry { Type = FieldTypes.Boolean, Label = "Photo ID Required" },
                    ["InterpreterRequired"] = new FieldSchemaEntry { Type = FieldTypes.Boolean, Label = "Interpreter Requested" }
                },
                // Fields inside the child region repeat once per party (PartyData properties).
                [ChildRegion] = new Dictionary<string, FieldSchemaEntry>
                {
                    ["FirstName"] = new FieldSchemaEntry { Type = FieldTypes.String, Label = "First Name",      FromData = true },
                    ["LastName"]  = new FieldSchemaEntry { Type = FieldTypes.String, Label = "Last Name",       FromData = true },
                    ["Address"]   = new FieldSchemaEntry { Type = FieldTypes.String, Label = "Mailing Address", FromData = true }
                }
            };
        }

        /// <summary>Looks a field up by name across all regions; null when it is not in the schema.</summary>
        public static FieldSchemaEntry Find(string fieldName)
        {
            if (string.IsNullOrEmpty(fieldName)) { return null; }
            foreach (var group in Schema.Values)
            {
                foreach (var kv in group)
                {
                    if (string.Equals(kv.Key, fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        return kv.Value;
                    }
                }
            }
            return null;
        }

        /// <summary>Names of the fields in a region whose values come from the data.</summary>
        public static List<string> DataFields(string region)
        {
            return Schema.TryGetValue(region, out var fields)
                ? fields.Where(kv => kv.Value.FromData).Select(kv => kv.Key).ToList()
                : new List<string>();
        }
    }
}
