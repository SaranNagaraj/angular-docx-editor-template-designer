using System.Collections.Generic;
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
    /// One merge field entry in the schema. Carries the type tag plus an optional
    /// human-readable label and description for the UI.
    /// </summary>
    public class FieldSchemaEntry
    {
        [JsonPropertyName("type")]
        public string Type { get; set; } = FieldTypes.String;

        [JsonPropertyName("label")]
        public string Label { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }
    }

    /// <summary>
    /// Merge field schema served by GET /api/schema.
    /// The Angular Merge Field Explorer AND the missing-field dialog are built
    /// entirely from this response, so the front end has no hardcoded field
    /// definitions.
    /// </summary>
    public class FieldSchema
    {
        /// <summary>Case-level merge fields (one value per case).</summary>
        public Dictionary<string, FieldSchemaEntry> Case { get; set; }

        /// <summary>Party-level merge fields (repeat for every party in a case).</summary>
        public Dictionary<string, FieldSchemaEntry> Parties { get; set; }
    }
}
