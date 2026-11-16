using System.Text.Json.Serialization;

namespace EJ2APIServices.Models
{
    /// <summary>
    /// A supporting document (cover page, addendum, signature page, ...) available
    /// in the configured templates folder. Returned by GET /api/templates.
    /// </summary>
    public class SupportingDocumentInfo
    {
        [JsonPropertyName("id")]
        public string Id { get; set; }

        [JsonPropertyName("displayName")]
        public string DisplayName { get; set; }

        /// <summary>Order in which this document will be appended to the merged output.</summary>
        [JsonPropertyName("order")]
        public int Order { get; set; }

        /// <summary>
        /// True for the default main template ("Notice of Hearing"). The main template
        /// is the document that gets mail-merged, so it is offered in the template
        /// picker but NOT in the "Documents to Merge" list.
        /// </summary>
        [JsonPropertyName("isMainTemplate")]
        public bool IsMainTemplate { get; set; }
    }

    /// <summary>
    /// SFDT body for a supporting document, returned by GET /api/templates/{id}.
    /// The editor can open it directly.
    /// </summary>
    public class SupportingDocumentSfdtResponse
    {
        [JsonPropertyName("sfdt")]
        public string Sfdt { get; set; }
    }

    /// <summary>
    /// Response of POST /api/pdf/convert. Carries the converted document twice so the
    /// client has both representations available:
    ///   - Sfdt:        for the editor (lossy round-trip).
    ///   - DocxBase64:  the converted DOCX bytes base64-encoded, for layout-faithful
    ///                  download (preserves tables, headers, bold/italic, etc.).
    /// </summary>
    public class ConvertPdfResponse
    {
        [JsonPropertyName("sfdt")]
        public string Sfdt { get; set; }

        [JsonPropertyName("docxBase64")]
        public string DocxBase64 { get; set; }

        [JsonPropertyName("fileName")]
        public string FileName { get; set; }

        [JsonPropertyName("fileSizeBytes")]
        public long FileSizeBytes { get; set; }
    }

    /// <summary>
    /// Body of POST /api/templates/{id}/save: the editor's current document
    /// serialized as SFDT. The server converts it back to DOCX and writes it over
    /// the file in the templates folder.
    /// </summary>
    public class SaveTemplateRequest
    {
        [JsonPropertyName("sfdt")]
        public string Sfdt { get; set; }
    }

    /// <summary>
    /// Result of a successful save.
    /// </summary>
    public class SaveTemplateResponse
    {
        [JsonPropertyName("id")]
        public string Id { get; set; }

        [JsonPropertyName("displayName")]
        public string DisplayName { get; set; }

        /// <summary>Absolute path of the written file (server-side detail).</summary>
        [JsonPropertyName("path")]
        public string Path { get; set; }
    }
}
