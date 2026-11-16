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
}
