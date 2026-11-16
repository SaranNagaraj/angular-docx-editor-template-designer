using System;
using System.IO;
using EJ2APIServices.Models;
using EJ2APIServices.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.SmartDataExtractor;

namespace EJ2APIServices.Controllers
{
    /// <summary>
    /// PDF -> Word conversion API:
    /// POST /api/pdf/convert   (multipart upload; returns the converted document as SFDT
    ///                           for the editor AND as base64 DOCX for high-fidelity
    ///                           download)
    ///
    /// The conversion follows the canonical pattern from
    /// https://help.syncfusion.com/document-processing/data-extraction/net/conversions/pdf-to-word#convert-pdf-or-image-to-word-document
    ///
    ///   DataExtractor extractor = new DataExtractor();
    ///   WordDocument document = extractor.ExtractDataAsWordDocument(stream);
    ///   document.Save(saveStream, FormatType.Docx);
    ///
    /// The converted DOCX is then re-serialised as SFDT so the Angular Document
    /// Editor can open it directly. The SFDT round-trip is lossy (it cannot carry
    /// arbitrary text-frame positioning, run-level styling or floating shapes), so
    /// the response also carries the original DOCX bytes base64-encoded - the
    /// client offers a "Download DOCX" link that preserves the full PDF layout.
    /// </summary>
    [Route("api/[controller]")]
    public class PdfController : Controller
    {
        [HttpPost("convert")]
        [RequestSizeLimit(50 * 1024 * 1024)] // 50 MB
        public IActionResult Convert(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { message = "No PDF file was uploaded." });
            }

            // Defence in depth on top of the accept="application/pdf" attribute on the
            // client-side file input: check the file really starts with the PDF header.
            if (!LooksLikePdf(file))
            {
                return BadRequest(new { message = "Only PDF files are accepted." });
            }

            try
            {
                // Open the input PDF file as a stream.
                using (var pdfStream = file.OpenReadStream())
                {
                    // Initialize the Data Extractor.
                    var extractor = new DataExtractor();

                    // Extract data as WordDocument.
                    using (var wordDoc = extractor.ExtractDataAsWordDocument(pdfStream))
                    {
                        // Save the extracted Word data into an output stream. This
                        // DOCX preserves as much of the original PDF layout as
                        // Syncfusion's SmartDataExtractor can reconstruct (tables,
                        // lists, headers, bold/italic at the paragraph level, ...).
                        byte[] docxBytes;
                        using (var saveStream = new MemoryStream())
                        {
                            wordDoc.Save(saveStream, FormatType.Docx);
                            docxBytes = saveStream.ToArray();
                        }

                        // Re-open the converted DOCX and serialise it as SFDT so the
                        // Angular Document Editor can open it directly. NOTE: the
                        // SFDT round-trip is lossy compared with the DOCX above;
                        // for layout-faithful work the user should download the
                        // DOCX instead.
                        string sfdt;
                        using (var converted = new WordDocument(new MemoryStream(docxBytes), FormatType.Docx))
                        {
                            sfdt = DocumentMergeService.ToSfdt(converted);
                        }

                        var displayName = StripDocxExtension(SafeFileName(file.FileName)) + ".docx";
                        return Ok(new ConvertPdfResponse
                        {
                            Sfdt = sfdt,
                            DocxBase64 = Convert.ToBase64String(docxBytes),
                            FileName = displayName,
                            FileSizeBytes = docxBytes.Length
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Could not convert the PDF: {ex.Message}" });
            }
        }

        private static string SafeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "converted";
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new System.Text.StringBuilder(name.Length);
            foreach (var ch in name)
            {
                sb.Append(System.Array.IndexOf(invalid, ch) >= 0 ? '_' : ch);
            }
            return sb.ToString().Trim();
        }

        private static string StripDocxExtension(string fileName)
        {
            return fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                ? fileName.Substring(0, fileName.Length - ".docx".Length)
                : fileName;
        }

        /// <summary>
        /// Checks that the uploaded file is a PDF by sniffing the first four bytes for
        /// the "%PDF" magic header. The browser's accept attribute is not a security
        /// boundary, so the content is verified server-side.
        /// </summary>
        private static bool LooksLikePdf(IFormFile file)
        {
            if (file.ContentType != null
                && !file.ContentType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase)
                && !file.ContentType.Equals("application/octet-stream", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            using (var s = file.OpenReadStream())
            {
                Span<byte> head = stackalloc byte[4];
                int read = s.Read(head);
                return read == 4
                    && head[0] == (byte)'%'
                    && head[1] == (byte)'P'
                    && head[2] == (byte)'D'
                    && head[3] == (byte)'F';
            }
        }
    }
}
