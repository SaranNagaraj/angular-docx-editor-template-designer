using System;
using System.IO;
using EJ2APIServices.Models;
using EJ2APIServices.Services;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO.DLS;
using Syncfusion.PdfToDocConverter;

namespace EJ2APIServices.Controllers
{
    /// <summary>
    /// PDF -> DOCX conversion API:
    /// POST /api/pdf/convert            (multipart upload; returns SFDT for the editor)
    /// POST /api/pdf/convert-and-save   (multipart upload; saves to templates folder and
    ///                                  returns the resulting template id + SFDT)
    ///
    /// The conversion uses Syncfusion's PdfToDocConverter. The output WordDocument is
    /// re-serialised as SFDT so the Angular editor can open it directly without ever
    /// dealing with a DOCX file on disk.
    /// </summary>
    [Route("api/[controller]")]
    public class PdfController : Controller
    {
        private readonly IWebHostEnvironment _env;
        private readonly Microsoft.Extensions.Configuration.IConfiguration _config;

        public PdfController(IWebHostEnvironment env, Microsoft.Extensions.Configuration.IConfiguration config)
        {
            _env = env;
            _config = config;
        }

        [HttpPost("convert")]
        [RequestSizeLimit(50 * 1024 * 1024)] // 50 MB
        public IActionResult Convert(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { message = "No PDF file was uploaded." });
            }

            // Make sure the uploaded file is a PDF (defence in depth on top of the
            // accept="application/pdf" attribute on the file input).
            if (!LooksLikePdf(file))
            {
                return BadRequest(new { message = "Only PDF files are accepted." });
            }

            try
            {
                using (var pdfStream = file.OpenReadStream())
                {
                    var converter = new PdfToDocConverter();
                    using (var wordDoc = converter.Convert(pdfStream))
                    {
                        var sfdt = DocumentMergeService.ToSfdt(wordDoc);
                        return Ok(new SupportingDocumentSfdtResponse { Sfdt = sfdt });
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Could not convert the PDF: {ex.Message}" });
            }
        }

        /// <summary>
        /// Converts the uploaded PDF and writes the result to the templates folder, so
        /// the user can save and reuse it like any other template. Returns the saved
        /// template id, its display name, and the converted SFDT so the editor can open
        /// it without an extra round trip.
        /// </summary>
        [HttpPost("convert-and-save")]
        [RequestSizeLimit(50 * 1024 * 1024)]
        public IActionResult ConvertAndSave(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { message = "No PDF file was uploaded." });
            }
            if (!LooksLikePdf(file))
            {
                return BadRequest(new { message = "Only PDF files are accepted." });
            }

            // Derive a destination file name: original name minus the .pdf extension,
            // plus ".docx". Sanitised so the file system is safe and the server's path
            // traversal protection (id must be a plain file name with no path
            // separators) is satisfied.
            var originalName = Path.GetFileNameWithoutExtension(file.FileName ?? "uploaded");
            var baseName = SanitiseFileName(originalName);
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = "uploaded";
            }

            var folder = Data.SupportingDocumentRepository.GetFolder(
                _env.ContentRootPath,
                _config["TemplatesFolder"]);

            // Ensure the templates folder exists; PdfToDocConverter does not need it,
            // but the save target must be writable.
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            var targetPath = Path.Combine(folder, baseName + ".docx");
            targetPath = AvoidCollision(targetPath);

            try
            {
                using (var pdfStream = file.OpenReadStream())
                {
                    var converter = new PdfToDocConverter();
                    using (var wordDoc = converter.Convert(pdfStream))
                    {
                        // Persist the DOCX so the user can re-load it later and so the
                        // Save Document action (which writes back as DOCX) keeps working.
                        wordDoc.Save(targetPath, FormatType.Docx);
                        var sfdt = DocumentMergeService.ToSfdt(wordDoc);

                        return Ok(new SaveTemplateResponse
                        {
                            Id = Path.GetFileName(targetPath),
                            DisplayName = Path.GetFileNameWithoutExtension(targetPath),
                            Path = targetPath,
                            Sfdt = sfdt
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Could not convert and save the PDF: {ex.Message}" });
            }
        }

        private static bool LooksLikePdf(IFormFile file)
        {
            if (file.ContentType != null &&
                !file.ContentType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase) &&
                !file.ContentType.Equals("application/octet-stream", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // Sniff the first 4 bytes for the PDF magic header "%PDF".
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

        private static string SanitiseFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return string.Empty;
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new System.Text.StringBuilder(name.Length);
            foreach (var ch in name)
            {
                sb.Append(Array.IndexOf(invalid, ch) >= 0 ? '_' : ch);
            }
            return sb.ToString().Trim();
        }

        /// <summary>
        /// If a file with the same name already exists, append " (n)" before the
        /// extension so the upload never silently overwrites an existing template.
        /// </summary>
        private static string AvoidCollision(string path)
        {
            if (!File.Exists(path)) return path;

            var dir = Path.GetDirectoryName(path) ?? string.Empty;
            var stem = Path.GetFileNameWithoutExtension(path);
            var ext = Path.GetExtension(path);
            for (int i = 1; i < 1000; i++)
            {
                var candidate = Path.Combine(dir, $"{stem} ({i}){ext}");
                if (!File.Exists(candidate)) return candidate;
            }
            return path;
        }
    }
}
