using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EJ2APIServices.Models;

namespace EJ2APIServices.Data
{
    /// <summary>
    /// Lists and loads supporting documents (templates) from a configured folder on disk.
    /// Files are read-only and addressed by their file name (id).
    ///
    /// If the configured folder does not exist, the repository seeds a default folder
    /// containing five plain DOCX files (Cover Page, Addendum A, Addendum B,
    /// Terms & Conditions, Signature Page) so the sample runs out-of-the-box without
    /// any manual setup.
    /// </summary>
    public static class SupportingDocumentRepository
    {
        private const string DefaultRelativeFolder = "App_Data/SupportingDocuments";

        // Stable order used when the folder is auto-seeded and when callers
        // do not pass an explicit order. Real folder listings fall back to alphabetical.
        private static readonly (string FileName, string DisplayName, int Order)[] SeededDocuments =
        {
            ("Cover Page.docx",            "Cover Page",          1),
            ("Addendum A.docx",            "Addendum A",          2),
            ("Addendum B.docx",            "Addendum B",          3),
            ("Terms and Conditions.docx",  "Terms & Conditions",  4),
            ("Signature Page.docx",        "Signature Page",      5),
        };

        /// <summary>
        /// Returns the absolute path to the configured folder, seeding it on first use
        /// so the sample always has something to list.
        /// </summary>
        public static string EnsureFolder(string contentRootPath, string configuredFolder)
        {
            var folder = string.IsNullOrWhiteSpace(configuredFolder)
                ? Path.Combine(contentRootPath, DefaultRelativeFolder)
                : (Path.IsPathRooted(configuredFolder)
                    ? configuredFolder
                    : Path.Combine(contentRootPath, configuredFolder));

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
                SeedDefaultDocuments(folder);
            }

            return folder;
        }

        public static List<SupportingDocumentInfo> List(string folder)
        {
            if (!Directory.Exists(folder))
            {
                return new List<SupportingDocumentInfo>();
            }

            var files = Directory.EnumerateFiles(folder, "*.docx", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileName)
                .Where(name => !string.IsNullOrEmpty(name))
                .ToList();

            return files
                .Select(name =>
                {
                    var seeded = SeededDocuments.FirstOrDefault(s =>
                        string.Equals(s.FileName, name, StringComparison.OrdinalIgnoreCase));
                    return new SupportingDocumentInfo
                    {
                        Id = name,
                        DisplayName = seeded.FileName != null ? seeded.DisplayName : StripDocxExtension(name),
                        Order = seeded.FileName != null ? seeded.Order : int.MaxValue
                    };
                })
                .OrderBy(d => d.Order)
                .ThenBy(d => d.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public static string GetFilePath(string folder, string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return null;
            }

            // Reject path traversal: id must be a plain file name with no separators.
            if (id.IndexOfAny(new[] { '/', '\\' }) >= 0)
            {
                return null;
            }

            var path = Path.Combine(folder, id);
            return File.Exists(path) ? path : null;
        }

        private static string StripDocxExtension(string fileName)
        {
            return fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                ? fileName.Substring(0, fileName.Length - ".docx".Length)
                : fileName;
        }

        /// <summary>
        /// Creates five very simple DOCX files via DocIO so the folder has real content
        /// even on a clean checkout. Seeding is best-effort: a missing or invalid
        /// DocIO license (or any other failure) is logged and the folder is left empty
        /// rather than failing the whole request.
        /// </summary>
        private static void SeedDefaultDocuments(string folder)
        {
            foreach (var entry in SeededDocuments)
            {
                var path = Path.Combine(folder, entry.FileName);
                if (File.Exists(path))
                {
                    continue;
                }

                try
                {
                    using (var doc = new Syncfusion.DocIO.DLS.WordDocument())
                    {
                        var section = doc.AddSection();
                        var heading = section.AddParagraph();
                        heading.AppendText(entry.DisplayName);
                        heading.ApplyStyle(Syncfusion.DocIO.DLS.BuiltinStyle.Heading1);
                        heading.ParagraphFormat.HorizontalAlignment =
                            Syncfusion.DocIO.DLS.HorizontalAlignment.Center;

                        section.AddParagraph().AppendText(GetBodyText(entry.DisplayName));
                        section.AddParagraph().AppendText();
                        section.AddParagraph().AppendText();

                        using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write))
                        {
                            doc.Save(stream, Syncfusion.DocIO.FormatType.Docx);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[seed] Could not seed '{entry.FileName}': {ex.Message}");
                }
            }
        }

        private static string GetBodyText(string displayName)
        {
            switch (displayName)
            {
                case "Cover Page":
                    return "Court of Record — Cover sheet for the consolidated case file.";
                case "Addendum A":
                    return "Addendum A — Supporting exhibits and referenced documents.";
                case "Addendum B":
                    return "Addendum B — Witness list and statement summaries.";
                case "Terms & Conditions":
                    return "Standard terms and conditions governing the proceedings described in the main notice.";
                case "Signature Page":
                    return "Signature page — sign and date below to acknowledge receipt.";
                default:
                    return string.Empty;
            }
        }
    }
}
