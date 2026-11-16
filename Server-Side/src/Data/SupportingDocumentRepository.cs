using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EJ2APIServices.Models;

namespace EJ2APIServices.Data
{
    /// <summary>
    /// Reads supporting documents (templates) from the configured folder on disk.
    /// Files are addressed by their file name (id) and are expected to be real DOCX
    /// files authored by the user - this repository does NOT generate any document
    /// via DocIO.
    ///
    /// Folder layout (default: Server-Side/src/wwwroot/Templates):
    ///
    ///   Notice of Hearing.docx     - main template, loaded by default on app start
    ///   Cover Page.docx            - supporting document
    ///   Addendum A.docx            - supporting document
    ///   Addendum B.docx            - supporting document
    ///   Terms and Conditions.docx  - supporting document
    ///   Signature Page.docx        - supporting document
    ///
    /// Any *.docx dropped into the folder is listed by GET /api/templates. The
    /// "Notice of Hearing" file is also served via GET /api/template as the
    /// default main template.
    /// </summary>
    public static class SupportingDocumentRepository
    {
        /// <summary>Default folder relative to the content root (Server-Side/src).</summary>
        public const string DefaultRelativeFolder = "wwwroot/Templates";

        /// <summary>Default main template file name (must live in the templates folder).</summary>
        public const string DefaultMainTemplateFileName = "Notice of Hearing.docx";

        // Stable display order for the supporting documents when they exist.
        // Files not in this list fall back to alphabetical ordering.
        private static readonly (string FileName, string DisplayName, int Order)[] KnownSupportingDocuments =
        {
            ("Cover Page.docx",            "Cover Page",          1),
            ("Addendum A.docx",            "Addendum A",          2),
            ("Addendum B.docx",            "Addendum B",          3),
            ("Terms and Conditions.docx",  "Terms & Conditions",  4),
            ("Signature Page.docx",        "Signature Page",      5),
        };

        /// <summary>
        /// Returns the absolute path to the configured folder. The folder is NOT
        /// auto-created here: callers must arrange for the folder to exist (the
        /// repo ships a .gitkeep placeholder so the directory is tracked).
        /// </summary>
        public static string GetFolder(string contentRootPath, string configuredFolder)
        {
            var folder = string.IsNullOrWhiteSpace(configuredFolder)
                ? Path.Combine(contentRootPath, DefaultRelativeFolder)
                : (Path.IsPathRooted(configuredFolder)
                    ? configuredFolder
                    : Path.Combine(contentRootPath, configuredFolder));

            return folder;
        }

        /// <summary>
        /// Lists every .docx in the folder except the default main template (which is
        /// served separately via GET /api/template). Returned in stable order:
        /// known supporting documents first (in their defined order), then any
        /// additional files in alphabetical order.
        /// </summary>
        public static List<SupportingDocumentInfo> ListSupporting(string folder)
        {
            if (!Directory.Exists(folder))
            {
                return new List<SupportingDocumentInfo>();
            }

            var files = Directory.EnumerateFiles(folder, "*.docx", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileName)
                .Where(name => !string.IsNullOrEmpty(name))
                .Where(name => !string.Equals(name, DefaultMainTemplateFileName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            return files
                .Select(name =>
                {
                    var known = KnownSupportingDocuments.FirstOrDefault(s =>
                        string.Equals(s.FileName, name, StringComparison.OrdinalIgnoreCase));
                    return new SupportingDocumentInfo
                    {
                        Id = name,
                        DisplayName = known.FileName != null ? known.DisplayName : StripDocxExtension(name),
                        Order = known.FileName != null ? known.Order : int.MaxValue
                    };
                })
                .OrderBy(d => d.Order)
                .ThenBy(d => d.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>
        /// Returns the absolute path to the default main template (Notice of Hearing)
        /// inside the configured folder, or null if the file does not exist.
        /// </summary>
        public static string GetMainTemplatePath(string folder)
        {
            if (!Directory.Exists(folder))
            {
                return null;
            }

            var path = Path.Combine(folder, DefaultMainTemplateFileName);
            return File.Exists(path) ? path : null;
        }

        /// <summary>
        /// Returns the absolute path to the supporting document with the given id,
        /// or null if no such file exists or the id is invalid (path-traversal
        /// protection: id must be a plain file name with no path separators).
        /// </summary>
        public static string GetSupportingFilePath(string folder, string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return null;
            }

            if (id.IndexOfAny(new[] { '/', '\\' }) >= 0)
            {
                return null;
            }

            if (!Directory.Exists(folder))
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
    }
}
