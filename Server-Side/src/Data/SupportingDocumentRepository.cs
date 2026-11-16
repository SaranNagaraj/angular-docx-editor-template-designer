using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EJ2APIServices.Models;

namespace EJ2APIServices.Data
{
    /// <summary>
    /// Reads templates from the configured templates folder on disk. Nothing is
    /// hardcoded per file: every .docx in the folder is listed, so adding or removing
    /// a file is reflected in the UI on the next load.
    ///
    /// Folder (default, relative to Server-Side/src): wwwroot/Templates
    ///
    ///   Notice of Hearing.docx     - main template (optional; a built-in sample is used if absent)
    ///   Instructions.docx          - supporting document
    ///   Interpreter.docx           - supporting document
    ///   Terms and conditions.docx  - supporting document
    ///   (any other *.docx is also treated as a supporting document)
    /// </summary>
    public static class SupportingDocumentRepository
    {
        /// <summary>Default folder relative to the content root (Server-Side/src).</summary>
        public const string DefaultRelativeFolder = "wwwroot/Templates";

        /// <summary>Default main template file name (must live in the templates folder).</summary>
        public const string DefaultMainTemplateFileName = "Notice of Hearing.docx";

        /// <summary>Default XSL/XSLT stylesheet converted by POST /api/templates/xsl-to-word.</summary>
        public const string DefaultXslTemplateFileName = "InputXSL.txt";

        /// <summary>Absolute path to the configured folder.</summary>
        public static string GetFolder(string contentRootPath, string configuredFolder)
        {
            return string.IsNullOrWhiteSpace(configuredFolder)
                ? Path.Combine(contentRootPath, DefaultRelativeFolder)
                : (Path.IsPathRooted(configuredFolder)
                    ? configuredFolder
                    : Path.Combine(contentRootPath, configuredFolder));
        }

        /// <summary>
        /// Lists every .docx in the folder: the main template first (if present), then
        /// all other files alphabetically. The folder is read on every call.
        /// </summary>
        public static List<SupportingDocumentInfo> ListSupporting(string folder)
        {
            var items = new List<SupportingDocumentInfo>();
            if (!Directory.Exists(folder))
            {
                return items;
            }

            var files = Directory.EnumerateFiles(folder, "*.docx", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileName)
                // Word creates "~$name.docx" lock files while a document is open.
                .Where(n => !string.IsNullOrEmpty(n) && !n.StartsWith("~$", StringComparison.Ordinal))
                .ToList();

            return files
                .Select(name =>
                {
                    var isMain = string.Equals(name, DefaultMainTemplateFileName, StringComparison.OrdinalIgnoreCase);
                    return new SupportingDocumentInfo
                    {
                        Id = name,
                        DisplayName = StripDocxExtension(name),
                        IsMainTemplate = isMain,
                        Order = isMain ? 0 : 1
                    };
                })
                .OrderBy(d => d.Order)
                .ThenBy(d => d.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>Path to the main template, or null if the file does not exist.</summary>
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
        /// Resolves a template id (plain .docx file name) to its path inside the folder.
        /// Path-traversal protection: no directory separators, no "..", and the resolved
        /// path must stay inside the folder. Returns null if invalid or missing.
        ///
        /// When no extension is supplied only ".docx" is accepted, which keeps the DOCX
        /// endpoints locked down to Word documents. Callers can widen this explicitly
        /// (e.g. the XSL endpoint passes ".txt").
        /// </summary>
        public static string ResolveTemplatePath(string folder, string id, params string[] allowedExtensions)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return null;
            }

            if (id.IndexOfAny(new[] { '/', '\\' }) >= 0
                || id.Contains("..", StringComparison.Ordinal))
            {
                return null;
            }

            var extensions = (allowedExtensions == null || allowedExtensions.Length == 0)
                ? new[] { ".docx" }
                : allowedExtensions;

            bool extensionAllowed = extensions.Any(ext =>
                id.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
            if (!extensionAllowed)
            {
                return null;
            }

            var fullFolder = Path.GetFullPath(folder);
            var fullPath = Path.GetFullPath(Path.Combine(folder, id));

            if (!fullPath.StartsWith(fullFolder, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return File.Exists(fullPath) ? fullPath : null;
        }

        /// <summary>Path of a supporting document (same lookup as any template).</summary>
        public static string GetSupportingFilePath(string folder, string id)
        {
            return ResolveTemplatePath(folder, id);
        }

        private static string StripDocxExtension(string fileName)
        {
            return fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                ? fileName.Substring(0, fileName.Length - ".docx".Length)
                : fileName;
        }
    }
}
