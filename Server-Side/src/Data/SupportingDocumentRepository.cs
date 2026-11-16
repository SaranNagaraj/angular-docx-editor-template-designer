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
        /// The file name of the main template. Configurable through the
        /// "MainTemplateFileName" setting so the sample works with any template pack;
        /// falls back to the built-in default when not configured.
        /// </summary>
        public static string GetMainTemplateFileName(string configuredFileName)
        {
            return string.IsNullOrWhiteSpace(configuredFileName)
                ? DefaultMainTemplateFileName
                : configuredFileName.Trim();
        }

        /// <summary>
        /// Lists every .docx in the folder: the main template first (if present), then
        /// all other files alphabetically. The folder is read on every call.
        ///
        /// When the main template file does not exist yet, a main-template entry is
        /// still returned - flagged with IsMainTemplate and Exists=false - so the UI
        /// knows the document it opened is the main template and can save it to disk
        /// (creating the file) instead of leaving the Save action permanently
        /// disabled.
        /// </summary>
        public static List<SupportingDocumentInfo> ListSupporting(
            string folder,
            string mainTemplateFileName = null,
            bool includeMissingMainTemplate = false)
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

            var mainName = GetMainTemplateFileName(mainTemplateFileName);

            var result = files
                .Select(name =>
                {
                    var isMain = string.Equals(name, mainName, StringComparison.OrdinalIgnoreCase);
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

            // The main template file may not exist yet (e.g. the folder holds only the
            // supporting documents). Still report it so the UI can treat the opened
            // document as the main template and offer to save it - creating the file.
            bool mainExists = result.Any(d => d.IsMainTemplate);
            if (!mainExists && includeMissingMainTemplate)
            {
                result.Insert(0, new SupportingDocumentInfo
                {
                    Id = mainName,
                    DisplayName = StripDocxExtension(mainName),
                    IsMainTemplate = true,
                    Order = 0,
                    Exists = false
                });
            }

            return result;
        }

        /// <summary>
        /// Path to the main template, or null if the file does not exist.
        /// </summary>
        public static string GetMainTemplatePath(string folder, string mainTemplateFileName = null)
        {
            if (!Directory.Exists(folder))
            {
                return null;
            }

            var path = Path.Combine(folder, GetMainTemplateFileName(mainTemplateFileName));
            return File.Exists(path) ? path : null;
        }

        /// <summary>
        /// Resolves a template id to the path it SHOULD have inside the folder, without
        /// requiring the file to exist. Used when saving a template for the first time
        /// (the main template may never have been written to disk).
        ///
        /// The same path-traversal and extension checks as ResolveTemplatePath apply:
        /// the id must be a plain .docx file name and the resolved path must stay
        /// inside the templates folder.
        /// </summary>
        public static string ResolveTemplateTargetPath(string folder, string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return null;
            }

            if (id.IndexOfAny(new[] { '/', '\\' }) >= 0
                || id.Contains("..", StringComparison.Ordinal)
                || !id.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var fullFolder = Path.GetFullPath(folder);
            var fullPath = Path.GetFullPath(Path.Combine(folder, id));

            if (!fullPath.StartsWith(fullFolder, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return fullPath;
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
