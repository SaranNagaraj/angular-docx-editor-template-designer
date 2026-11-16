using System;
using System.IO;
using EJ2APIServices.Data;
using EJ2APIServices.Models;
using EJ2APIServices.Services;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

namespace EJ2APIServices.Controllers
{
    /// <summary>
    /// Supporting documents (templates) API:
    /// GET  /api/templates              - list available templates
    /// GET  /api/templates/{id}         - load a template as SFDT
    /// POST /api/templates/{id}/save    - save the edited template back to disk
    ///
    /// The templates live in a folder on disk configured by "TemplatesFolder" in
    /// appsettings.json (defaults to "wwwroot/Templates" under the content root).
    /// Saving writes the editor's SFDT payload back as a DOCX file in that same
    /// folder, overwriting the file the user opened.
    /// </summary>
    [Route("api/[controller]")]
    public class TemplatesController : Controller
    {
        private readonly IWebHostEnvironment _env;
        private readonly IConfiguration _config;

        public TemplatesController(IWebHostEnvironment env, IConfiguration config)
        {
            _env = env;
            _config = config;
        }

        private string GetFolder()
        {
            return SupportingDocumentRepository.GetFolder(
                _env.ContentRootPath,
                _config["TemplatesFolder"]);
        }

        [HttpGet]
        public IActionResult List()
        {
            var folder = GetFolder();
            var items = SupportingDocumentRepository.ListSupporting(folder);
            return Ok(items);
        }

        [HttpGet("{id}")]
        public IActionResult Get(string id)
        {
            var folder = GetFolder();
            var path = SupportingDocumentRepository.ResolveTemplatePath(folder, id);
            if (path == null)
            {
                return NotFound(new { message = $"Template '{id}' was not found." });
            }

            try
            {
                var sfdt = DocumentMergeService.LoadDocxAsSfdt(path);
                return Ok(new SupportingDocumentSfdtResponse { Sfdt = sfdt });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Could not load template '{id}': {ex.Message}" });
            }
        }

        /// <summary>
        /// Saves the template currently open in the editor back to the templates folder.
        ///
        /// The client sends the editor document serialized as SFDT. The server converts
        /// it back to DOCX and overwrites the file it came from, so the template pack
        /// in wwwroot/Templates always reflects the latest edits.
        ///
        /// Only files that already exist in the templates folder can be overwritten -
        /// this endpoint does not create new templates.
        /// </summary>
        [HttpPost("{id}/save")]
        public IActionResult Save(string id, [FromBody] SaveTemplateRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Sfdt))
            {
                return BadRequest(new { message = "The document payload is empty." });
            }

            var folder = GetFolder();
            var path = SupportingDocumentRepository.ResolveTemplatePath(folder, id);
            if (path == null)
            {
                return NotFound(new { message = $"Template '{id}' was not found." });
            }

            try
            {
                DocumentMergeService.SaveSfdtAsDocx(request.Sfdt, path);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Could not save template '{id}': {ex.Message}" });
            }

            var displayName = Path.GetFileNameWithoutExtension(path);
            return Ok(new SaveTemplateResponse
            {
                Id = Path.GetFileName(path),
                DisplayName = displayName,
                Path = path
            });
        }
    }
}
