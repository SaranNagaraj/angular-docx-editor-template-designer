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
    /// GET /api/templates              - list available supporting documents
    /// GET /api/templates/{id}         - load a supporting document as SFDT
    ///
    /// The supporting documents live in a folder on disk configured by
    /// "TemplatesFolder" in appsettings.json (defaults to "App_Data/SupportingDocuments"
    /// under the content root, which is auto-seeded on first run).
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
            var path = SupportingDocumentRepository.GetSupportingFilePath(folder, id);
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
    }
}
