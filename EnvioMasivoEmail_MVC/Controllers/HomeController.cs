using System.Diagnostics;
using EnvioMasivoEmail_MVC.Models;
using Microsoft.AspNetCore.Mvc;
using Servicios.Interfaz;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace EnvioMasivoEmail_MVC.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IEmailService _emailService;

        public HomeController(ILogger<HomeController> logger, IEmailService emailService)
        {
            _logger = logger;
            _emailService = emailService;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> UploadEmails(IFormFile emailFile)
        {
            if (emailFile == null || emailFile.Length == 0)
                return Json(new { success = false, message = "Archivo no válido." });

            var filePath = Path.GetTempFileName();
            using (var stream = System.IO.File.Create(filePath))
            {
                await emailFile.CopyToAsync(stream);
            }
            var emails = _emailService.LeerCorreosDesdeTxt(filePath);
            System.IO.File.Delete(filePath);
            return Json(new { success = true, emails });
        }

        [HttpPost]
        public async Task<IActionResult> SendBulkEmails([FromBody] List<string> emails)
        {
            if (emails == null || emails.Count == 0)
                return Json(new { success = false, message = "No se recibieron correos" });

            // Cargar plantilla HTML base (si existe) para reutilizar contenido
            string bodyHtml = "<b>Mensaje de prueba</b>"; // fallback
            var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Plantillas", "miPlantilla.html");
            if (System.IO.File.Exists(templatePath))
            {
                bodyHtml = await System.IO.File.ReadAllTextAsync(templatePath);
            }

            var logs = await _emailService.SendBulkEmailsAsync(emails, "Asunto de prueba", bodyHtml, null);
            return Json(new { success = true, logs });
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
