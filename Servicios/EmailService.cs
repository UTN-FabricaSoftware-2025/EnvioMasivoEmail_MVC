using ClosedXML.Excel;
using MimeKit;
using Servicios.DTOs;
using Servicios.Interfaz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Extensions.Options;

using MimeKit.Utils;

namespace Servicios
{
    public class EmailService : IEmailService
    {
        private readonly EmailSettings _settings;
        public EmailService(EmailSettings settings)
        {
            _settings = settings;
        }


        public async Task SendEmailWithAttachmentsAsync(
            string toEmail,
            string subject,
            string bodyHtml,
            List<AttachmentDto> attachments)
        {
            var msg = new MimeMessage();
            msg.From.Add(new MailboxAddress(_settings.FromName, _settings.FromEmail));
            //msg.Bcc.Add(new MailboxAddress("mjara@mutualistaimbabura.com", "mjara@mutualistaimbabura.com"));
            //msg.Bcc.Add(new MailboxAddress("gtoledo@mutualistaimbabura.com", "gtoledo@mutualistaimbabura.com"));
            //msg.Headers.Add("Disposition-Notification-To", _settings.FromEmail); // leído
            //msg.Headers.Add("Return-Receipt-To", _settings.FromEmail);           // recibido
            //limpio espacio en correos
            toEmail = toEmail.Trim();
            if (string.IsNullOrWhiteSpace(toEmail))
            {
                throw new ArgumentException("El correo electrónico del destinatario no puede estar vacío.", nameof(toEmail));
            }
            try
            {
                msg.To.Add(new MailboxAddress(toEmail, toEmail));
            }
            catch (MimeKit.ParseException ex)
            {
                Console.WriteLine($"Correo inválido ignorado: {toEmail} - {ex.Message}");
                return; // Sale del método y no intenta enviar el correo
            }
            msg.Subject = subject;

            var builder = new BodyBuilder();

            // Adjunta la imagen como recurso embebido (opcional)
            var logoPath = Path.Combine(AppContext.BaseDirectory, "Templates", "logo.png");
            //if (File.Exists(logoPath))
            //{
            //    var image = builder.LinkedResources.Add(logoPath);
            //    image.ContentId = MimeUtils.GenerateMessageId();
            //    bodyHtml = bodyHtml.Replace("{{Logo}}", $"<img src=\"cid:{image.ContentId}\" />");
            //}
            // Agregar la imagen para {{img2}}
            var bannerPath = Path.Combine(AppContext.BaseDirectory, "Templates", "elecciones.jpg");
            if (File.Exists(bannerPath))
            {
                var bannerImage = builder.LinkedResources.Add(bannerPath);
                bannerImage.ContentId = MimeUtils.GenerateMessageId();
                bodyHtml = bodyHtml.Replace("{{img2}}", $"<img src=\"cid:{bannerImage.ContentId}\" style=\"max-width:100%;height:auto;\" />");
            }
            else
            {
                // Si no existe, puedes dejar el marcador vacío o poner una imagen por defecto
                bodyHtml = bodyHtml.Replace("{{img2}}", "");
            }
            builder.HtmlBody = bodyHtml;

            // Adjuntar los PDFs
            //if (attachments != null)
            //{
            //    foreach (var att in attachments)
            //    {
            //        builder.Attachments.Add(att.FileName, att.Content, new ContentType("application", "pdf"));
            //    }
            //}

            msg.Body = builder.ToMessageBody();

            using var client = new SmtpClient();
            await client.ConnectAsync(_settings.SmtpHost, _settings.SmtpPort, MailKit.Security.SecureSocketOptions.StartTls);
            await client.AuthenticateAsync(_settings.Username, _settings.Password);


            try
            {
                await client.SendAsync(msg);
                Console.WriteLine($"Correo enviado con exito a {toEmail}");

            }
            catch (Exception ex)
            {
                // Log o manejo de error aquí
                Console.WriteLine($"Error al enviar correo a {toEmail}: {ex.Message}");
            }
            await Task.Delay(1000);
            await client.DisconnectAsync(true);
        }






        // TEST SIN ERRORES
        //    public async Task SendEmailWithAttachmentsAsync(
        //string toEmail,
        //string subject,
        //string bodyHtml,
        //List<AttachmentDto> attachments)
        //    {
        //        var msg = new MimeMessage();
        //        msg.From.Add(new MailboxAddress(_settings.FromName, _settings.FromEmail));
        //        toEmail = toEmail.Trim();
        //        if (string.IsNullOrWhiteSpace(toEmail))
        //        {
        //            throw new ArgumentException("El correo electrónico del destinatario no puede estar vacío.", nameof(toEmail));
        //        }
        //        try
        //        {
        //            msg.To.Add(new MailboxAddress(toEmail, toEmail));
        //        }
        //        catch (MimeKit.ParseException ex)
        //        {
        //            Console.WriteLine($"Correo inválido ignorado: {toEmail} - {ex.Message}");
        //            return;
        //        }
        //        msg.Subject = subject;

        //        var builder = new BodyBuilder();
        //        var logoPath = Path.Combine(AppContext.BaseDirectory, "Templates", "logo.png");
        //        //if (File.Exists(logoPath))
        //        //{
        //        //    var image = builder.LinkedResources.Add(logoPath);
        //        //    image.ContentId = MimeUtils.GenerateMessageId();
        //        //    bodyHtml = bodyHtml.Replace("{{Logo}}", $"<img src=\"cid:{image.ContentId}\" />");
        //        //}

        //        // Agregar la imagen para {{img2}}
        //        var bannerPath = Path.Combine(AppContext.BaseDirectory, "Templates", "elecciones.jpg");
        //        if (File.Exists(bannerPath))
        //        {
        //            var bannerImage = builder.LinkedResources.Add(bannerPath);
        //            bannerImage.ContentId = MimeUtils.GenerateMessageId();
        //            bodyHtml = bodyHtml.Replace("{{img2}}", $"<img src=\"cid:{bannerImage.ContentId}\" style=\"max-width:100%;height:auto;\" />");
        //        }
        //        else
        //        {
        //            // Si no existe, puedes dejar el marcador vacío o poner una imagen por defecto
        //            bodyHtml = bodyHtml.Replace("{{img2}}", "");
        //        }
        //        builder.HtmlBody = bodyHtml;
        //        if (attachments != null)
        //        {
        //            foreach (var att in attachments)
        //            {
        //                builder.Attachments.Add(att.FileName, att.Content, new ContentType("application", "pdf"));
        //            }
        //        }
        //        msg.Body = builder.ToMessageBody();

        //        using var client = new SmtpClient();
        //        try
        //        {
        //            await client.ConnectAsync(_settings.SmtpHost, _settings.SmtpPort, MailKit.Security.SecureSocketOptions.StartTls);
        //            await client.AuthenticateAsync(_settings.Username, _settings.Password);
        //            await client.SendAsync(msg);
        //            Console.WriteLine($"Correo enviado con exito a {toEmail}");
        //        }
        //        catch (Exception ex)
        //        {
        //            // Aquí solo registramos el error y continuamos
        //            Console.WriteLine($"Error al enviar correo a {toEmail}: {ex.Message}");
        //        }
        //        finally
        //        {
        //            try
        //            {
        //                await Task.Delay(500);
        //                await client.DisconnectAsync(true);
        //            }
        //            catch { /* Ignorar cualquier error aquí también */ }
        //        }
        //    }


        #region EVIAR CORREO ANTERIOR
        //public async Task SendBulkEmailAsync(
        //    IEnumerable<EmailRecipient> recipients,
        //    string templatePath,
        //    object modelCommon)
        //{
        //    // Carga plantilla HTML
        //    var template = await File.ReadAllTextAsync(templatePath);

        //    using var client = new SmtpClient();
        //    try
        //    {
        //        if (string.IsNullOrWhiteSpace(_settings.SmtpHost) ||
        //            _settings.SmtpPort == 0 ||
        //            string.IsNullOrWhiteSpace(_settings.Username) ||
        //            string.IsNullOrWhiteSpace(_settings.Password))
        //        {
        //            throw new InvalidOperationException("La configuración SMTP es inválida o está incompleta.");
        //        }

        //        await client.ConnectAsync(_settings.SmtpHost, _settings.SmtpPort, MailKit.Security.SecureSocketOptions.StartTls);
        //        await client.AuthenticateAsync(_settings.Username, _settings.Password);
        //    }
        //    catch (Exception ex)
        //    {
        //        // Aquí puedes registrar el error o relanzarlo según tu necesidad
        //        throw new ApplicationException("Error al conectar o autenticar con el servidor SMTP.", ex);
        //    }

        //    foreach (var r in recipients)
        //    {
        //        string body = template
        //         .Replace("{{Name}}", r.Name)
        //         .Replace("{{Year}}", DateTime.Now.Year.ToString());

        //        var msg = new MimeMessage();
        //        msg.From.Add(new MailboxAddress(_settings.FromName, _settings.FromEmail));
        //        msg.To.Add(new MailboxAddress(r.Name, r.Email));
        //        msg.Subject = $"Tu informe anual {DateTime.Now.Year}";

        //        var builder = new BodyBuilder();

        //        // Adjunta la imagen como recurso embebido
        //        var logoPath = Path.Combine(AppContext.BaseDirectory, "Templates", "logo.png");
        //        var image = builder.LinkedResources.Add(logoPath);
        //        image.ContentId = MimeUtils.GenerateMessageId();

        //        // Referencia la imagen en el HTML usando el Content-Id
        //        builder.HtmlBody = body.Replace("{{Logo}}", $"<img src=\"cid:{image.ContentId}\" />");

        //        msg.Body = builder.ToMessageBody();
        //        try
        //        {
        //            await client.SendAsync(msg);
        //            Console.WriteLine($"✔️ Correo enviado a {r.Email}");

        //        }
        //        catch (Exception ex)
        //        {
        //            // Log o manejo de error aquí
        //        }
        //        await Task.Delay(200);
        //    }

        //    await client.DisconnectAsync(true);
        //}
        #endregion

        #region EXCEL NO SE USA
        //public List<DataUsersDto> LeerUsuariosDesdeExcel(string rutaExcel)
        //{
        //    var lista = new List<DataUsersDto>();

        //    using var workbook = new XLWorkbook(rutaExcel);
        //    var worksheet = workbook.Worksheet(1); // Primera hoja
        //    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Salta encabezado

        //    foreach (var row in rows)
        //    {
        //        var usuario = new DataUsersDto
        //        {
        //            Cedula = row.Cell(0).GetString(),
        //            Nombres = row.Cell(1).GetString(),
        //            Email = row.Cell(2).GetString()
        //        };
        //        lista.Add(usuario);
        //    }

        //    return lista;
        //}
        public List<DataUsersDto> LeerUsuariosDesdeExcel(string rutaExcel)
        {
            var lista = new List<DataUsersDto>();

            using var workbook = new XLWorkbook(rutaExcel);
            var worksheet = workbook.Worksheet(1); // Primera hoja
            var rows = worksheet.RangeUsed().RowsUsed().Skip(4); // Salta encabezado

            foreach (var row in rows)
            {
                // Validar que la fila tenga al menos 3 celdas no vacías
                if (row.CellsUsed().Count() >= 3)
                {
                    var usuario = new DataUsersDto
                    {
                        Cedula = row.Cell(1).GetString(),
                        Nombres = row.Cell(2).GetString(),
                        Email = row.Cell(3).GetString()
                    };
                    lista.Add(usuario);
                }
            }

            return lista;
        }
        #endregion
        public List<string> LeerCorreosDesdeTxt(string rutaTxt)
        {
            var listaCorreos = new List<string>();
            var regexCorreo = new Regex(@"^[^@\s]+@[^@\s]+\.[^@\s]+$", RegexOptions.Compiled);

            foreach (var linea in File.ReadLines(rutaTxt))
            {
                var correo = linea.Trim();
                if (!string.IsNullOrEmpty(correo) && regexCorreo.IsMatch(correo))
                {
                    listaCorreos.Add(correo);
                }
                else
                {
                    // Si quieres registrar los correos inválidos, puedes hacerlo aquí
                    Console.WriteLine($"Correo inválido ignorado: {correo}");
                }
            }

            return listaCorreos;
        }


    }
}
