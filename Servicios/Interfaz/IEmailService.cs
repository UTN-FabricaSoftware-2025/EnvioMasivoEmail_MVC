using Servicios.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Servicios.Interfaz
{
    public interface IEmailService
    {

        Task SendEmailWithAttachmentsAsync(
            string toEmail,
            string subject,
            string bodyHtml,
            List<AttachmentDto> attachments);
        //Task SendBulkEmailAsync(
        //IEnumerable<EmailRecipient> recipients,
        //string templatePath,
        //object modelCommon);

        List<DataUsersDto> LeerUsuariosDesdeExcel(string rutaExcel);
        List<string> LeerCorreosDesdeTxt(string rutaTxt);

    }
}
