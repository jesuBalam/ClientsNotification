using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace ClientsNotification
{
    class EmailUtils
    {   
        public static void EmailSender(string toEmail, string attachmentPath = "")
        {
            try
            {
                var smtpClient = new SmtpClient("smtp.gmail.com")
                {
                    Port = 587,
                    Credentials = new NetworkCredential(ConfigurationManager.AppSettings["EmailSender"].ToString(), ConfigurationManager.AppSettings["PasswordEmail"].ToString()),
                    EnableSsl = true,
                };

                var mailMessage = new MailMessage
                {
                    From = new MailAddress(ConfigurationManager.AppSettings["EmailSender"].ToString()),
                    Subject = "Recordatorio",
                    Body = "<h1>Recordatorio</h1>" +
                    "</br>" +
                    "<h2> Recuerda subir la información de tu estación. </h2>",
                    IsBodyHtml = true,
                };
                mailMessage.To.Add(toEmail);

                if(!string.IsNullOrEmpty(attachmentPath))
                {
                    var attachment = new Attachment(attachmentPath);
                    mailMessage.Attachments.Add(attachment);
                    mailMessage.Body = "<h1>Reporte final</h1>" +
                    "</br>" +
                    "<h2> En este correo se adjunta el reporte final. </h2>";
                }

                smtpClient.Send(mailMessage);
                Console.WriteLine("Notifications sent");
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }
           
        }

        public static void EmailSenderByPackage(List<string> emails, string attachmentPath = "")
        {
            try
            {
                var smtpClient = new SmtpClient("smtp.gmail.com")
                {
                    Port = 587,
                    Credentials = new NetworkCredential(ConfigurationManager.AppSettings["EmailSender"].ToString(), ConfigurationManager.AppSettings["PasswordEmail"].ToString()),
                    EnableSsl = true,
                };

                var mailMessage = new MailMessage
                {
                    From = new MailAddress(ConfigurationManager.AppSettings["EmailSender"].ToString()),
                    Subject = "Recordatorio",
                    Body = "<h1>Recordatorio</h1>" +
                    "</br>" +
                    "<h2> Recuerda subir la información de tu estación. </h2>",
                    IsBodyHtml = true,
                };

                foreach(string mail in emails)
                {
                    mailMessage.To.Add(mail);
                }

                if (!string.IsNullOrEmpty(attachmentPath))
                {
                    var attachment = new Attachment(attachmentPath);
                    mailMessage.Attachments.Add(attachment);
                    mailMessage.Body = "<h1>Reporte final</h1>" +
                    "</br>" +
                    "<h2> En este correo se adjunta el reporte final. </h2>";
                }

                smtpClient.Send(mailMessage);
                Console.WriteLine("Notifications sent");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
    }
}
