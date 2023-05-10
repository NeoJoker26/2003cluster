using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;
using System.Net;
using System.Net.Mail;

namespace _2003v5.Pages
{
    public class EmailcodeModel : PageModel
    {
        private string smtpServer = "smtp.gmail.com";
        private int smtpPort = 465;
        private string emailUsername = "greensonscreen1@gmail.com";
        private string emailPassword = "pptudnqvwynssrid";

        public void SendEmail(string toAddress, string ccAddress, string bccAddress, string fromAddress, string emailSubject, string emailBody)
        {
            using (SmtpClient client = new SmtpClient(smtpServer, smtpPort))
            {
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(emailUsername, emailPassword);
                client.EnableSsl = true;
                client.Timeout = 60000;

                using (MailMessage mailMessage = new MailMessage())
                {
                    mailMessage.To.Add(toAddress);
                    if (!string.IsNullOrEmpty(ccAddress))
                    {
                        mailMessage.CC.Add(ccAddress);
                    }

                    if (!string.IsNullOrEmpty(bccAddress))
                    {
                        mailMessage.Bcc.Add(bccAddress);
                    }

                    mailMessage.From = new MailAddress(fromAddress);
                    mailMessage.Subject = emailSubject;
                    mailMessage.Body = emailBody;
                    mailMessage.IsBodyHtml = true;

                    try
                    {
                        client.Send(mailMessage);
                    }
                    catch (Exception ex)
                    {
                       
                    }
                }
            }
        }
    }
}


