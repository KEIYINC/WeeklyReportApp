using System;
using System.Net;
using System.Net.Mail;
using System.IO;
using WeeklyReportApp.Models;

namespace WeeklyReportApp.Services
{
    public class EmailService
    {
        public static void SendReport(UserInfo userInfo, string pdfPath)
        {
            using (var message = new MailMessage())
            {
                message.From = new MailAddress(userInfo.Email);
                message.To.Add(userInfo.TargetEmail);
                message.Subject = $"Weekly Report - {DateTime.Now:dd.MM.yyyy}";
                message.Body = $"Please find attached the weekly report in PDF format.";

                var attachment = new Attachment(pdfPath);
                message.Attachments.Add(attachment);

                using (var smtpClient = new SmtpClient("smtp.gmail.com", 587))
                {
                smtpClient.Credentials = new NetworkCredential(userInfo.Email, userInfo.EmailPassword);
                    smtpClient.EnableSsl = true;
                    smtpClient.Send(message);
                }
            }
        }
    }
}
