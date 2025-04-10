using System;
using System.Net;
using System.Net.Mail;
using System.IO;
using WeeklyReportApp.Models;

namespace WeeklyReportApp.Services
{
    public class EmailService
    {
        public static void SendReport(UserInfo userInfo, string reportPath)
        {
            using (var message = new MailMessage())
            {
                message.From = new MailAddress(userInfo.Email);
                message.To.Add(userInfo.TargetEmail);
                message.Subject = $"Weekly Report - {DateTime.Now:dd.MM.yyyy}";
                message.Body = $"Please find attached the weekly report for {DateTime.Now:dd.MM.yyyy}.";

                // Attach the report
                var attachment = new Attachment(reportPath);
                message.Attachments.Add(attachment);

                using (var smtpClient = new SmtpClient("smtp.gmail.com", 587))
                {
                    // Replace "YOUR_APP_PASSWORD" with the 16-digit app password you generated
                    smtpClient.Credentials = new NetworkCredential(userInfo.Email, "egxe zkby ebhh lzfl");
                    smtpClient.EnableSsl = true;
                    smtpClient.Send(message);
                }
            }
        }
    }
} 