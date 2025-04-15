using System;
using System.IO;
using System.Net;
using System.Net.Mail;
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
                message.Body = "Please find attached the weekly report in PDF format.";

                var attachment = new Attachment(pdfPath);
                message.Attachments.Add(attachment);

                // SMTP Client Ayarları
                string smtpHost = userInfo.SmtpServer;
                int smtpPort = userInfo.SmtpPort;
                bool enableSsl = userInfo.EnableSsl;

                // Gmail veya Outlook için sabit ayar kontrolü
                if (userInfo.SmtpServer.Contains("gmail"))
                {
                    smtpHost = "smtp.gmail.com";
                    smtpPort = 587;
                    enableSsl = true;
                }
                else if (userInfo.SmtpServer.Contains("office365") || userInfo.SmtpServer.Contains("outlook"))
                {
                    smtpHost = "smtp.office365.com";
                    smtpPort = 587;
                    enableSsl = true;
                }

                using (var smtpClient = new SmtpClient(smtpHost, smtpPort))
                {
                    smtpClient.Credentials = new NetworkCredential(userInfo.Email, userInfo.EmailPassword);
                    smtpClient.EnableSsl = enableSsl;

                    try
                    {
                        smtpClient.Send(message);
                    }
                    catch (SmtpException smtpEx)
                    {
                        throw new Exception($"SMTP Error: {smtpEx.StatusCode} - {smtpEx.Message}");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Unexpected Error: {ex.Message}");
                    }
                }
            }
        }
    }
}
