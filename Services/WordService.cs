using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WeeklyReportApp.Models;

namespace WeeklyReportApp.Services
{
    public class WordService
    {
        public static void UpdateTemplate(UserInfo userInfo, string completedActivities, string ongoingActivities, string plannedActivities)
        {
            string tempPath = Path.GetTempFileName();
            File.Copy(userInfo.TemplatePath, tempPath, true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(tempPath, true))
            {
                var body = doc.MainDocumentPart.Document.Body;

                // Replace placeholders
                ReplacePlaceholder(body, "{Ad}", userInfo.FullName);
                ReplacePlaceholder(body, "{Tarih}", DateTime.Now.ToString("dd.MM.yyyy"));
                ReplacePlaceholder(body, "{TamamlananFaaliyetler}", completedActivities);
                ReplacePlaceholder(body, "{DevamEdenFaaliyetler}", ongoingActivities);
                ReplacePlaceholder(body, "{PlanlananFaaliyetler}", plannedActivities);

                doc.MainDocumentPart.Document.Save();
            }

            // Save the updated document
            string outputPath = Path.Combine(
                Path.GetDirectoryName(userInfo.TemplatePath),
                $"WeeklyReport_{DateTime.Now:yyyyMMdd}.docx");

            File.Copy(tempPath, outputPath, true);
            File.Delete(tempPath);
        }

        private static void ReplacePlaceholder(Body body, string placeholder, string value)
        {
            if (body == null) return;

            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                string paragraphText = GetParagraphText(paragraph);
                if (paragraphText.Contains(placeholder))
                {
                    // Tüm metni birleştir ve değiştir
                    StringBuilder newText = new StringBuilder();
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        foreach (var text in run.Descendants<Text>())
                        {
                            if (text.Text.Contains(placeholder))
                            {
                                text.Text = text.Text.Replace(placeholder, value);
                                Console.WriteLine($"Replaced {placeholder} with {value}");
                            }
                            newText.Append(text.Text);
                        }
                    }

                    // Eğer hala placeholder varsa, tüm paragrafı yeniden oluştur
                    if (newText.ToString().Contains(placeholder))
                    {
                        paragraph.RemoveAllChildren();
                        paragraph.AppendChild(new Run(new Text(newText.ToString().Replace(placeholder, value))));
                    }
                }
            }
        }

        private static string GetParagraphText(Paragraph paragraph)
        {
            StringBuilder text = new StringBuilder();
            foreach (var run in paragraph.Descendants<Run>())
            {
                foreach (var t in run.Descendants<Text>())
                {
                    text.Append(t.Text);
                }
            }
            return text.ToString();
        }
    }
} 