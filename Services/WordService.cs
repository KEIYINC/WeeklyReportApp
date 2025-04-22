using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WeeklyReportApp.Models;

namespace WeeklyReportApp.Services
{
    public class WordService
    {
        public static string GenerateReport(UserInfo userInfo, string completedActivities, string ongoingActivities, string plannedActivities)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string newDocPath = Path.Combine(
                Path.GetDirectoryName(userInfo.TemplatePath),
                $"WeeklyReport_{timestamp}.docx");

            File.Copy(userInfo.TemplatePath, newDocPath, true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(newDocPath, true))
            {
                var body = doc.MainDocumentPart?.Document?.Body;

                if (body == null)
                    throw new InvalidOperationException("Word belgesi hatalı: Body kısmı bulunamadı.");

                ReplacePlaceholder(body, "{Ad}", userInfo.FullName);
                ReplacePlaceholder(body, "{Tarih}", GetCurrentWeekDateRange());
                ReplacePlaceholder(body, "{TamamlananFaaliyetler}", completedActivities);
                ReplacePlaceholder(body, "{DevamEdenFaaliyetler}", ongoingActivities);
                ReplacePlaceholder(body, "{PlanlananFaaliyetler}", plannedActivities);

                doc.MainDocumentPart.Document.Save();
            }

            return newDocPath;
        }

        public static string ConvertDocxToPdf(string docxPath)
        {
            string pdfPath = Path.ChangeExtension(docxPath, ".pdf");

            // Spire.Doc ile Office gerekmeden PDF dönüşümü
            var document = new Spire.Doc.Document();
            document.LoadFromFile(docxPath);
            document.SaveToFile(pdfPath, Spire.Doc.FileFormat.PDF);

            return pdfPath;
        }
        private static string GetCurrentWeekDateRange()
        {
            // Haftanın bugünü
            DateTime today = DateTime.Today;

            // Bugünün haftanın kaçıncı günü olduğunu bul (Pazartesi = 1)
            int diffToMonday = ((int)today.DayOfWeek + 6) % 7;

            // Haftanın Pazartesi ve Cuma gününü bul
            DateTime monday = today.AddDays(-diffToMonday);
            DateTime friday = monday.AddDays(4);

            // Türkçe ay ismini büyük harfle yaz
            string start = monday.ToString("dd MMMM yyyy", new System.Globalization.CultureInfo("tr-TR")).ToUpper();
            string end = friday.ToString("dd MMMM yyyy", new System.Globalization.CultureInfo("tr-TR")).ToUpper();

            return $"Dönemi: {start} – {end}";
        }


        private static void ReplacePlaceholder(Body body, string placeholder, string value)
        {
            foreach (var paragraph in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
            {
                string paragraphText = GetParagraphText(paragraph);
                if (paragraphText.Contains(placeholder))
                {
                    StringBuilder newText = new StringBuilder();
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        foreach (var text in run.Descendants<Text>())
                        {
                            if (text.Text.Contains(placeholder))
                            {
                                text.Text = text.Text.Replace(placeholder, value);
                            }
                            newText.Append(text.Text);
                        }
                    }

                    if (newText.ToString().Contains(placeholder))
                    {
                        paragraph.RemoveAllChildren();
                        paragraph.AppendChild(new Run(new Text(newText.ToString().Replace(placeholder, value))));
                    }
                }
            }
        }

        private static string GetParagraphText(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
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
