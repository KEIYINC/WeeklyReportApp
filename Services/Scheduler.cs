using System;
using System.Windows.Forms;
using WeeklyReportApp.Forms;

namespace WeeklyReportApp.Services
{
    public class Scheduler
    {
        private static System.Windows.Forms.Timer timer;
        private static MainForm mainForm;
        private static bool hasOpenedToday = false;

        public static void Start(MainForm form)
        {
            mainForm = form;
            timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000; // Her saniye kontrol et
            timer.Tick += Timer_Tick;
            timer.Start();

            // İlk kontrolü hemen yap
            CheckAndOpenForm();
        }

        private static void Timer_Tick(object sender, EventArgs e)
        {
            CheckAndOpenForm();
        }

        private static void CheckAndOpenForm()
        {
            DateTime now = DateTime.Now;

            // Debug bilgisi
            Console.WriteLine($"Current time: {now:dd.MM.yyyy hh:mm:ss tt}, Day: {now.DayOfWeek}");

            // Her gün başında hasOpenedToday'i sıfırla
            if (now.Hour == 0 && now.Minute == 0)
            {
                hasOpenedToday = false;
            }

            // Perşembe günü ve saat 10:58 PM ise ve bugün açılmadıysa
            if (now.DayOfWeek == DayOfWeek.Thursday &&
                now.Hour == 23 && // 10 PM
                now.Minute == 02 &&
                !hasOpenedToday)
            {
                Console.WriteLine("Opening form at " + now.ToString("dd.MM.yyyy hh:mm:ss tt"));

                mainForm.Invoke((MethodInvoker)delegate
                {
                    mainForm.Show();
                    mainForm.WindowState = FormWindowState.Normal;
                    mainForm.Activate();
                });

                hasOpenedToday = true;
            }
        }
    }
}