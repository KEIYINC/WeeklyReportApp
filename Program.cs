using System;
using System.Windows.Forms;
using WeeklyReportApp.Forms;
using WeeklyReportApp.Services;
using WeeklyReportApp.Utils;

namespace WeeklyReportApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Load user info
            var userInfo = ConfigManager.LoadUserInfo();

            if (userInfo == null)
            {
                // Show setup form if no user info exists
                var setupForm = new SetupForm();
                if (setupForm.ShowDialog() == DialogResult.OK)
                {
                    userInfo = ConfigManager.LoadUserInfo();
                }
                else
                {
                    return;
                }
            }

            // Create and show main form
            var mainForm = new MainForm();
            
            // Start scheduler with main form
            Scheduler.Start(mainForm);
            
            Application.Run(mainForm);
        }
    }
} 