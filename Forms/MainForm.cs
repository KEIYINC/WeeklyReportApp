using System;
using System.Windows.Forms;
using System.IO;
using WeeklyReportApp.Models;
using WeeklyReportApp.Services;
using WeeklyReportApp.Utils;

namespace WeeklyReportApp.Forms
{
    public partial class MainForm : Form
    {
        private UserInfo _userInfo;
        private TextBox txtCompletedActivities;
        private TextBox txtOngoingActivities;
        private TextBox txtPlannedActivities;
        private Button btnGenerateReport;
        private Button btnSendEmail;
        private Button btnSettings;

        public MainForm()
        {
            InitializeComponent();
            _userInfo = ConfigManager.LoadUserInfo();
        }

        private void InitializeComponent()
        {
            this.txtCompletedActivities = new TextBox();
            this.txtOngoingActivities = new TextBox();
            this.txtPlannedActivities = new TextBox();
            this.btnGenerateReport = new Button();
            this.btnSendEmail = new Button();
            this.btnSettings = new Button();

            // Form setup
            this.Text = "Weekly Report - Activity Input";
            this.Size = new System.Drawing.Size(600, 400);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Labels
            var lblCompleted = new Label
            {
                Text = "Completed Activities:",
                Location = new System.Drawing.Point(20, 20),
                AutoSize = true
            };

            var lblOngoing = new Label
            {
                Text = "Ongoing Activities:",
                Location = new System.Drawing.Point(20, 120),
                AutoSize = true
            };

            var lblPlanned = new Label
            {
                Text = "Planned Activities:",
                Location = new System.Drawing.Point(20, 220),
                AutoSize = true
            };

            // TextBoxes
            this.txtCompletedActivities.Multiline = true;
            this.txtCompletedActivities.Location = new System.Drawing.Point(20, 40);
            this.txtCompletedActivities.Size = new System.Drawing.Size(550, 70);

            this.txtOngoingActivities.Multiline = true;
            this.txtOngoingActivities.Location = new System.Drawing.Point(20, 140);
            this.txtOngoingActivities.Size = new System.Drawing.Size(550, 70);

            this.txtPlannedActivities.Multiline = true;
            this.txtPlannedActivities.Location = new System.Drawing.Point(20, 240);
            this.txtPlannedActivities.Size = new System.Drawing.Size(550, 70);

            // Buttons
            this.btnGenerateReport.Text = "Generate Report";
            this.btnGenerateReport.Location = new System.Drawing.Point(20, 320);
            this.btnGenerateReport.Size = new System.Drawing.Size(120, 30);
            this.btnGenerateReport.Click += BtnGenerateReport_Click;

            this.btnSendEmail.Text = "Send Email";
            this.btnSendEmail.Location = new System.Drawing.Point(150, 320);
            this.btnSendEmail.Size = new System.Drawing.Size(120, 30);
            this.btnSendEmail.Click += BtnSendEmail_Click;

            this.btnSettings.Text = "Settings";
            this.btnSettings.Location = new System.Drawing.Point(280, 320);
            this.btnSettings.Size = new System.Drawing.Size(120, 30);
            this.btnSettings.Click += BtnSettings_Click;

            // Add controls to form
            this.Controls.AddRange(new Control[] {
                lblCompleted, txtCompletedActivities,
                lblOngoing, txtOngoingActivities,
                lblPlanned, txtPlannedActivities,
                btnGenerateReport, btnSendEmail, btnSettings
            });
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            var settingsForm = new SetupForm();
            settingsForm.ShowDialog();
            // Reload user info after settings are updated
            _userInfo = ConfigManager.LoadUserInfo();
        }

        private void BtnGenerateReport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCompletedActivities.Text) ||
                string.IsNullOrWhiteSpace(txtOngoingActivities.Text) ||
                string.IsNullOrWhiteSpace(txtPlannedActivities.Text))
            {
                MessageBox.Show("Please fill in all activity fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                WordService.UpdateTemplate(
                    _userInfo,
                    txtCompletedActivities.Text,
                    txtOngoingActivities.Text,
                    txtPlannedActivities.Text);

                MessageBox.Show("Report generated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating report: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSendEmail_Click(object sender, EventArgs e)
        {
            try
            {
                string reportPath = Path.Combine(
                    Path.GetDirectoryName(_userInfo.TemplatePath),
                    $"WeeklyReport_{DateTime.Now:yyyyMMdd}.docx");

                if (!File.Exists(reportPath))
                {
                    MessageBox.Show("Please generate the report first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                EmailService.SendReport(_userInfo, reportPath);
                MessageBox.Show("Email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error sending email: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
} 