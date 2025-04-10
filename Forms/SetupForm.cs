using System;
using System.Windows.Forms;
using WeeklyReportApp.Models;
using WeeklyReportApp.Utils;

namespace WeeklyReportApp.Forms
{
    public partial class SetupForm : Form
    {
        private TextBox txtFullName;
        private TextBox txtEmail;
        private TextBox txtTargetEmail;
        private Button btnSelectTemplate;
        private Button btnSave;
        private Label lblFullName;
        private Label lblEmail;
        private Label lblTargetEmail;
        private Label lblTemplate;
        private TextBox txtTemplatePath;
        private UserInfo _currentUserInfo;

        public SetupForm()
        {
            InitializeComponent();
            LoadCurrentSettings();
        }

        private void LoadCurrentSettings()
        {
            _currentUserInfo = ConfigManager.LoadUserInfo();
            if (_currentUserInfo != null)
            {
                txtFullName.Text = _currentUserInfo.FullName;
                txtEmail.Text = _currentUserInfo.Email;
                txtTargetEmail.Text = _currentUserInfo.TargetEmail;
                txtTemplatePath.Text = _currentUserInfo.TemplatePath;
            }
        }

        private void InitializeComponent()
        {
            this.txtFullName = new TextBox();
            this.txtEmail = new TextBox();
            this.txtTargetEmail = new TextBox();
            this.btnSelectTemplate = new Button();
            this.btnSave = new Button();
            this.lblFullName = new Label();
            this.lblEmail = new Label();
            this.lblTargetEmail = new Label();
            this.lblTemplate = new Label();
            this.txtTemplatePath = new TextBox();

            // Form setup
            this.Text = "Weekly Report App - Settings";
            this.Size = new System.Drawing.Size(500, 300);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Labels
            this.lblFullName.Text = "Full Name:";
            this.lblFullName.Location = new System.Drawing.Point(20, 20);
            this.lblFullName.AutoSize = true;

            this.lblEmail.Text = "Email:";
            this.lblEmail.Location = new System.Drawing.Point(20, 60);
            this.lblEmail.AutoSize = true;

            this.lblTargetEmail.Text = "Target Email:";
            this.lblTargetEmail.Location = new System.Drawing.Point(20, 100);
            this.lblTargetEmail.AutoSize = true;

            this.lblTemplate.Text = "Word Template:";
            this.lblTemplate.Location = new System.Drawing.Point(20, 140);
            this.lblTemplate.AutoSize = true;

            // TextBoxes
            this.txtFullName.Location = new System.Drawing.Point(120, 20);
            this.txtFullName.Size = new System.Drawing.Size(350, 20);

            this.txtEmail.Location = new System.Drawing.Point(120, 60);
            this.txtEmail.Size = new System.Drawing.Size(350, 20);

            this.txtTargetEmail.Location = new System.Drawing.Point(120, 100);
            this.txtTargetEmail.Size = new System.Drawing.Size(350, 20);

            this.txtTemplatePath.Location = new System.Drawing.Point(120, 140);
            this.txtTemplatePath.Size = new System.Drawing.Size(250, 20);
            this.txtTemplatePath.ReadOnly = true;

            // Buttons
            this.btnSelectTemplate.Text = "Browse";
            this.btnSelectTemplate.Location = new System.Drawing.Point(380, 140);
            this.btnSelectTemplate.Size = new System.Drawing.Size(90, 23);
            this.btnSelectTemplate.Click += BtnSelectTemplate_Click;

            this.btnSave.Text = "Save";
            this.btnSave.Location = new System.Drawing.Point(200, 180);
            this.btnSave.Size = new System.Drawing.Size(100, 30);
            this.btnSave.Click += BtnSave_Click;

            // Add controls to form
            this.Controls.AddRange(new Control[] {
                this.lblFullName, this.txtFullName,
                this.lblEmail, this.txtEmail,
                this.lblTargetEmail, this.txtTargetEmail,
                this.lblTemplate, this.txtTemplatePath,
                this.btnSelectTemplate, this.btnSave
            });
        }

        private void BtnSelectTemplate_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Word Documents|*.docx";
                openFileDialog.Title = "Select Word Template";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtTemplatePath.Text = openFileDialog.FileName;
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFullName.Text) ||
                string.IsNullOrWhiteSpace(txtEmail.Text) ||
                string.IsNullOrWhiteSpace(txtTargetEmail.Text) ||
                string.IsNullOrWhiteSpace(txtTemplatePath.Text))
            {
                MessageBox.Show("Please fill in all fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var userInfo = new UserInfo
            {
                FullName = txtFullName.Text,
                Email = txtEmail.Text,
                TargetEmail = txtTargetEmail.Text,
                TemplatePath = txtTemplatePath.Text,
                LastReportDate = _currentUserInfo?.LastReportDate ?? DateTime.Now
            };

            ConfigManager.SaveUserInfo(userInfo);
            MessageBox.Show("Settings saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
    }
} 