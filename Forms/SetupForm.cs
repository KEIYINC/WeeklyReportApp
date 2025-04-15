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
        private TextBox txtTemplatePath;
        private TextBox txtEmailPassword;
        private TextBox txtSmtpServer;
        private TextBox txtSmtpPort;
        private CheckBox chkEnableSsl;
        private Label lblFullName;
        private Label lblEmail;
        private Label lblTargetEmail;
        private Label lblTemplate;
        private Label lblEmailPassword;
        private Label lblSmtpServer;
        private Label lblSmtpPort;
        private Label lblEnableSsl;
        private Button btnSelectTemplate;
        private Button btnSave;
        private UserInfo _currentUserInfo;
        private ComboBox cmbSmtpProvider;
        private Label lblSmtpProvider;
        


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
                txtEmailPassword.Text = _currentUserInfo.EmailPassword;
                txtSmtpServer.Text = _currentUserInfo.SmtpServer;
                txtSmtpPort.Text = _currentUserInfo.SmtpPort.ToString();
                chkEnableSsl.Checked = _currentUserInfo.EnableSsl;

                // SMTP provider seçimini otomatik olarak geri yükle
                if (!string.IsNullOrWhiteSpace(_currentUserInfo.SmtpServer))
                {
                    if (_currentUserInfo.SmtpServer.Contains("gmail"))
                        cmbSmtpProvider.SelectedItem = "Gmail";
                    else if (_currentUserInfo.SmtpServer.Contains("office365"))
                        cmbSmtpProvider.SelectedItem = "Outlook";
                }
            }
        }


        private void InitializeComponent()
{
    // Initialize Controls
    txtFullName = new TextBox();
    txtEmail = new TextBox();
    txtTargetEmail = new TextBox();
    txtTemplatePath = new TextBox();
    txtEmailPassword = new TextBox();
    txtSmtpServer = new TextBox();
    txtSmtpPort = new TextBox();
    chkEnableSsl = new CheckBox();
    cmbSmtpProvider = new ComboBox();

    lblFullName = new Label();
    lblEmail = new Label();
    lblTargetEmail = new Label();
    lblTemplate = new Label();
    lblEmailPassword = new Label();
    lblSmtpServer = new Label();
    lblSmtpPort = new Label();
    lblEnableSsl = new Label();
    lblSmtpProvider = new Label();

    btnSelectTemplate = new Button();
    btnSave = new Button();

    // Form setup
    this.Text = "Weekly Report App - Settings";
    this.Size = new System.Drawing.Size(550, 480);
    this.StartPosition = FormStartPosition.CenterScreen;

    int labelX = 20;
    int inputX = 150;
    int width = 350;
    int lineHeight = 30;

    // Labels
    lblFullName.Text = "Full Name:";
    lblFullName.Location = new System.Drawing.Point(labelX, 20);
    lblFullName.AutoSize = true;

    lblEmail.Text = "Email:";
    lblEmail.Location = new System.Drawing.Point(labelX, 20 + lineHeight * 1);
    lblEmail.AutoSize = true;

    lblTargetEmail.Text = "Target Email:";
    lblTargetEmail.Location = new System.Drawing.Point(labelX, 20 + lineHeight * 2);
    lblTargetEmail.AutoSize = true;

    lblTemplate.Text = "Word Template:";
    lblTemplate.Location = new System.Drawing.Point(labelX, 20 + lineHeight * 3);
    lblTemplate.AutoSize = true;

    lblEmailPassword.Text = "Email Password:";
    lblEmailPassword.Location = new System.Drawing.Point(labelX, 20 + lineHeight * 4);
    lblEmailPassword.AutoSize = true;

    lblSmtpServer.Text = "SMTP Server:";
    lblSmtpServer.Location = new System.Drawing.Point(labelX, 20 + lineHeight * 5);
    lblSmtpServer.AutoSize = true;

    lblSmtpPort.Text = "SMTP Port:";
    lblSmtpPort.Location = new System.Drawing.Point(labelX, 20 + lineHeight * 6);
    lblSmtpPort.AutoSize = true;

    lblEnableSsl.Text = "Enable SSL:";
    lblEnableSsl.Location = new System.Drawing.Point(labelX, 20 + lineHeight * 7);
    lblEnableSsl.AutoSize = true;

    lblSmtpProvider.Text = "Email Provider:";
    lblSmtpProvider.Location = new System.Drawing.Point(labelX, 20 + lineHeight * 8);
    lblSmtpProvider.AutoSize = true;

    // Inputs
    txtFullName.Location = new System.Drawing.Point(inputX, 20);
    txtFullName.Size = new System.Drawing.Size(width, 20);

    txtEmail.Location = new System.Drawing.Point(inputX, 20 + lineHeight * 1);
    txtEmail.Size = new System.Drawing.Size(width, 20);

    txtTargetEmail.Location = new System.Drawing.Point(inputX, 20 + lineHeight * 2);
    txtTargetEmail.Size = new System.Drawing.Size(width, 20);

    txtTemplatePath.Location = new System.Drawing.Point(inputX, 20 + lineHeight * 3);
    txtTemplatePath.Size = new System.Drawing.Size(width - 100, 20);
    txtTemplatePath.ReadOnly = true;

    txtEmailPassword.Location = new System.Drawing.Point(inputX, 20 + lineHeight * 4);
    txtEmailPassword.Size = new System.Drawing.Size(width, 20);
    txtEmailPassword.PasswordChar = '*';

    txtSmtpServer.Location = new System.Drawing.Point(inputX, 20 + lineHeight * 5);
    txtSmtpServer.Size = new System.Drawing.Size(width, 20);

    txtSmtpPort.Location = new System.Drawing.Point(inputX, 20 + lineHeight * 6);
    txtSmtpPort.Size = new System.Drawing.Size(width, 20);

    chkEnableSsl.Location = new System.Drawing.Point(inputX, 20 + lineHeight * 7);

    cmbSmtpProvider.Location = new System.Drawing.Point(inputX, 20 + lineHeight * 8);
    cmbSmtpProvider.Size = new System.Drawing.Size(200, 21);
    cmbSmtpProvider.DropDownStyle = ComboBoxStyle.DropDownList;
    cmbSmtpProvider.Items.AddRange(new object[] { "Gmail", "Outlook" });
    cmbSmtpProvider.SelectedIndexChanged += CmbSmtpProvider_SelectedIndexChanged;

    // Buttons
    btnSelectTemplate.Text = "Browse";
    btnSelectTemplate.Location = new System.Drawing.Point(inputX + width - 90, 20 + lineHeight * 3);
    btnSelectTemplate.Size = new System.Drawing.Size(80, 23);
    btnSelectTemplate.Click += BtnSelectTemplate_Click;

    btnSave.Text = "Save";
    btnSave.Location = new System.Drawing.Point(inputX, 20 + lineHeight * 10);
    btnSave.Size = new System.Drawing.Size(100, 30);
    btnSave.Click += BtnSave_Click;

    // Add controls to form
    this.Controls.AddRange(new Control[] {
        lblFullName, txtFullName,
        lblEmail, txtEmail,
        lblTargetEmail, txtTargetEmail,
        lblTemplate, txtTemplatePath, btnSelectTemplate,
        lblEmailPassword, txtEmailPassword,
        lblSmtpServer, txtSmtpServer,
        lblSmtpPort, txtSmtpPort,
        lblEnableSsl, chkEnableSsl,
        lblSmtpProvider, cmbSmtpProvider,
        btnSave
    });
}


        private void CmbSmtpProvider_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selected = cmbSmtpProvider.SelectedItem.ToString();
            if (selected == "Gmail")
            {
                txtSmtpServer.Text = "smtp.gmail.com";
                txtSmtpPort.Text = "587";
                chkEnableSsl.Checked = true;
            }
            else if (selected == "Outlook")
            {
                txtSmtpServer.Text = "smtp.office365.com";
                txtSmtpPort.Text = "587";
                chkEnableSsl.Checked = true;
            }


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
                string.IsNullOrWhiteSpace(txtTemplatePath.Text) ||
                string.IsNullOrWhiteSpace(txtEmailPassword.Text) ||
                string.IsNullOrWhiteSpace(txtSmtpServer.Text) ||
                string.IsNullOrWhiteSpace(txtSmtpPort.Text))
            {
                MessageBox.Show("Please fill in all fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!int.TryParse(txtSmtpPort.Text, out int smtpPort))
            {
                MessageBox.Show("SMTP Port must be a valid number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var userInfo = new UserInfo
            {
                FullName = txtFullName.Text,
                Email = txtEmail.Text,
                TargetEmail = txtTargetEmail.Text,
                TemplatePath = txtTemplatePath.Text,
                LastReportDate = _currentUserInfo?.LastReportDate ?? DateTime.Now,
                EmailPassword = txtEmailPassword.Text,
                SmtpServer = txtSmtpServer.Text,
                SmtpPort = smtpPort,
                EnableSsl = chkEnableSsl.Checked
            };

            ConfigManager.SaveUserInfo(userInfo);
            MessageBox.Show("Settings saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
    }
}
