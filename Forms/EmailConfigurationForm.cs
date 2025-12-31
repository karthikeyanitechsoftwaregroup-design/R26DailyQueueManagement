using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using R26_DailyQueueWinForm.Models;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;

namespace R26_DailyQueueWinForm.Forms
{
    public partial class EmailConfigurationForm : Form
    {
        private EmailConfiguration _emailConfig;
        private string _configFilePath = "appsettings.json";
        private IConfiguration _configuration;

        public EmailConfigurationForm()
        {
            // Load configuration
            _configuration = new Microsoft.Extensions.Configuration.ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false)
                .Build();

            InitializeComponent();
            LoadConfiguration();
        }

        private void InitializeComponent()
        {
            // Form Properties
            this.Text = "Email Configuration";
            this.Size = new Size(700, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Title Label
            Label lblTitle = new Label
            {
                Text = "Email Configuration Settings",
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(20, 20),
                AutoSize = true
            };

            // SMTP Settings Group
            GroupBox grpSmtp = new GroupBox
            {
                Text = "SMTP Settings",
                Location = new Point(20, 60),
                Size = new Size(640, 220),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            Label lblHost = new Label { Text = "Host:", Location = new Point(20, 30), AutoSize = true };
            txtHost = new TextBox { Location = new Point(150, 28), Size = new Size(450, 25), Font = new Font("Segoe UI", 10F) };

            Label lblPort = new Label { Text = "Port:", Location = new Point(20, 65), AutoSize = true };
            txtPort = new TextBox { Location = new Point(150, 63), Size = new Size(100, 25), Font = new Font("Segoe UI", 10F) };

            Label lblEnableSsl = new Label { Text = "Enable SSL:", Location = new Point(270, 65), AutoSize = true };
            chkEnableSsl = new CheckBox { Location = new Point(370, 65), Checked = true };

            Label lblUserName = new Label { Text = "Username:", Location = new Point(20, 100), AutoSize = true };
            txtUserName = new TextBox { Location = new Point(150, 98), Size = new Size(450, 25), Font = new Font("Segoe UI", 10F) };

            Label lblPassword = new Label { Text = "Password:", Location = new Point(20, 135), AutoSize = true };
            txtPassword = new TextBox { Location = new Point(150, 133), Size = new Size(450, 25), Font = new Font("Segoe UI", 10F), UseSystemPasswordChar = true };

            Label lblFromEmail = new Label { Text = "From Email:", Location = new Point(20, 170), AutoSize = true };
            txtFromEmail = new TextBox { Location = new Point(150, 168), Size = new Size(450, 25), Font = new Font("Segoe UI", 10F) };

            grpSmtp.Controls.AddRange(new Control[] { lblHost, txtHost, lblPort, txtPort, lblEnableSsl, chkEnableSsl, lblUserName, txtUserName, lblPassword, txtPassword, lblFromEmail, txtFromEmail });

            // Recipients Group
            GroupBox grpRecipients = new GroupBox
            {
                Text = "Email Recipients",
                Location = new Point(20, 290),
                Size = new Size(640, 180),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            Label lblToMails = new Label { Text = "To Emails:", Location = new Point(20, 30), AutoSize = true };
            txtToMails = new TextBox { Location = new Point(20, 55), Size = new Size(600, 45), Multiline = true, Font = new Font("Segoe UI", 9F) };
            Label lblToMailsInfo = new Label { Text = "(Comma-separated)", Location = new Point(20, 105), AutoSize = true, ForeColor = Color.Gray, Font = new Font("Segoe UI", 8F) };

            Label lblCcEmails = new Label { Text = "CC Emails:", Location = new Point(20, 125), AutoSize = true };
            txtCcEmails = new TextBox { Location = new Point(120, 123), Size = new Size(500, 25), Font = new Font("Segoe UI", 9F) };

            grpRecipients.Controls.AddRange(new Control[] { lblToMails, txtToMails, lblToMailsInfo, lblCcEmails, txtCcEmails });

            // Buttons
            btnSave = new Button
            {
                Text = "Save Configuration",
                Location = new Point(350, 490),
                Size = new Size(150, 40),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnSave.Click += BtnSave_Click;

            btnContinue = new Button
            {
                Text = "Continue to App",
                Location = new Point(510, 490),
                Size = new Size(150, 40),
                BackColor = Color.FromArgb(76, 175, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnContinue.Click += BtnContinue_Click;

            btnTest = new Button
            {
                Text = "Test Email",
                Location = new Point(200, 490),
                Size = new Size(140, 40),
                BackColor = Color.FromArgb(255, 140, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnTest.Click += BtnTest_Click;

            // Status Label
            lblStatus = new Label
            {
                Text = "",
                Location = new Point(20, 545),
                Size = new Size(640, 40),
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.Green,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // Add controls to form
            this.Controls.AddRange(new Control[] { lblTitle, grpSmtp, grpRecipients, btnTest, btnSave, btnContinue, lblStatus });
        }

        private TextBox txtHost, txtPort, txtUserName, txtPassword, txtFromEmail, txtToMails, txtCcEmails;
        private CheckBox chkEnableSsl;
        private Button btnSave, btnContinue, btnTest;
        private Label lblStatus;

        private void LoadConfiguration()
        {
            _emailConfig = GetDefaultConfiguration();

            txtHost.Text = _emailConfig.SmtpSettings.Host;
            txtPort.Text = _emailConfig.SmtpSettings.Port.ToString();
            chkEnableSsl.Checked = _emailConfig.SmtpSettings.EnableSsl;
            txtUserName.Text = _emailConfig.SmtpSettings.UserName;
            txtPassword.Text = _emailConfig.SmtpSettings.Password;
            txtFromEmail.Text = _emailConfig.SmtpSettings.FromEmail;
            txtToMails.Text = _emailConfig.ToMails;
            txtCcEmails.Text = _emailConfig.ErrorNotifications.CcEmails;
        }

        private EmailConfiguration GetDefaultConfiguration()
        {
            return new EmailConfiguration
            {
                SmtpSettings = new SmtpSettings
                {
                    Host = _configuration["SmtpSettings:Host"],
                    Port = int.Parse(_configuration["SmtpSettings:Port"]),
                    EnableSsl = bool.Parse(_configuration["SmtpSettings:EnableSsl"]),
                    UserName = _configuration["SmtpSettings:UserName"],
                    Password = _configuration["SmtpSettings:Password"],
                    FromEmail = _configuration["SmtpSettings:FromEmail"],
                    FromName = _configuration["SmtpSettings:FromName"]
                },
                ToMails = _configuration["EmailRecipients:ToMails"],
                ErrorNotifications = new ErrorNotifications
                {
                    CcEmails = _configuration["EmailRecipients:CcEmails"]
                }
            };
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // Update configuration object
                _emailConfig.SmtpSettings.Host = txtHost.Text.Trim();
                _emailConfig.SmtpSettings.Port = int.Parse(txtPort.Text);
                _emailConfig.SmtpSettings.EnableSsl = chkEnableSsl.Checked;
                _emailConfig.SmtpSettings.UserName = txtUserName.Text.Trim();
                _emailConfig.SmtpSettings.Password = txtPassword.Text;
                _emailConfig.SmtpSettings.FromEmail = txtFromEmail.Text.Trim();
                _emailConfig.ToMails = txtToMails.Text.Trim();
                _emailConfig.ErrorNotifications.CcEmails = txtCcEmails.Text.Trim();

                // Save to file
                string json = JsonSerializer.Serialize(_emailConfig, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(_configFilePath, json);

                lblStatus.Text = "✓ Configuration saved successfully!";
                lblStatus.ForeColor = Color.Green;

                MessageBox.Show("Email configuration saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                lblStatus.Text = "✗ Error saving configuration: " + ex.Message;
                lblStatus.ForeColor = Color.Red;
                MessageBox.Show($"Error saving configuration: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnTest_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Test email functionality would be implemented here.\nThis would send a test email using the configured settings.",
                "Test Email", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnContinue_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        public EmailConfiguration GetEmailConfiguration()
        {
            return _emailConfig;
        }
    }
}