using R26_DailyQueueWinForm.Models;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace R26_DailyQueueWinForm.Forms
{
    // ============================================================
    // MAIN FORM CLASS
    // ============================================================
    public partial class SamsDeliveryReportForm : Form
    {
        private WebBrowser webBrowser;
        private MenuStrip mainMenuStrip;
        private ToolStripMenuItem menuItemDeliveryReport;
        private ToolStripMenuItem menuItemR26Queue;
        private ToolStripMenuItem menuItemExit;
        private Button btnSendEmail;
        private Button btnRefresh;

        private DateTime _targetDate;
        private EmailConfiguration _emailConfig;
        private string _connectionString;
        private string _storedProcedureName;
        private string _cachedHtmlContent;

        // ✅ FIXED: Declare _r26Form field (was missing, causing CS0103 errors)
        private R26QueueForm _r26Form;

        // ✅ FIXED: Added _r26QueueForm field for menu navigation
        private R26QueueForm _r26QueueForm;

        private SamsDeliveryDatabaseService _databaseService;
        private SamsDeliveryEmailService _emailService;

        public SamsDeliveryReportForm(
            DateTime targetDate,
            EmailConfiguration emailConfig,
            string connectionString,
            string storedProcedureName = "SDgetSAMSReportStatus",
            R26QueueForm r26Form = null)
        {
            _targetDate = targetDate;
            _emailConfig = emailConfig;
            _connectionString = connectionString;
            _storedProcedureName = storedProcedureName;
            _r26Form = r26Form; // ✅ FIXED: Now this field exists
            _r26QueueForm = r26Form; // ✅ FIXED: Initialize both references

            // Initialize services
            _databaseService = new SamsDeliveryDatabaseService(_connectionString, _storedProcedureName);
            _emailService = new SamsDeliveryEmailService(_emailConfig);

            InitializeComponent();
            LoadReportFromDatabase();

            // ✅ FIXED: Add FormClosing event handler
            this.FormClosing += SamsDeliveryReportForm_FormClosing;
        }

        private void InitializeComponent()
        {
            this.Text = $"Sam's Delivery Report Status - {_targetDate:yyyy-MM-dd}";
            this.Size = new Size(1400, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;

            // Menu Strip
            mainMenuStrip = new MenuStrip
            {
                Dock = DockStyle.None,
                Font = new Font("Segoe UI", 10F),
                GripStyle = ToolStripGripStyle.Hidden,
                RenderMode = ToolStripRenderMode.Professional,
                Padding = new Padding(10, 5, 0, 5)
            };

            menuItemDeliveryReport = new ToolStripMenuItem("Sam's Delivery Report Status")
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            menuItemR26Queue = new ToolStripMenuItem("R26 Daily Queue Management");
            menuItemR26Queue.Click += MenuItemR26Queue_Click;

            menuItemExit = new ToolStripMenuItem("Exit");
            menuItemExit.Click += (s, e) => Application.Exit();

            mainMenuStrip.Items.Add(menuItemDeliveryReport);
            mainMenuStrip.Items.Add(menuItemR26Queue);
            mainMenuStrip.Items.Add(menuItemExit);
            mainMenuStrip.Location = new Point(0, 5);
            this.MainMenuStrip = mainMenuStrip;

            // Title Label
            Label lblTitle = new Label
            {
                Text = $"Sam's Delivery Report Status - {_targetDate:MMMM dd, yyyy}",
                Font = new Font("Segoe UI", 14F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(20, 40),
                AutoSize = true
            };

            // Refresh Button
            btnRefresh = new Button
            {
                Text = "Refresh Report",
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(140, 40),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnRefresh.FlatAppearance.BorderSize = 0;
            btnRefresh.Click += BtnRefresh_Click;

            // Send Email Button
            btnSendEmail = new Button
            {
                Text = "Send Email",
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(150, 40),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnSendEmail.FlatAppearance.BorderSize = 0;
            btnSendEmail.Click += BtnSendEmail_Click;

            // Web Browser
            webBrowser = new WebBrowser
            {
                Location = new Point(0, 80),
                Size = new Size(this.ClientSize.Width, this.ClientSize.Height - 80),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                ScriptErrorsSuppressed = true
            };

            this.Controls.Add(mainMenuStrip);
            this.Controls.Add(lblTitle);
            this.Controls.Add(btnRefresh);
            this.Controls.Add(btnSendEmail);
            this.Controls.Add(webBrowser);

            this.Resize += (s, e) =>
            {
                CenterMenuStrip();
                PositionButtons();
            };
            this.Shown += (s, e) =>
            {
                CenterMenuStrip();
                PositionButtons();
            };
        }

        private void PositionButtons()
        {
            if (btnSendEmail != null && btnRefresh != null)
            {
                btnSendEmail.Location = new Point(
                    this.ClientSize.Width - btnSendEmail.Width - 20,
                    10
                );
                btnRefresh.Location = new Point(
                    btnSendEmail.Location.X - btnRefresh.Width - 10,
                    10
                );
            }
        }

        private void CenterMenuStrip()
        {
            if (mainMenuStrip != null)
            {
                int x = (this.ClientSize.Width - mainMenuStrip.Width) / 2;
                mainMenuStrip.Location = new Point(x, 5);
            }
        }

        private async void LoadReportFromDatabase()
        {
            try
            {
                // Show loading message
                webBrowser.DocumentText = @"
                    <html>
                        <body style='font-family: Segoe UI, Arial, sans-serif;'>
                            <div style='text-align:center; margin-top: 100px;'>
                                <h2 style='color:#0078D4;'>Loading Sam's Delivery Report...</h2>
                                <p style='color:#666;'>Please wait while we fetch the data from the database.</p>
                            </div>
                        </body>
                    </html>";

                btnSendEmail.Enabled = false;
                btnRefresh.Enabled = false;

                // Fetch HTML content from database
                _cachedHtmlContent = await _databaseService.GetHtmlContentFromStoredProcedure(_targetDate);

                if (!string.IsNullOrWhiteSpace(_cachedHtmlContent))
                {
                    webBrowser.DocumentText = _cachedHtmlContent;
                    btnSendEmail.Enabled = true;
                }
                else
                {
                    webBrowser.DocumentText = @"
                        <html>
                            <body style='font-family: Segoe UI, Arial, sans-serif;'>
                                <div style='text-align:center; margin-top: 100px;'>
                                    <h2 style='color: #FF8C00;'>No Data Available</h2>
                                    <p style='color:#666;'>The stored procedure returned no data for the selected date.</p>
                                    <p style='color:#999; font-size: 14px;'>Date: " + _targetDate.ToString("MMMM dd, yyyy") + @"</p>
                                </div>
                            </body>
                        </html>";
                    btnSendEmail.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                webBrowser.DocumentText = $@"
                    <html>
                        <body style='font-family: Segoe UI, Arial, sans-serif;'>
                            <div style='text-align:center; margin-top: 100px;'>
                                <h2 style='color: #DC3545;'>Error Loading Report</h2>
                                <p style='color:#666;'>Failed to load report from database.</p>
                                <div style='background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; padding: 15px; margin: 20px auto; max-width: 600px; text-align: left;'>
                                    <strong>Error Details:</strong><br/>
                                    {ex.Message}
                                </div>
                            </div>
                        </body>
                    </html>";

                MessageBox.Show(
                    $"Error loading report from database:\n\n{ex.Message}",
                    "Database Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                btnSendEmail.Enabled = false;
            }
            finally
            {
                btnRefresh.Enabled = true;
            }
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            LoadReportFromDatabase();
        }

        private async void BtnSendEmail_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate HTML content
                if (string.IsNullOrWhiteSpace(_cachedHtmlContent))
                {
                    MessageBox.Show(
                        "No HTML content available to send.\n\nPlease ensure the report is loaded successfully before sending.",
                        "No Content",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                btnSendEmail.Enabled = false;
                btnSendEmail.Text = "Sending...";

                // Prepare email details
                string subject = $"Sam's Delivery Report Status - {_targetDate:MMMM dd, yyyy}";

                // Get recipient emails from configuration
                string[] toEmails = _emailConfig.ToMails
                    .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(email => email.Trim())
                    .Where(email => !string.IsNullOrWhiteSpace(email))
                    .ToArray();

                if (toEmails.Length == 0)
                {
                    MessageBox.Show(
                        "No recipient email addresses configured.\n\nPlease check the 'ToMails' setting in appsettings.json.",
                        "Configuration Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Send email
                bool success = await _emailService.SendEmailAsync(
                    toEmails,
                    subject,
                    _cachedHtmlContent,
                    isHtml: true
                );

                if (success)
                {
                    string ccInfo = !string.IsNullOrWhiteSpace(_emailConfig.ErrorNotifications.CcEmails)
                        ? $"\n\nCC: {_emailConfig.ErrorNotifications.CcEmails}"
                        : "";

                    MessageBox.Show(
                        $"Email sent successfully!\n\nRecipients:\n{string.Join("\n", toEmails)}{ccInfo}",
                        "Email Sent",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to send email:\n\n{ex.Message}",
                    "Email Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                btnSendEmail.Enabled = true;
                btnSendEmail.Text = "Send Email";
            }
        }

        // ✅ FIXED: FormClosing handler with proper field references
        private void SamsDeliveryReportForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // If this is the startup form (no R26 form exists) and user is closing
            if (_r26Form == null && e.CloseReason == CloseReason.UserClosing)
            {
                var result = MessageBox.Show(
                    "Are you sure you want to exit the application?",
                    "Confirm Exit",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    Application.Exit();
                }
                else
                {
                    e.Cancel = true;
                }
            }
            // If R26 form exists and user is closing, show R26 form
            else if (_r26Form != null && !_r26Form.IsDisposed && e.CloseReason == CloseReason.UserClosing)
            {
                _r26Form.Show();
                _r26Form.BringToFront();
            }
        }

        private void MenuItemR26Queue_Click(object sender, EventArgs e)
        {
            if (_r26QueueForm != null && !_r26QueueForm.IsDisposed)
            {
                _r26QueueForm.ShowR26Form();
                this.Hide();
            }
            else
            {
                _r26QueueForm = new R26QueueForm(_emailConfig);
                _r26QueueForm.FormClosed += (s, args) =>
                {
                    _r26QueueForm = null;
                    // ✅ Show Sam's form again when R26 form is closed
                    this.Show();
                    this.BringToFront();
                };
                _r26QueueForm.Show();
                this.Hide();
            }
        }

        public void ShowSamsDeliveryForm()
        {
            this.Show();
            this.BringToFront();
            this.Focus();
        }
    }

    // ============================================================
    // DATABASE SERVICE CLASS
    // ============================================================
    internal class SamsDeliveryDatabaseService
    {
        private readonly string _connectionString;
        private readonly string _storedProcedureName;

        public SamsDeliveryDatabaseService(string connectionString, string storedProcedureName)
        {
            _connectionString = connectionString;
            _storedProcedureName = storedProcedureName;
        }

        public async Task<string> GetHtmlContentFromStoredProcedure(DateTime targetDate)
        {
            try
            {
                using (var connection = new SqlConnection(_connectionString))
                using (var command = new SqlCommand(_storedProcedureName, connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandTimeout = 300; // 5 minutes timeout

                    // Add parameter for target date
                    command.Parameters.AddWithValue("@TargetDate", targetDate);

                    await connection.OpenAsync();

                    // Execute and get the HTML content
                    var result = await command.ExecuteScalarAsync();
                    return result?.ToString() ?? string.Empty;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error executing stored procedure '{_storedProcedureName}': {ex.Message}", ex);
            }
        }
    }

    // ============================================================
    // EMAIL SERVICE CLASS
    // ============================================================
    internal class SamsDeliveryEmailService
    {
        private readonly EmailConfiguration _config;

        public SamsDeliveryEmailService(EmailConfiguration config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        public async Task<bool> SendEmailAsync(string[] toEmails, string subject, string body, bool isHtml = true)
        {
            try
            {
                using (var client = new SmtpClient(_config.SmtpSettings.Host, _config.SmtpSettings.Port))
                {
                    client.EnableSsl = _config.SmtpSettings.EnableSsl;
                    client.UseDefaultCredentials = false;
                    client.Credentials = new NetworkCredential(
                        _config.SmtpSettings.UserName,
                        _config.SmtpSettings.Password
                    );
                    client.Timeout = 30000; // 30 seconds

                    using (var message = new MailMessage())
                    {
                        message.From = new MailAddress(
                            _config.SmtpSettings.FromEmail,
                            _config.SmtpSettings.FromName
                        );

                        // Add TO recipients
                        foreach (var email in toEmails)
                        {
                            if (!string.IsNullOrWhiteSpace(email))
                            {
                                message.To.Add(email.Trim());
                            }
                        }

                        // Add CC recipients if configured
                        if (!string.IsNullOrWhiteSpace(_config.ErrorNotifications.CcEmails))
                        {
                            var ccEmails = _config.ErrorNotifications.CcEmails.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var ccEmail in ccEmails)
                            {
                                if (!string.IsNullOrWhiteSpace(ccEmail))
                                {
                                    message.CC.Add(ccEmail.Trim());
                                }
                            }
                        }

                        message.Subject = subject;
                        message.Body = body;
                        message.IsBodyHtml = isHtml;
                        message.Priority = MailPriority.Normal;

                        await client.SendMailAsync(message);
                        return true;
                    }
                }
            }
            catch (SmtpException smtpEx)
            {
                throw new Exception($"SMTP Error: {smtpEx.Message}", smtpEx);
            }
            catch (Exception ex)
            {
                throw new Exception($"Email Error: {ex.Message}", ex);
            }
        }
    }
}