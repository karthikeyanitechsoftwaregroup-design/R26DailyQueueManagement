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
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace R26_DailyQueueWinForm.Forms
{
    [ComVisible(true)]
    public partial class SamsDeliveryReportForm : Form
    {
        private WebBrowser webBrowser;
        private MenuStrip mainMenuStrip;
        private ToolStripMenuItem menuItemMain;
        private ToolStripMenuItem menuItemR26Queue;
        private ToolStripMenuItem menuItemDeliveryReport;
        private Button btnSendEmail;
        private Button btnRefresh;
        private Panel pnlOneDriveCounts;
        private Label lblCountTotal;
        private Label lblCountCompleted;
        private Label lblCountPending;

        private DateTime _targetDate;
        private EmailConfiguration _emailConfig;
        private string _connectionString;
        private string _storedProcedureName;
        private string _cachedHtmlContent;
        private string _currentSortOrder = "none";
        private System.Windows.Forms.Timer _countUpdateTimer;

        private R26QueueForm _r26Form;
        private R26QueueForm _r26QueueForm;

        private SamsDeliveryDatabaseService _databaseService;
        private SamsDeliveryEmailService _emailService;

        // NEW: Flag to prevent re-entry into exit confirmation
        private bool _isExiting = false;

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
            _r26Form = r26Form;
            _r26QueueForm = r26Form;

            _databaseService = new SamsDeliveryDatabaseService(_connectionString, _storedProcedureName);
            _emailService = new SamsDeliveryEmailService(_emailConfig);

            _countUpdateTimer = new System.Windows.Forms.Timer();
            _countUpdateTimer.Interval = 2000;
            //_countUpdateTimer.Tick += CountUpdateTimer_Tick;

            InitializeComponent();
            LoadReportFromDatabase();

            this.FormClosing += SamsDeliveryReportForm_FormClosing;
        }

        private void InitializeComponent()
        {
            this.Text = $"Sam's Delivery Report Status - {_targetDate:yyyy-MM-dd}";
            this.Size = new Size(1400, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;

            mainMenuStrip = new MenuStrip
            {
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI", 10F),
                BackColor = Color.White,
                Height = 35,
                Padding = new Padding(10, 5, 0, 5)
            };

            menuItemMain = new ToolStripMenuItem("Menu")
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.Black
            };

            menuItemR26Queue = new ToolStripMenuItem("R26 Daily Queue Management");
            menuItemR26Queue.Click += MenuItemR26Queue_Click;

            menuItemDeliveryReport = new ToolStripMenuItem("Sam's Delivery Report Status")
            {
                Checked = true
            };

            var menuItemSample1 = new ToolStripMenuItem("Sample 1");
            menuItemSample1.Click += MenuItemSample1_Click;

            var menuItemSample2 = new ToolStripMenuItem("Sample 2");
            menuItemSample2.Click += MenuItemSample2_Click;

            menuItemMain.DropDownItems.Add(menuItemR26Queue);
            menuItemMain.DropDownItems.Add(menuItemDeliveryReport);
            menuItemMain.DropDownItems.Add(new ToolStripSeparator()); 
            menuItemMain.DropDownItems.Add(menuItemSample1);
            menuItemMain.DropDownItems.Add(menuItemSample2);

            mainMenuStrip.Items.Add(menuItemMain);
            this.MainMenuStrip = mainMenuStrip;

            Label lblTitle = new Label
            {
                Text = $"Sam's Delivery Report Status - {_targetDate:MMMM dd, yyyy}",
                Font = new Font("Segoe UI", 14F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(20, 45),
                AutoSize = true
            };

            pnlOneDriveCounts = new Panel
            {
                Location = new Point(730, 45),
                Size = new Size(400, 30),
                BackColor = Color.White,
                Visible = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };

            lblCountTotal = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.Black,
                Location = new Point(0, 5),
                Text = "Total: 0"
            };

            lblCountCompleted = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.Green,
                Location = new Point(100, 5),
                Text = "Completed: 0"
            };

            lblCountPending = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.Red,
                Location = new Point(250, 5),
                Text = "Pending: 0"
            };

            pnlOneDriveCounts.Controls.Add(lblCountTotal);
            pnlOneDriveCounts.Controls.Add(lblCountCompleted);
            pnlOneDriveCounts.Controls.Add(lblCountPending);

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

            webBrowser = new WebBrowser
            {
                Location = new Point(0, 85),
                Size = new Size(this.ClientSize.Width, this.ClientSize.Height - 85),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                ScriptErrorsSuppressed = true
            };

            webBrowser.DocumentCompleted += WebBrowser_DocumentCompleted;
            webBrowser.ObjectForScripting = this;

            this.Controls.Add(mainMenuStrip);
            this.Controls.Add(lblTitle);
            this.Controls.Add(pnlOneDriveCounts);
            this.Controls.Add(btnRefresh);
            this.Controls.Add(btnSendEmail);
            this.Controls.Add(webBrowser);

            pnlOneDriveCounts.BringToFront();

            this.Resize += (s, e) => PositionButtons();
            this.Shown += (s, e) => PositionButtons();
        }

        private void PositionButtons()
        {
            if (btnSendEmail != null && btnRefresh != null)
            {
                btnSendEmail.Location = new Point(
                    this.ClientSize.Width - btnSendEmail.Width - 20,
                    40
                );
                btnRefresh.Location = new Point(
                    btnSendEmail.Location.X - btnRefresh.Width - 10,
                    40
                );
            }
        }

        //private void CountUpdateTimer_Tick(object sender, EventArgs e)
        //{
        //    if (webBrowser.Document != null && webBrowser.ReadyState == WebBrowserReadyState.Complete)
        //    {
        //        UpdateOneDriveLinkCounts();
        //    }
        //}

        private void WebBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser.Document != null)
            {
                UpdateOneDriveLinkCounts();
                _countUpdateTimer.Start();

                try
                {
                    HtmlElementCollection tables = webBrowser.Document.GetElementsByTagName("table");

                    foreach (HtmlElement table in tables)
                    {
                        HtmlElementCollection headers = table.GetElementsByTagName("th");
                        foreach (HtmlElement header in headers)
                        {
                            if (header.InnerText != null && header.InnerText.Trim().Contains("OneDrive Link"))
                            {
                                header.Style = "cursor: pointer; user-select: none;";
                                header.Click -= OneDriveLinkHeader_Click;
                                header.Click += OneDriveLinkHeader_Click;
                                break;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    // Silently handle errors
                }
            }
        }


        private void OneDriveLinkHeader_Click(object sender, HtmlElementEventArgs e)
        {
            SortOneDriveLinks();
        }

        public void OneDriveLinkHeaderClicked()
        {
            SortOneDriveLinks();
        }

        private void SortOneDriveLinks()
        {
            if (webBrowser.Document == null) return;

            try
            {
                // Toggle between original order and sorted (no-links first)
                if (_currentSortOrder == "none")
                    _currentSortOrder = "sorted"; // No-links first
                else
                    _currentSortOrder = "none"; // Original order (links already first from SP)

                HtmlElementCollection tables = webBrowser.Document.GetElementsByTagName("table");
                HtmlElement targetTable = null;

                foreach (HtmlElement table in tables)
                {
                    HtmlElementCollection headers = table.GetElementsByTagName("th");
                    foreach (HtmlElement header in headers)
                    {
                        if (header.InnerText != null && header.InnerText.Contains("OneDrive Link"))
                        {
                            targetTable = table;
                            break;
                        }
                    }
                    if (targetTable != null) break;
                }

                if (targetTable == null) return;

                HtmlElementCollection tbody = targetTable.GetElementsByTagName("tbody");
                if (tbody.Count == 0) return;

                HtmlElement tableBody = tbody[0];
                HtmlElementCollection rows = tableBody.GetElementsByTagName("tr");

                if (rows.Count <= 1) return;

                var rowDataList = new System.Collections.Generic.List<RowData>();

                foreach (HtmlElement row in rows)
                {
                    HtmlElementCollection cells = row.GetElementsByTagName("td");
                    if (cells.Count >= 4)
                    {
                        string oneDriveLinkHtml = cells[3].InnerHtml ?? "";
                        bool hasLink = oneDriveLinkHtml.Contains("<a") || oneDriveLinkHtml.Contains("📂");

                        rowDataList.Add(new RowData
                        {
                            RowHtml = row.OuterHtml,
                            HasLink = hasLink
                        });
                    }
                }

                // Sort only when in "sorted" mode - rows WITHOUT links appear first
                if (_currentSortOrder == "sorted")
                {
                    rowDataList = rowDataList.OrderBy(r => r.HasLink).ToList(); // false (no link) comes first
                }
                // When _currentSortOrder is "none", keep original order (links already first from SP)

                // ✅ FIX: Reorder rows without losing any data
                // This rebuilds the table with ALL existing rows, just in a different order
                string newTableBodyHtml = "";
                foreach (var rowData in rowDataList)
                {
                    newTableBodyHtml += rowData.RowHtml; // Keeps ALL row data intact
                }

                tableBody.InnerHtml = newTableBodyHtml; // Replace with reordered rows

                // Update header with sort indicator
                HtmlElementCollection tableHeaders = targetTable.GetElementsByTagName("th");
                foreach (HtmlElement header in tableHeaders)
                {
                    if (header.InnerText != null && header.InnerText.Contains("OneDrive Link"))
                    {
                        // No indicator needed - user can see the order change
                        string headerText = "OneDrive Link";
                        header.InnerText = headerText;
                    }
                }

                // ✅ Re-attach click handlers after sorting (they get lost when InnerHtml is replaced)
                HtmlElementCollection newRows = tableBody.GetElementsByTagName("tr");
                foreach (HtmlElement row in newRows)
                {
                    HtmlElementCollection cells = row.GetElementsByTagName("td");
                    if (cells.Count >= 4)
                    {
                        // Optional: You can add click handlers to cells here if needed
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error sorting OneDrive links: {ex.Message}",
                    "Sort Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void MenuItemSample1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Sample 1 menu item clicked!\n\nYou can implement your custom functionality here.",
                "Sample 1",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void MenuItemSample2_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Sample 2 menu item clicked!\n\nYou can implement your custom functionality here.",
                "Sample 2",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        private void UpdateOneDriveLinkCounts()
        {
            if (webBrowser.Document == null)
            {
                pnlOneDriveCounts.Visible = false;
                return;
            }

            try
            {
                int totalCount = 0;
                int completedCount = 0;
                int pendingCount = 0;

                HtmlElementCollection tables = webBrowser.Document.GetElementsByTagName("table");
                HtmlElement targetTable = null;
                int oneDriveLinkColumnIndex = -1;

                foreach (HtmlElement table in tables)
                {
                    HtmlElementCollection tableHeaders = table.GetElementsByTagName("th");
                    for (int i = 0; i < tableHeaders.Count; i++)
                    {
                        HtmlElement header = tableHeaders[i];
                        if (header.InnerText != null && header.InnerText.Contains("OneDrive Link"))
                        {
                            targetTable = table;
                            oneDriveLinkColumnIndex = i;
                            break;
                        }
                    }
                    if (targetTable != null) break;
                }

                if (targetTable != null && oneDriveLinkColumnIndex >= 0)
                {
                    HtmlElementCollection tbody = targetTable.GetElementsByTagName("tbody");
                    if (tbody.Count > 0)
                    {
                        HtmlElementCollection rows = tbody[0].GetElementsByTagName("tr");

                        foreach (HtmlElement row in rows)
                        {
                            HtmlElementCollection cells = row.GetElementsByTagName("td");

                            if (cells.Count == 0)
                                continue;

                            totalCount++;

                            if (cells.Count > oneDriveLinkColumnIndex)
                            {
                                HtmlElement cell = cells[oneDriveLinkColumnIndex];
                                string cellContent = cell.InnerHtml ?? "";
                                string cellText = cell.InnerText ?? "";

                                bool hasLink = cellContent.Contains("href=") ||
                                              cellContent.Contains("<a ") ||
                                              cellContent.Contains("<a>") ||
                                              cellContent.Contains("📂") ||
                                              (cellText.Contains("📂"));

                                if (hasLink)
                                {
                                    completedCount++;
                                }
                                else if (cellText.Trim() == "-" || string.IsNullOrWhiteSpace(cellText))
                                {
                                    pendingCount++;
                                }
                                else
                                {
                                    if (!string.IsNullOrWhiteSpace(cellText) && cellText.Trim() != "-")
                                    {
                                        completedCount++;
                                    }
                                    else
                                    {
                                        pendingCount++;
                                    }
                                }
                            }
                        }
                    }

                    lblCountTotal.Text = $"Total: {totalCount}";
                    lblCountCompleted.Text = $"| Completed: {completedCount}";
                    lblCountPending.Text = $"| Pending: {pendingCount}";

                    lblCountTotal.Location = new Point(5, 5);
                    lblCountCompleted.Location = new Point(lblCountTotal.Right + 10, 5);
                    lblCountPending.Location = new Point(lblCountCompleted.Right + 10, 5);

                    pnlOneDriveCounts.Size = new Size(lblCountPending.Right + 10, 30);

                    pnlOneDriveCounts.Visible = true;
                    pnlOneDriveCounts.BringToFront();
                }
                else
                {
                    pnlOneDriveCounts.Visible = false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Count error: {ex.Message}");
                pnlOneDriveCounts.Visible = false;
            }
        }

        private async void LoadReportFromDatabase()
        {
            try
            {
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
            _currentSortOrder = "none";
            _countUpdateTimer.Stop();
            LoadReportFromDatabase();
        }

        private async void BtnSendEmail_Click(object sender, EventArgs e)
        {
            try
            {
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

                string subject = $"Sam's Delivery Report Status - {_targetDate:MMMM dd, yyyy}";

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

        private void SamsDeliveryReportForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // ✅ Stop and dispose timer
            if (_countUpdateTimer != null)
            {
                _countUpdateTimer.Stop();
                _countUpdateTimer.Dispose();
            }

            // ✅ Check if already exiting to prevent re-entry
            if (_isExiting)
            {
                return;
            }

            // ✅ If this is triggered by Application.Exit() from another form, don't show confirmation
            if (e.CloseReason == CloseReason.ApplicationExitCall)
            {
                return;
            }

            // Only handle user closing (X button)
            if (e.CloseReason == CloseReason.UserClosing)
            {
                // ✅ Set flag BEFORE showing dialog to prevent re-entry
                _isExiting = true;

                // Always show exit confirmation when user clicks close button
                var result = MessageBox.Show(
                    "Are you sure you want to exit the application?",
                    "Confirm Exit",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Close R26 form if it exists (without triggering their events)
                    if (_r26Form != null && !_r26Form.IsDisposed)
                    {
                        // ✅ Set their exit flag to prevent their dialog
                        _r26Form._isExiting = true;
                        _r26Form.Close();
                    }

                    if (_r26QueueForm != null && !_r26QueueForm.IsDisposed && _r26QueueForm != _r26Form)
                    {
                        // ✅ Set their exit flag to prevent their dialog
                        _r26QueueForm._isExiting = true;
                        _r26QueueForm.Close();
                    }

                    // Exit the entire application
                    Application.Exit();
                }
                else
                {
                    // ✅ Reset flag if user cancels
                    _isExiting = false;
                    // Cancel the close operation
                    e.Cancel = true;
                }
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

        private class RowData
        {
            public string RowHtml { get; set; }
            public bool HasLink { get; set; }
        }
    }

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
                    command.CommandTimeout = 300;

                    command.Parameters.AddWithValue("@TargetDate", targetDate);

                    await connection.OpenAsync();

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
                    client.Timeout = 30000;

                    using (var message = new MailMessage())
                    {
                        message.From = new MailAddress(
                            _config.SmtpSettings.FromEmail,
                            _config.SmtpSettings.FromName
                        );

                        foreach (var email in toEmails)
                        {
                            if (!string.IsNullOrWhiteSpace(email))
                            {
                                message.To.Add(email.Trim());
                            }
                        }

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