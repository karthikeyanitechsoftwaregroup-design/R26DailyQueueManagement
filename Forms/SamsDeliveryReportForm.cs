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
using System.Collections.Generic;
using System.Text;

namespace R26_DailyQueueWinForm.Forms
{
    [ComVisible(true)]
    public partial class SamsDeliveryReportForm : Form
    {
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);

        private const int WM_SETREDRAW = 0x000B;
        private const int WM_VSCROLL = 0x0115;
        private const int SB_LINEDOWN = 1;
        private const int SB_LINEUP = 0;
        private DataGridView dgvDeliveryStatus;
        private DataGridView dgvOneDriveLinks;
        private Label lblDeliveryStatusTitle;
        private Label lblOneDriveLinksTitle;
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
        private ComboBox cmbFilterReportName;
        private ComboBox cmbFilterCompanyName;
        private Label lblFilterReportName;
        private Label lblFilterCompanyName;
        private Label lblFilterOneDriveLink;
        private ComboBox cmbFilterOneDriveLink;
        private string _lastSortedColumn = "";
        private bool _lastSortAscending = true;

        private DateTime _targetDate;
        private EmailConfiguration _emailConfig;
        private string _connectionString;
        private string _storedProcedureName;
        private DataTable _originalDeliveryStatusTable;
        private DataTable _originalOneDriveLinksTable;
        private DataTable _filteredOneDriveLinksTable;
        private string _cachedHtmlContent;
        private string _currentSortOrder = "none";
        private System.Windows.Forms.Timer _countUpdateTimer;

        private R26QueueForm _r26Form;
        private R26QueueForm _r26QueueForm;
        private RpaScheduleQueueDetailForm _rpaScheduleQueueDetailForm;

        private SamsDeliveryDatabaseService _databaseService;
        private SamsDeliveryEmailService _emailService;

        public bool _isExiting = false;

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
            _countUpdateTimer.Tick += CountUpdateTimer_Tick;

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

            var menuItemProduction = new ToolStripMenuItem("Production");
            menuItemProduction.Click += MenuItemProduction_Click;

            menuItemMain.DropDownItems.Add(menuItemR26Queue);
            menuItemMain.DropDownItems.Add(menuItemDeliveryReport);
            menuItemMain.DropDownItems.Add(menuItemProduction);

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
                Text = "| Completed: 0"
            };

            lblCountPending = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.Red,
                Location = new Point(250, 5),
                Text = "| Pending: 0"
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

            // Delivery Status Table Title
            lblDeliveryStatusTitle = new Label
            {
                Text = $"Delivery Status - {_targetDate:dd MMM yyyy}",
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                Location = new Point(20, 90),
                AutoSize = true,
                BackColor = Color.LightGray,
                Padding = new Padding(5)
            };

            // Delivery Status DataGridView
            dgvDeliveryStatus = new DataGridView
            {
                Location = new Point(20, 130),
                Size = new Size(this.ClientSize.Width - 40, 142),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                EnableHeadersVisualStyles = false,
                AllowUserToResizeRows = false,
                RowTemplate = { Height = 20 },
                ColumnHeadersHeight = 40
            };

            dgvDeliveryStatus.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 58, 85);
            dgvDeliveryStatus.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(30, 58, 85);
            dgvDeliveryStatus.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDeliveryStatus.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dgvDeliveryStatus.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDeliveryStatus.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;

            dgvDeliveryStatus.DefaultCellStyle.Font = new Font("Segoe UI", 8F);
            dgvDeliveryStatus.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 212);
            dgvDeliveryStatus.DefaultCellStyle.SelectionForeColor = Color.White;
            dgvDeliveryStatus.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);

            EnableSmoothScrolling(dgvDeliveryStatus);

            // OneDrive Links Table Title
            lblOneDriveLinksTitle = new Label
            {
                Text = $"OneDrive Links For Processed Files - {_targetDate:dd MMM yyyy}",
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                Location = new Point(20, 280),
                AutoSize = true,
                BackColor = Color.LightGray,
                Padding = new Padding(5),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };

            // Filter labels and comboboxes
            lblFilterReportName = new Label
            {
                Text = "Report Name:",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 323),
                AutoSize = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };

            cmbFilterReportName = new ComboBox
            {
                Font = new Font("Segoe UI", 9F),
                Location = new Point(110, 320),
                Size = new Size(470, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            cmbFilterReportName.SelectedIndexChanged += FilterComboBox_SelectedIndexChanged;

            lblFilterCompanyName = new Label
            {
                Text = "Company Name:",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(600, 323),
                AutoSize = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };

            cmbFilterCompanyName = new ComboBox
            {
                Font = new Font("Segoe UI", 9F),
                Location = new Point(705, 320),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            cmbFilterCompanyName.SelectedIndexChanged += FilterComboBox_SelectedIndexChanged;

            lblFilterOneDriveLink = new Label
            {
                Text = "OneDrive Status:",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(925, 323),
                AutoSize = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };

            cmbFilterOneDriveLink = new ComboBox
            {
                Font = new Font("Segoe UI", 9F),
                Location = new Point(1035, 320),
                Size = new Size(150, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            cmbFilterOneDriveLink.SelectedIndexChanged += FilterComboBox_SelectedIndexChanged;


            // OneDrive Links DataGridView
            dgvOneDriveLinks = new DataGridView
            {
                Location = new Point(20, 350),
                Size = new Size(this.ClientSize.Width - 150, this.ClientSize.Height - 361),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                EnableHeadersVisualStyles = false,
                AllowUserToResizeRows = false,
                RowTemplate = { Height = 30 },
                ColumnHeadersHeight = 40
            };

            dgvOneDriveLinks.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 58, 85);
            dgvOneDriveLinks.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvOneDriveLinks.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dgvOneDriveLinks.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvOneDriveLinks.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(30, 58, 85);

            dgvOneDriveLinks.DefaultCellStyle.Font = new Font("Segoe UI", 9F);
            dgvOneDriveLinks.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 212);
            dgvOneDriveLinks.DefaultCellStyle.SelectionForeColor = Color.White;
            dgvOneDriveLinks.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);

            dgvOneDriveLinks.CellContentClick += DgvOneDriveLinks_CellContentClick;
            dgvOneDriveLinks.ColumnHeaderMouseClick += DgvOneDriveLinks_ColumnHeaderMouseClick;
            dgvOneDriveLinks.CellFormatting += DgvOneDriveLinks_CellFormatting;

            EnableSmoothScrolling(dgvOneDriveLinks);

            this.Controls.Add(mainMenuStrip);
            this.Controls.Add(lblTitle);
            this.Controls.Add(pnlOneDriveCounts);
            this.Controls.Add(btnRefresh);
            this.Controls.Add(btnSendEmail);
            this.Controls.Add(lblDeliveryStatusTitle);
            this.Controls.Add(dgvDeliveryStatus);
            this.Controls.Add(lblOneDriveLinksTitle);
            this.Controls.Add(lblFilterReportName);
            this.Controls.Add(cmbFilterReportName);
            this.Controls.Add(lblFilterCompanyName);
            this.Controls.Add(cmbFilterCompanyName);
            this.Controls.Add(lblFilterOneDriveLink);
            this.Controls.Add(cmbFilterOneDriveLink);
            this.Controls.Add(dgvOneDriveLinks);

            pnlOneDriveCounts.BringToFront();

            this.Resize += (s, e) => PositionButtons();
            this.Shown += (s, e) => PositionButtons();
        }

        private void DgvOneDriveLinks_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgvOneDriveLinks.Columns[e.ColumnIndex].Name == "OneDriveLink" && e.RowIndex >= 0)
            {
                var cellValue = e.Value?.ToString();
                if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith("http"))
                {
                    e.Value = "📂";
                    e.FormattingApplied = true;
                }
            }
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

        private void CountUpdateTimer_Tick(object sender, EventArgs e)
        {
            UpdateOneDriveLinkCounts();
        }

        private void DgvOneDriveLinks_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (_filteredOneDriveLinksTable == null || _filteredOneDriveLinksTable.Rows.Count == 0) return;

            try
            {
                string columnName = dgvOneDriveLinks.Columns[e.ColumnIndex].Name;
                bool ascending = true;

                // Determine sort direction
                if (_lastSortedColumn == columnName)
                {
                    ascending = !_lastSortAscending;
                }

                _lastSortedColumn = columnName;
                _lastSortAscending = ascending;

                // Sort the data
                if (columnName == "OneDriveLink")
                {
                    // Special sorting for OneDrive Link column
                    var sortedRows = _filteredOneDriveLinksTable.AsEnumerable()
                        .OrderBy(row => {
                            string link = row["OneDrive Link"]?.ToString() ?? "";
                            bool hasLink = !string.IsNullOrEmpty(link) && link != "-" &&
                                         (link.Contains("http") || link.Contains("📂"));

                            if (ascending)
                                return hasLink ? 0 : 1;
                            else
                                return hasLink ? 1 : 0;
                        })
                        .ThenBy(row => row["Company Name"]?.ToString() ?? "")
                        .ThenBy(row => row["Report Name"]?.ToString() ?? "")
                        .ToList();

                    DataTable sortedTable = _filteredOneDriveLinksTable.Clone();
                    foreach (var row in sortedRows)
                    {
                        sortedTable.ImportRow(row);
                    }

                    dgvOneDriveLinks.DataSource = sortedTable;
                    _filteredOneDriveLinksTable = sortedTable;
                }
                else
                {
                    // Standard sorting for other columns
                    string sortColumn = dgvOneDriveLinks.Columns[e.ColumnIndex].DataPropertyName;
                    DataView dv = _filteredOneDriveLinksTable.DefaultView;
                    dv.Sort = $"[{sortColumn}] {(ascending ? "ASC" : "DESC")}";

                    _filteredOneDriveLinksTable = dv.ToTable();
                    dgvOneDriveLinks.DataSource = _filteredOneDriveLinksTable;
                }

                // Update column header to show sort direction
                foreach (DataGridViewColumn col in dgvOneDriveLinks.Columns)
                {
                    col.HeaderText = col.DataPropertyName.Replace("OneDriveLink", "OneDrive Link");
                }

                dgvOneDriveLinks.Columns[e.ColumnIndex].HeaderText =
                    dgvOneDriveLinks.Columns[e.ColumnIndex].DataPropertyName.Replace("OneDriveLink", "OneDrive Link") +
                    (ascending ? " ▲" : " ▼");

                UpdateOneDriveLinkCounts();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error sorting column: {ex.Message}",
                    "Sort Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DgvOneDriveLinks_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (dgvOneDriveLinks.Columns[e.ColumnIndex].Name == "OneDriveLink")
                {
                    var cellValue = dgvOneDriveLinks.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith("http"))
                    {
                        try
                        {
                            System.Diagnostics.Process.Start(cellValue);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Failed to open link: {ex.Message}", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private void FilterComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void PopulateFilterDropdowns()
        {
            if (_originalOneDriveLinksTable == null) return;

            try
            {
                // Store current selections
                string currentReportName = cmbFilterReportName.SelectedItem?.ToString();
                string currentCompanyName = cmbFilterCompanyName.SelectedItem?.ToString();

                // Populate Report Name dropdown
                cmbFilterReportName.Items.Clear();
                cmbFilterReportName.Items.Add("-- All Reports --");

                var reportNames = _originalOneDriveLinksTable.AsEnumerable()
                    .Select(row => row["Report Name"]?.ToString())
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Distinct()
                    .OrderBy(name => name)
                    .ToList();

                foreach (var name in reportNames)
                {
                    cmbFilterReportName.Items.Add(name);
                }

                // Populate Company Name dropdown
                cmbFilterCompanyName.Items.Clear();
                cmbFilterCompanyName.Items.Add("-- All Companies --");

                var companyNames = _originalOneDriveLinksTable.AsEnumerable()
                    .Select(row => row["Company Name"]?.ToString())
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Distinct()
                    .OrderBy(name => name)
                    .ToList();

                foreach (var name in companyNames)
                {
                    cmbFilterCompanyName.Items.Add(name);
                }

                // Restore previous selections or set to "All"
                if (!string.IsNullOrEmpty(currentReportName) && cmbFilterReportName.Items.Contains(currentReportName))
                {
                    cmbFilterReportName.SelectedItem = currentReportName;
                }
                else
                {
                    cmbFilterReportName.SelectedIndex = 0;
                }

                if (!string.IsNullOrEmpty(currentCompanyName) && cmbFilterCompanyName.Items.Contains(currentCompanyName))
                {
                    cmbFilterCompanyName.SelectedItem = currentCompanyName;
                }
                else
                {
                    cmbFilterCompanyName.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error populating dropdowns: {ex.Message}");
            }

            // Populate OneDrive Link Status dropdown
            cmbFilterOneDriveLink.Items.Clear();
            cmbFilterOneDriveLink.Items.Add("-- All Status --");
            cmbFilterOneDriveLink.Items.Add("Completed");
            cmbFilterOneDriveLink.Items.Add("Pending");

            string currentOneDriveFilter = cmbFilterOneDriveLink.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(currentOneDriveFilter) && cmbFilterOneDriveLink.Items.Contains(currentOneDriveFilter))
            {
                cmbFilterOneDriveLink.SelectedItem = currentOneDriveFilter;
            }
            else
            {
                cmbFilterOneDriveLink.SelectedIndex = 0;
            }
        }

        private void ApplyFilters()
        {
            if (_originalOneDriveLinksTable == null) return;

            try
            {
                string reportNameFilter = cmbFilterReportName.SelectedItem?.ToString();
                string companyNameFilter = cmbFilterCompanyName.SelectedItem?.ToString();
                string oneDriveLinkFilter = cmbFilterOneDriveLink.SelectedItem?.ToString();

                DataView dv = _originalOneDriveLinksTable.DefaultView;
                List<string> filters = new List<string>();

                if (!string.IsNullOrEmpty(reportNameFilter) &&
                    reportNameFilter != "-- All Reports --")
                {
                    filters.Add($"[Report Name] = '{reportNameFilter.Replace("'", "''")}'");
                }

                if (!string.IsNullOrEmpty(companyNameFilter) &&
                    companyNameFilter != "-- All Companies --")
                {
                    filters.Add($"[Company Name] = '{companyNameFilter.Replace("'", "''")}'");
                }

                dv.RowFilter = filters.Count > 0 ? string.Join(" AND ", filters) : "";

                _filteredOneDriveLinksTable = dv.ToTable();

                // Apply OneDrive Link filter (after converting to DataTable because we need custom logic)
                if (!string.IsNullOrEmpty(oneDriveLinkFilter) && oneDriveLinkFilter != "-- All Status --")
                {
                    DataTable tempTable = _filteredOneDriveLinksTable.Clone();

                    foreach (DataRow row in _filteredOneDriveLinksTable.Rows)
                    {
                        string oneDriveLink = row["OneDrive Link"]?.ToString() ?? "";
                        bool hasLink = !string.IsNullOrEmpty(oneDriveLink) && oneDriveLink != "-" &&
                                      (oneDriveLink.Contains("http") || oneDriveLink.Contains("📂"));

                        if ((oneDriveLinkFilter == "Completed" && hasLink) ||
                            (oneDriveLinkFilter == "Pending" && !hasLink))
                        {
                            tempTable.ImportRow(row);
                        }
                    }

                    _filteredOneDriveLinksTable = tempTable;
                }

                dgvOneDriveLinks.DataSource = _filteredOneDriveLinksTable;

                UpdateOneDriveLinkCounts();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying filters: {ex.Message}", "Filter Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void MenuItemProduction_Click(object sender, EventArgs e)
        {
            var existingRpaForm = Application.OpenForms.OfType<RpaScheduleQueueDetailForm>()
                .FirstOrDefault();

            if (existingRpaForm != null && !existingRpaForm.IsDisposed)
            {
                existingRpaForm.SetR26QueueForm(_r26QueueForm);
                existingRpaForm.SetSamsDeliveryForm(this);
                existingRpaForm.Show();
                existingRpaForm.BringToFront();
                this.Hide();
                return;
            }

            if (_rpaScheduleQueueDetailForm != null && !_rpaScheduleQueueDetailForm.IsDisposed)
            {
                _rpaScheduleQueueDetailForm.SetR26QueueForm(_r26QueueForm);
                _rpaScheduleQueueDetailForm.SetSamsDeliveryForm(this);
                _rpaScheduleQueueDetailForm.Show();
                _rpaScheduleQueueDetailForm.BringToFront();
                this.Hide();
                return;
            }

            _rpaScheduleQueueDetailForm = new RpaScheduleQueueDetailForm(
                _emailConfig,
                _r26QueueForm,
                this
            );

            _rpaScheduleQueueDetailForm.FormClosed += (s, args) =>
            {
                _rpaScheduleQueueDetailForm = null;
                if (!this.IsDisposed && !_isExiting)
                {
                    this.Show();
                    this.BringToFront();
                }
            };

            _rpaScheduleQueueDetailForm.Show();
            this.Hide();
        }

        private void UpdateOneDriveLinkCounts()
        {
            if (_filteredOneDriveLinksTable == null || _filteredOneDriveLinksTable.Rows.Count == 0)
            {
                pnlOneDriveCounts.Visible = false;
                return;
            }

            try
            {
                int totalCount = _filteredOneDriveLinksTable.Rows.Count;
                int completedCount = 0;
                int pendingCount = 0;

                foreach (DataRow row in _filteredOneDriveLinksTable.Rows)
                {
                    string oneDriveLink = row["OneDrive Link"]?.ToString() ?? "";

                    bool hasLink = !string.IsNullOrEmpty(oneDriveLink) && oneDriveLink != "-" &&
                                  (oneDriveLink.Contains("http") || oneDriveLink.Contains("📂"));

                    if (hasLink)
                    {
                        completedCount++;
                    }
                    else
                    {
                        pendingCount++;
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
                dgvDeliveryStatus.DataSource = null;
                dgvOneDriveLinks.DataSource = null;
                btnSendEmail.Enabled = false;
                btnRefresh.Enabled = false;

                _cachedHtmlContent = await _databaseService.GetHtmlContentFromStoredProcedure(_targetDate);

                if (!string.IsNullOrWhiteSpace(_cachedHtmlContent))
                {
                    var tables = ParseHtmlToDataTables(_cachedHtmlContent);

                    if (tables.Count >= 1)
                    {
                        _originalDeliveryStatusTable = tables[0];
                        SetupDeliveryStatusGrid();
                        dgvDeliveryStatus.DataSource = _originalDeliveryStatusTable;
                    }

                    if (tables.Count >= 2)
                    {
                        _originalOneDriveLinksTable = tables[1];
                        _filteredOneDriveLinksTable = _originalOneDriveLinksTable.Copy();
                        SetupOneDriveLinksGrid();
                        dgvOneDriveLinks.DataSource = _filteredOneDriveLinksTable;

                        // Populate filter dropdowns
                        PopulateFilterDropdowns();

                        UpdateOneDriveLinkCounts();
                        _countUpdateTimer.Start();
                    }

                    btnSendEmail.Enabled = true;
                }
                else
                {
                    MessageBox.Show(
                        $"The stored procedure returned no data for the selected date.\n\nDate: {_targetDate:MMMM dd, yyyy}",
                        "No Data Available",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    btnSendEmail.Enabled = false;
                }
            }
            catch (Exception ex)
            {
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

        private List<DataTable> ParseHtmlToDataTables(string html)
        {
            List<DataTable> tables = new List<DataTable>();

            try
            {
                var tableMatches = Regex.Matches(html, @"<table[^>]*>(.*?)</table>", RegexOptions.Singleline | RegexOptions.IgnoreCase);

                foreach (Match tableMatch in tableMatches)
                {
                    DataTable dt = new DataTable();
                    string tableContent = tableMatch.Groups[1].Value;

                    // Extract headers
                    var headerMatches = Regex.Matches(tableContent, @"<th[^>]*>(.*?)</th>", RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    foreach (Match header in headerMatches)
                    {
                        string headerText = Regex.Replace(header.Groups[1].Value, @"<[^>]+>", "").Trim();
                        headerText = System.Net.WebUtility.HtmlDecode(headerText);

                        if (string.IsNullOrWhiteSpace(headerText))
                        {
                            headerText = $"Column{dt.Columns.Count + 1}";
                        }

                        string uniqueColumnName = headerText;
                        int counter = 1;
                        while (dt.Columns.Contains(uniqueColumnName))
                        {
                            uniqueColumnName = $"{headerText}{counter}";
                            counter++;
                        }

                        dt.Columns.Add(uniqueColumnName);
                    }

                    // Extract body
                    var bodyMatch = Regex.Match(tableContent, @"<tbody[^>]*>(.*?)</tbody>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                    string bodyContent = bodyMatch.Success ? bodyMatch.Groups[1].Value : tableContent;

                    var rowMatches = Regex.Matches(bodyContent, @"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    foreach (Match rowMatch in rowMatches)
                    {
                        string rowContent = rowMatch.Groups[1].Value;

                        if (rowContent.Contains("<th")) continue;

                        var cellMatches = Regex.Matches(rowContent, @"<td[^>]*>(.*?)</td>", RegexOptions.Singleline | RegexOptions.IgnoreCase);

                        if (cellMatches.Count > 0 && cellMatches.Count <= dt.Columns.Count)
                        {
                            DataRow dr = dt.NewRow();

                            for (int i = 0; i < cellMatches.Count; i++)
                            {
                                string cellContent = cellMatches[i].Groups[1].Value;

                                var linkMatch = Regex.Match(cellContent, @"href=['""]([^'""]+)['""]", RegexOptions.IgnoreCase);
                                if (linkMatch.Success)
                                {
                                    dr[i] = linkMatch.Groups[1].Value;
                                }
                                else
                                {
                                    string cellText = Regex.Replace(cellContent, @"<[^>]+>", "").Trim();
                                    cellText = System.Net.WebUtility.HtmlDecode(cellText);
                                    cellText = cellText.Replace("&nbsp;", " ").Trim();
                                    dr[i] = string.IsNullOrWhiteSpace(cellText) ? "" : cellText;
                                }
                            }

                            dt.Rows.Add(dr);
                        }
                    }

                    if (dt.Columns.Count > 0)
                    {
                        tables.Add(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error parsing HTML: {ex.Message}", ex);
            }

            return tables;
        }

        private void EnableSmoothScrolling(DataGridView dgv)
        {
            if (dgv == null) return;

            try
            {
                // Enable double buffering to reduce flicker
                typeof(DataGridView).InvokeMember("DoubleBuffered",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
                    null, dgv, new object[] { true });

                // Attach mouse wheel event for smooth scrolling
                dgv.MouseWheel += Dgv_MouseWheel;
            }
            catch
            {
                // If smooth scrolling setup fails, continue with default behavior
            }
        }

private void Dgv_MouseWheel(object sender, MouseEventArgs e)
{
    try
    {
        DataGridView grid = sender as DataGridView;
        if (grid == null || grid.RowCount == 0) return;

        // Don't interfere if no rows to scroll
        if (grid.Rows.Count == 0)
        {
            return;
        }

        int currentIndex = grid.FirstDisplayedScrollingRowIndex;
        if (currentIndex < 0) return;

        int scrollLines = SystemInformation.MouseWheelScrollLines;
        int scrollAmount = Math.Max(1, scrollLines / 3);

        int newIndex = currentIndex;

        if (e.Delta > 0) // Scroll up
        {
            newIndex = Math.Max(0, currentIndex - scrollAmount);
        }
        else // Scroll down
        {
            int maxIndex = Math.Max(0, grid.RowCount - 1);
            newIndex = Math.Min(maxIndex, currentIndex + scrollAmount);
        }

        // Only set if valid and different
        if (newIndex >= 0 && newIndex < grid.RowCount && newIndex != currentIndex)
        {
            grid.FirstDisplayedScrollingRowIndex = newIndex;
        }

        // Mark the event as handled
        ((HandledMouseEventArgs)e).Handled = true;
    }
    catch
    {
        // Silently ignore scroll errors
    }
}
        private void SetupDeliveryStatusGrid()
        {
            dgvDeliveryStatus.Columns.Clear();
            dgvDeliveryStatus.AutoGenerateColumns = false;

            foreach (DataColumn col in _originalDeliveryStatusTable.Columns)
            {
                DataGridViewTextBoxColumn textColumn = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = col.ColumnName,
                    Name = col.ColumnName,
                    HeaderText = col.ColumnName,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                };

                if (col.ColumnName == "No")
                {
                    textColumn.Width = 50;
                    textColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else if (col.ColumnName == "Code")
                {
                    textColumn.Width = 60;
                    textColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else if (col.ColumnName == "Frequency")
                {
                    textColumn.Width = 100;
                    textColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else if (col.ColumnName == "Report Name")
                {
                    textColumn.Width = 400;
                    textColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                }
                else
                {
                    textColumn.Width = 120;
                    textColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }

                dgvDeliveryStatus.Columns.Add(textColumn);
            }
        }

        private void SetupOneDriveLinksGrid()
        {
            dgvOneDriveLinks.Columns.Clear();
            dgvOneDriveLinks.AutoGenerateColumns = false;

            foreach (DataColumn col in _originalOneDriveLinksTable.Columns)
            {
                if (col.ColumnName == "OneDrive Link")
                {
                    DataGridViewLinkColumn linkColumn = new DataGridViewLinkColumn
                    {
                        DataPropertyName = col.ColumnName,
                        Name = "OneDriveLink",
                        HeaderText = "OneDrive Link",
                        Width = 250,
                        SortMode = DataGridViewColumnSortMode.Programmatic,
                        LinkBehavior = LinkBehavior.HoverUnderline,
                        LinkColor = Color.Black,
                        ActiveLinkColor = Color.FromArgb(0, 100, 180),
                        VisitedLinkColor = Color.FromArgb(0, 120, 212),
                        DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
                    };
                    dgvOneDriveLinks.Columns.Add(linkColumn);
                }
                else
                {
                    DataGridViewTextBoxColumn textColumn = new DataGridViewTextBoxColumn
                    {
                        DataPropertyName = col.ColumnName,
                        Name = col.ColumnName,
                        HeaderText = col.ColumnName,
                        SortMode = DataGridViewColumnSortMode.NotSortable,
                        AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                    };

                    if (col.ColumnName == "No")
                    {
                        textColumn.Width = 50;
                        textColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }
                    else if (col.ColumnName == "Company Name")
                    {
                        textColumn.Width = 150;
                    }
                    else if (col.ColumnName == "Report Name")
                    {
                        textColumn.Width = 250;
                        textColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    }

                    dgvOneDriveLinks.Columns.Add(textColumn);
                }
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
                        "No data available to send.\n\nPlease ensure the report is loaded successfully before sending.",
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
            if (_countUpdateTimer != null)
            {
                _countUpdateTimer.Stop();
                _countUpdateTimer.Dispose();
            }

            if (_isExiting)
            {
                return;
            }

            if (e.CloseReason == CloseReason.ApplicationExitCall)
            {
                return;
            }

            if (e.CloseReason == CloseReason.UserClosing)
            {
                _isExiting = true;

                var result = MessageBox.Show(
                    "Are you sure you want to exit the application?",
                    "Confirm Exit",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    if (_r26Form != null && !_r26Form.IsDisposed)
                    {
                        _r26Form._isExiting = true;
                        _r26Form.Close();
                    }

                    if (_r26QueueForm != null && !_r26QueueForm.IsDisposed && _r26QueueForm != _r26Form)
                    {
                        _r26QueueForm._isExiting = true;
                        _r26QueueForm.Close();
                    }

                    if (_rpaScheduleQueueDetailForm != null && !_rpaScheduleQueueDetailForm.IsDisposed)
                    {
                        _rpaScheduleQueueDetailForm._isExiting = true;
                        _rpaScheduleQueueDetailForm.Close();
                    }

                    Application.Exit();
                }
                else
                {
                    _isExiting = false;
                    e.Cancel = true;
                }
            }
        }

        public void SetR26QueueForm(R26QueueForm r26Form)
        {
            _r26QueueForm = r26Form;
            _r26Form = r26Form;
        }

        public void SetRpaScheduleQueueDetailForm(RpaScheduleQueueDetailForm rpaForm)
        {
            _rpaScheduleQueueDetailForm = rpaForm;
        }

        private void MenuItemR26Queue_Click(object sender, EventArgs e)
        {
            var existingR26Form = Application.OpenForms.OfType<R26QueueForm>()
                .FirstOrDefault();

            if (existingR26Form != null && !existingR26Form.IsDisposed)
            {
                existingR26Form.SetSamsDeliveryForm(this);
                existingR26Form.Show();
                existingR26Form.BringToFront();
                this.Hide();
                return;
            }

            if (_r26QueueForm != null && !_r26QueueForm.IsDisposed)
            {
                _r26QueueForm.SetSamsDeliveryForm(this);
                _r26QueueForm.Show();
                _r26QueueForm.BringToFront();
                this.Hide();
                return;
            }

            _r26QueueForm = new R26QueueForm(_emailConfig);
            _r26Form = _r26QueueForm;
            _r26QueueForm.SetSamsDeliveryForm(this);

            _r26QueueForm.FormClosed += (s, args) =>
            {
                _r26QueueForm = null;
                _r26Form = null;
                if (!this.IsDisposed && !_isExiting)
                {
                    this.Show();
                    this.BringToFront();
                }
            };

            _r26QueueForm.Show();
            this.Hide();
        }

        public void ShowSamsDeliveryForm()
        {
            this.Show();
            this.BringToFront();
            this.Focus();
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