using Microsoft.Extensions.Configuration;
using R26_DailyQueueWinForm.Data;
using R26_DailyQueueWinForm.Forms;
using R26_DailyQueueWinForm.Models;
using R26DailyQueueWinForm;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace R26_DailyQueueWinForm
{
    public partial class RpaScheduleQueueDetailForm : Form
    {
        private readonly RpaScheduleQueueDetailRepository _repository;
        private List<RpaScheduleQueueDetailModel> _queueData;
        private DataTable _dataTable;
        private Dictionary<int, string> _individualStatusChanges;
        private BindingSource _bindingSource;
        private EmailConfiguration _emailConfig;
        public bool _isExiting = false;

        private R26QueueForm _r26QueueForm;
        private SamsDeliveryReportForm _samsDeliveryForm;

        // Filter controls
        private Label lblCompanyFilter;
        private ComboBox cmbCompanyFilter;
        private Label lblStatusFilter;
        private ComboBox cmbStatusFilter;

        // Menu controls
        private MenuStrip mainMenuStrip;
        private ToolStripMenuItem menuItemR26;
        private ToolStripMenuItem menuItemReports;
        private ToolStripMenuItem menuItemProduction;
        private ToolStripMenuItem menuItemFileType;
        private ToolStripMenuItem menuItemRawFile;
        private ToolStripMenuItem menuItemBulkFile;
        private ToolStripMenuItem menuItemProcessedFile;

        // Main UI controls
        private DataGridView dgvQueue;
        private Button btnBulkUpdate;
        private Button btnSaveIndividual;
        private Button btnRefresh;
        private Label lblTitle;
        private Panel panelSide;
        private Label lblSideTitle;
        private DataGridView dgvSelectedRecords;
        private Label lblBulkStatus;
        private ComboBox cmbBulkStatus;
        private Label lblRecordCount;
        private Label lblSelectedCount;
        private Label lblIndividualChanges;
        private ProgressBar progressBar;
        private TextBox txtSearch;
        private Label lblSearch;

        // Loading panel
        private Panel loadingPanel;

        private string _systemName;
        private List<string> _statusList;

        private string GetSystemName()
        {
            try
            {
                return Environment.MachineName;
            }
            catch
            {
                return "Unknown";
            }
        }

        public RpaScheduleQueueDetailForm(
    EmailConfiguration emailConfig,
    R26QueueForm r26Form = null,
    SamsDeliveryReportForm samsForm = null)
        {
            InitializeComponent();

            _emailConfig = emailConfig;
            _systemName = GetSystemName();
            _r26QueueForm = r26Form;      // This is the ACTUAL R26QueueForm
            _samsDeliveryForm = samsForm;

            string connectionString = Program.DbConnectionString;
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                MessageBox.Show("Database connection string is empty!", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                var builder = new System.Data.SqlClient.SqlConnectionStringBuilder(connectionString);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Invalid connection string format:\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _repository = new RpaScheduleQueueDetailRepository(connectionString);
            _individualStatusChanges = new Dictionary<int, string>();
            _bindingSource = new BindingSource();

            LoadStatusesFromDatabase();

            if (mainMenuStrip != null && lblIndividualChanges != null)
            {
                mainMenuStrip.Dock = DockStyle.None;
            }

            this.Load += RpaScheduleQueueDetailForm_Load;
            this.FormClosing += RpaScheduleQueueDetailForm_FormClosing;
        }

        public void SetR26QueueForm(R26QueueForm r26Form)
        {
            _r26QueueForm = r26Form;
        }

        public void SetSamsDeliveryForm(SamsDeliveryReportForm samsForm)
        {
            _samsDeliveryForm = samsForm;
        }

        private void LoadStatusesFromDatabase()
        {
            try
            {
                _statusList = _repository.GetAllStatuses();

                if (_statusList == null || _statusList.Count == 0)
                {
                    MessageBox.Show("No statuses found in database.",
                        "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _statusList = new List<string>();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading statuses: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusList = new List<string>();
            }
        }

        private void InitializeComponent()
        {
            this.dgvQueue = new DataGridView();
            this.btnBulkUpdate = new Button();
            this.btnSaveIndividual = new Button();
            this.btnRefresh = new Button();
            this.lblTitle = new Label();
            this.panelSide = new Panel();
            this.lblSideTitle = new Label();
            this.dgvSelectedRecords = new DataGridView();
            this.lblBulkStatus = new Label();
            this.cmbBulkStatus = new ComboBox();
            this.lblRecordCount = new Label();
            this.lblSelectedCount = new Label();
            this.lblIndividualChanges = new Label();
            this.progressBar = new ProgressBar();
            this.txtSearch = new TextBox();
            this.lblSearch = new Label();
            this.mainMenuStrip = new MenuStrip();
            this.menuItemR26 = new ToolStripMenuItem();
            this.menuItemReports = new ToolStripMenuItem();
            this.menuItemProduction = new ToolStripMenuItem();
            this.lblCompanyFilter = new Label();
            this.cmbCompanyFilter = new ComboBox();
            this.lblStatusFilter = new Label();
            this.cmbStatusFilter = new ComboBox();

            ((ISupportInitialize)(this.dgvQueue)).BeginInit();
            ((ISupportInitialize)(this.dgvSelectedRecords)).BeginInit();
            this.panelSide.SuspendLayout();
            this.SuspendLayout();

            // MenuStrip
            this.mainMenuStrip.BackColor = Color.White;
            this.mainMenuStrip.Dock = DockStyle.Top;
            this.mainMenuStrip.Font = new Font("Segoe UI", 10F);
            this.mainMenuStrip.Height = 35;
            this.mainMenuStrip.GripStyle = ToolStripGripStyle.Hidden;
            this.mainMenuStrip.RenderMode = ToolStripRenderMode.Professional;
            this.mainMenuStrip.Padding = new Padding(10, 5, 0, 5);

            var menuItemMain = new ToolStripMenuItem("Menu")
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.Black
            };

            this.menuItemR26.Text = "R26 Daily Queue Management";
            this.menuItemR26.Click += MenuItemR26_Click;

            this.menuItemReports.Text = "Sam's Delivery Report Status";
            this.menuItemReports.Click += MenuItemReports_Click;

            this.menuItemProduction.Text = "Production";
            this.menuItemProduction.Checked = true;
            this.menuItemProduction.Click += MenuItemProduction_Click;

            menuItemMain.DropDownItems.Add(this.menuItemR26);
            menuItemMain.DropDownItems.Add(this.menuItemReports);
            menuItemMain.DropDownItems.Add(this.menuItemProduction);

            this.mainMenuStrip.Items.Add(menuItemMain);

            this.menuItemFileType = new ToolStripMenuItem("Raw File")  // Display current file name
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.Black,
                Margin = new Padding(20, 0, 0, 0)
            };

            this.menuItemRawFile = new ToolStripMenuItem();
            this.menuItemRawFile.Text = "Raw File";
            this.menuItemRawFile.Checked = true;
            this.menuItemRawFile.Click += MenuItemRawFile_Click;

            this.menuItemBulkFile = new ToolStripMenuItem();
            this.menuItemBulkFile.Text = "Bulk File";
            this.menuItemBulkFile.Click += MenuItemBulkFile_Click;

            this.menuItemProcessedFile = new ToolStripMenuItem();
            this.menuItemProcessedFile.Text = "Processed File";
            this.menuItemProcessedFile.Click += MenuItemProcessedFile_Click;

            this.menuItemFileType.DropDownItems.Add(this.menuItemRawFile);
            this.menuItemFileType.DropDownItems.Add(this.menuItemBulkFile);
            this.menuItemFileType.DropDownItems.Add(this.menuItemProcessedFile);

            this.mainMenuStrip.Items.Add(this.menuItemFileType);

            // Title
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new Font("Segoe UI", 18F, FontStyle.Bold);
            this.lblTitle.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblTitle.Location = new Point(13, 45);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new Size(500, 32);
            this.lblTitle.Text = "RPA Schedule Queue Detail Management";

            // Record count
            this.lblRecordCount.AutoSize = true;
            this.lblRecordCount.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblRecordCount.ForeColor = Color.Black;
            this.lblRecordCount.Location = new Point(15, 100);
            this.lblRecordCount.Name = "lblRecordCount";
            this.lblRecordCount.Text = "Total Records: 0";

            this.lblSelectedCount.AutoSize = true;
            this.lblSelectedCount.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblSelectedCount.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblSelectedCount.Location = new Point(250, 97);
            this.lblSelectedCount.Name = "lblSelectedCount";
            this.lblSelectedCount.Text = "Selected: 0";

            this.lblIndividualChanges.AutoSize = true;
            this.lblIndividualChanges.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblIndividualChanges.ForeColor = Color.FromArgb(255, 140, 0);
            this.lblIndividualChanges.Location = new Point(450, 95);
            this.lblIndividualChanges.Name = "lblIndividualChanges";
            this.lblIndividualChanges.Text = "Individual Changes: 0";

            // Search
            this.lblSearch.AutoSize = true;
            this.lblSearch.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            this.lblSearch.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblSearch.Location = new Point(680, 95);
            this.lblSearch.Name = "lblSearch";
            this.lblSearch.Text = "Search";

            this.txtSearch.Font = new Font("Segoe UI", 11F);
            this.txtSearch.Location = new Point(760, 90);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new Size(300, 27);
            this.txtSearch.BorderStyle = BorderStyle.FixedSingle;
            this.txtSearch.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtSearch.TextChanged += TxtSearch_TextChanged;

            // Loading panel + progress bar
            this.progressBar = new ProgressBar();
            this.progressBar.Style = ProgressBarStyle.Marquee;
            this.progressBar.Size = new Size(200, 30);
            this.progressBar.Visible = false;

            this.loadingPanel = new Panel();
            this.loadingPanel.BackColor = Color.White;
            this.loadingPanel.Size = new Size(250, 100);
            this.loadingPanel.Visible = false;
            this.loadingPanel.Name = "loadingPanel";

            var lblLoading = new Label();
            lblLoading.Text = "LOADING...";
            lblLoading.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            lblLoading.ForeColor = Color.FromArgb(52, 73, 94);
            lblLoading.AutoSize = true;
            lblLoading.Location = new Point(75, 55);

            this.progressBar.Location = new Point(25, 20);
            this.loadingPanel.Controls.Add(this.progressBar);
            this.loadingPanel.Controls.Add(lblLoading);
            this.Controls.Add(this.loadingPanel);

            // Filters
            this.lblCompanyFilter.AutoSize = true;
            this.lblCompanyFilter.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            this.lblCompanyFilter.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblCompanyFilter.Location = new Point(15, 148);
            this.lblCompanyFilter.Name = "lblCompanyFilter";
            this.lblCompanyFilter.Text = "Company Name";

            this.cmbCompanyFilter.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbCompanyFilter.Font = new Font("Segoe UI", 9.5F);
            this.cmbCompanyFilter.Location = new Point(165, 145);
            this.cmbCompanyFilter.Size = new Size(260, 28);
            this.cmbCompanyFilter.Name = "cmbCompanyFilter";
            this.cmbCompanyFilter.SelectedIndexChanged += FilterChanged;

            this.lblStatusFilter.AutoSize = true;
            this.lblStatusFilter.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            this.lblStatusFilter.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblStatusFilter.Location = new Point(450, 149);
            this.lblStatusFilter.Name = "lblStatusFilter";
            this.lblStatusFilter.Text = "Status";

            this.cmbStatusFilter.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbStatusFilter.Font = new Font("Segoe UI", 9.5F);
            this.cmbStatusFilter.Location = new Point(520, 145);
            this.cmbStatusFilter.Size = new Size(180, 28);
            this.cmbStatusFilter.Name = "cmbStatusFilter";
            this.cmbStatusFilter.SelectedIndexChanged += FilterChanged;

            // In the InitializeComponent method, find the panelSide section and replace with:

            // DataGridView - EXTENDED WIDTH
            this.dgvQueue.AllowUserToAddRows = false;
            this.dgvQueue.AllowUserToDeleteRows = false;
            this.dgvQueue.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            this.dgvQueue.BackgroundColor = Color.White;
            this.dgvQueue.BorderStyle = BorderStyle.Fixed3D;
            this.dgvQueue.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvQueue.ColumnHeadersHeight = 45;
            this.dgvQueue.EnableHeadersVisualStyles = false;
            this.dgvQueue.Location = new Point(12, 195);
            this.dgvQueue.Name = "dgvQueue";
            this.dgvQueue.RowHeadersWidth = 30;
            this.dgvQueue.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.dgvQueue.Size = new Size(1170, 550);  // Changed from 1050 to 1170 (added 120 pixels)
            this.dgvQueue.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.dgvQueue.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 58, 85);
            this.dgvQueue.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            this.dgvQueue.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.dgvQueue.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.dgvQueue.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(30, 58, 85);
            this.dgvQueue.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.White;
            this.dgvQueue.DefaultCellStyle.SelectionBackColor = Color.FromArgb(173, 216, 230);
            this.dgvQueue.DefaultCellStyle.SelectionForeColor = Color.Black;
            this.dgvQueue.VirtualMode = false;
            this.dgvQueue.AllowUserToResizeRows = false;
            this.dgvQueue.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            this.dgvQueue.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            this.dgvQueue.DoubleBuffered(true);

            // Right side panel - MOVED TO RIGHT EDGE
            this.panelSide.BackColor = Color.FromArgb(240, 240, 240);
            this.panelSide.BorderStyle = BorderStyle.FixedSingle;
            this.panelSide.Location = new Point(1192, 195);  // Moved further right
            this.panelSide.Name = "panelSide";
            this.panelSide.Size = new Size(380, 550);
            this.panelSide.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;

            this.lblSideTitle.AutoSize = true;
            this.lblSideTitle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            this.lblSideTitle.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblSideTitle.Location = new Point(10, 6);
            this.lblSideTitle.Name = "lblSideTitle";
            this.lblSideTitle.Text = "Bulk Status Update";

            this.dgvSelectedRecords.AllowUserToAddRows = false;
            this.dgvSelectedRecords.AllowUserToDeleteRows = false;
            this.dgvSelectedRecords.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;  // Changed from AllCells to Fill
            this.dgvSelectedRecords.BackgroundColor = Color.White;
            this.dgvSelectedRecords.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSelectedRecords.Location = new Point(10, 40);
            this.dgvSelectedRecords.Name = "dgvSelectedRecords";
            this.dgvSelectedRecords.ReadOnly = true;
            this.dgvSelectedRecords.RowHeadersWidth = 30;
            this.dgvSelectedRecords.Size = new Size(358, 260);  // Changed from 478 to 358
            this.dgvSelectedRecords.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.dgvSelectedRecords.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;  // Auto-size rows
            this.dgvSelectedRecords.DefaultCellStyle.WrapMode = DataGridViewTriState.True;  // Enable text wrapping
            this.dgvSelectedRecords.RowTemplate.Height = 25;  // Set minimum row height
            this.dgvSelectedRecords.ScrollBars = ScrollBars.Vertical;  // Only vertical scrollbar
            this.dgvSelectedRecords.RowHeadersVisible = false;  // Hide row headers to save space

            this.lblBulkStatus.AutoSize = true;
            this.lblBulkStatus.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblBulkStatus.Location = new Point(10, 310);
            this.lblBulkStatus.Name = "lblBulkStatus";
            this.lblBulkStatus.Text = "Select New Status";
            this.lblBulkStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            this.cmbBulkStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbBulkStatus.Font = new Font("Segoe UI", 11F);
            this.cmbBulkStatus.FormattingEnabled = true;
            this.cmbBulkStatus.Location = new Point(10, 340);
            this.cmbBulkStatus.Name = "cmbBulkStatus";
            this.cmbBulkStatus.Size = new Size(358, 28);  // Changed from 478 to 358
            this.cmbBulkStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            this.btnBulkUpdate.BackColor = Color.FromArgb(0, 120, 212);
            this.btnBulkUpdate.FlatStyle = FlatStyle.Flat;
            this.btnBulkUpdate.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.btnBulkUpdate.ForeColor = Color.White;
            this.btnBulkUpdate.Location = new Point(10, 387);
            this.btnBulkUpdate.Name = "btnBulkUpdate";
            this.btnBulkUpdate.Size = new Size(358, 43);  // Changed from 478 to 358
            this.btnBulkUpdate.Text = "Update Selected (Bulk)";
            this.btnBulkUpdate.UseVisualStyleBackColor = false;
            this.btnBulkUpdate.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.btnBulkUpdate.Click += BtnBulkUpdate_Click;

            this.btnSaveIndividual.BackColor = Color.FromArgb(255, 140, 0);
            this.btnSaveIndividual.FlatStyle = FlatStyle.Flat;
            this.btnSaveIndividual.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.btnSaveIndividual.ForeColor = Color.White;
            this.btnSaveIndividual.Location = new Point(10, 437);
            this.btnSaveIndividual.Name = "btnSaveIndividual";
            this.btnSaveIndividual.Size = new Size(358, 43);  // Changed from 478 to 358
            this.btnSaveIndividual.Text = "Save Individual Changes";
            this.btnSaveIndividual.UseVisualStyleBackColor = false;
            this.btnSaveIndividual.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.btnSaveIndividual.Click += BtnSaveIndividual_Click;

            this.btnRefresh.BackColor = Color.FromArgb(76, 175, 80);
            this.btnRefresh.FlatStyle = FlatStyle.Flat;
            this.btnRefresh.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.btnRefresh.ForeColor = Color.White;
            this.btnRefresh.Location = new Point(10, 487);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new Size(358, 43);  // Changed from 478 to 358
            this.btnRefresh.Text = "Refresh Data";
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.btnRefresh.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.btnRefresh.Click += BtnRefresh_Click;

            this.panelSide.Controls.Add(this.lblSideTitle);
            this.panelSide.Controls.Add(this.dgvSelectedRecords);
            this.panelSide.Controls.Add(this.lblBulkStatus);
            this.panelSide.Controls.Add(this.cmbBulkStatus);
            this.panelSide.Controls.Add(this.btnBulkUpdate);
            this.panelSide.Controls.Add(this.btnSaveIndividual);
            this.panelSide.Controls.Add(this.btnRefresh);

            // Form
            this.AutoScaleDimensions = new SizeF(8F, 20F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1584, 761);
            this.Controls.Add(this.mainMenuStrip);
            this.Controls.Add(this.txtSearch);
            this.Controls.Add(this.lblSearch);
            this.Controls.Add(this.lblCompanyFilter);
            this.Controls.Add(this.cmbCompanyFilter);
            this.Controls.Add(this.lblStatusFilter);
            this.Controls.Add(this.cmbStatusFilter);
            this.Controls.Add(this.panelSide);
            this.Controls.Add(this.lblIndividualChanges);
            this.Controls.Add(this.lblSelectedCount);
            this.Controls.Add(this.lblRecordCount);
            this.Controls.Add(this.dgvQueue);
            this.Controls.Add(this.lblTitle);
            this.MainMenuStrip = this.mainMenuStrip;
            this.MinimumSize = new Size(1400, 700);
            this.Name = "RpaScheduleQueueDetailForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "RPA Schedule Queue Detail - Status Management";
            this.WindowState = FormWindowState.Maximized;

            ((ISupportInitialize)(this.dgvQueue)).EndInit();
            ((ISupportInitialize)(this.dgvSelectedRecords)).EndInit();
            this.panelSide.ResumeLayout(false);
            this.panelSide.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private async void RpaScheduleQueueDetailForm_Load(object sender, EventArgs e)
        {
            await LoadQueueDataAsync(isInitialLoad: true);
        }

        private async System.Threading.Tasks.Task LoadQueueDataAsync(bool isInitialLoad = false)
        {
            try
            {
                ShowLoadingIndicator(true);
                btnRefresh.Enabled = false;
                btnBulkUpdate.Enabled = false;
                btnSaveIndividual.Enabled = false;

                dgvQueue.SuspendLayout();
                this.SuspendLayout();
                dgvQueue.DataSource = null;

                await System.Threading.Tasks.Task.Run(() =>
                {
                    _queueData = _repository.GetAllRpaScheduleQueueDetail();
                });

                _dataTable = new DataTable();

                _dataTable.Columns.Add("Select", typeof(bool));
                _dataTable.Columns.Add("RpaScheduleQueueDetailUid", typeof(int));
                _dataTable.Columns.Add("CompanyName", typeof(string));
                // REMOVED: _dataTable.Columns.Add("RawReportUid", typeof(int));
                _dataTable.Columns.Add("RawReportName", typeof(string));
                _dataTable.Columns.Add("RawFilePath", typeof(string));
                _dataTable.Columns.Add("Status", typeof(string));
                _dataTable.Columns.Add("BotStartTime", typeof(DateTime));
                _dataTable.Columns.Add("BotEndTime", typeof(DateTime));
                _dataTable.Columns.Add("Duration", typeof(string));
                _dataTable.Columns.Add("CreatedDate", typeof(DateTime));
                _dataTable.Columns.Add("CreatedBy", typeof(string));
                _dataTable.Columns.Add("ModifiedDate", typeof(DateTime));
                _dataTable.Columns.Add("ModifiedBy", typeof(string));

                _dataTable.BeginLoadData();
                foreach (var queue in _queueData)
                {
                    _dataTable.Rows.Add(
                        false,
                        queue.RpaScheduleQueueDetailUid,
                        queue.CompanyName ?? "",
                        // REMOVED: queue.RawReportUid ?? (object)DBNull.Value,
                        queue.RawReportName ?? "",
                        queue.RawFilePath ?? "",
                        queue.Status ?? "Pending",
                        queue.BotStartTime ?? (object)DBNull.Value,
                        queue.BotEndTime ?? (object)DBNull.Value,
                        queue.Duration ?? "",
                        queue.CreatedDate ?? (object)DBNull.Value,
                        queue.CreatedBy ?? "",
                        queue.ModifiedDate ?? (object)DBNull.Value,
                        queue.ModifiedBy ?? ""
                    );
                }
                _dataTable.EndLoadData();

                _bindingSource.DataSource = _dataTable;

                if (isInitialLoad)
                {
                    LoadCompanyFilter();
                    LoadStatusListsFromData();
                    cmbStatusFilter.SelectedIndex = 0;
                    txtSearch.Text = "";
                }

                dgvQueue.DataSource = _bindingSource;
                FormatDataGridView();

                ApplyFilters();
                lblIndividualChanges.Text = $"Individual Changes: {_individualStatusChanges.Count}";
                UpdateSelectedRecordsGrid();

                dgvQueue.ResumeLayout();
                this.ResumeLayout();

                ShowLoadingIndicator(false);
                btnRefresh.Enabled = true;
                btnBulkUpdate.Enabled = true;
                btnSaveIndividual.Enabled = true;

                MessageBox.Show($"Loaded {_bindingSource.Count:N0} records successfully!",
                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                dgvQueue.ResumeLayout();
                this.ResumeLayout();
                ShowLoadingIndicator(false);

                MessageBox.Show($"Error loading queue data: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void MenuItemRawFile_Click(object sender, EventArgs e)
        {
            // Update checked state
            menuItemRawFile.Checked = true;
            menuItemBulkFile.Checked = false;
            menuItemProcessedFile.Checked = false;

            // Update the parent menu text to show current file
            menuItemFileType.Text = "Raw File";

            // Already on Raw File form, no action needed
            this.Text = "RPA Schedule Queue Detail - Status Management (Raw File)";
        }

        private void MenuItemBulkFile_Click(object sender, EventArgs e)
        {
            menuItemRawFile.Checked = false;
            menuItemBulkFile.Checked = true;
            menuItemProcessedFile.Checked = false;
            menuItemFileType.Text = "Bulk File";

            // ✅ ALWAYS check Application.OpenForms FIRST
            var existingBulkForm = Application.OpenForms.OfType<SDReportProcessScheduleQueueForm>()
                .FirstOrDefault();

            if (existingBulkForm != null && !existingBulkForm.IsDisposed)
            {
                existingBulkForm.Show();
                existingBulkForm.BringToFront();
                existingBulkForm.Focus();
                this.Hide();
                return;
            }

            // ✅ Create new Bulk File form ONLY if none exists
            var sdReportForm = new SDReportProcessScheduleQueueForm(
                _emailConfig,
                _r26QueueForm,
                this,
                _samsDeliveryForm
            );

            sdReportForm.FormClosed += (s, args) =>
            {
                if (!this.IsDisposed && !_isExiting)
                {
                    this.Show();
                    this.BringToFront();
                    menuItemRawFile.Checked = true;
                    menuItemBulkFile.Checked = false;
                    menuItemProcessedFile.Checked = false;
                    menuItemFileType.Text = "Raw File";
                    this.Text = "RPA Schedule Queue Detail - Status Management (Raw File)";
                }
            };

            sdReportForm.Show();
            this.Hide();
        }

        private void MenuItemProcessedFile_Click(object sender, EventArgs e)
        {
            menuItemRawFile.Checked = false;
            menuItemBulkFile.Checked = false;
            menuItemProcessedFile.Checked = true;
            menuItemFileType.Text = "Processed File";

            // ✅ ALWAYS check Application.OpenForms FIRST
            var existingProcessedForm = Application.OpenForms.OfType<ReportProcessScheduleQueueForm>()
                .FirstOrDefault();

            if (existingProcessedForm != null && !existingProcessedForm.IsDisposed)
            {
                existingProcessedForm.Show();
                existingProcessedForm.BringToFront();
                existingProcessedForm.Focus();
                this.Hide();
                return;
            }

            // ✅ Create new Processed File form ONLY if none exists
            var processedForm = new ReportProcessScheduleQueueForm(
                _emailConfig,
                _r26QueueForm,
                this,
                _samsDeliveryForm,
                null
            );

            processedForm.FormClosed += (s, args) =>
            {
                if (!this.IsDisposed && !_isExiting)
                {
                    this.Show();
                    this.BringToFront();
                    menuItemRawFile.Checked = true;
                    menuItemBulkFile.Checked = false;
                    menuItemProcessedFile.Checked = false;
                    menuItemFileType.Text = "Raw File";
                    this.Text = "RPA Schedule Queue Detail - Status Management (Raw File)";
                }
            };

            processedForm.Show();
            this.Hide();
        }

        // Update OnVisibleChanged to reset the menu text and filters:
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);

            if (this.Visible)
            {
                // Reset menu selection to Raw File when this form becomes visible
                menuItemRawFile.Checked = true;
                menuItemBulkFile.Checked = false;
                menuItemProcessedFile.Checked = false;
                menuItemFileType.Text = "Raw File"; // Update parent menu text
                this.Text = "RPA Schedule Queue Detail - Status Management (Raw File)";

                // Reset filters to initial state
                if (cmbCompanyFilter != null && cmbCompanyFilter.Items.Count > 0)
                {
                    cmbCompanyFilter.SelectedIndex = 0; // "All Companies"
                }

                if (cmbStatusFilter != null && cmbStatusFilter.Items.Count > 0)
                {
                    cmbStatusFilter.SelectedIndex = 0; // "All Status"
                }

                if (txtSearch != null)
                {
                    txtSearch.Text = ""; // Clear search text
                }

                // Apply the reset filters
                ApplyFilters();
            }
        }


        private void ShowLoadingIndicator(bool show)
        {
            if (loadingPanel == null || dgvQueue == null)
                return;

            if (show)
            {
                int x = dgvQueue.Location.X + (dgvQueue.Width - loadingPanel.Width) / 2;
                int y = dgvQueue.Location.Y + (dgvQueue.Height - loadingPanel.Height) / 2;
                loadingPanel.Location = new Point(x, y);
                loadingPanel.Visible = true;
                progressBar.Visible = true;
                loadingPanel.BringToFront();
            }
            else
            {
                loadingPanel.Visible = false;
                progressBar.Visible = false;
            }
        }

        private void LoadCompanyFilter()
        {
            cmbCompanyFilter.BeginUpdate();
            cmbCompanyFilter.Items.Clear();
            cmbCompanyFilter.Items.Add("All Companies");

            if (_dataTable != null)
            {
                var companies = _dataTable.AsEnumerable()
                    .Select(r => r.Field<string>("CompanyName"))
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();

                foreach (var company in companies)
                    cmbCompanyFilter.Items.Add(company);
            }

            cmbCompanyFilter.SelectedIndex = 0;
            cmbCompanyFilter.EndUpdate();
        }

        private void LoadStatusListsFromData()
        {
            // Status filter - only DB statuses, no hardcoded values
            cmbStatusFilter.BeginUpdate();
            cmbStatusFilter.Items.Clear();
            cmbStatusFilter.Items.Add("All Status");

            foreach (var s in _statusList)
                cmbStatusFilter.Items.Add(s);

            cmbStatusFilter.SelectedIndex = 0;
            cmbStatusFilter.EndUpdate();

            // Bulk status dropdown - includes all DB statuses
            cmbBulkStatus.BeginUpdate();
            cmbBulkStatus.Items.Clear();

            foreach (var s in _statusList)
                cmbBulkStatus.Items.Add(s);

            if (cmbBulkStatus.Items.Count > 0)
                cmbBulkStatus.SelectedIndex = 0;
            cmbBulkStatus.EndUpdate();
        }

        private void ConfigureStatusColumn()
        {
            if (!dgvQueue.Columns.Contains("Status"))
                return;

            int statusIndex = dgvQueue.Columns["Status"].Index;
            dgvQueue.Columns.RemoveAt(statusIndex);

            var statusColumn = new DataGridViewComboBoxColumn
            {
                HeaderText = "Status",
                Name = "Status",
                DataPropertyName = "Status",
                Width = 150,
                MinimumWidth = 150,
                FlatStyle = FlatStyle.Flat,
                DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
                DisplayStyleForCurrentCellOnly = false,
                ReadOnly = false
            };

            if (_statusList != null && _statusList.Count > 0)
            {
                foreach (var s in _statusList)
                    statusColumn.Items.Add(s);
            }

            dgvQueue.Columns.Insert(statusIndex, statusColumn);
        }

        private void FormatDataGridView()
        {
            if (!dgvQueue.Columns.Contains("Select"))
                return;

            dgvQueue.Columns["Select"].Width = 60;
            dgvQueue.Columns["Select"].HeaderText = "Select";
            dgvQueue.Columns["Select"].ReadOnly = false;
            dgvQueue.Columns["Select"].Frozen = true;

            if (dgvQueue.Columns.Contains("RpaScheduleQueueDetailUid"))
                dgvQueue.Columns["RpaScheduleQueueDetailUid"].Visible = false;

            dgvQueue.Columns["CompanyName"].HeaderText = "Company Name";
            dgvQueue.Columns["CompanyName"].Width = 200;
            dgvQueue.Columns["CompanyName"].Frozen = true;

            // REMOVED: Raw Report UID column formatting

            dgvQueue.Columns["RawReportName"].HeaderText = "Report Name";
            dgvQueue.Columns["RawReportName"].Width = 200;

            dgvQueue.Columns["RawFilePath"].HeaderText = "Raw File Path";
            dgvQueue.Columns["RawFilePath"].Width = 250;

            dgvQueue.Columns["Status"].HeaderText = "Status";
            dgvQueue.Columns["Status"].Width = 150;

            dgvQueue.Columns["BotStartTime"].HeaderText = "Bot Start Time";
            dgvQueue.Columns["BotStartTime"].Width = 170;
            dgvQueue.Columns["BotStartTime"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

            dgvQueue.Columns["BotEndTime"].HeaderText = "Bot End Time";
            dgvQueue.Columns["BotEndTime"].Width = 170;
            dgvQueue.Columns["BotEndTime"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

            dgvQueue.Columns["Duration"].HeaderText = "Duration";
            dgvQueue.Columns["Duration"].Width = 100;

            dgvQueue.Columns["CreatedDate"].HeaderText = "Created Date";
            dgvQueue.Columns["CreatedDate"].Width = 170;
            dgvQueue.Columns["CreatedDate"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

            dgvQueue.Columns["CreatedBy"].HeaderText = "Created By";
            dgvQueue.Columns["CreatedBy"].Width = 130;

            dgvQueue.Columns["ModifiedDate"].HeaderText = "Modified Date";
            dgvQueue.Columns["ModifiedDate"].Width = 170;
            dgvQueue.Columns["ModifiedDate"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

            dgvQueue.Columns["ModifiedBy"].HeaderText = "Modified By";
            dgvQueue.Columns["ModifiedBy"].Width = 130;

            ConfigureStatusColumn();

            dgvQueue.ColumnHeaderMouseClick -= DgvQueue_ColumnHeaderMouseClick;
            dgvQueue.ColumnHeaderMouseClick += DgvQueue_ColumnHeaderMouseClick;

            dgvQueue.CellContentClick -= DgvQueue_CellContentClick;
            dgvQueue.CellContentClick += DgvQueue_CellContentClick;

            dgvQueue.CellValueChanged -= DgvQueue_CellValueChanged;
            dgvQueue.CellValueChanged += DgvQueue_CellValueChanged;

            dgvQueue.CurrentCellDirtyStateChanged -= DgvQueue_CurrentCellDirtyStateChanged;
            dgvQueue.CurrentCellDirtyStateChanged += DgvQueue_CurrentCellDirtyStateChanged;

            dgvQueue.RowPrePaint -= DgvQueue_RowPrePaint;
            dgvQueue.RowPrePaint += DgvQueue_RowPrePaint;

            dgvQueue.EditingControlShowing -= DgvQueue_EditingControlShowing;
            dgvQueue.EditingControlShowing += DgvQueue_EditingControlShowing;

            dgvQueue.DataError -= DgvQueue_DataError;
            dgvQueue.DataError += DgvQueue_DataError;

            foreach (DataGridViewColumn col in dgvQueue.Columns)
            {
                if (col.Name != "Select" && col.Name != "Status")
                    col.ReadOnly = true;
            }
        }

        private void DgvQueue_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgvQueue.CurrentCell != null &&
                dgvQueue.CurrentCell.OwningColumn.Name == "Status")
            {
                ComboBox combo = e.Control as ComboBox;
                if (combo != null)
                {
                    combo.DropDownStyle = ComboBoxStyle.DropDownList;
                    combo.SelectedIndexChanged -= StatusCombo_SelectedIndexChanged;
                    combo.SelectedIndexChanged += StatusCombo_SelectedIndexChanged;
                }
            }
        }

        private void StatusCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dgvQueue.IsCurrentCellDirty)
                dgvQueue.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void DgvQueue_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (dgvQueue.Columns[e.ColumnIndex].Name == "Status")
            {
                e.ThrowException = false;
                e.Cancel = false;
            }
        }

        private void DgvQueue_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvQueue.IsCurrentCellDirty)
                dgvQueue.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void DgvQueue_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            if (dgvQueue.Columns[e.ColumnIndex].Name == "Select")
            {
                dgvQueue.EndEdit();
                UpdateSelectedRecordsGrid();
            }
        }

        private void DgvQueue_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            string columnName = dgvQueue.Columns[e.ColumnIndex].Name;

            if (columnName == "Status")
            {
                var row = dgvQueue.Rows[e.RowIndex];
                int queueId = Convert.ToInt32(row.Cells["RpaScheduleQueueDetailUid"].Value);
                string newStatusStr = row.Cells["Status"].Value?.ToString() ?? "Pending";

                var originalRecord = _queueData.FirstOrDefault(q => q.RpaScheduleQueueDetailUid == queueId);
                string originalStatus = originalRecord?.Status ?? "Pending";

                if (newStatusStr != originalStatus)
                {
                    _individualStatusChanges[queueId] = newStatusStr;
                }
                else
                {
                    _individualStatusChanges.Remove(queueId);
                }

                if (_dataTable != null)
                {
                    var dataRow = _dataTable.AsEnumerable()
                        .FirstOrDefault(r => r.Field<int>("RpaScheduleQueueDetailUid") == queueId);
                    if (dataRow != null)
                        dataRow["Status"] = newStatusStr;
                }

                lblIndividualChanges.Text = $"Individual Changes: {_individualStatusChanges.Count}";
                RefreshGridHighlighting();
            }
            else if (columnName == "Select")
            {
                UpdateSelectedRecordsGrid();
            }
        }

        private void DgvQueue_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            var row = dgvQueue.Rows[e.RowIndex];
            object idValue = row.Cells["RpaScheduleQueueDetailUid"].Value;
            if (idValue == null)
                return;

            if (int.TryParse(idValue.ToString(), out int queueId))
            {
                if (_individualStatusChanges.ContainsKey(queueId))
                    row.DefaultCellStyle.BackColor = Color.LightYellow;
                else
                    row.DefaultCellStyle.BackColor = Color.White;
            }
        }

        private void RefreshGridHighlighting()
        {
            foreach (DataGridViewRow row in dgvQueue.Rows)
            {
                if (row.Cells["RpaScheduleQueueDetailUid"].Value != null &&
                    int.TryParse(row.Cells["RpaScheduleQueueDetailUid"].Value.ToString(), out int queueId))
                {
                    if (_individualStatusChanges.ContainsKey(queueId))
                        row.DefaultCellStyle.BackColor = Color.LightYellow;
                    else
                        row.DefaultCellStyle.BackColor = Color.White;
                }
            }
        }

        private void DgvQueue_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == 0 &&
                dgvQueue.Columns[e.ColumnIndex].Name == "Select")
            {
                return;
            }
        }

        private void FilterChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            if (_dataTable == null || _bindingSource == null)
                return;

            var filters = new List<string>();

            if (cmbCompanyFilter.SelectedIndex > 0)
            {
                string companyName = cmbCompanyFilter.SelectedItem.ToString().Replace("'", "''");
                filters.Add($"CompanyName = '{companyName}'");
            }

            if (cmbStatusFilter.SelectedIndex > 0)
            {
                string status = cmbStatusFilter.SelectedItem.ToString().Replace("'", "''");
                filters.Add($"Status = '{status}'");
            }

            string searchText = txtSearch.Text.Trim().Replace("'", "''");
            if (!string.IsNullOrEmpty(searchText))
            {
                filters.Add(
                    $"(Convert(RpaScheduleQueueDetailUid, System.String) LIKE '%{searchText}%' " +
                    $"OR CompanyName LIKE '%{searchText}%' " +
                    $"OR RawReportName LIKE '%{searchText}%')");
            }

            string finalFilter = filters.Count == 0 ? null : string.Join(" AND ", filters);
            _bindingSource.Filter = finalFilter;

            int displayedCount = _bindingSource.Count;
            lblRecordCount.Text = $"Total Records: {displayedCount:N0}";
            dgvQueue.Refresh();
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void UpdateSelectedRecordsGrid()
        {
            if (_dataTable == null)
                return;

            var selectedRows = _dataTable.AsEnumerable()
                .Where(r => r.Field<bool>("Select"))
                .Select(r => new
                {
                    QueueDetailID = r.Field<int>("RpaScheduleQueueDetailUid"),
                    CompanyName = r.Field<string>("CompanyName"),
                    ReportName = r.Field<string>("RawReportName"),
                    Status = r.Field<string>("Status")
                })
                .ToList();

            var selectedTable = new DataTable();
            selectedTable.Columns.Add("Company Name", typeof(string));
            selectedTable.Columns.Add("Report Name", typeof(string));
            selectedTable.Columns.Add("Status", typeof(string));

            foreach (var item in selectedRows)
            {
                selectedTable.Rows.Add(item.CompanyName, item.ReportName, item.Status);
            }

            dgvSelectedRecords.DataSource = selectedTable;
            lblSelectedCount.Text = $"Selected: {selectedRows.Count:N0}";
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            _ = LoadQueueDataAsync(isInitialLoad: false);
        }

        private void BtnSaveIndividual_Click(object sender, EventArgs e)
        {
            if (_individualStatusChanges.Count == 0)
            {
                MessageBox.Show("No individual status changes to save.",
                    "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show(
                $"Are you sure you want to save {_individualStatusChanges.Count} individual status changes?",
                "Confirm Individual Updates",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    ShowLoadingIndicator(true);
                    btnSaveIndividual.Enabled = false;

                    int updatedCount = _repository.UpdateStatuses(_individualStatusChanges, _systemName);

                    ShowLoadingIndicator(false);
                    btnSaveIndividual.Enabled = true;

                    MessageBox.Show($"Successfully updated {updatedCount} individual records!",
                        "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    _individualStatusChanges.Clear();
                    lblIndividualChanges.Text = "Individual Changes: 0";

                    _ = LoadQueueDataAsync(isInitialLoad: false);
                }
                catch (Exception ex)
                {
                    ShowLoadingIndicator(false);
                    btnSaveIndividual.Enabled = true;

                    MessageBox.Show($"Error updating individual statuses: {ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnBulkUpdate_Click(object sender, EventArgs e)
        {
            var selectedRows = dgvQueue.Rows
                .Cast<DataGridViewRow>()
                .Where(r => r.Cells["Select"].Value != null && (bool)r.Cells["Select"].Value == true)
                .ToList();

            if (selectedRows.Count == 0)
            {
                MessageBox.Show("Please select at least one record to update.",
                    "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (cmbBulkStatus.SelectedItem == null)
            {
                MessageBox.Show("Please select a status from the dropdown.",
                    "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                string newStatus = cmbBulkStatus.SelectedItem.ToString();

                var result = MessageBox.Show(
                    $"Are you sure you want to update {selectedRows.Count} records to status '{newStatus}'?",
                    "Confirm Bulk Update",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        ShowLoadingIndicator(true);
                        btnBulkUpdate.Enabled = false;

                        var updates = new Dictionary<int, string>();

                        foreach (var row in selectedRows)
                        {
                            int queueId = Convert.ToInt32(row.Cells["RpaScheduleQueueDetailUid"].Value);
                            updates[queueId] = newStatus;
                        }

                        int updatedCount = _repository.UpdateStatuses(updates, _systemName);

                        ShowLoadingIndicator(false);
                        btnBulkUpdate.Enabled = true;

                        MessageBox.Show($"Successfully updated {updatedCount} records via bulk update!",
                            "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        foreach (DataGridViewRow row in dgvQueue.Rows)
                        {
                            if (row.Cells["Select"].Value != null)
                                row.Cells["Select"].Value = false;
                        }

                        if (_dataTable != null)
                        {
                            foreach (DataRow dataRow in _dataTable.Rows)
                            {
                                dataRow["Select"] = false;
                            }
                        }

                        UpdateSelectedRecordsGrid();
                        _ = LoadQueueDataAsync(isInitialLoad: false);
                    }
                    catch (Exception ex)
                    {
                        ShowLoadingIndicator(false);
                        btnBulkUpdate.Enabled = true;

                        MessageBox.Show($"Error updating bulk statuses: {ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void MenuItemR26_Click(object sender, EventArgs e)
        {
            // ✅ ALWAYS check Application.OpenForms FIRST
            var existingR26Form = Application.OpenForms.OfType<R26QueueForm>()
                .FirstOrDefault();

            if (existingR26Form != null && !existingR26Form.IsDisposed)
            {
                // ✅ Update references in existing form
                if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
                {
                    existingR26Form.SetSamsDeliveryForm(_samsDeliveryForm);
                }
                existingR26Form.Show();
                existingR26Form.BringToFront();
                existingR26Form.Focus();
                this.Hide();
                return;
            }

            // ✅ Check stored reference only if not found in OpenForms
            if (_r26QueueForm != null && !_r26QueueForm.IsDisposed)
            {
                if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
                {
                    _r26QueueForm.SetSamsDeliveryForm(_samsDeliveryForm);
                }
                _r26QueueForm.Show();
                _r26QueueForm.BringToFront();
                _r26QueueForm.Focus();
                this.Hide();
                return;
            }

            // ✅ Only create NEW form if none exists
            _r26QueueForm = new R26QueueForm(_emailConfig);
            if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
            {
                _r26QueueForm.SetSamsDeliveryForm(_samsDeliveryForm);
            }

            _r26QueueForm.FormClosed += (s, args) =>
            {
                _r26QueueForm = null;
                if (!this.IsDisposed && !_isExiting)
                {
                    this.Show();
                    this.BringToFront();
                }
            };

            _r26QueueForm.Show();
            this.Hide();
        }

        private void MenuItemReports_Click(object sender, EventArgs e)
        {
            // ✅ ALWAYS check Application.OpenForms FIRST
            var existingSamsForm = Application.OpenForms.OfType<SamsDeliveryReportForm>()
                .FirstOrDefault();

            if (existingSamsForm != null && !existingSamsForm.IsDisposed)
            {
                // ✅ Update references in existing form
                existingSamsForm.SetR26QueueForm(_r26QueueForm);
                existingSamsForm.SetRpaScheduleQueueDetailForm(this);
                existingSamsForm.Show();
                existingSamsForm.BringToFront();
                existingSamsForm.Focus();
                this.Hide();
                return;
            }

            // ✅ Check stored reference only if not found in OpenForms
            if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
            {
                _samsDeliveryForm.SetR26QueueForm(_r26QueueForm);
                _samsDeliveryForm.SetRpaScheduleQueueDetailForm(this);
                _samsDeliveryForm.Show();
                _samsDeliveryForm.BringToFront();
                _samsDeliveryForm.Focus();
                this.Hide();
                return;
            }

            // ✅ Only create NEW form if none exists
            DateTime targetDate = DateTime.Today;
            string connectionString = Program.DbConnectionString;
            string storedProcedureName = "SDgetSAMSReportStatus";

            _samsDeliveryForm = new SamsDeliveryReportForm(
                targetDate,
                _emailConfig,
                connectionString,
                storedProcedureName,
                _r26QueueForm
            );

            _samsDeliveryForm.SetRpaScheduleQueueDetailForm(this);

            _samsDeliveryForm.FormClosed += (s, args) =>
            {
                _samsDeliveryForm = null;
                if (!this.IsDisposed && !_isExiting)
                {
                    this.Show();
                    this.BringToFront();
                }
            };

            _samsDeliveryForm.Show();
            this.Hide();
        }

        private void MenuItemProduction_Click(object sender, EventArgs e)
        {
            // Already in Production menu - show message
            MessageBox.Show("You are already in the Production menu.\n\nUse the 'File Type' dropdown to switch between:\n• Raw File (current)\n• Bulk File\n• Processed File",
                "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RpaScheduleQueueDetailForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // ✅ If already exiting, just return (prevents duplicate dialogs)
            if (_isExiting)
            {
                return;
            }

            // ✅ If closed by Application.Exit() from another form, don't show dialog
            if (e.CloseReason == CloseReason.ApplicationExitCall)
            {
                return;
            }

            // Only show confirmation if user clicked the X button
            if (e.CloseReason == CloseReason.UserClosing)
            {
                // Set flag immediately to prevent re-entry
                _isExiting = true;

                // Count unsaved changes
                int individualChangesCount = _individualStatusChanges?.Count ?? 0;
                int selectedRowsCount = 0;

                if (dgvQueue?.Rows != null)
                {
                    selectedRowsCount = dgvQueue.Rows
                        .Cast<DataGridViewRow>()
                        .Count(r => r.Cells["Select"].Value != null && (bool)r.Cells["Select"].Value == true);
                }

                int totalPendingChanges = individualChangesCount + selectedRowsCount;

                DialogResult result;

                if (totalPendingChanges > 0)
                {
                    string message = "You have unsaved changes:\n\n";

                    if (individualChangesCount > 0)
                    {
                        message += $"• Individual status changes: {individualChangesCount}\n";
                    }

                    if (selectedRowsCount > 0)
                    {
                        message += $"• Selected records for bulk update: {selectedRowsCount}\n";
                    }

                    message += "\nAre you sure you want to exit without saving?";

                    result = MessageBox.Show(
                        message,
                        "Unsaved Changes - Confirm Exit",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                }
                else
                {
                    result = MessageBox.Show(
                        "Are you sure you want to exit?",
                        "Confirm Exit",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                }

                if (result == DialogResult.Yes)
                {
                    // ✅ Set _isExiting flag on other forms BEFORE closing them
                    if (_r26QueueForm != null && !_r26QueueForm.IsDisposed)
                    {
                        _r26QueueForm._isExiting = true;
                        _r26QueueForm.Close();
                    }

                    if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
                    {
                        _samsDeliveryForm._isExiting = true;
                        _samsDeliveryForm.Close();
                    }

                    // Exit the entire application
                    Application.Exit();
                }
                else
                {
                    // User clicked No - reset flag and cancel close
                    _isExiting = false;
                    e.Cancel = true;
                }
            }
        }
    }
}