using Microsoft.Extensions.Configuration;
using R26_DailyQueueWinForm;
using R26_DailyQueueWinForm.Data;
using R26_DailyQueueWinForm.Forms;
using R26_DailyQueueWinForm.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace R26DailyQueueWinForm
{
    public partial class ReportProcessScheduleQueueForm : Form
    {
        private readonly ReportProcessScheduleQueueRepository _repository;
        private List<ReportProcessScheduleQueueModel> _queueData;
        private DataTable _dataTable;
        private Dictionary<int, string> _individualStatusChanges;
        private BindingSource _bindingSource;
        private EmailConfiguration _emailConfig;

        public bool _isExiting = false;

        private R26QueueForm _actualR26QueueForm;           // The REAL R26QueueForm
        private RpaScheduleQueueDetailForm _rawFileForm;     // The Raw File form
        private SamsDeliveryReportForm _samsDeliveryForm;
        private SDReportProcessScheduleQueueForm _bulkFileForm;


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

        public ReportProcessScheduleQueueForm(
    EmailConfiguration emailConfig,
    R26QueueForm r26QueueForm,              // ACTUAL R26QueueForm
    RpaScheduleQueueDetailForm rawFileForm, // Raw File form
    SamsDeliveryReportForm samsDeliveryForm,
    SDReportProcessScheduleQueueForm bulkFileForm)
        {
            InitializeComponent();

            _emailConfig = emailConfig;
            _systemName = GetSystemName();
            _actualR26QueueForm = r26QueueForm;     // Store ACTUAL R26QueueForm
            _rawFileForm = rawFileForm;              // Store Raw File form
            _samsDeliveryForm = samsDeliveryForm;
            _bulkFileForm = bulkFileForm;

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
                MessageBox.Show($"Invalid connection string format: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _repository = new ReportProcessScheduleQueueRepository(connectionString);
            _individualStatusChanges = new Dictionary<int, string>();
            _bindingSource = new BindingSource();

            LoadStatusesFromDatabase();

            this.Load += ReportProcessScheduleQueueForm_Load;
            this.FormClosing += ReportProcessScheduleQueueForm_FormClosing;
        }

        private void LoadStatusesFromDatabase()
        {
            try
            {
                _statusList = _repository.GetAllStatuses();
                if (_statusList == null || _statusList.Count == 0)
                {
                    MessageBox.Show("No statuses found in database.", "Warning",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _statusList = new List<string>();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading statuses: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusList = new List<string>();
            }
        }

        private void InitializeComponent()
        {
            this.dgvQueue = new DataGridView();
            this.btnBulkUpdate = new Button();
            this.btnSaveIndividual = new Button();
            this.btnRefresh = new Button();
            this.mainMenuStrip = new MenuStrip();
            this.menuItemR26 = new ToolStripMenuItem();
            this.menuItemReports = new ToolStripMenuItem();
            this.menuItemProduction = new ToolStripMenuItem();
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
            this.lblCompanyFilter = new Label();
            this.cmbCompanyFilter = new ComboBox();
            this.lblStatusFilter = new Label();
            this.cmbStatusFilter = new ComboBox();

            ((ISupportInitialize)this.dgvQueue).BeginInit();
            ((ISupportInitialize)this.dgvSelectedRecords).BeginInit();
            this.panelSide.SuspendLayout();
            this.mainMenuStrip.SuspendLayout();
            this.SuspendLayout();

            // MenuStrip
            this.mainMenuStrip.BackColor = Color.White;
            this.mainMenuStrip.Dock = DockStyle.Top;
            this.mainMenuStrip.Font = new Font("Segoe UI", 10F);
            this.mainMenuStrip.Height = 35;
            this.mainMenuStrip.GripStyle = ToolStripGripStyle.Hidden;
            this.mainMenuStrip.RenderMode = ToolStripRenderMode.Professional;
            this.mainMenuStrip.Padding = new Padding(10, 5, 0, 5);
            this.mainMenuStrip.Name = "mainMenuStrip";

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

            this.menuItemFileType = new ToolStripMenuItem("Processed File")  // Display current file name
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.Black,
                Margin = new Padding(20, 0, 0, 0)
            };

            this.menuItemRawFile = new ToolStripMenuItem();
            this.menuItemRawFile.Text = "Raw File";
            this.menuItemRawFile.Click += MenuItemRawFile_Click_Processed;

            this.menuItemBulkFile = new ToolStripMenuItem();
            this.menuItemBulkFile.Text = "Bulk File";
            this.menuItemBulkFile.Click += MenuItemBulkFile_Click_Processed;

            this.menuItemProcessedFile = new ToolStripMenuItem();
            this.menuItemProcessedFile.Text = "Processed File";
            this.menuItemProcessedFile.Checked = true;  // This is the Processed File form
            this.menuItemProcessedFile.Click += MenuItemProcessedFile_Click_Processed;

            this.menuItemFileType.DropDownItems.Add(this.menuItemRawFile);
            this.menuItemFileType.DropDownItems.Add(this.menuItemBulkFile);
            this.menuItemFileType.DropDownItems.Add(this.menuItemProcessedFile);

            this.mainMenuStrip.Items.Add(this.menuItemFileType);

            // Title
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new Font("Segoe UI", 18F, FontStyle.Bold);
            this.lblTitle.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblTitle.Location = new Point(15, 50);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new Size(500, 32);
            this.lblTitle.Text = "Report Process Schedule Queue Management";

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
            this.lblSelectedCount.Location = new Point(250, 100);
            this.lblSelectedCount.Name = "lblSelectedCount";
            this.lblSelectedCount.Text = "Selected: 0";

            this.lblIndividualChanges.AutoSize = true;
            this.lblIndividualChanges.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblIndividualChanges.ForeColor = Color.FromArgb(255, 140, 0);
            this.lblIndividualChanges.Location = new Point(450, 100);
            this.lblIndividualChanges.Name = "lblIndividualChanges";
            this.lblIndividualChanges.Text = "Individual Changes: 0";

            // Search
            this.lblSearch.AutoSize = true;
            this.lblSearch.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            this.lblSearch.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblSearch.Location = new Point(680, 100);
            this.lblSearch.Name = "lblSearch";
            this.lblSearch.Text = "Search";

            this.txtSearch.Font = new Font("Segoe UI", 11F);
            this.txtSearch.Location = new Point(760, 95);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new Size(300, 27);
            this.txtSearch.BorderStyle = BorderStyle.FixedSingle;
            this.txtSearch.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtSearch.TextChanged += TxtSearch_TextChanged;

            // Loading panel + progress
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

            // In the InitializeComponent method, find and replace the DataGridView section:

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
            this.dgvQueue.Size = new Size(1170, 550);  // Changed from 1050 to 1170
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

            // Right side panel - REDUCED WIDTH AND MOVED RIGHT
            this.panelSide.BackColor = Color.FromArgb(240, 240, 240);
            this.panelSide.BorderStyle = BorderStyle.FixedSingle;
            this.panelSide.Location = new Point(1192, 195);  // Changed from 1070 to 1192
            this.panelSide.Name = "panelSide";
            this.panelSide.Size = new Size(380, 550);  // Changed from 500 to 380
            this.panelSide.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;

            this.lblSideTitle.AutoSize = true;
            this.lblSideTitle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            this.lblSideTitle.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblSideTitle.Location = new Point(10, 6);
            this.lblSideTitle.Name = "lblSideTitle";
            this.lblSideTitle.Text = "Bulk Status Update";

            this.dgvSelectedRecords.AllowUserToAddRows = false;
            this.dgvSelectedRecords.AllowUserToDeleteRows = false;
            this.dgvSelectedRecords.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;  // Changed from AllCells
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
            this.MainMenuStrip = this.mainMenuStrip;
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
            this.MinimumSize = new Size(1400, 700);
            this.Name = "ReportProcessScheduleQueueForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Report Process Schedule Queue - Status Management (Processed File)";
            this.WindowState = FormWindowState.Maximized;

            ((ISupportInitialize)this.dgvQueue).EndInit();
            ((ISupportInitialize)this.dgvSelectedRecords).EndInit();
            this.panelSide.ResumeLayout(false);
            this.panelSide.PerformLayout();
            this.mainMenuStrip.ResumeLayout(false);
            this.mainMenuStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private async void ReportProcessScheduleQueueForm_Load(object sender, EventArgs e)
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
                    _queueData = _repository.GetAllReportProcessScheduleQueue();
                });

                _dataTable = new DataTable();
                _dataTable.Columns.Add("Select", typeof(bool));
                _dataTable.Columns.Add("ReportScheduleQueueUid", typeof(int));
                _dataTable.Columns.Add("CompanyUid", typeof(int));
                _dataTable.Columns.Add("CompanyName", typeof(string));
                _dataTable.Columns.Add("ReportTypeUid", typeof(int));
                _dataTable.Columns.Add("ReportTypeName", typeof(string));
                _dataTable.Columns.Add("Status", typeof(string));
                _dataTable.Columns.Add("FrequencyUid", typeof(int));
                _dataTable.Columns.Add("ScheduleDate", typeof(DateTime));
                _dataTable.Columns.Add("ScheduleTime", typeof(TimeSpan));
                _dataTable.Columns.Add("ExecutionDuration", typeof(string));
                _dataTable.Columns.Add("RawFilePath", typeof(string));
                _dataTable.Columns.Add("ProcessedFilePath", typeof(string));
                _dataTable.Columns.Add("Timezone", typeof(string));
                _dataTable.Columns.Add("ReportStartTime", typeof(DateTime));
                _dataTable.Columns.Add("ReportEndTime", typeof(DateTime));
                _dataTable.Columns.Add("SLAComplianceFlag", typeof(string));
                _dataTable.Columns.Add("CreatedDate", typeof(DateTime));
                _dataTable.Columns.Add("CreatedBy", typeof(string));
                _dataTable.Columns.Add("ModifiedDate", typeof(DateTime));
                _dataTable.Columns.Add("ModifiedBy", typeof(string));

                _dataTable.BeginLoadData();
                foreach (var queue in _queueData)
                {
                    _dataTable.Rows.Add(
                        false,
                        queue.ReportScheduleQueueUid,
                        queue.CompanyUid ?? (object)DBNull.Value,
                        queue.CompanyName ?? "",
                        queue.ReportTypeUid ?? (object)DBNull.Value,
                        queue.ReportTypeName ?? "",
                        queue.Status ?? "Pending",
                        queue.FrequencyUid ?? (object)DBNull.Value,
                        queue.ScheduleDate ?? (object)DBNull.Value,
                        queue.ScheduleTime ?? (object)DBNull.Value,
                        queue.ExecutionDuration ?? "",
                        queue.RawFilePath ?? "",
                        queue.ProcessedFilePath ?? "",
                        queue.Timezone ?? "",
                        queue.ReportStartTime ?? (object)DBNull.Value,
                        queue.ReportEndTime ?? (object)DBNull.Value,
                        queue.SLAComplianceFlag ?? "",
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
                    txtSearch.Text = string.Empty;
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
                MessageBox.Show($"Error loading queue data: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            cmbStatusFilter.BeginUpdate();
            cmbStatusFilter.Items.Clear();
            cmbStatusFilter.Items.Add("All Status");

            foreach (var s in _statusList)
                cmbStatusFilter.Items.Add(s);

            cmbStatusFilter.SelectedIndex = 0;
            cmbStatusFilter.EndUpdate();

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

            if (dgvQueue.Columns.Contains("ReportScheduleQueueUid"))
                dgvQueue.Columns["ReportScheduleQueueUid"].Visible = false;
            if (dgvQueue.Columns.Contains("CompanyUid"))
                dgvQueue.Columns["CompanyUid"].Visible = false;
            if (dgvQueue.Columns.Contains("ReportTypeUid"))
                dgvQueue.Columns["ReportTypeUid"].Visible = false;
            if (dgvQueue.Columns.Contains("FrequencyUid"))
                dgvQueue.Columns["FrequencyUid"].Visible = false;

            dgvQueue.Columns["ReportTypeName"].HeaderText = "Report TypeName";
            dgvQueue.Columns["ReportTypeName"].Width = 470;
            dgvQueue.Columns["ReportTypeName"].ReadOnly = true;


            dgvQueue.Columns["Status"].HeaderText = "Status";
            dgvQueue.Columns["Status"].Width = 150;

            dgvQueue.Columns["CompanyName"].HeaderText = "Company Name";
            dgvQueue.Columns["CompanyName"].Width = 150;
            dgvQueue.Columns["CompanyName"].Frozen = true;

            dgvQueue.Columns["ScheduleDate"].HeaderText = "Schedule Date";
            dgvQueue.Columns["ScheduleDate"].Width = 120;
            dgvQueue.Columns["ScheduleDate"].DefaultCellStyle.Format = "yyyy-MM-dd";

            dgvQueue.Columns["ScheduleTime"].HeaderText = "Schedule Time";
            dgvQueue.Columns["ScheduleTime"].Width = 120;

            dgvQueue.Columns["ExecutionDuration"].HeaderText = "Execution Duration";
            dgvQueue.Columns["ExecutionDuration"].Width = 140;

            dgvQueue.Columns["RawFilePath"].HeaderText = "Raw File Path";
            dgvQueue.Columns["RawFilePath"].Width = 250;

            dgvQueue.Columns["ProcessedFilePath"].HeaderText = "Processed File Path";
            dgvQueue.Columns["ProcessedFilePath"].Width = 300;

            dgvQueue.Columns["Timezone"].HeaderText = "Timezone";
            dgvQueue.Columns["Timezone"].Width = 120;

            dgvQueue.Columns["ReportStartTime"].HeaderText = "Report Start Time";
            dgvQueue.Columns["ReportStartTime"].Width = 170;
            dgvQueue.Columns["ReportStartTime"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

            dgvQueue.Columns["ReportEndTime"].HeaderText = "Report End Time";
            dgvQueue.Columns["ReportEndTime"].Width = 170;
            dgvQueue.Columns["ReportEndTime"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

            dgvQueue.Columns["SLAComplianceFlag"].HeaderText = "SLA Compliance";
            dgvQueue.Columns["SLAComplianceFlag"].Width = 130;

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

            dgvQueue.ColumnHeaderMouseClick -= DgvQueueColumnHeaderMouseClick;
            dgvQueue.ColumnHeaderMouseClick += DgvQueueColumnHeaderMouseClick;

            dgvQueue.CellContentClick -= DgvQueueCellContentClick;
            dgvQueue.CellContentClick += DgvQueueCellContentClick;

            dgvQueue.CellValueChanged -= DgvQueueCellValueChanged;
            dgvQueue.CellValueChanged += DgvQueueCellValueChanged;

            dgvQueue.CurrentCellDirtyStateChanged -= DgvQueueCurrentCellDirtyStateChanged;
            dgvQueue.CurrentCellDirtyStateChanged += DgvQueueCurrentCellDirtyStateChanged;

            dgvQueue.RowPrePaint -= DgvQueueRowPrePaint;
            dgvQueue.RowPrePaint += DgvQueueRowPrePaint;

            dgvQueue.EditingControlShowing -= DgvQueueEditingControlShowing;
            dgvQueue.EditingControlShowing += DgvQueueEditingControlShowing;

            dgvQueue.DataError -= DgvQueueDataError;
            dgvQueue.DataError += DgvQueueDataError;

            foreach (DataGridViewColumn col in dgvQueue.Columns)
            {
                if (col.Name != "Select" && col.Name != "Status")
                    col.ReadOnly = true;
            }
        }

        private void DgvQueueEditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgvQueue.CurrentCell != null &&
                dgvQueue.CurrentCell.OwningColumn.Name == "Status")
            {
                ComboBox combo = e.Control as ComboBox;
                if (combo != null)
                {
                    combo.DropDownStyle = ComboBoxStyle.DropDownList;
                    combo.SelectedIndexChanged -= StatusComboSelectedIndexChanged;
                    combo.SelectedIndexChanged += StatusComboSelectedIndexChanged;
                }
            }
        }

        private void StatusComboSelectedIndexChanged(object sender, EventArgs e)
        {
            if (dgvQueue.IsCurrentCellDirty)
                dgvQueue.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void DgvQueueDataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (dgvQueue.Columns[e.ColumnIndex].Name == "Status")
            {
                e.ThrowException = false;
                e.Cancel = false;
            }
        }

        private void DgvQueueCurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvQueue.IsCurrentCellDirty)
                dgvQueue.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void DgvQueueCellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            if (dgvQueue.Columns[e.ColumnIndex].Name == "Select")
            {
                dgvQueue.EndEdit();
                UpdateSelectedRecordsGrid();
            }
        }

        private void MenuItemR26_Click(object sender, EventArgs e)
        {
            // ✅ ALWAYS check Application.OpenForms FIRST
            var existingR26Form = Application.OpenForms.OfType<R26QueueForm>()
                .FirstOrDefault();

            if (existingR26Form != null && !existingR26Form.IsDisposed)
            {
                existingR26Form.Show();
                existingR26Form.BringToFront();
                this.Hide();
                return;
            }

            // ✅ Check stored reference
            if (_actualR26QueueForm != null && !_actualR26QueueForm.IsDisposed)
            {
                _actualR26QueueForm.Show();
                _actualR26QueueForm.BringToFront();
                this.Hide();
                return;
            }

            // ✅ Create new form if none exists
            _actualR26QueueForm = new R26QueueForm(_emailConfig);

            _actualR26QueueForm.FormClosed += (s, args) =>
            {
                _actualR26QueueForm = null;
                if (!this.IsDisposed && !_isExiting)
                {
                    this.Show();
                    this.BringToFront();
                }
            };

            _actualR26QueueForm.Show();
            this.Hide();
        }

        private void MenuItemReports_Click(object sender, EventArgs e)
        {
            // ✅ ALWAYS check Application.OpenForms FIRST
            var existingSamsForm = Application.OpenForms.OfType<SamsDeliveryReportForm>()
                .FirstOrDefault();

            if (existingSamsForm != null && !existingSamsForm.IsDisposed)
            {
                existingSamsForm.Show();
                existingSamsForm.BringToFront();
                existingSamsForm.Focus();
                this.Hide();
                return;
            }

            // ✅ Check stored reference
            if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
            {
                _samsDeliveryForm.Show();
                _samsDeliveryForm.BringToFront();
                _samsDeliveryForm.Focus();
                this.Hide();
                return;
            }

            // ✅ Create new form if none exists
            DateTime targetDate = DateTime.Today;
            string connectionString = Program.DbConnectionString;
            string storedProcedureName = "SDgetSAMSReportStatus";

            _samsDeliveryForm = new SamsDeliveryReportForm(
                targetDate,
                _emailConfig,
                connectionString,
                storedProcedureName,
                _actualR26QueueForm
            );

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
            MessageBox.Show("You are already in the Production menu.\n\nUse the 'File Type' dropdown to switch between:\n• Raw File\n• Bulk File\n• Processed File (current)",
                "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void DgvQueueCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            string columnName = dgvQueue.Columns[e.ColumnIndex].Name;

            if (columnName == "Status")
            {
                var row = dgvQueue.Rows[e.RowIndex];
                int queueId = Convert.ToInt32(row.Cells["ReportScheduleQueueUid"].Value);
                string newStatusStr = row.Cells["Status"].Value?.ToString() ?? "Pending";

                var originalRecord = _queueData.FirstOrDefault(q => q.ReportScheduleQueueUid == queueId);
                string originalStatus = originalRecord?.Status ?? "Pending";

                if (newStatusStr != originalStatus)
                    _individualStatusChanges[queueId] = newStatusStr;
                else
                    _individualStatusChanges.Remove(queueId);

                if (_dataTable != null)
                {
                    var dataRow = _dataTable.AsEnumerable()
                        .FirstOrDefault(r => r.Field<int>("ReportScheduleQueueUid") == queueId);
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

        private void MenuItemRawFile_Click_Processed(object sender, EventArgs e)
        {
            menuItemRawFile.Checked = true;
            menuItemBulkFile.Checked = false;
            menuItemProcessedFile.Checked = false;
            menuItemFileType.Text = "Raw File";

            // ✅ ALWAYS check Application.OpenForms FIRST
            var existingRawForm = Application.OpenForms.OfType<RpaScheduleQueueDetailForm>()
                .FirstOrDefault();

            if (existingRawForm != null && !existingRawForm.IsDisposed)
            {
                existingRawForm.SetR26QueueForm(_actualR26QueueForm);
                existingRawForm.SetSamsDeliveryForm(_samsDeliveryForm);
                existingRawForm.Show();
                existingRawForm.BringToFront();
                existingRawForm.Focus();
                this.Hide();
                return;
            }

            // ✅ Check stored reference
            if (_rawFileForm != null && !_rawFileForm.IsDisposed)
            {
                _rawFileForm.SetR26QueueForm(_actualR26QueueForm);
                _rawFileForm.SetSamsDeliveryForm(_samsDeliveryForm);
                _rawFileForm.Show();
                _rawFileForm.BringToFront();
                _rawFileForm.Focus();
                this.Hide();
                return;
            }

            // ✅ Create new form if none exists
            _rawFileForm = new RpaScheduleQueueDetailForm(
                _emailConfig,
                _actualR26QueueForm,
                _samsDeliveryForm
            );

            _rawFileForm.FormClosed += (s, args) =>
            {
                _rawFileForm = null;
                if (!this.IsDisposed && !_isExiting)
                {
                    this.Show();
                    this.BringToFront();
                    menuItemRawFile.Checked = false;
                    menuItemBulkFile.Checked = false;
                    menuItemProcessedFile.Checked = true;
                    menuItemFileType.Text = "Processed File";
                }
            };

            _rawFileForm.Show();
            this.Hide();
        }

        private void MenuItemBulkFile_Click_Processed(object sender, EventArgs e)
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

            // ✅ Check stored reference
            if (_bulkFileForm != null && !_bulkFileForm.IsDisposed)
            {
                _bulkFileForm.Show();
                _bulkFileForm.BringToFront();
                _bulkFileForm.Focus();
                this.Hide();
                return;
            }

            // ✅ Create new Bulk File form if not found
            _bulkFileForm = new SDReportProcessScheduleQueueForm(
                _emailConfig,
                _actualR26QueueForm,
                _rawFileForm,
                _samsDeliveryForm
            );

            _bulkFileForm.FormClosed += (s, args) =>
            {
                _bulkFileForm = null;
                if (!this.IsDisposed && !_isExiting)
                {
                    this.Show();
                    this.BringToFront();
                    menuItemRawFile.Checked = false;
                    menuItemBulkFile.Checked = false;
                    menuItemProcessedFile.Checked = true;
                    menuItemFileType.Text = "Processed File";
                }
            };

            _bulkFileForm.Show();
            this.Hide();
        }

        private void MenuItemProcessedFile_Click_Processed(object sender, EventArgs e)
        {
            menuItemRawFile.Checked = false;
            menuItemBulkFile.Checked = false;
            menuItemProcessedFile.Checked = true;

            // Update parent menu text
            menuItemFileType.Text = "Processed File";
            this.Text = "Report Process Schedule Queue - Status Management (Processed File)";
        }

        // Update OnVisibleChanged:
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);

            if (this.Visible)
            {
                // Reset menu selection to Processed File when this form becomes visible
                menuItemRawFile.Checked = false;
                menuItemBulkFile.Checked = false;
                menuItemProcessedFile.Checked = true;
                menuItemFileType.Text = "Processed File"; // Update parent menu text
                this.Text = "Report Process Schedule Queue - Status Management (Processed File)";

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

        private void DgvQueueRowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            var row = dgvQueue.Rows[e.RowIndex];
            object idValue = row.Cells["ReportScheduleQueueUid"].Value;
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
                if (row.Cells["ReportScheduleQueueUid"].Value != null &&
                    int.TryParse(row.Cells["ReportScheduleQueueUid"].Value.ToString(), out int queueId))
                {
                    if (_individualStatusChanges.ContainsKey(queueId))
                        row.DefaultCellStyle.BackColor = Color.LightYellow;
                    else
                        row.DefaultCellStyle.BackColor = Color.White;
                }
            }
        }

        private void DgvQueueColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == 0 &&
                dgvQueue.Columns[e.ColumnIndex].Name == "Select")
                return;
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
                    $"Convert(ReportScheduleQueueUid, 'System.String') LIKE '%{searchText}%' OR " +
                    $"CompanyName LIKE '%{searchText}%' OR " +
                    $"RawFilePath LIKE '%{searchText}%' OR " +
                    $"ProcessedFilePath LIKE '%{searchText}%'");
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
                    ID = r.Field<int>("ReportScheduleQueueUid"),
                    CompanyName = r.Field<string>("CompanyName"),
                    ProcessedFilePath = r.Field<string>("ProcessedFilePath"),
                    Status = r.Field<string>("Status")
                })
                .ToList();

            var selectedTable = new DataTable();
            selectedTable.Columns.Add("Company Name", typeof(string));
            selectedTable.Columns.Add("Processed File Path", typeof(string));
            selectedTable.Columns.Add("Status", typeof(string));

            foreach (var item in selectedRows)
                selectedTable.Rows.Add(item.CompanyName, item.ProcessedFilePath, item.Status);

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
                MessageBox.Show("No individual status changes to save.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    MessageBox.Show(
                        $"Successfully updated {updatedCount} individual records!",
                        "Update Successful",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    _individualStatusChanges.Clear();
                    lblIndividualChanges.Text = "Individual Changes: 0";

                    _ = LoadQueueDataAsync(isInitialLoad: false);
                }
                catch (Exception ex)
                {
                    ShowLoadingIndicator(false);
                    btnSaveIndividual.Enabled = true;
                    MessageBox.Show($"Error updating individual statuses: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnBulkUpdate_Click(object sender, EventArgs e)
        {
            var selectedRows = dgvQueue.Rows
                .Cast<DataGridViewRow>()
                .Where(r => r.Cells["Select"].Value != null &&
                            (bool)r.Cells["Select"].Value == true)
                .ToList();

            if (selectedRows.Count == 0)
            {
                MessageBox.Show("Please select at least one record to update.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (cmbBulkStatus.SelectedItem == null)
            {
                MessageBox.Show("Please select a status from the dropdown.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                            int queueId = Convert.ToInt32(row.Cells["ReportScheduleQueueUid"].Value);
                            updates[queueId] = newStatus;
                        }

                        int updatedCount = _repository.UpdateStatuses(updates, _systemName);

                        ShowLoadingIndicator(false);
                        btnBulkUpdate.Enabled = true;

                        MessageBox.Show(
                            $"Successfully updated {updatedCount} records via bulk update!",
                            "Update Successful",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

                        foreach (DataGridViewRow row in dgvQueue.Rows)
                        {
                            if (row.Cells["Select"].Value != null)
                                row.Cells["Select"].Value = false;
                        }

                        if (_dataTable != null)
                        {
                            foreach (DataRow dataRow in _dataTable.Rows)
                                dataRow["Select"] = false;
                        }

                        UpdateSelectedRecordsGrid();
                        _ = LoadQueueDataAsync(isInitialLoad: false);
                    }
                    catch (Exception ex)
                    {
                        ShowLoadingIndicator(false);
                        btnBulkUpdate.Enabled = true;
                        MessageBox.Show($"Error updating bulk statuses: {ex.Message}", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void MenuItemMenu_Click(object sender, EventArgs e)
        {
            // Same behaviour as other forms – just close and go back.
            _isExiting = true;
            this.Close();
        }

        private void ReportProcessScheduleQueueForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_isExiting)
                return;
            if (e.CloseReason == CloseReason.ApplicationExitCall)
                return;
            if (e.CloseReason == CloseReason.UserClosing)
            {
                int individualChangesCount = _individualStatusChanges?.Count ?? 0;
                int selectedRowsCount = 0;
                if (dgvQueue?.Rows != null)
                {
                    selectedRowsCount = dgvQueue.Rows
                        .Cast<DataGridViewRow>()
                        .Count(r => r.Cells["Select"].Value != null &&
                                    (bool)r.Cells["Select"].Value == true);
                }
                int totalPendingChanges = individualChangesCount + selectedRowsCount;
                if (totalPendingChanges > 0)
                {
                    string message = "You have unsaved changes:\n\n";
                    if (individualChangesCount > 0)
                        message += $"• Individual status changes: {individualChangesCount}\n";
                    if (selectedRowsCount > 0)
                        message += $"• Selected records for bulk update: {selectedRowsCount}\n";

                    message += "\nAre you sure you want to exit without saving?";
                    var result = MessageBox.Show(message, "Unsaved Changes - Confirm Exit",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.Yes)
                    {
                        if (_actualR26QueueForm != null && !_actualR26QueueForm.IsDisposed)
                        {
                            _actualR26QueueForm._isExiting = true;
                            _actualR26QueueForm.Close();
                        }
                        if (_rawFileForm != null && !_rawFileForm.IsDisposed)
                        {
                            _rawFileForm._isExiting = true;
                            _rawFileForm.Close();
                        }
                        if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
                        {
                            _samsDeliveryForm._isExiting = true;
                            _samsDeliveryForm.Close();
                        }
                        if (_bulkFileForm != null && !_bulkFileForm.IsDisposed)
                        {
                            _bulkFileForm._isExiting = true;
                            _bulkFileForm.Close();
                        }
                        Application.Exit();
                    }
                    else
                    {
                        _isExiting = false;
                        e.Cancel = true;
                    }
                }
                else
                {
                    var result = MessageBox.Show("Are you sure you want to exit?", "Confirm Exit",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        if (_actualR26QueueForm != null && !_actualR26QueueForm.IsDisposed)
                        {
                            _actualR26QueueForm._isExiting = true;
                            _actualR26QueueForm.Close();
                        }
                        if (_rawFileForm != null && !_rawFileForm.IsDisposed)
                        {
                            _rawFileForm._isExiting = true;
                            _rawFileForm.Close();
                        }
                        if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
                        {
                            _samsDeliveryForm._isExiting = true;
                            _samsDeliveryForm.Close();
                        }
                        if (_bulkFileForm != null && !_bulkFileForm.IsDisposed)
                        {
                            _bulkFileForm._isExiting = true;
                            _bulkFileForm.Close();
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
        }
    }
}