using Microsoft.Extensions.Configuration;
using R26_DailyQueueWinForm.Data;
using R26_DailyQueueWinForm.Forms;
using R26_DailyQueueWinForm.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrayNotify;


namespace R26_DailyQueueWinForm
{
    public partial class R26QueueForm : Form
    {
        private readonly R26QueueRepository _repository;
        private List<R26QueueModel> _queueData;
        private DataTable _dataTable;
        private Dictionary<int, string> _individualStatusChanges;
        private BindingSource _bindingSource;
        private EmailConfiguration _emailConfig;

        // Reference to Sam's Delivery Report form for toggling
        private SamsDeliveryReportForm _samsDeliveryForm;

        // Filter controls
        private Label lblCompanyFilter;
        private ComboBox cmbCompanyFilter;
        private Label lblStatusFilter;
        private ComboBox cmbStatusFilter;
        private Label lblCreatedDateFilter;
        private Button btnCreatedDatePicker;
        private Label lblCreatedDateValue;
        private Label lblModifiedDateFilter;
        private Button btnModifiedDatePicker;
        private Label lblModifiedDateValue;

        private DateTime? _selectedCreatedDate = null;
        private DateTime? _selectedModifiedDate = null;

        // menu controls
        private MenuStrip mainMenuStrip;
        private ToolStripMenuItem menuItemR26;
        private ToolStripMenuItem menuItemReports;

        // main UI controls
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

        // Cache for status list from DB
        private List<string> _statusList;

        public R26QueueForm(EmailConfiguration emailConfig)
        {
            InitializeComponent();
            _emailConfig = emailConfig;

            // ✅ Use decrypted DB connection
            string connectionString = Program.DbConnectionString;

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                MessageBox.Show("Database connection string is empty!", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Try to parse it to verify it's valid
            try
            {
                var builder = new System.Data.SqlClient.SqlConnectionStringBuilder(connectionString);
                // Connection string is valid
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Invalid connection string format:\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _repository = new R26QueueRepository(connectionString);

            _individualStatusChanges = new Dictionary<int, string>();
            _bindingSource = new BindingSource();

            // Load statuses from database once during initialization
            LoadStatusesFromDatabase();

            if (mainMenuStrip != null && lblIndividualChanges != null)
            {
                mainMenuStrip.Dock = DockStyle.None;
                mainMenuStrip.Location = new Point(
                    lblIndividualChanges.Right + 20,
                    lblIndividualChanges.Top - 2
                );
            }

            this.Load += R26QueueForm_Load;
            this.FormClosing += R26QueueForm_FormClosing;
        }


        private void LoadStatusesFromDatabase()
        {
            try
            {
                // Fetch statuses from database and cache them
                _statusList = _repository.GetAllStatuses();

                if (_statusList == null || _statusList.Count == 0)
                {
                    MessageBox.Show("No statuses found in database. Please add statuses to the Status table.",
                        "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _statusList = new List<string>();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading statuses from database: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusList = new List<string>();
            }
        }

        private void InitializeComponent()
        {
            int filterY = 115;
            int filterHeight = 28;
            int labelVerticalOffset = 7;

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

            // Filter controls
            this.lblCompanyFilter = new Label();
            this.cmbCompanyFilter = new ComboBox();
            this.lblStatusFilter = new Label();
            this.cmbStatusFilter = new ComboBox();
            this.lblCreatedDateFilter = new Label();
            this.btnCreatedDatePicker = new Button();
            this.lblCreatedDateValue = new Label();
            this.lblModifiedDateFilter = new Label();
            this.btnModifiedDatePicker = new Button();
            this.lblModifiedDateValue = new Label();

            ((ISupportInitialize)(this.dgvQueue)).BeginInit();
            ((ISupportInitialize)(this.dgvSelectedRecords)).BeginInit();
            this.panelSide.SuspendLayout();
            this.SuspendLayout();

            // mainMenuStrip
            this.mainMenuStrip.BackColor = Color.White;
            this.mainMenuStrip.Dock = DockStyle.None;
            this.mainMenuStrip.Font = new Font("Segoe UI", 10F);
            this.mainMenuStrip.GripStyle = ToolStripGripStyle.Hidden;
            this.mainMenuStrip.RenderMode = ToolStripRenderMode.Professional;
            this.mainMenuStrip.Padding = new Padding(10, 5, 0, 5);

            // menuItemR26
            this.menuItemR26.Text = "R26 Daily Queue";
            this.menuItemR26.Font = new Font("Segoe UI", 10F, FontStyle.Bold);

            // menuItemReports
            this.menuItemReports.Text = "Sam's Delivery Report Status";
            this.menuItemReports.Click += MenuItemReports_Click;

            this.mainMenuStrip.Items.AddRange(new ToolStripItem[]
            {
                this.menuItemR26,
                this.menuItemReports
            });

            // lblTitle
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new Font("Segoe UI", 18F, FontStyle.Bold);
            this.lblTitle.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblTitle.Location = new Point(12, 10);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new Size(400, 32);
            this.lblTitle.Text = "R26 Daily Queue Management";

            // lblRecordCount
            this.lblRecordCount.AutoSize = true;
            this.lblRecordCount.Font = new Font("Segoe UI", 10F);
            this.lblRecordCount.Location = new Point(12, 60);
            this.lblRecordCount.Name = "lblRecordCount";
            this.lblRecordCount.Size = new Size(150, 19);
            this.lblRecordCount.Text = "Total Records: 0";

            // lblSelectedCount
            this.lblSelectedCount.AutoSize = true;
            this.lblSelectedCount.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblSelectedCount.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblSelectedCount.Location = new Point(250, 60);
            this.lblSelectedCount.Name = "lblSelectedCount";
            this.lblSelectedCount.Size = new Size(150, 19);
            this.lblSelectedCount.Text = "Selected: 0";

            // lblIndividualChanges
            this.lblIndividualChanges.AutoSize = true;
            this.lblIndividualChanges.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblIndividualChanges.ForeColor = Color.FromArgb(255, 140, 0);
            this.lblIndividualChanges.Location = new Point(450, 60);
            this.lblIndividualChanges.Name = "lblIndividualChanges";
            this.lblIndividualChanges.Size = new Size(180, 19);
            this.lblIndividualChanges.Text = "Individual Changes: 0";

            // lblSearch
            this.lblSearch.AutoSize = true;
            this.lblSearch.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            this.lblSearch.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblSearch.Location = new Point(850, 15);
            this.lblSearch.Name = "lblSearch";
            this.lblSearch.Size = new Size(70, 20);
            this.lblSearch.Text = "Search";
            this.lblSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            // txtSearch
            this.txtSearch.Font = new Font("Segoe UI", 11F);
            this.txtSearch.Location = new Point(930, 13);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new Size(350, 27);
            this.txtSearch.BorderStyle = BorderStyle.FixedSingle;
            this.txtSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.txtSearch.TextChanged += TxtSearch_TextChanged;

            // progressBar
            this.progressBar.Location = new Point(12, 85);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new Size(400, 23);
            this.progressBar.Style = ProgressBarStyle.Marquee;
            this.progressBar.Visible = false;

            // lblCompanyFilter
            this.lblCompanyFilter.AutoSize = true;
            this.lblCompanyFilter.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            this.lblCompanyFilter.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblCompanyFilter.Location = new Point(20, filterY + labelVerticalOffset);
            this.lblCompanyFilter.Name = "lblCompanyFilter";
            this.lblCompanyFilter.Text = "Company Name";

            // cmbCompanyFilter
            this.cmbCompanyFilter.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbCompanyFilter.Font = new Font("Segoe UI", 9.5F);
            this.cmbCompanyFilter.Location = new Point(170, filterY);
            this.cmbCompanyFilter.Size = new Size(260, filterHeight);
            this.cmbCompanyFilter.Name = "cmbCompanyFilter";
            this.cmbCompanyFilter.SelectedIndexChanged += FilterChanged;

            // lblStatusFilter
            this.lblStatusFilter.AutoSize = true;
            this.lblStatusFilter.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            this.lblStatusFilter.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblStatusFilter.Location = new Point(450, filterY + 5);
            this.lblStatusFilter.Name = "lblStatusFilter";
            this.lblStatusFilter.Text = "Status";

            this.cmbStatusFilter.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbStatusFilter.Font = new Font("Segoe UI", 9.5F);
            this.cmbStatusFilter.Location = new Point(520, filterY);
            this.cmbStatusFilter.Size = new Size(180, filterHeight);
            this.cmbStatusFilter.Name = "cmbStatusFilter";
            this.cmbStatusFilter.SelectedIndexChanged += FilterChanged;

            // lblCreatedDateFilter
            this.lblCreatedDateFilter.AutoSize = true;
            this.lblCreatedDateFilter.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            this.lblCreatedDateFilter.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblCreatedDateFilter.Location = new Point(720, filterY + 5);
            this.lblCreatedDateFilter.Name = "lblCreatedDateFilter";
            this.lblCreatedDateFilter.Text = "Created Date";

            this.lblCreatedDateValue.Font = new Font("Segoe UI", 9.5F);
            this.lblCreatedDateValue.ForeColor = Color.FromArgb(100, 100, 100);
            this.lblCreatedDateValue.Location = new Point(850, filterY);
            this.lblCreatedDateValue.Size = new Size(120, filterHeight);
            this.lblCreatedDateValue.TextAlign = ContentAlignment.MiddleLeft;
            this.lblCreatedDateValue.Name = "lblCreatedDateValue";
            this.lblCreatedDateValue.Text = "Not Selected";

            this.btnCreatedDatePicker.BackColor = Color.FromArgb(0, 120, 212);
            this.btnCreatedDatePicker.FlatStyle = FlatStyle.Flat;
            this.btnCreatedDatePicker.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            this.btnCreatedDatePicker.ForeColor = Color.White;
            this.btnCreatedDatePicker.Location = new Point(975, filterY - 2);
            this.btnCreatedDatePicker.Size = new Size(90, 32);
            this.btnCreatedDatePicker.FlatAppearance.BorderSize = 0;
            this.btnCreatedDatePicker.Name = "btnCreatedDatePicker";
            this.btnCreatedDatePicker.Text = "Select";
            this.btnCreatedDatePicker.TextAlign = ContentAlignment.MiddleCenter;
            this.btnCreatedDatePicker.UseVisualStyleBackColor = false;
            this.btnCreatedDatePicker.Padding = new Padding(0, 0, 0, 0);
            this.btnCreatedDatePicker.AutoSize = false;
            this.btnCreatedDatePicker.Click += BtnCreatedDatePicker_Click;

            // lblModifiedDateFilter
            this.lblModifiedDateFilter.AutoSize = true;
            this.lblModifiedDateFilter.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            this.lblModifiedDateFilter.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblModifiedDateFilter.Location = new Point(1070, filterY + 5);
            this.lblModifiedDateFilter.Name = "lblModifiedDateFilter";
            this.lblModifiedDateFilter.Text = "Modified Date";

            this.lblModifiedDateValue.Font = new Font("Segoe UI", 9.5F);
            this.lblModifiedDateValue.ForeColor = Color.FromArgb(100, 100, 100);
            this.lblModifiedDateValue.Location = new Point(1200, filterY);
            this.lblModifiedDateValue.Size = new Size(120, filterHeight);
            this.lblModifiedDateValue.TextAlign = ContentAlignment.MiddleLeft;
            this.lblModifiedDateValue.Name = "lblModifiedDateValue";
            this.lblModifiedDateValue.Text = "Not Selected";

            this.btnModifiedDatePicker.BackColor = Color.FromArgb(0, 120, 212);
            this.btnModifiedDatePicker.FlatStyle = FlatStyle.Flat;
            this.btnModifiedDatePicker.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            this.btnModifiedDatePicker.ForeColor = Color.White;
            this.btnModifiedDatePicker.Location = new Point(1320, filterY);
            this.btnModifiedDatePicker.Size = new Size(90, 32);
            this.btnModifiedDatePicker.FlatAppearance.BorderSize = 0;
            this.btnModifiedDatePicker.Name = "btnModifiedDatePicker";
            this.btnModifiedDatePicker.Text = "Select";
            this.btnModifiedDatePicker.TextAlign = ContentAlignment.MiddleCenter;
            this.btnModifiedDatePicker.UseVisualStyleBackColor = false;
            this.btnModifiedDatePicker.Padding = new Padding(0, 0, 0, 0);
            this.btnModifiedDatePicker.AutoSize = false;
            this.btnModifiedDatePicker.Click += BtnModifiedDatePicker_Click;

            // dgvQueue - SMOOTH SCROLLING OPTIMIZATIONS
            this.dgvQueue.AllowUserToAddRows = false;
            this.dgvQueue.AllowUserToDeleteRows = false;
            this.dgvQueue.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            this.dgvQueue.BackgroundColor = Color.White;
            this.dgvQueue.BorderStyle = BorderStyle.Fixed3D;
            this.dgvQueue.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvQueue.ColumnHeadersHeight = 45;
            this.dgvQueue.EnableHeadersVisualStyles = false;
            this.dgvQueue.Location = new Point(12, 190);
            this.dgvQueue.Name = "dgvQueue";
            this.dgvQueue.RowHeadersWidth = 30;
            this.dgvQueue.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.dgvQueue.Size = new Size(1050, 570);
            this.dgvQueue.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // Column header styling - Professional dark blue/teal header
            this.dgvQueue.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 58, 85);
            this.dgvQueue.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            this.dgvQueue.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.dgvQueue.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.dgvQueue.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(30, 58, 85);
            this.dgvQueue.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.White;

            // Row selection color - Light blue (different from header)
            this.dgvQueue.DefaultCellStyle.SelectionBackColor = Color.FromArgb(173, 216, 230);
            this.dgvQueue.DefaultCellStyle.SelectionForeColor = Color.Black;

            // CRITICAL: Disable VirtualMode for smooth scrolling
            this.dgvQueue.VirtualMode = false;
            this.dgvQueue.AllowUserToResizeRows = false;

            // Performance optimizations for smooth scrolling
            this.dgvQueue.DoubleBuffered(true);
            this.dgvQueue.RowTemplate.Height = 25;
            this.dgvQueue.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            this.dgvQueue.CellBorderStyle = DataGridViewCellBorderStyle.Single;

            // panelSide
            this.panelSide.BackColor = Color.FromArgb(240, 240, 240);
            this.panelSide.BorderStyle = BorderStyle.FixedSingle;
            this.panelSide.Location = new Point(1070, 190);
            this.panelSide.Name = "panelSide";
            this.panelSide.Size = new Size(500, 570);
            this.panelSide.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;

            // lblSideTitle
            this.lblSideTitle.AutoSize = true;
            this.lblSideTitle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            this.lblSideTitle.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblSideTitle.Location = new Point(10, 10);
            this.lblSideTitle.Name = "lblSideTitle";
            this.lblSideTitle.Size = new Size(200, 21);
            this.lblSideTitle.Text = "Bulk Status Update";

            // dgvSelectedRecords
            this.dgvSelectedRecords.AllowUserToAddRows = false;
            this.dgvSelectedRecords.AllowUserToDeleteRows = false;
            this.dgvSelectedRecords.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvSelectedRecords.BackgroundColor = Color.White;
            this.dgvSelectedRecords.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSelectedRecords.Location = new Point(10, 40);
            this.dgvSelectedRecords.Name = "dgvSelectedRecords";
            this.dgvSelectedRecords.ReadOnly = true;
            this.dgvSelectedRecords.RowHeadersWidth = 30;
            this.dgvSelectedRecords.Size = new Size(478, 280);
            this.dgvSelectedRecords.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // lblBulkStatus
            this.lblBulkStatus.AutoSize = true;
            this.lblBulkStatus.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblBulkStatus.Location = new Point(10, 335);
            this.lblBulkStatus.Name = "lblBulkStatus";
            this.lblBulkStatus.Size = new Size(140, 19);
            this.lblBulkStatus.Text = "Select New Status";
            this.lblBulkStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            // cmbBulkStatus
            this.cmbBulkStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbBulkStatus.Font = new Font("Segoe UI", 11F);
            this.cmbBulkStatus.FormattingEnabled = true;
            this.cmbBulkStatus.Location = new Point(10, 360);
            this.cmbBulkStatus.Name = "cmbBulkStatus";
            this.cmbBulkStatus.Size = new Size(478, 28);
            this.cmbBulkStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // btnBulkUpdate
            this.btnBulkUpdate.BackColor = Color.FromArgb(0, 120, 212);
            this.btnBulkUpdate.FlatStyle = FlatStyle.Flat;
            this.btnBulkUpdate.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.btnBulkUpdate.ForeColor = Color.White;
            this.btnBulkUpdate.Location = new Point(10, 400);
            this.btnBulkUpdate.Name = "btnBulkUpdate";
            this.btnBulkUpdate.Size = new Size(478, 40);
            this.btnBulkUpdate.Text = "Update Selected (Bulk)";
            this.btnBulkUpdate.UseVisualStyleBackColor = false;
            this.btnBulkUpdate.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.btnBulkUpdate.Click += BtnBulkUpdate_Click;

            // btnSaveIndividual
            this.btnSaveIndividual.BackColor = Color.FromArgb(255, 140, 0);
            this.btnSaveIndividual.FlatStyle = FlatStyle.Flat;
            this.btnSaveIndividual.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.btnSaveIndividual.ForeColor = Color.White;
            this.btnSaveIndividual.Location = new Point(10, 450);
            this.btnSaveIndividual.Name = "btnSaveIndividual";
            this.btnSaveIndividual.Size = new Size(478, 40);
            this.btnSaveIndividual.Text = "Save Individual Changes";
            this.btnSaveIndividual.UseVisualStyleBackColor = false;
            this.btnSaveIndividual.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.btnSaveIndividual.Click += BtnSaveIndividual_Click;

            // btnRefresh
            this.btnRefresh.BackColor = Color.FromArgb(76, 175, 80);
            this.btnRefresh.FlatStyle = FlatStyle.Flat;
            this.btnRefresh.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.btnRefresh.ForeColor = Color.White;
            this.btnRefresh.Location = new Point(10, 500);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new Size(478, 40);
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
            this.Controls.Add(this.lblCreatedDateFilter);
            this.Controls.Add(this.lblCreatedDateValue);
            this.Controls.Add(this.btnCreatedDatePicker);
            this.Controls.Add(this.lblModifiedDateFilter);
            this.Controls.Add(this.lblModifiedDateValue);
            this.Controls.Add(this.btnModifiedDatePicker);
            this.Controls.Add(this.panelSide);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblIndividualChanges);
            this.Controls.Add(this.lblSelectedCount);
            this.Controls.Add(this.lblRecordCount);
            this.Controls.Add(this.dgvQueue);
            this.Controls.Add(this.lblTitle);
            this.MainMenuStrip = this.mainMenuStrip;
            this.MinimumSize = new Size(1400, 700);
            this.Name = "R26QueueForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "R26 Daily Queue - Status Management";
            this.WindowState = FormWindowState.Maximized;

            ((ISupportInitialize)(this.dgvQueue)).EndInit();
            ((ISupportInitialize)(this.dgvSelectedRecords)).EndInit();
            this.panelSide.ResumeLayout(false);
            this.panelSide.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private async void R26QueueForm_Load(object sender, EventArgs e)
        {
            await LoadQueueDataAsync();
        }

        private async System.Threading.Tasks.Task LoadQueueDataAsync()
        {
            try
            {
                progressBar.Visible = true;
                btnRefresh.Enabled = false;
                btnBulkUpdate.Enabled = false;
                btnSaveIndividual.Enabled = false;

                // ✅ Suspend layout for better performance (DON'T hide the grid)
                this.SuspendLayout();
                dgvQueue.SuspendLayout();

                // ✅ Temporarily detach DataSource to prevent UI updates during data load
                dgvQueue.DataSource = null;

                // ✅ Load data from database in background thread
                await System.Threading.Tasks.Task.Run(() =>
                {
                    _queueData = _repository.GetAllR26Queue();
                });

                _dataTable = new DataTable();
                _individualStatusChanges.Clear();

                // Define columns
                _dataTable.Columns.Add("Select", typeof(bool));
                _dataTable.Columns.Add("r26queueuid", typeof(int));
                _dataTable.Columns.Add("CompanyName", typeof(string));
                _dataTable.Columns.Add("PatNumber", typeof(string));
                _dataTable.Columns.Add("PatFName", typeof(string));
                _dataTable.Columns.Add("PatMInitial", typeof(string));
                _dataTable.Columns.Add("PatLName", typeof(string));
                _dataTable.Columns.Add("PatSex", typeof(string));
                _dataTable.Columns.Add("PatBirthdate", typeof(string));
                _dataTable.Columns.Add("LocationNumber", typeof(string));
                _dataTable.Columns.Add("Date", typeof(string));
                _dataTable.Columns.Add("Time", typeof(string));
                _dataTable.Columns.Add("ApptType", typeof(string));
                _dataTable.Columns.Add("Reason", typeof(string));
                _dataTable.Columns.Add("PatAddress1", typeof(string));
                _dataTable.Columns.Add("PatCity", typeof(string));
                _dataTable.Columns.Add("PatState", typeof(string));
                _dataTable.Columns.Add("PatZip5", typeof(string));
                _dataTable.Columns.Add("HomePhone", typeof(string));
                _dataTable.Columns.Add("WorkPhone", typeof(string));
                _dataTable.Columns.Add("ResourceName", typeof(string));
                _dataTable.Columns.Add("ProviderName", typeof(string));
                _dataTable.Columns.Add("PrimaryInsuranceName", typeof(string));
                _dataTable.Columns.Add("BotName", typeof(string));
                _dataTable.Columns.Add("Status", typeof(string));
                _dataTable.Columns.Add("createdate", typeof(DateTime));
                _dataTable.Columns.Add("modifieddate", typeof(DateTime));

                // ✅ Batch add rows with BeginLoadData/EndLoadData
                _dataTable.BeginLoadData();
                foreach (var queue in _queueData)
                {
                    _dataTable.Rows.Add(
                        false,
                        queue.R26QueueUid,
                        queue.CompanyName ?? "Company " + queue.CompanyUid,
                        queue.PatNumber?.ToString() ?? "",
                        queue.PatFName ?? "",
                        queue.PatMInitial ?? "",
                        queue.PatLName ?? "",
                        queue.PatSex ?? "",
                        queue.PatBirthdate?.ToString("yyyy-MM-dd") ?? "",
                        queue.LocationNumber?.ToString() ?? "",
                        queue.Date?.ToString("yyyy-MM-dd") ?? "",
                        queue.Time?.ToString(@"hh\:mm") ?? "",
                        queue.ApptType ?? "",
                        queue.Reason ?? "",
                        queue.PatAddress1 ?? "",
                        queue.PatCity ?? "",
                        queue.PatState ?? "",
                        queue.PatZip5 ?? "",
                        queue.HomePhone ?? "",
                        queue.WorkPhone ?? "",
                        queue.ResourceName ?? "",
                        queue.ProviderName ?? "",
                        queue.PrimaryInsuranceName ?? "",
                        queue.BotName ?? "",
                        queue.Status ?? "Pending",
                        queue.CreatedDate ?? (object)DBNull.Value,
                        queue.ModifiedDate ?? (object)DBNull.Value
                    );
                }
                _dataTable.EndLoadData();

                // ✅ Set binding source BEFORE attaching to DataGridView
                _bindingSource.DataSource = _dataTable;

                // ✅ Load filters and status lists BEFORE binding to grid
                LoadCompanyFilter();
                LoadStatusListsFromData();

                // ✅ NOW attach the data source (this is the expensive operation)
                dgvQueue.DataSource = _bindingSource;

                // ✅ Format AFTER data is bound (only once)
                FormatDataGridView();

                lblRecordCount.Text = $"Total Records: {_queueData.Count:N0}";
                lblIndividualChanges.Text = "Individual Changes: 0";
                UpdateSelectedRecordsGrid();

                // ✅ Resume layout
                dgvQueue.ResumeLayout();
                this.ResumeLayout();

                progressBar.Visible = false;
                btnRefresh.Enabled = true;
                btnBulkUpdate.Enabled = true;
                btnSaveIndividual.Enabled = true;

                // ✅ SUCCESS MESSAGE RETAINED - Shows AFTER UI is fully loaded
                MessageBox.Show($"Loaded {_queueData.Count:N0} records successfully!",
                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                dgvQueue.ResumeLayout();
                this.ResumeLayout();
                progressBar.Visible = false;
                btnRefresh.Enabled = true;
                btnBulkUpdate.Enabled = true;
                btnSaveIndividual.Enabled = true;

                MessageBox.Show($"Error loading queue data: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadCompanyFilter()
        {
            cmbCompanyFilter.BeginUpdate(); // ✅ Suspend painting
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
            cmbCompanyFilter.EndUpdate(); // ✅ Resume painting
        }

        private void LoadStatusListsFromData()
        {
            if (_dataTable == null || _statusList == null) return;

            // Status filter combo - use cached status list from DB
            cmbStatusFilter.BeginUpdate(); // ✅ Suspend painting
            cmbStatusFilter.Items.Clear();
            cmbStatusFilter.Items.Add("All Status");
            foreach (var s in _statusList)
                cmbStatusFilter.Items.Add(s);
            cmbStatusFilter.SelectedIndex = 0;
            cmbStatusFilter.EndUpdate(); // ✅ Resume painting

            // Bulk status combo - use cached status list from DB
            cmbBulkStatus.BeginUpdate(); // ✅ Suspend painting
            cmbBulkStatus.Items.Clear();
            foreach (var s in _statusList)
                cmbBulkStatus.Items.Add(s);
            if (cmbBulkStatus.Items.Count > 0)
                cmbBulkStatus.SelectedIndex = 0;
            cmbBulkStatus.EndUpdate(); // ✅ Resume painting

            // Grid status column combo - use cached status list from DB
            // ✅ Only create the column if it doesn't already exist as ComboBoxColumn
            if (dgvQueue.Columns.Contains("Status") && !(dgvQueue.Columns["Status"] is DataGridViewComboBoxColumn))
            {
                int statusIndex = dgvQueue.Columns["Status"].Index;
                dgvQueue.Columns.RemoveAt(statusIndex);

                var statusColumn = new DataGridViewComboBoxColumn
                {
                    HeaderText = "Status",
                    Name = "Status",
                    DataPropertyName = "Status",
                    Width = 130,
                    MinimumWidth = 130,
                    FlatStyle = FlatStyle.Flat
                };
                statusColumn.Items.AddRange(_statusList.Cast<object>().ToArray());
                dgvQueue.Columns.Insert(statusIndex, statusColumn);
            }
        }


        private void DgvQueue_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // If user clicks on the "Select" column header, do nothing (prevent select all)
            if (e.ColumnIndex >= 0 && dgvQueue.Columns[e.ColumnIndex].Name == "Select")
            {
                // Do nothing - user must select rows individually
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

            // Company filter
            if (cmbCompanyFilter.SelectedIndex > 0)
            {
                string companyName = cmbCompanyFilter.SelectedItem.ToString().Replace("'", "''");
                filters.Add($"CompanyName = '{companyName}'");
            }

            // Status filter
            if (cmbStatusFilter.SelectedIndex > 0)
            {
                string status = cmbStatusFilter.SelectedItem.ToString().Replace("'", "''");
                filters.Add($"Status = '{status}'");
            }

            // Created date filter
            if (_selectedCreatedDate.HasValue)
            {
                DateTime d = _selectedCreatedDate.Value.Date;
                filters.Add($"createdate >= '{d:yyyy-MM-dd}' AND createdate < '{d.AddDays(1):yyyy-MM-dd}'");
            }

            // Modified date filter
            if (_selectedModifiedDate.HasValue)
            {
                DateTime d = _selectedModifiedDate.Value.Date;
                filters.Add($"modifieddate >= '{d:yyyy-MM-dd}' AND modifieddate < '{d.AddDays(1):yyyy-MM-dd}'");
            }

            // Search filter
            string searchText = txtSearch.Text.Trim().Replace("'", "''");
            if (!string.IsNullOrEmpty(searchText))
            {
                var searchColumns = new[]
                {
                    "r26queueuid", "CompanyName", "PatNumber", "PatFName", "PatMInitial",
                    "PatLName", "PatSex", "PatBirthdate", "LocationNumber", "Date",
                    "Time", "ApptType", "Reason", "PatAddress1", "PatCity", "PatState",
                    "PatZip5", "HomePhone", "WorkPhone", "ResourceName", "ProviderName",
                    "PrimaryInsuranceName", "BotName", "Status"
                };

                var searchConditions = new List<string>();
                foreach (string column in searchColumns)
                {
                    searchConditions.Add($"Convert([{column}], System.String) LIKE '%{searchText}%'");
                }

                filters.Add("(" + string.Join(" OR ", searchConditions) + ")");
            }

            string finalFilter = filters.Count == 0 ? null : string.Join(" AND ", filters);
            _bindingSource.Filter = finalFilter;

            lblRecordCount.Text = finalFilter == null
                ? $"Total Records: {_dataTable.Rows.Count:N0}"
                : $"Total Records: {_bindingSource.Count:N0}";
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void FormatDataGridView()
        {
            if (!dgvQueue.Columns.Contains("Select")) return;

            dgvQueue.Columns["Select"].Width = 60;
            dgvQueue.Columns["Select"].HeaderText = "Select";
            dgvQueue.Columns["Select"].ReadOnly = false;
            dgvQueue.Columns["Select"].Frozen = true;

            dgvQueue.Columns["r26queueuid"].HeaderText = "Queue ID";
            dgvQueue.Columns["r26queueuid"].Width = 85;
            dgvQueue.Columns["r26queueuid"].Frozen = true;

            dgvQueue.Columns["CompanyName"].HeaderText = "Company Name";
            dgvQueue.Columns["CompanyName"].Width = 180;
            dgvQueue.Columns["CompanyName"].Frozen = true;

            dgvQueue.Columns["PatNumber"].HeaderText = "Pat Number";
            dgvQueue.Columns["PatNumber"].Width = 95;
            dgvQueue.Columns["PatNumber"].Frozen = true;

            // Set proper widths for all other columns to make headers fully visible
            if (dgvQueue.Columns.Contains("PatFName"))
                dgvQueue.Columns["PatFName"].Width = 100;

            if (dgvQueue.Columns.Contains("PatMInitial"))
                dgvQueue.Columns["PatMInitial"].Width = 85;

            if (dgvQueue.Columns.Contains("PatLName"))
                dgvQueue.Columns["PatLName"].Width = 100;

            if (dgvQueue.Columns.Contains("PatSex"))
                dgvQueue.Columns["PatSex"].Width = 70;

            if (dgvQueue.Columns.Contains("PatBirthdate"))
                dgvQueue.Columns["PatBirthdate"].Width = 110;

            if (dgvQueue.Columns.Contains("LocationNumber"))
                dgvQueue.Columns["LocationNumber"].Width = 125;

            if (dgvQueue.Columns.Contains("Date"))
                dgvQueue.Columns["Date"].Width = 100;

            if (dgvQueue.Columns.Contains("Time"))
                dgvQueue.Columns["Time"].Width = 80;

            if (dgvQueue.Columns.Contains("ApptType"))
                dgvQueue.Columns["ApptType"].Width = 90;

            if (dgvQueue.Columns.Contains("Reason"))
                dgvQueue.Columns["Reason"].Width = 120;

            if (dgvQueue.Columns.Contains("PatAddress1"))
                dgvQueue.Columns["PatAddress1"].Width = 150;

            if (dgvQueue.Columns.Contains("PatCity"))
                dgvQueue.Columns["PatCity"].Width = 100;

            if (dgvQueue.Columns.Contains("PatState"))
                dgvQueue.Columns["PatState"].Width = 80;

            if (dgvQueue.Columns.Contains("PatZip5"))
                dgvQueue.Columns["PatZip5"].Width = 80;

            if (dgvQueue.Columns.Contains("HomePhone"))
                dgvQueue.Columns["HomePhone"].Width = 115;

            if (dgvQueue.Columns.Contains("WorkPhone"))
                dgvQueue.Columns["WorkPhone"].Width = 115;

            if (dgvQueue.Columns.Contains("ResourceName"))
                dgvQueue.Columns["ResourceName"].Width = 140;

            if (dgvQueue.Columns.Contains("ProviderName"))
                dgvQueue.Columns["ProviderName"].Width = 140;

            if (dgvQueue.Columns.Contains("PrimaryInsuranceName"))
                dgvQueue.Columns["PrimaryInsuranceName"].Width = 160;

            if (dgvQueue.Columns.Contains("BotName"))
                dgvQueue.Columns["BotName"].Width = 110;

            dgvQueue.Columns["createdate"].HeaderText = "Created Date";
            dgvQueue.Columns["createdate"].Width = 165;
            dgvQueue.Columns["createdate"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

            dgvQueue.Columns["modifieddate"].HeaderText = "Modified Date";
            dgvQueue.Columns["modifieddate"].Width = 165;
            dgvQueue.Columns["modifieddate"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

            // DISABLE SELECT ALL - Prevent clicking on column header to select all
            dgvQueue.ColumnHeaderMouseClick -= DgvQueue_ColumnHeaderMouseClick;
            dgvQueue.ColumnHeaderMouseClick += DgvQueue_ColumnHeaderMouseClick;

            // Attach events once
            dgvQueue.CellContentClick -= DgvQueueCellContentClick;
            dgvQueue.CellValueChanged -= DgvQueueCellValueChanged;
            dgvQueue.CurrentCellDirtyStateChanged -= DgvQueueCurrentCellDirtyStateChanged;
            dgvQueue.RowPrePaint -= DgvQueueRowPrePaint;

            dgvQueue.CellContentClick += DgvQueueCellContentClick;
            dgvQueue.CellValueChanged += DgvQueueCellValueChanged;
            dgvQueue.CurrentCellDirtyStateChanged += DgvQueueCurrentCellDirtyStateChanged;
            dgvQueue.RowPrePaint += DgvQueueRowPrePaint;

            // Mark all non-select/status columns read-only
            foreach (DataGridViewColumn col in dgvQueue.Columns)
            {
                if (col.Name != "Select" && col.Name != "Status")
                    col.ReadOnly = true;
            }
        }

        private void DgvQueueCurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvQueue.IsCurrentCellDirty)
            {
                dgvQueue.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void DgvQueueCellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvQueue.Columns[e.ColumnIndex].Name == "Select")
            {
                dgvQueue.EndEdit();
                UpdateSelectedRecordsGrid();
            }
        }

        private void DgvQueueCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            string columnName = dgvQueue.Columns[e.ColumnIndex].Name;
            if (columnName == "Status")
            {
                var row = dgvQueue.Rows[e.RowIndex];
                int queueId = Convert.ToInt32(row.Cells["r26queueuid"].Value);
                string newStatusStr = row.Cells["Status"].Value?.ToString() ?? "";

                _individualStatusChanges[queueId] = newStatusStr;
                lblIndividualChanges.Text = $"Individual Changes: {_individualStatusChanges.Count}";
                dgvQueue.InvalidateRow(e.RowIndex);
            }
            else if (columnName == "Select")
            {
                UpdateSelectedRecordsGrid();
            }
        }

        private void DgvQueueRowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = dgvQueue.Rows[e.RowIndex];
            object idValue = row.Cells["r26queueuid"].Value;
            if (idValue == null) return;

            if (int.TryParse(idValue.ToString(), out int queueId))
            {
                row.DefaultCellStyle.BackColor = _individualStatusChanges.ContainsKey(queueId)
                    ? Color.LightYellow
                    : Color.White;
            }
        }

        private void UpdateSelectedRecordsGrid()
        {
            if (_dataTable == null) return;

            var selectedRows = _dataTable.AsEnumerable()
                .Where(r => r.Field<bool>("Select"))
                .Select(r => new
                {
                    QueueID = r.Field<int>("r26queueuid"),
                    CompanyName = r.Field<string>("CompanyName"),
                    PatNumber = r.Field<string>("PatNumber"),
                    Status = r.Field<string>("Status")
                })
                .ToList();

            var selectedTable = new DataTable();
            selectedTable.Columns.Add("Queue ID", typeof(int));
            selectedTable.Columns.Add("Company Name", typeof(string));
            selectedTable.Columns.Add("Pat Number", typeof(string));
            selectedTable.Columns.Add("Status", typeof(string));

            foreach (var item in selectedRows)
            {
                selectedTable.Rows.Add(item.QueueID, item.CompanyName, item.PatNumber, item.Status);
            }

            dgvSelectedRecords.DataSource = selectedTable;
            lblSelectedCount.Text = $"Selected: {selectedRows.Count:N0}";
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
                    progressBar.Visible = true;
                    btnSaveIndividual.Enabled = false;

                    int updatedCount = _repository.UpdateStatuses(_individualStatusChanges);

                    progressBar.Visible = false;
                    btnSaveIndividual.Enabled = true;

                    MessageBox.Show($"Successfully updated {updatedCount} individual records!",
                        "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    _individualStatusChanges.Clear();
                    lblIndividualChanges.Text = "Individual Changes: 0";
                    _ = LoadQueueDataAsync();
                }
                catch (Exception ex)
                {
                    progressBar.Visible = false;
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
                return;
            }

            if (cmbBulkStatus.SelectedItem == null)
            {
                MessageBox.Show("Please select a status from the dropdown.",
                    "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

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
                    progressBar.Visible = true;
                    btnBulkUpdate.Enabled = false;

                    var updates = new Dictionary<int, string>();
                    foreach (var row in selectedRows)
                    {
                        int queueId = Convert.ToInt32(row.Cells["r26queueuid"].Value);
                        updates[queueId] = newStatus;
                    }

                    int updatedCount = _repository.UpdateStatuses(updates);

                    progressBar.Visible = false;
                    btnBulkUpdate.Enabled = true;

                    MessageBox.Show($"Successfully updated {updatedCount} records via bulk update!",
                        "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    _ = LoadQueueDataAsync();
                }
                catch (Exception ex)
                {
                    progressBar.Visible = false;
                    btnBulkUpdate.Enabled = true;

                    MessageBox.Show($"Error updating bulk statuses: {ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            _ = LoadQueueDataAsync();
        }

        private void BtnCreatedDatePicker_Click(object sender, EventArgs e)
        {
            using (var calendarForm = new Form())
            {
                calendarForm.Text = "Select Created Date";
                calendarForm.Size = new Size(300, 300);
                calendarForm.StartPosition = FormStartPosition.CenterParent;
                calendarForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                calendarForm.MaximizeBox = false;
                calendarForm.MinimizeBox = false;

                var calendar = new MonthCalendar
                {
                    Location = new Point(20, 20)
                };

                if (_selectedCreatedDate.HasValue)
                {
                    calendar.SelectionStart = _selectedCreatedDate.Value;
                }

                var btnOK = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Location = new Point(60, 220),
                    Size = new Size(80, 30)
                };

                var btnClear = new Button
                {
                    Text = "Clear",
                    DialogResult = DialogResult.Retry,
                    Location = new Point(150, 220),
                    Size = new Size(80, 30)
                };

                calendarForm.Controls.Add(calendar);
                calendarForm.Controls.Add(btnOK);
                calendarForm.Controls.Add(btnClear);
                calendarForm.AcceptButton = btnOK;

                var result = calendarForm.ShowDialog(this);

                if (result == DialogResult.OK)
                {
                    _selectedCreatedDate = calendar.SelectionStart.Date;
                    lblCreatedDateValue.Text = _selectedCreatedDate.Value.ToString("yyyy-MM-dd");
                    lblCreatedDateValue.ForeColor = Color.FromArgb(0, 120, 212);
                    ApplyFilters();
                }
                else if (result == DialogResult.Retry)
                {
                    _selectedCreatedDate = null;
                    lblCreatedDateValue.Text = "Not Selected";
                    lblCreatedDateValue.ForeColor = Color.FromArgb(100, 100, 100);
                    ApplyFilters();
                }
            }
        }

        private void BtnModifiedDatePicker_Click(object sender, EventArgs e)
        {
            using (var calendarForm = new Form())
            {
                calendarForm.Text = "Select Modified Date";
                calendarForm.Size = new Size(300, 300);
                calendarForm.StartPosition = FormStartPosition.CenterParent;
                calendarForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                calendarForm.MaximizeBox = false;
                calendarForm.MinimizeBox = false;

                var calendar = new MonthCalendar
                {
                    Location = new Point(20, 20)
                };

                if (_selectedModifiedDate.HasValue)
                {
                    calendar.SelectionStart = _selectedModifiedDate.Value;
                }

                var btnOK = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Location = new Point(60, 220),
                    Size = new Size(80, 30)
                };

                var btnClear = new Button
                {
                    Text = "Clear",
                    DialogResult = DialogResult.Retry,
                    Location = new Point(150, 220),
                    Size = new Size(80, 30)
                };

                calendarForm.Controls.Add(calendar);
                calendarForm.Controls.Add(btnOK);
                calendarForm.Controls.Add(btnClear);
                calendarForm.AcceptButton = btnOK;

                var result = calendarForm.ShowDialog(this);

                if (result == DialogResult.OK)
                {
                    _selectedModifiedDate = calendar.SelectionStart.Date;
                    lblModifiedDateValue.Text = _selectedModifiedDate.Value.ToString("yyyy-MM-dd");
                    lblModifiedDateValue.ForeColor = Color.FromArgb(0, 120, 212);
                    ApplyFilters();
                }
                else if (result == DialogResult.Retry)
                {
                    _selectedModifiedDate = null;
                    lblModifiedDateValue.Text = "Not Selected";
                    lblModifiedDateValue.ForeColor = Color.FromArgb(100, 100, 100);
                    ApplyFilters();
                }
            }
        }

        private void R26QueueForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Check if user is closing the form (not just hiding it)
            if (e.CloseReason == CloseReason.UserClosing)
            {
                var result = MessageBox.Show(
                    "Are you sure you want to exit the application?",
                    "Confirm Exit",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Close Sam's Delivery form if it exists
                    if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
                    {
                        _samsDeliveryForm.Close();
                    }

                    // Exit the entire application
                    Application.Exit();
                }
                else
                {
                    // Cancel the close operation
                    e.Cancel = true;
                }
            }
        }

        private void MenuItemReports_Click(object sender, EventArgs e)
        {
            // Check if Sam's Delivery form already exists and is not disposed
            if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
            {
                // Show the existing form and bring it to front
                _samsDeliveryForm.ShowSamsDeliveryForm();
                this.Hide(); // ✅ Keep hiding (correct)
            }
            else
            {
                // Create new form with proper parameters
                DateTime targetDate = DateTime.Today;
                string connectionString = Program.DbConnectionString;
                string storedProcedureName = "SDgetSAMSReportStatus";

                _samsDeliveryForm = new SamsDeliveryReportForm(
                    targetDate,
                    _emailConfig,
                    connectionString,
                    storedProcedureName,
                    this
                );

                _samsDeliveryForm.FormClosed += (s, args) =>
                {
                    _samsDeliveryForm = null;
                    // ✅ ADD THIS: Show R26 form again when Sam's form is closed
                    this.Show();
                    this.BringToFront();
                };

                _samsDeliveryForm.Show();
                this.Hide(); // ✅ Keep hiding (correct)
            }
        }

        // Method to be called from HtmlViewerForm to show this form
        public void ShowR26Form()
        {
            this.Show();
            this.BringToFront();
            this.Focus();
        }
    }

    // Extension method for enabling double buffering on DataGridView
    public static class DataGridViewExtensions
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            System.Reflection.PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }
    }
}