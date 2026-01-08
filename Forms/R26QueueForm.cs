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
using System.Text;
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
        public bool _isExiting = false;

        // Reference to Sam's Delivery Report form
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
        private Button btnExport;

        private string _systemName;
        // Cache for status list from DB
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

        public R26QueueForm(EmailConfiguration emailConfig)
        {
            InitializeComponent();
            _emailConfig = emailConfig;
            _systemName = GetSystemName();

            // Use decrypted DB connection
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
            }

            this.Load += R26QueueForm_Load;
            this.FormClosing += R26QueueForm_FormClosing;
        }

        // Add this method to R26QueueForm class
        public void SetSamsDeliveryForm(SamsDeliveryReportForm samsForm)
        {
            _samsDeliveryForm = samsForm;
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

            // Menu Strip - Positioned at top with dropdown
            this.mainMenuStrip.BackColor = Color.White;
            this.mainMenuStrip.Dock = DockStyle.Top;
            this.mainMenuStrip.Font = new Font("Segoe UI", 10F);
            this.mainMenuStrip.Height = 35;
            this.mainMenuStrip.GripStyle = ToolStripGripStyle.Hidden;
            this.mainMenuStrip.RenderMode = ToolStripRenderMode.Professional;
            this.mainMenuStrip.Padding = new Padding(10, 5, 0, 5);

            // Create main dropdown menu
            var menuItemMain = new ToolStripMenuItem("Menu")
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.Black
            };

            // menuItemR26
            this.menuItemR26.Text = "R26 Daily Queue Management";
            this.menuItemR26.Checked = true;

            // menuItemReports
            this.menuItemReports.Text = "Sam's Delivery Report Status";
            this.menuItemReports.Click += MenuItemReports_Click;

            var menuItemSample1 = new ToolStripMenuItem("Sample 1");
            menuItemSample1.Click += MenuItemSample1_Click;

            var menuItemSample2 = new ToolStripMenuItem("Sample 2");
            menuItemSample2.Click += MenuItemSample2_Click;

            // Add submenu items to main menu
            menuItemMain.DropDownItems.Add(this.menuItemR26);
            menuItemMain.DropDownItems.Add(this.menuItemReports);
            menuItemMain.DropDownItems.Add(new ToolStripSeparator());
            menuItemMain.DropDownItems.Add(menuItemSample1);
            menuItemMain.DropDownItems.Add(menuItemSample2);

            this.mainMenuStrip.Items.Add(menuItemMain);

            // lblTitle
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new Font("Segoe UI", 18F, FontStyle.Bold);
            this.lblTitle.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblTitle.Location = new Point(13, 45);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new Size(400, 32);
            this.lblTitle.Text = "R26 Daily Queue Management";

            // lblRecordCount
            this.lblRecordCount.AutoSize = true;
            this.lblRecordCount.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblRecordCount.ForeColor = Color.Black;
            this.lblRecordCount.Location = new Point(15, 100);
            this.lblRecordCount.Name = "lblRecordCount";
            this.lblRecordCount.Size = new Size(150, 19);
            this.lblRecordCount.Text = "Total Records: 0";

            // lblSelectedCount
            this.lblSelectedCount.AutoSize = true;
            this.lblSelectedCount.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblSelectedCount.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblSelectedCount.Location = new Point(250, 97);
            this.lblSelectedCount.Name = "lblSelectedCount";
            this.lblSelectedCount.Size = new Size(150, 19);
            this.lblSelectedCount.Text = "Selected: 0";

            // lblIndividualChanges
            this.lblIndividualChanges.AutoSize = true;
            this.lblIndividualChanges.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblIndividualChanges.ForeColor = Color.FromArgb(255, 140, 0);
            this.lblIndividualChanges.Location = new Point(450, 95);
            this.lblIndividualChanges.Name = "lblIndividualChanges";
            this.lblIndividualChanges.Size = new Size(180, 19);
            this.lblIndividualChanges.Text = "Individual Changes: 0";

            // lblSearch
            this.lblSearch.AutoSize = true;
            this.lblSearch.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            this.lblSearch.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblSearch.Location = new Point(680, 95);
            this.lblSearch.Name = "lblSearch";
            this.lblSearch.Size = new Size(70, 20);
            this.lblSearch.Text = "Search";

            // txtSearch
            this.txtSearch.Font = new Font("Segoe UI", 11F);
            this.txtSearch.Location = new Point(760, 90);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new Size(300, 27);
            this.txtSearch.BorderStyle = BorderStyle.FixedSingle;
            this.txtSearch.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtSearch.TextChanged += TxtSearch_TextChanged;

            // Loading Panel - centered in DataGridView
            this.progressBar = new ProgressBar();
            this.progressBar.Style = ProgressBarStyle.Marquee;
            this.progressBar.Size = new Size(200, 30);
            this.progressBar.Visible = false;

            // Create a panel to hold loading indicator
            Panel loadingPanel = new Panel();
            loadingPanel.BackColor = Color.White;
            loadingPanel.Size = new Size(250, 100);
            loadingPanel.Visible = false;
            loadingPanel.Name = "loadingPanel";

            Label lblLoading = new Label();
            lblLoading.Text = "LOADING...";
            lblLoading.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            lblLoading.ForeColor = Color.FromArgb(52, 73, 94);
            lblLoading.AutoSize = true;
            lblLoading.Location = new Point(75, 55);

            this.progressBar.Location = new Point(25, 20);

            loadingPanel.Controls.Add(this.progressBar);
            loadingPanel.Controls.Add(lblLoading);

            this.Controls.Add(loadingPanel);
            loadingPanel.BringToFront();

            // lblCompanyFilter
            this.lblCompanyFilter.AutoSize = true;
            this.lblCompanyFilter.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            this.lblCompanyFilter.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblCompanyFilter.Location = new Point(15, 148);
            this.lblCompanyFilter.Name = "lblCompanyFilter";
            this.lblCompanyFilter.Text = "Company Name";

            // cmbCompanyFilter
            this.cmbCompanyFilter.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbCompanyFilter.Font = new Font("Segoe UI", 9.5F);
            this.cmbCompanyFilter.Location = new Point(165, 145);
            this.cmbCompanyFilter.Size = new Size(260, 28);
            this.cmbCompanyFilter.Name = "cmbCompanyFilter";
            this.cmbCompanyFilter.SelectedIndexChanged += FilterChanged;

            // lblStatusFilter
            this.lblStatusFilter.AutoSize = true;
            this.lblStatusFilter.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            this.lblStatusFilter.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblStatusFilter.Location = new Point(450, 149);
            this.lblStatusFilter.Name = "lblStatusFilter";
            this.lblStatusFilter.Text = "Status";

            // cmbStatusFilter
            this.cmbStatusFilter.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbStatusFilter.Font = new Font("Segoe UI", 9.5F);
            this.cmbStatusFilter.Location = new Point(520, 145);
            this.cmbStatusFilter.Size = new Size(180, 28);
            this.cmbStatusFilter.Name = "cmbStatusFilter";
            this.cmbStatusFilter.SelectedIndexChanged += FilterChanged;

            // lblCreatedDateFilter
            this.lblCreatedDateFilter.AutoSize = true;
            this.lblCreatedDateFilter.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            this.lblCreatedDateFilter.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblCreatedDateFilter.Location = new Point(730, 151);
            this.lblCreatedDateFilter.Name = "lblCreatedDateFilter";
            this.lblCreatedDateFilter.Text = "Created Date";

            // lblCreatedDateValue
            this.lblCreatedDateValue.Font = new Font("Segoe UI", 9.5F);
            this.lblCreatedDateValue.ForeColor = Color.FromArgb(100, 100, 100);
            this.lblCreatedDateValue.Location = new Point(860, 149);
            this.lblCreatedDateValue.Size = new Size(120, 28);
            this.lblCreatedDateValue.TextAlign = ContentAlignment.MiddleLeft;
            this.lblCreatedDateValue.Name = "lblCreatedDateValue";
            this.lblCreatedDateValue.Text = "Not Selected";

            // btnCreatedDatePicker
            this.btnCreatedDatePicker.BackColor = Color.Transparent;
            this.btnCreatedDatePicker.FlatStyle = FlatStyle.Flat;
            this.btnCreatedDatePicker.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.btnCreatedDatePicker.ForeColor = Color.Blue;
            this.btnCreatedDatePicker.Location = new Point(980, 144);
            this.btnCreatedDatePicker.Size = new Size(40, 42);
            this.btnCreatedDatePicker.FlatAppearance.BorderSize = 0;
            this.btnCreatedDatePicker.FlatAppearance.MouseOverBackColor = Color.LightGray;
            this.btnCreatedDatePicker.Name = "btnCreatedDatePicker";
            this.btnCreatedDatePicker.Text = "📅";
            this.btnCreatedDatePicker.TextAlign = ContentAlignment.MiddleCenter;
            this.btnCreatedDatePicker.UseVisualStyleBackColor = true;
            this.btnCreatedDatePicker.Padding = new Padding(0, 0, 0, 0);
            this.btnCreatedDatePicker.AutoSize = false;
            this.btnCreatedDatePicker.Cursor = Cursors.Hand;
            this.btnCreatedDatePicker.Click += BtnCreatedDatePicker_Click;

            // lblModifiedDateFilter
            this.lblModifiedDateFilter.AutoSize = true;
            this.lblModifiedDateFilter.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            this.lblModifiedDateFilter.ForeColor = Color.FromArgb(52, 73, 94);
            this.lblModifiedDateFilter.Location = new Point(1035, 153);
            this.lblModifiedDateFilter.Name = "lblModifiedDateFilter";
            this.lblModifiedDateFilter.Text = "Modified Date";

            // lblModifiedDateValue
            this.lblModifiedDateValue.Font = new Font("Segoe UI", 9.5F);
            this.lblModifiedDateValue.ForeColor = Color.FromArgb(100, 100, 100);
            this.lblModifiedDateValue.Location = new Point(1175, 153);
            this.lblModifiedDateValue.Size = new Size(120, 28);
            this.lblModifiedDateValue.TextAlign = ContentAlignment.MiddleLeft;
            this.lblModifiedDateValue.Name = "lblModifiedDateValue";
            this.lblModifiedDateValue.Text = "Not Selected";

            // btnModifiedDatePicker
            this.btnModifiedDatePicker.BackColor = Color.Transparent;
            this.btnModifiedDatePicker.FlatStyle = FlatStyle.Flat;
            this.btnModifiedDatePicker.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.btnModifiedDatePicker.ForeColor = Color.Blue;
            this.btnModifiedDatePicker.Location = new Point(1295, 146);
            this.btnModifiedDatePicker.Size = new Size(40, 40);
            this.btnModifiedDatePicker.FlatAppearance.BorderSize = 0;
            this.btnModifiedDatePicker.FlatAppearance.MouseOverBackColor = Color.LightGray;
            this.btnModifiedDatePicker.Name = "btnModifiedDatePicker";
            this.btnModifiedDatePicker.Text = "📅";
            this.btnModifiedDatePicker.TextAlign = ContentAlignment.MiddleCenter;
            this.btnModifiedDatePicker.UseVisualStyleBackColor = true;
            this.btnModifiedDatePicker.Padding = new Padding(0, 0, 0, 0);
            this.btnModifiedDatePicker.AutoSize = false;
            this.btnModifiedDatePicker.Cursor = Cursors.Hand;
            this.btnModifiedDatePicker.Click += BtnModifiedDatePicker_Click;

            // btnExport
            this.btnExport = new Button();
            this.btnExport.BackColor = Color.FromArgb(46, 125, 50);
            this.btnExport.FlatStyle = FlatStyle.Flat;
            this.btnExport.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.btnExport.ForeColor = Color.White;
            this.btnExport.Location = new Point(1420, 133);
            this.btnExport.Size = new Size(140, 50);
            this.btnExport.FlatAppearance.BorderSize = 0;
            this.btnExport.Name = "btnExport";
            this.btnExport.Text = "Export";
            this.btnExport.TextAlign = ContentAlignment.MiddleCenter;
            this.btnExport.UseVisualStyleBackColor = false;
            this.btnExport.Cursor = Cursors.Hand;
            this.btnExport.Click += BtnExport_Click;
            this.btnExport.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            // dgvQueue
            this.dgvQueue.AllowUserToAddRows = false;
            this.dgvQueue.AllowUserToDeleteRows = false;
            this.dgvQueue.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            this.dgvQueue.BackgroundColor = Color.White;
            this.dgvQueue.BorderStyle = BorderStyle.Fixed3D;
            this.dgvQueue.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvQueue.ColumnHeadersHeight = 45;
            this.dgvQueue.EnableHeadersVisualStyles = false;
            this.dgvQueue.Location = new Point(12, 210);
            this.dgvQueue.Name = "dgvQueue";
            this.dgvQueue.RowHeadersWidth = 30;
            this.dgvQueue.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.dgvQueue.Size = new Size(1050, 550);
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
            this.dgvQueue.DoubleBuffered(true);
            this.dgvQueue.RowTemplate.Height = 25;
            this.dgvQueue.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            this.dgvQueue.CellBorderStyle = DataGridViewCellBorderStyle.Single;

            // panelSide
            this.panelSide.BackColor = Color.FromArgb(240, 240, 240);
            this.panelSide.BorderStyle = BorderStyle.FixedSingle;
            this.panelSide.Location = new Point(1070, 210);
            this.panelSide.Name = "panelSide";
            this.panelSide.Size = new Size(500, 550);
            this.panelSide.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;

            // lblSideTitle
            this.lblSideTitle.AutoSize = true;
            this.lblSideTitle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            this.lblSideTitle.ForeColor = Color.FromArgb(0, 120, 212);
            this.lblSideTitle.Location = new Point(10, 6);
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
            this.dgvSelectedRecords.Size = new Size(478, 260);
            this.dgvSelectedRecords.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // lblBulkStatus
            this.lblBulkStatus.AutoSize = true;
            this.lblBulkStatus.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.lblBulkStatus.Location = new Point(10, 310);
            this.lblBulkStatus.Name = "lblBulkStatus";
            this.lblBulkStatus.Size = new Size(140, 19);
            this.lblBulkStatus.Text = "Select New Status";
            this.lblBulkStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            // cmbBulkStatus
            this.cmbBulkStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbBulkStatus.Font = new Font("Segoe UI", 11F);
            this.cmbBulkStatus.FormattingEnabled = true;
            this.cmbBulkStatus.Location = new Point(10, 340);
            this.cmbBulkStatus.Name = "cmbBulkStatus";
            this.cmbBulkStatus.Size = new Size(478, 28);
            this.cmbBulkStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // btnBulkUpdate
            this.btnBulkUpdate.BackColor = Color.FromArgb(0, 120, 212);
            this.btnBulkUpdate.FlatStyle = FlatStyle.Flat;
            this.btnBulkUpdate.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.btnBulkUpdate.ForeColor = Color.White;
            this.btnBulkUpdate.Location = new Point(10, 387);
            this.btnBulkUpdate.Name = "btnBulkUpdate";
            this.btnBulkUpdate.Size = new Size(478, 43);
            this.btnBulkUpdate.Text = "Update Selected (Bulk)";
            this.btnBulkUpdate.UseVisualStyleBackColor = false;
            this.btnBulkUpdate.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.btnBulkUpdate.Click += BtnBulkUpdate_Click;

            // btnSaveIndividual
            this.btnSaveIndividual.BackColor = Color.FromArgb(255, 140, 0);
            this.btnSaveIndividual.FlatStyle = FlatStyle.Flat;
            this.btnSaveIndividual.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.btnSaveIndividual.ForeColor = Color.White;
            this.btnSaveIndividual.Location = new Point(10, 437);
            this.btnSaveIndividual.Name = "btnSaveIndividual";
            this.btnSaveIndividual.Size = new Size(478, 43);
            this.btnSaveIndividual.Text = "Save Individual Changes";
            this.btnSaveIndividual.UseVisualStyleBackColor = false;
            this.btnSaveIndividual.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.btnSaveIndividual.Click += BtnSaveIndividual_Click;

            // btnRefresh
            this.btnRefresh.BackColor = Color.FromArgb(76, 175, 80);
            this.btnRefresh.FlatStyle = FlatStyle.Flat;
            this.btnRefresh.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.btnRefresh.ForeColor = Color.White;
            this.btnRefresh.Location = new Point(10, 487);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new Size(478, 43);
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
            this.Controls.Add(this.lblIndividualChanges);
            this.Controls.Add(this.lblSelectedCount);
            this.Controls.Add(this.lblRecordCount);
            this.Controls.Add(this.dgvQueue);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.btnExport);
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
            await LoadQueueDataAsync(isInitialLoad: true);
        }

        private async System.Threading.Tasks.Task LoadQueueDataAsync(bool isInitialLoad = false)
        {
            try
            {
                // SAVE CURRENT STATE BEFORE REFRESH
                var savedState = new GridViewState();

                if (!isInitialLoad)
                    if (!isInitialLoad)
                    {
                        // Save filter states
                        savedState.CompanyFilterIndex = cmbCompanyFilter.SelectedIndex;
                        savedState.StatusFilterIndex = cmbStatusFilter.SelectedIndex;
                        savedState.CreatedDate = _selectedCreatedDate;
                        savedState.ModifiedDate = _selectedModifiedDate;
                        savedState.SearchText = txtSearch.Text;

                        // Save scroll position
                        savedState.FirstDisplayedScrollingRowIndex = dgvQueue.FirstDisplayedScrollingRowIndex;
                        savedState.FirstDisplayedScrollingColumnIndex = dgvQueue.FirstDisplayedScrollingColumnIndex;

                        // Save selected rows (by Queue ID)
                        savedState.SelectedQueueIds = new HashSet<int>();
                        foreach (DataGridViewRow row in dgvQueue.Rows)
                        {
                            if (row.Cells["Select"].Value != null && (bool)row.Cells["Select"].Value == true)
                            {
                                int queueId = Convert.ToInt32(row.Cells["r26queueuid"].Value);
                                savedState.SelectedQueueIds.Add(queueId);
                            }
                        }

                        // Save individual status changes
                        savedState.IndividualChanges = new Dictionary<int, string>(_individualStatusChanges);
                    }
                    else
                    {
                        // INITIAL LOAD - Set default state
                        savedState.CompanyFilterIndex = -1; 
                        savedState.StatusFilterIndex = 0;
                        savedState.CreatedDate = null;
                        savedState.ModifiedDate = null;
                        savedState.SearchText = "";
                        savedState.FirstDisplayedScrollingRowIndex = 0;
                        savedState.FirstDisplayedScrollingColumnIndex = 0;
                        savedState.SelectedQueueIds = new HashSet<int>();
                        savedState.IndividualChanges = new Dictionary<int, string>();

                        // Clear existing changes on initial load
                        _individualStatusChanges.Clear();
                    }

                ShowLoadingIndicator(true);
                btnRefresh.Enabled = false;
                btnBulkUpdate.Enabled = false;
                btnSaveIndividual.Enabled = false;

                this.SuspendLayout();
                dgvQueue.SuspendLayout();
                dgvQueue.DataSource = null;

                // Load data from database
                await System.Threading.Tasks.Task.Run(() =>
                {
                    _queueData = _repository.GetAllR26Queue();
                });

                _dataTable = new DataTable();

                // Define columns
                _dataTable.Columns.Add("Select", typeof(bool));
                _dataTable.Columns.Add("r26queueuid", typeof(int));
                _dataTable.Columns.Add("CompanyName", typeof(string));
                _dataTable.Columns.Add("PatNumber", typeof(string));
                _dataTable.Columns.Add("Status", typeof(string));
                _dataTable.Columns.Add("createdate", typeof(DateTime));
                _dataTable.Columns.Add("modifieddate", typeof(DateTime));

                // Add rows
                _dataTable.BeginLoadData();
                foreach (var queue in _queueData)
                {
                    _dataTable.Rows.Add(
                        false,
                        queue.R26QueueUid,
                        queue.CompanyName ?? "Company " + queue.CompanyUid,
                        queue.PatNumber?.ToString() ?? "",
                        queue.Status ?? "Pending",
                        queue.CreatedDate ?? (object)DBNull.Value,
                        queue.ModifiedDate ?? (object)DBNull.Value
                    );
                }
                _dataTable.EndLoadData();

                _bindingSource.DataSource = _dataTable;

                // Load filters based on initial load or refresh
                if (isInitialLoad)
                {
                    // INITIAL LOAD - Load company filter and auto-select Texashvi
                    LoadCompanyFilter(); 
                    LoadStatusListsFromData();
                }
                else
                {
                    // REFRESH - Load filters without changing selection
                    LoadCompanyFilterWithoutChange(savedState.CompanyFilterIndex);
                    LoadStatusListsFromData();
                }

                // Bind data source
                dgvQueue.DataSource = _bindingSource;

                // Format grid
                FormatDataGridView();

                // RESTORE FILTER STATES based on load type
                if (isInitialLoad)
                {
                    // INITIAL LOAD - Texashvi is already selected by LoadCompanyFilter()
                    cmbStatusFilter.SelectedIndex = 0;
                    _selectedCreatedDate = null;
                    _selectedModifiedDate = null;
                    txtSearch.Text = "";

                    lblCreatedDateValue.Text = "Not Selected";
                    lblCreatedDateValue.ForeColor = Color.FromArgb(100, 100, 100);
                    lblModifiedDateValue.Text = "Not Selected";
                    lblModifiedDateValue.ForeColor = Color.FromArgb(100, 100, 100);
                }
                else
                {
                    // REFRESH - Restore all saved filter states
                    cmbCompanyFilter.SelectedIndex = savedState.CompanyFilterIndex;
                    cmbStatusFilter.SelectedIndex = savedState.StatusFilterIndex;
                    _selectedCreatedDate = savedState.CreatedDate;
                    _selectedModifiedDate = savedState.ModifiedDate;
                    txtSearch.Text = savedState.SearchText;

                    // Update date labels
                    lblCreatedDateValue.Text = _selectedCreatedDate.HasValue
                        ? _selectedCreatedDate.Value.ToString("yyyy-MM-dd")
                        : "Not Selected";
                    lblCreatedDateValue.ForeColor = _selectedCreatedDate.HasValue
                        ? Color.FromArgb(0, 120, 212)
                        : Color.FromArgb(100, 100, 100);

                    lblModifiedDateValue.Text = _selectedModifiedDate.HasValue
                        ? _selectedModifiedDate.Value.ToString("yyyy-MM-dd")
                        : "Not Selected";
                    lblModifiedDateValue.ForeColor = _selectedModifiedDate.HasValue
                        ? Color.FromArgb(0, 120, 212)
                        : Color.FromArgb(100, 100, 100);
                }

                // RESTORE SELECTIONS AND INDIVIDUAL CHANGES (only on refresh)
                if (!isInitialLoad)
                {
                    _individualStatusChanges = savedState.IndividualChanges;

                    foreach (DataGridViewRow row in dgvQueue.Rows)
                    {
                        int queueId = Convert.ToInt32(row.Cells["r26queueuid"].Value);

                        // Restore selection state
                        if (savedState.SelectedQueueIds.Contains(queueId))
                        {
                            row.Cells["Select"].Value = true;
                        }

                        // Restore individual status changes to both grid AND DataTable
                        if (_individualStatusChanges.ContainsKey(queueId))
                        {
                            string changedStatus = _individualStatusChanges[queueId];
                            row.Cells["Status"].Value = changedStatus;

                            // Also update DataTable
                            var dataRow = _dataTable.AsEnumerable()
                                .FirstOrDefault(r => r.Field<int>("r26queueuid") == queueId);
                            if (dataRow != null)
                            {
                                dataRow["Status"] = changedStatus;
                            }
                        }
                    }
                }

                RefreshGridHighlighting();

                // Apply filters (this will update record count)
                ApplyFilters();

                // RESTORE SCROLL POSITION (only on refresh)
                if (!isInitialLoad)
                {
                    try
                    {
                        if (savedState.FirstDisplayedScrollingRowIndex >= 0 &&
                            savedState.FirstDisplayedScrollingRowIndex < dgvQueue.RowCount)
                        {
                            dgvQueue.FirstDisplayedScrollingRowIndex = savedState.FirstDisplayedScrollingRowIndex;
                        }
                    }
                    catch { /* Ignore scroll position errors */ }
                }

                lblIndividualChanges.Text = $"Individual Changes: {_individualStatusChanges.Count}";
                UpdateSelectedRecordsGrid();

                dgvQueue.ResumeLayout();
                this.ResumeLayout();

                ShowLoadingIndicator(false);
                btnRefresh.Enabled = true;
                btnBulkUpdate.Enabled = true;
                btnSaveIndividual.Enabled = true;

                int displayedRecordCount = _bindingSource.Count;

                string messageText = isInitialLoad
                    ? $"Loaded {displayedRecordCount:N0} records successfully!"
                    : $"Refreshed {displayedRecordCount:N0} records successfully!";

                MessageBox.Show(messageText, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                dgvQueue.ResumeLayout();
                this.ResumeLayout();
                ShowLoadingIndicator(false);
                btnRefresh.Enabled = true;
                btnBulkUpdate.Enabled = true;
                btnSaveIndividual.Enabled = true;

                MessageBox.Show($"Error loading queue data: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private class GridViewState
        {
            public int CompanyFilterIndex { get; set; }
            public int StatusFilterIndex { get; set; }
            public DateTime? CreatedDate { get; set; }
            public DateTime? ModifiedDate { get; set; }
            public string SearchText { get; set; }
            public int FirstDisplayedScrollingRowIndex { get; set; }
            public int FirstDisplayedScrollingColumnIndex { get; set; }
            public HashSet<int> SelectedQueueIds { get; set; }
            public Dictionary<int, string> IndividualChanges { get; set; }
        }


        private void LoadCompanyFilterWithoutChange(int previousIndex)
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

            // Don't change selection - keep previous selection
            if (previousIndex >= 0 && previousIndex < cmbCompanyFilter.Items.Count)
            {
                cmbCompanyFilter.SelectedIndex = previousIndex;
            }
            else
            {
                cmbCompanyFilter.SelectedIndex = 0;
            }

            cmbCompanyFilter.EndUpdate();
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
        private void ShowLoadingIndicator(bool show)
        {
            Panel loadingPanel = this.Controls.Find("loadingPanel", false).FirstOrDefault() as Panel;

            if (loadingPanel != null)
            {
                if (show)
                {
                    // Center the panel in the DataGridView
                    int x = dgvQueue.Location.X + (dgvQueue.Width - loadingPanel.Width) / 2;
                    int y = dgvQueue.Location.Y + (dgvQueue.Height - loadingPanel.Height) / 2;
                    loadingPanel.Location = new Point(x, y);
                    loadingPanel.Visible = true;
                    progressBar.Visible = true;
                }
                else
                {
                    loadingPanel.Visible = false;
                    progressBar.Visible = false;
                }
            }
        }
        private void BtnExport_Click(object sender, EventArgs e)
        {
            try
            {
                // Check if there's data to export
                if (_bindingSource == null || _bindingSource.Count == 0)
                {
                    MessageBox.Show("No data available to export.",
                        "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Create SaveFileDialog
                using (var saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.FileName = $"R26_Queue_Export_{DateTime.Now:yyyyMMdd_HHmmss}";
                    saveFileDialog.Title = "Export R26 Queue Data";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ShowLoadingIndicator(true);
                        btnExport.Enabled = false;

                        // Get queue IDs from visible/filtered rows in BindingSource
                        var visibleQueueIds = new HashSet<int>();
                        foreach (DataRowView rowView in _bindingSource)
                        {
                            DataRow row = rowView.Row;
                            visibleQueueIds.Add((int)row["r26queueuid"]);
                        }

                        // Filter _queueData to get only visible records with ALL columns
                        var filteredData = _queueData
                            .Where(q => visibleQueueIds.Contains(q.R26QueueUid))
                            .ToList();

                        // Create export table with ALL columns from R26QueueModel
                        var dataToExport = new DataTable();

                        // Add ALL columns from your model
                        dataToExport.Columns.Add("Queue ID", typeof(string));
                        dataToExport.Columns.Add("Company UID", typeof(string));
                        dataToExport.Columns.Add("Company Name", typeof(string));
                        dataToExport.Columns.Add("Pat Number", typeof(string));
                        dataToExport.Columns.Add("Pat First Name", typeof(string));
                        dataToExport.Columns.Add("Pat Middle Initial", typeof(string));
                        dataToExport.Columns.Add("Pat Last Name", typeof(string));
                        dataToExport.Columns.Add("Pat Sex", typeof(string));
                        dataToExport.Columns.Add("Pat Birthdate", typeof(string));
                        dataToExport.Columns.Add("Location Number", typeof(string));
                        dataToExport.Columns.Add("Date", typeof(string));
                        dataToExport.Columns.Add("Time", typeof(string));
                        dataToExport.Columns.Add("Appt Type", typeof(string));
                        dataToExport.Columns.Add("Reason", typeof(string));
                        dataToExport.Columns.Add("Pat Address1", typeof(string));
                        dataToExport.Columns.Add("Pat Address2", typeof(string));
                        dataToExport.Columns.Add("Pat City", typeof(string));
                        dataToExport.Columns.Add("Pat State", typeof(string));
                        dataToExport.Columns.Add("Pat Zip5", typeof(string));
                        dataToExport.Columns.Add("Home Phone", typeof(string));
                        dataToExport.Columns.Add("Work Phone", typeof(string));
                        dataToExport.Columns.Add("Ticket Number", typeof(string));
                        dataToExport.Columns.Add("User Field1", typeof(string));
                        dataToExport.Columns.Add("User Field2", typeof(string));
                        dataToExport.Columns.Add("User Field3", typeof(string));
                        dataToExport.Columns.Add("User Field4", typeof(string));
                        dataToExport.Columns.Add("User Field5", typeof(string));
                        dataToExport.Columns.Add("RDr Number", typeof(string));
                        dataToExport.Columns.Add("Dictation ID", typeof(string));
                        dataToExport.Columns.Add("Admit Date", typeof(string));
                        dataToExport.Columns.Add("Discharge Date", typeof(string));
                        dataToExport.Columns.Add("Two Day Ago", typeof(string));
                        dataToExport.Columns.Add("Tomorrow", typeof(string));
                        dataToExport.Columns.Add("Two Days Ago Date Only", typeof(string));
                        dataToExport.Columns.Add("Tomorrow Date Only", typeof(string));
                        dataToExport.Columns.Add("Resource Name", typeof(string));
                        dataToExport.Columns.Add("Provider Name", typeof(string));
                        dataToExport.Columns.Add("Primary Insurance Name", typeof(string));
                        dataToExport.Columns.Add("Primary Ins Subscriber No", typeof(string));
                        dataToExport.Columns.Add("Primary Ins Group No", typeof(string));
                        dataToExport.Columns.Add("Primary Ins Copays", typeof(string));
                        dataToExport.Columns.Add("Patient Balance", typeof(string));
                        dataToExport.Columns.Add("Account Balance", typeof(string));
                        dataToExport.Columns.Add("Bot Details", typeof(string));
                        dataToExport.Columns.Add("Bot Name", typeof(string));
                        dataToExport.Columns.Add("Webhook", typeof(string));
                        dataToExport.Columns.Add("Created Date", typeof(string));
                        dataToExport.Columns.Add("Status", typeof(string));
                        dataToExport.Columns.Add("Modified Date", typeof(string));

                        // Populate export table with ALL fields from filtered data
                        foreach (var record in filteredData)
                        {
                            dataToExport.Rows.Add(
                                record.R26QueueUid.ToString(),
                                record.CompanyUid.ToString(),
                                record.CompanyName ?? "",
                                record.PatNumber?.ToString() ?? "",
                                record.PatFName ?? "",
                                record.PatMInitial ?? "",
                                record.PatLName ?? "",
                                record.PatSex ?? "",
                                record.PatBirthdate?.ToString("yyyy-MM-dd") ?? "",
                                record.LocationNumber?.ToString() ?? "",
                                record.Date?.ToString("yyyy-MM-dd") ?? "",
                                record.Time?.ToString(@"hh\:mm\:ss") ?? "",
                                record.ApptType ?? "",
                                record.Reason ?? "",
                                record.PatAddress1 ?? "",
                                record.PatAddress2 ?? "",
                                record.PatCity ?? "",
                                record.PatState ?? "",
                                record.PatZip5 ?? "",
                                record.HomePhone ?? "",
                                record.WorkPhone ?? "",
                                record.TicketNumber?.ToString() ?? "",
                                record.UserField1 ?? "",
                                record.UserField2 ?? "",
                                record.UserField3 ?? "",
                                record.UserField4 ?? "",
                                record.UserField5 ?? "",
                                record.RDrNumber?.ToString() ?? "",
                                record.DictationID?.ToString() ?? "",
                                record.AdmitDate ?? "",
                                record.DischargeDate ?? "",
                                record.TwoDayAgo?.ToString("yyyy-MM-dd HH:mm:ss") ?? "",
                                record.Tomorrow?.ToString("yyyy-MM-dd HH:mm:ss") ?? "",
                                record.TwoDaysAgoDateOnly?.ToString("yyyy-MM-dd") ?? "",
                                record.TomorrowDateOnly?.ToString("yyyy-MM-dd") ?? "",
                                record.ResourceName ?? "",
                                record.ProviderName ?? "",
                                record.PrimaryInsuranceName ?? "",
                                record.PrimaryInsSubscriberNo ?? "",
                                record.PrimaryInsGroupNo ?? "",
                                record.PrimaryInsCopays ?? "",
                                record.PatientBalance ?? "",
                                record.AccountBalance ?? "",
                                record.BotDetails ?? "",
                                record.BotName ?? "",
                                record.Webhook ?? "",
                                record.CreatedDate?.ToString("yyyy-MM-dd HH:mm:ss") ?? "",
                                record.Status ?? "",
                                record.ModifiedDate?.ToString("yyyy-MM-dd HH:mm:ss") ?? ""
                            );
                        }

                        // Export to CSV
                        ExportDataTableToCsv(dataToExport, saveFileDialog.FileName);

                        ShowLoadingIndicator(false);
                        btnExport.Enabled = true;

                        // Show success message
                        MessageBox.Show(
                            "Successfully Exported",
                            "Export Successful",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

                        // Ask user if they want to open the file
                        DialogResult openResult = MessageBox.Show(
                            "Would you like to open the exported file?",
                            "Open File",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (openResult == DialogResult.Yes)
                        {
                            try
                            {
                                System.Diagnostics.Process.Start(saveFileDialog.FileName);
                            }
                            catch (Exception openEx)
                            {
                                MessageBox.Show(
                                    $"Could not open file: {openEx.Message}",
                                    "Error Opening File",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowLoadingIndicator(false);
                btnExport.Enabled = true;

                MessageBox.Show($"Error exporting data: {ex.Message}",
                    "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportDataTableToCsv(DataTable dataTable, string filePath)
        {
            var csvContent = new StringBuilder();

            // Add header row
            var columnNames = dataTable.Columns.Cast<DataColumn>()
                .Select(column => EscapeCsvField(column.ColumnName));
            csvContent.AppendLine(string.Join(",", columnNames));

            // Add data rows
            foreach (DataRow row in dataTable.Rows)
            {
                var fields = row.ItemArray.Select(field =>
                    EscapeCsvField(field?.ToString() ?? ""));
                csvContent.AppendLine(string.Join(",", fields));
            }

            // Write to file
            File.WriteAllText(filePath, csvContent.ToString(), Encoding.UTF8);
        }

        private string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "\"\"";

            // If field contains comma, quote, or newline, wrap in quotes and escape quotes
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            }

            return field;
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

            int cnplcIndex = cmbCompanyFilter.Items.IndexOf("Texashvi");
            if (cnplcIndex != -1)
            {
                cmbCompanyFilter.SelectedIndex = cnplcIndex;
                ApplyFilters();
            }
            else
            {
                cmbCompanyFilter.SelectedIndex = 0; // All Companies
                                                    // Update record count for "All Companies"
                if (_dataTable != null)
                {
                    lblRecordCount.Text = $"Total Records: {_dataTable.Rows.Count:N0}";
                }
            }

            cmbCompanyFilter.EndUpdate();
        }

        private void ConfigureStatusColumn()
        {
            if (dgvQueue.Columns.Contains("Status"))
            {
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

                // CHANGE: Add "Pending" first, then add all DB statuses
                statusColumn.Items.Add("Pending");

                if (_statusList != null && _statusList.Count > 0)
                {
                    foreach (var status in _statusList)
                    {
                        if (status != "Pending") // Avoid duplicate if "Pending" exists in DB
                            statusColumn.Items.Add(status);
                    }
                }

                dgvQueue.Columns.Insert(statusIndex, statusColumn);
            }
        }


        private void RefreshStatusListFromDatabase()
        {
            try
            {
                // Re-fetch statuses from database to get latest statuses
                _statusList = _repository.GetAllStatuses();

                if (_statusList == null)
                {
                    _statusList = new List<string>();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing statuses from database: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusList = new List<string>();
            }
        }

        private void LoadStatusListsFromData()
        {
            // CRITICAL: Refresh status list from database to get latest statuses
            RefreshStatusListFromDatabase();

            if (_dataTable == null || _statusList == null) return;

            // Status filter combo - fetch fresh statuses from DB
            cmbStatusFilter.BeginUpdate();
            cmbStatusFilter.Items.Clear();
            cmbStatusFilter.Items.Add("All Status");

            // Add all statuses from DB (including newly added ones like "Pending")
            foreach (var s in _statusList)
                cmbStatusFilter.Items.Add(s);

            cmbStatusFilter.SelectedIndex = 0;
            cmbStatusFilter.EndUpdate();

            // Bulk status combo - "Pending" + all DB statuses
            cmbBulkStatus.BeginUpdate();
            cmbBulkStatus.Items.Clear();

            // Add "Pending" first (hardcoded)
            cmbBulkStatus.Items.Add("Pending");

            // Then add all other statuses from DB (avoid duplicates)
            foreach (var s in _statusList)
            {
                if (s != "Pending") // Don't add Pending twice
                    cmbBulkStatus.Items.Add(s);
            }

            cmbBulkStatus.SelectedIndex = 0;
            cmbBulkStatus.EndUpdate();
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
            "r26queueuid", "CompanyName", "PatNumber", "Status"
        };

                var searchConditions = new List<string>();
                foreach (string column in searchColumns)
                {
                    searchConditions.Add($"Convert([{column}], System.String) LIKE '%{searchText}%'");
                }

                filters.Add("(" + string.Join(" OR ", searchConditions) + ")");
            }

            string finalFilter = filters.Count == 0 ? null : string.Join(" AND ", filters);

            // Store current filter
            _bindingSource.Filter = finalFilter;

            // Sort individual changes to top after applying filters
            SortIndividualChangesToTop();

            // Update record count based on filtered view
            int displayedCount = 0;
            if (_bindingSource.DataSource is DataView dv)
            {
                displayedCount = dv.Count;
            }
            else
            {
                displayedCount = _bindingSource.Count;
            }

            lblRecordCount.Text = $"Total Records: {displayedCount:N0}";

            dgvQueue.Refresh();
        }


        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void SortIndividualChangesToTop()
        {
            if (_bindingSource == null || _dataTable == null) return;

            // Suspend binding to avoid flickering
            _bindingSource.RaiseListChangedEvents = false;

            try
            {
                // Get the current filter to preserve it
                string currentFilter = _bindingSource.Filter;

                // Build sort expression: rows with individual changes first
                // Add a computed column if it doesn't exist
                if (!_dataTable.Columns.Contains("HasIndividualChange"))
                {
                    _dataTable.Columns.Add("HasIndividualChange", typeof(int));
                }

                // Update the HasIndividualChange flag for all rows
                foreach (DataRow row in _dataTable.Rows)
                {
                    int queueId = Convert.ToInt32(row["r26queueuid"]);
                    row["HasIndividualChange"] = _individualStatusChanges.ContainsKey(queueId) ? 1 : 0;
                }

                // Create DataView with sort: individual changes first (descending), then by Queue ID (descending)
                DataView dv = new DataView(_dataTable)
                {
                    RowFilter = currentFilter,
                    Sort = "HasIndividualChange DESC, r26queueuid DESC"
                };

                // Update binding source
                _bindingSource.DataSource = dv;

                // Refresh highlighting
                RefreshGridHighlighting();
            }
            finally
            {
                _bindingSource.RaiseListChangedEvents = true;
                _bindingSource.ResetBindings(false);
            }
        }
        private void FormatDataGridView()
        {
            if (!dgvQueue.Columns.Contains("Select")) return;

            dgvQueue.Columns["Select"].Width = 60;
            dgvQueue.Columns["Select"].HeaderText = "Select";
            dgvQueue.Columns["Select"].ReadOnly = false;
            dgvQueue.Columns["Select"].Frozen = true;

            dgvQueue.Columns["r26queueuid"].HeaderText = "Queue ID";
            dgvQueue.Columns["r26queueuid"].Width = 130;
            dgvQueue.Columns["r26queueuid"].Frozen = true;

            dgvQueue.Columns["CompanyName"].HeaderText = "Company Name";
            dgvQueue.Columns["CompanyName"].Width = 233;
            dgvQueue.Columns["CompanyName"].Frozen = true;

            dgvQueue.Columns["PatNumber"].HeaderText = "Pat Number";
            dgvQueue.Columns["PatNumber"].Width = 133;

            // ✅ Configure Status column as ComboBox AFTER data is bound
            ConfigureStatusColumn();

            dgvQueue.Columns["createdate"].HeaderText = "Created Date";
            dgvQueue.Columns["createdate"].Width = 190;
            dgvQueue.Columns["createdate"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

            dgvQueue.Columns["modifieddate"].HeaderText = "Modified Date";
            dgvQueue.Columns["modifieddate"].Width = 190;
            dgvQueue.Columns["modifieddate"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

            // DISABLE SELECT ALL
            dgvQueue.ColumnHeaderMouseClick -= DgvQueue_ColumnHeaderMouseClick;
            dgvQueue.ColumnHeaderMouseClick += DgvQueue_ColumnHeaderMouseClick;

            // Attach events once
            dgvQueue.CellContentClick -= DgvQueueCellContentClick;
            dgvQueue.CellValueChanged -= DgvQueueCellValueChanged;
            dgvQueue.CurrentCellDirtyStateChanged -= DgvQueueCurrentCellDirtyStateChanged;
            dgvQueue.RowPrePaint -= DgvQueueRowPrePaint;
            dgvQueue.EditingControlShowing -= DgvQueue_EditingControlShowing;
            dgvQueue.DataError -= DgvQueue_DataError;

            dgvQueue.CellContentClick += DgvQueueCellContentClick;
            dgvQueue.CellValueChanged += DgvQueueCellValueChanged;
            dgvQueue.CurrentCellDirtyStateChanged += DgvQueueCurrentCellDirtyStateChanged;
            dgvQueue.RowPrePaint += DgvQueueRowPrePaint;
            dgvQueue.EditingControlShowing += DgvQueue_EditingControlShowing;
            dgvQueue.DataError += DgvQueue_DataError;

            // Mark all non-select/status columns read-only
            foreach (DataGridViewColumn col in dgvQueue.Columns)
            {
                if (col.Name != "Select" && col.Name != "Status")
                    col.ReadOnly = true;
            }
        }


        private void DgvQueue_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            // Only apply to Status column
            if (dgvQueue.CurrentCell != null && dgvQueue.CurrentCell.OwningColumn.Name == "Status")
            {
                ComboBox combo = e.Control as ComboBox;
                if (combo != null)
                {
                    // CRITICAL: Set DropDownList style to prevent typing
                    // User can ONLY select from dropdown, cannot type
                    combo.DropDownStyle = ComboBoxStyle.DropDownList;

                    // Optional: Remove any previous event handlers to avoid duplicates
                    combo.SelectedIndexChanged -= StatusCombo_SelectedIndexChanged;
                    combo.SelectedIndexChanged += StatusCombo_SelectedIndexChanged;
                }
            }
        }

        private void StatusCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Commit the change immediately when user selects from dropdown
            if (dgvQueue.IsCurrentCellDirty)
            {
                dgvQueue.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void DgvQueue_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // Handle ComboBox data errors (e.g., value not in list)
            if (dgvQueue.Columns[e.ColumnIndex].Name == "Status")
            {
                e.ThrowException = false;
                e.Cancel = false;
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

                var originalRecord = _queueData.FirstOrDefault(q => q.R26QueueUid == queueId);
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
                        .FirstOrDefault(r => r.Field<int>("r26queueuid") == queueId);

                    if (dataRow != null)
                    {
                        dataRow["Status"] = newStatusStr;
                    }
                }

                lblIndividualChanges.Text = $"Individual Changes: {_individualStatusChanges.Count}";

                // Re-sort to move changed row to top
                SortIndividualChangesToTop();
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
                // ✅ Check if this queue ID has individual changes
                if (_individualStatusChanges.ContainsKey(queueId))
                {
                    row.DefaultCellStyle.BackColor = Color.LightYellow;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = Color.White;
                }
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
                $"Are you sure you want to save {_individualStatusChanges.Count} individual status changes?\n\n",
                "Confirm Individual Updates",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    ShowLoadingIndicator(true);
                    btnSaveIndividual.Enabled = false;

                    // Pass system name to repository
                    int updatedCount = _repository.UpdateStatuses(_individualStatusChanges, _systemName);

                    ShowLoadingIndicator(false);
                    btnSaveIndividual.Enabled = true;

                    MessageBox.Show(
                        $"Successfully updated {updatedCount} individual records!\n",
                        "Update Successful",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    // Clear individual changes after successful save
                    _individualStatusChanges.Clear();
                    lblIndividualChanges.Text = "Individual Changes: 0";

                    // Reload with isInitialLoad: false to maintain current view
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
                $"Are you sure you want to update {selectedRows.Count} records to status '{newStatus}'?\n\n",
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
                        int queueId = Convert.ToInt32(row.Cells["r26queueuid"].Value);
                        updates[queueId] = newStatus;
                    }

                    // Pass system name to repository
                    int updatedCount = _repository.UpdateStatuses(updates, _systemName);

                    ShowLoadingIndicator(false);
                    btnBulkUpdate.Enabled = true;

                    MessageBox.Show(
                        $"Successfully updated {updatedCount} records via bulk update!\n",
                        "Update Successful",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    // CLEAR ALL SELECTIONS BEFORE RELOAD
                    foreach (DataGridViewRow row in dgvQueue.Rows)
                    {
                        if (row.Cells["Select"].Value != null)
                        {
                            row.Cells["Select"].Value = false;
                        }
                    }

                    // Update the DataTable to reflect cleared selections
                    if (_dataTable != null)
                    {
                        foreach (DataRow dataRow in _dataTable.Rows)
                        {
                            dataRow["Select"] = false;
                        }
                    }

                    // Clear the selected records grid immediately
                    UpdateSelectedRecordsGrid();

                    // Now reload the data
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

        private void RefreshGridHighlighting()
        {
            foreach (DataGridViewRow row in dgvQueue.Rows)
            {
                if (row.Cells["r26queueuid"].Value != null)
                {
                    int queueId = Convert.ToInt32(row.Cells["r26queueuid"].Value);

                    if (_individualStatusChanges.ContainsKey(queueId))
                    {
                        row.DefaultCellStyle.BackColor = Color.LightYellow;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                }
            }
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            _ = LoadQueueDataAsync(isInitialLoad: false);
        }

        private void BtnCreatedDatePicker_Click(object sender, EventArgs e)
        {
            using (var calendarForm = new Form())
            {
                calendarForm.Text = "Select Created Date";
                calendarForm.Size = new Size(300, 290);  // Slightly taller to accommodate calendar
                calendarForm.StartPosition = FormStartPosition.CenterParent;
                calendarForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                calendarForm.MaximizeBox = true;
                calendarForm.MinimizeBox = true;

                var calendar = new MonthCalendar
                {
                    Location = new Point(28, 20),
                    ShowToday = false,           
                    ShowTodayCircle = true     
                };

                if (_selectedCreatedDate.HasValue)
                {
                    calendar.SelectionStart = _selectedCreatedDate.Value;
                }

                var btnOK = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Location = new Point(60, 210),  
                    Size = new Size(80, 30)
                };

                var btnClear = new Button
                {
                    Text = "Clear",
                    DialogResult = DialogResult.Retry,
                    Location = new Point(150, 210),  
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
                calendarForm.Size = new Size(300, 290);  
                calendarForm.StartPosition = FormStartPosition.CenterParent;
                calendarForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                calendarForm.MaximizeBox = true;
                calendarForm.MinimizeBox = true;

                var calendar = new MonthCalendar
                {
                    Location = new Point(27, 20),
                    ShowToday = false,           
                    ShowTodayCircle = true    
                };

                if (_selectedModifiedDate.HasValue)
                {
                    calendar.SelectionStart = _selectedModifiedDate.Value;
                }

                var btnOK = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Location = new Point(60, 210),  
                    Size = new Size(80, 30)
                };

                var btnClear = new Button
                {
                    Text = "Clear",
                    DialogResult = DialogResult.Retry,
                    Location = new Point(150, 210),  
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
            // Prevent re-entry
            if (_isExiting)
            {
                return;
            }

            // If this is triggered by Application.Exit() from another form, don't show confirmation
            if (e.CloseReason == CloseReason.ApplicationExitCall)
            {
                return;
            }

            // Check if user is closing the form (not just hiding it)
            if (e.CloseReason == CloseReason.UserClosing)
            {
                // Set flag to prevent re-entry
                _isExiting = true;

                // Count pending changes
                int individualChangesCount = _individualStatusChanges?.Count ?? 0;

                // Count selected rows (bulk update pending)
                int selectedRowsCount = 0;
                if (dgvQueue?.Rows != null)
                {
                    selectedRowsCount = dgvQueue.Rows
                        .Cast<DataGridViewRow>()
                        .Count(r => r.Cells["Select"].Value != null && (bool)r.Cells["Select"].Value == true);
                }

                int totalPendingChanges = individualChangesCount + selectedRowsCount;

                // If there are pending changes, show warning message
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

                    message += "\nAre you sure you want to exit without saving these changes?";

                    var result = MessageBox.Show(
                        message,
                        "Unsaved Changes - Confirm Exit",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

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
                        // Reset flag if user cancels
                        _isExiting = false;
                        // Cancel the close operation
                        e.Cancel = true;
                    }
                }
                else
                {
                    // No pending changes, show normal exit confirmation
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
                        // Reset flag if user cancels
                        _isExiting = false;
                        // Cancel the close operation
                        e.Cancel = true;
                    }
                }
            }
        }

        private void MenuItemReports_Click(object sender, EventArgs e)
        {
            // Check if Sam's Delivery form already exists and is not disposed
            if (_samsDeliveryForm != null && !_samsDeliveryForm.IsDisposed)
            {
                // Just show the existing form
                _samsDeliveryForm.Show();
                _samsDeliveryForm.BringToFront();
                this.Hide();
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

                // FIXED: Handle when Sam's form is closed - only show R26 form if not disposed
                _samsDeliveryForm.FormClosed += (s, args) =>
                {
                    _samsDeliveryForm = null;

                    // Only try to show this form if it hasn't been disposed
                    if (!this.IsDisposed && !_isExiting)
                    {
                        this.Show();
                        this.BringToFront();
                    }
                };

                _samsDeliveryForm.Show();
                this.Hide();
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