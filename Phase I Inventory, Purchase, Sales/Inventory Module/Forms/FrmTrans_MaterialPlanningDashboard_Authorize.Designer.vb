<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_Authorize
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslMP_CreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslMP_CreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslMP_ModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslMP_ModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel
        Me.ToolStripLabel2 = New System.Windows.Forms.ToolStripLabel
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.txtFilterDocNo = New System.Windows.Forms.TextBox
        Me.btnFilter = New System.Windows.Forms.Button
        Me.dgvUser = New System.Windows.Forms.DataGridView
        Me.col4No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col4UserCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col4UserName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col4Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtpFilterTo = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpFilterFrom = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.chkFilterDate = New System.Windows.Forms.CheckBox
        Me.btnFilterFindMP = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboFilterStatus = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvAuth = New System.Windows.Forms.DataGridView
        Me.col1No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col1DocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col1Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col1Auth = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.col1Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvCancel = New System.Windows.Forms.DataGridView
        Me.col2No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col2DocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col2Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col2CancelAuth = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.col2Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel4 = New System.Windows.Forms.LinkLabel
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvAll = New System.Windows.Forms.DataGridView
        Me.col3No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col3DocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col3Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col3Status = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col3Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.txtFilterPrimaryMP = New System.Windows.Forms.TextBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.gboMP = New System.Windows.Forms.GroupBox
        Me.dtpMP_Filter = New System.Windows.Forms.DateTimePicker
        Me.SplitContainer5 = New System.Windows.Forms.SplitContainer
        Me.TxtMP_DaysPerMonth = New System.Windows.Forms.TextBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label29 = New System.Windows.Forms.Label
        Me.dtpMP = New System.Windows.Forms.DateTimePicker
        Me.dtpMP_RefDate = New System.Windows.Forms.DateTimePicker
        Me.TxtMP_Primary = New System.Windows.Forms.TextBox
        Me.TxtMP_DocNo = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.chkClose = New System.Windows.Forms.CheckBox
        Me.TxtMP_Remarks = New System.Windows.Forms.TextBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.chkMP_RequiredOnly = New System.Windows.Forms.CheckBox
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.Label33 = New System.Windows.Forms.Label
        Me.TxtMP_QtyRMReq = New System.Windows.Forms.TextBox
        Me.Label31 = New System.Windows.Forms.Label
        Me.TxtMP_QtyPrOB = New System.Windows.Forms.TextBox
        Me.TxtMP_QtySafetyStock = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtMP_QtySOB = New System.Windows.Forms.TextBox
        Me.TxtMP_QtyOptimumStock = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtMP_QtyPendingDO = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.TxtMP_QtyFree = New System.Windows.Forms.TextBox
        Me.Label37 = New System.Windows.Forms.Label
        Me.TxtMP_QtyMOB = New System.Windows.Forms.TextBox
        Me.Label36 = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtMP_QtyRequired = New System.Windows.Forms.TextBox
        Me.TxtMP_QtyOnHand = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.TxtMP_QtyPOB = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtMP_QtyGIT = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtMP_AvgPerMonth = New System.Windows.Forms.TextBox
        Me.cboMPOrderStrategy = New System.Windows.Forms.ComboBox
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.TxtMP_DiscAmount = New System.Windows.Forms.TextBox
        Me.Label32 = New System.Windows.Forms.Label
        Me.LlbMP_UnitPrice = New System.Windows.Forms.LinkLabel
        Me.TxtMP_UnitPrice = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.TxtMP_NettPrice = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.TxtMP_PreferedSupplier = New System.Windows.Forms.TextBox
        Me.TxtMP_CurrencyCode = New System.Windows.Forms.TextBox
        Me.TxtMP_Discount = New System.Windows.Forms.TextBox
        Me.TxtMP_DiscFromPerct = New System.Windows.Forms.TextBox
        Me.TxtMP_PrimaryCurrency = New System.Windows.Forms.TextBox
        Me.TxtMP_PrimarySupplier = New System.Windows.Forms.TextBox
        Me.TxtMP_PreferedSupplierDesc = New System.Windows.Forms.TextBox
        Me.llbMP_Ratio = New System.Windows.Forms.LinkLabel
        Me.TxtMP_QtyRequired2 = New System.Windows.Forms.TextBox
        Me.LblMP_Unit1 = New System.Windows.Forms.Label
        Me.LblRatio = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.TxtMP_SubTotal = New System.Windows.Forms.TextBox
        Me.cboMP_OrderUnit = New System.Windows.Forms.ComboBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.TxtMP_QtyOrder = New System.Windows.Forms.TextBox
        Me.TxtMP_PrimaryOrderUnit = New System.Windows.Forms.TextBox
        Me.dgvMP = New System.Windows.Forms.DataGridView
        Me.ColMP_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOnHand = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyPOB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyGIT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyInquiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyPendingDO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtySOB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyPrOB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtySafetyStock = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyLeadTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyRequired = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColMP_QtyOrderMaximum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOrderMinimum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOrderOptimum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyMax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyMin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOpt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOrder = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_OrderUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_OrderPoint = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColMP_NextOrderPoint = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColMP_ProcurementType = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColMP_ProcurementStrategy = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColMP_OrderStrategy = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColMP_FollowUp = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColMP_PreferedSupplier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_Currency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_UnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_Disc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_NettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_SubTotal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_PrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_PrimarySupplier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DefPrecCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DefPrecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_PrimaryOrderUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DefPrecOrderUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_OrderPointDelay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_NextOrderPointDelay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_LeadTimeDays = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOrderMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_PrimaryUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DefPrecUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_RatioMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_RatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DiscFromPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DetailFollowUp = New System.Windows.Forms.DataGridViewLinkColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.StatusStrip1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvAuth, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvCancel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvAll, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboMP.SuspendLayout()
        Me.SplitContainer5.Panel1.SuspendLayout()
        Me.SplitContainer5.Panel2.SuspendLayout()
        Me.SplitContainer5.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        CType(Me.dgvMP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslMP_CreatedBy, Me.tslMP_CreatedDate, Me.tslMP_ModifiedBy, Me.tslMP_ModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 489)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(934, 22)
        Me.StatusStrip1.TabIndex = 1
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslMP_CreatedBy
        '
        Me.tslMP_CreatedBy.Name = "tslMP_CreatedBy"
        Me.tslMP_CreatedBy.Size = New System.Drawing.Size(229, 17)
        Me.tslMP_CreatedBy.Spring = True
        Me.tslMP_CreatedBy.Text = "Created By: "
        '
        'tslMP_CreatedDate
        '
        Me.tslMP_CreatedDate.Name = "tslMP_CreatedDate"
        Me.tslMP_CreatedDate.Size = New System.Drawing.Size(229, 17)
        Me.tslMP_CreatedDate.Spring = True
        Me.tslMP_CreatedDate.Text = "Created Date:"
        '
        'tslMP_ModifiedBy
        '
        Me.tslMP_ModifiedBy.Name = "tslMP_ModifiedBy"
        Me.tslMP_ModifiedBy.Size = New System.Drawing.Size(229, 17)
        Me.tslMP_ModifiedBy.Spring = True
        Me.tslMP_ModifiedBy.Text = "Modified By: "
        '
        'tslMP_ModifiedDate
        '
        Me.tslMP_ModifiedDate.Name = "tslMP_ModifiedDate"
        Me.tslMP_ModifiedDate.Size = New System.Drawing.Size(229, 17)
        Me.tslMP_ModifiedDate.Spring = True
        Me.tslMP_ModifiedDate.Text = "Modified Date: "
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel1, Me.ToolStripLabel2})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 464)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(934, 25)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ToolStripLabel1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripLabel1.IsLink = True
        Me.ToolStripLabel1.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(25, 22)
        Me.ToolStripLabel1.Text = "Exit"
        Me.ToolStripLabel1.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'ToolStripLabel2
        '
        Me.ToolStripLabel2.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ToolStripLabel2.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripLabel2.IsLink = True
        Me.ToolStripLabel2.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.ToolStripLabel2.Name = "ToolStripLabel2"
        Me.ToolStripLabel2.Size = New System.Drawing.Size(58, 22)
        Me.ToolStripLabel2.Text = "Authorize"
        Me.ToolStripLabel2.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtFilterDocNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvUser)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFilterFindMP)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboFilterStatus)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TabControl1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtFilterPrimaryMP)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnExit)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.gboMP)
        Me.SplitContainer1.Size = New System.Drawing.Size(934, 464)
        Me.SplitContainer1.SplitterDistance = 237
        Me.SplitContainer1.TabIndex = 3
        '
        'txtFilterDocNo
        '
        Me.txtFilterDocNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFilterDocNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtFilterDocNo.Location = New System.Drawing.Point(56, 283)
        Me.txtFilterDocNo.MaxLength = 150
        Me.txtFilterDocNo.Name = "txtFilterDocNo"
        Me.txtFilterDocNo.Size = New System.Drawing.Size(137, 20)
        Me.txtFilterDocNo.TabIndex = 2
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(158, 441)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(75, 21)
        Me.btnFilter.TabIndex = 0
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'dgvUser
        '
        Me.dgvUser.AllowUserToResizeColumns = False
        Me.dgvUser.AllowUserToResizeRows = False
        Me.dgvUser.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.col4No, Me.col4UserCode, Me.col4UserName, Me.col4Primary})
        Me.dgvUser.Location = New System.Drawing.Point(4, 380)
        Me.dgvUser.MultiSelect = False
        Me.dgvUser.Name = "dgvUser"
        Me.dgvUser.RowHeadersWidth = 25
        Me.dgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvUser.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvUser.Size = New System.Drawing.Size(229, 59)
        Me.dgvUser.TabIndex = 5
        '
        'col4No
        '
        Me.col4No.Frozen = True
        Me.col4No.HeaderText = "No"
        Me.col4No.Name = "col4No"
        Me.col4No.ReadOnly = True
        Me.col4No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.col4No.Width = 27
        '
        'col4UserCode
        '
        Me.col4UserCode.HeaderText = "User Code"
        Me.col4UserCode.Name = "col4UserCode"
        Me.col4UserCode.Width = 82
        '
        'col4UserName
        '
        Me.col4UserName.HeaderText = "Name"
        Me.col4UserName.Name = "col4UserName"
        Me.col4UserName.ReadOnly = True
        Me.col4UserName.Width = 60
        '
        'col4Primary
        '
        Me.col4Primary.HeaderText = "Primary User"
        Me.col4Primary.Name = "col4Primary"
        Me.col4Primary.ReadOnly = True
        Me.col4Primary.Visible = False
        Me.col4Primary.Width = 91
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dtpFilterTo)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.dtpFilterFrom)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.chkFilterDate)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 309)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(223, 64)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'dtpFilterTo
        '
        Me.dtpFilterTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFilterTo.CustomFormat = "dd MMM yyyy"
        Me.dtpFilterTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFilterTo.Location = New System.Drawing.Point(49, 39)
        Me.dtpFilterTo.Name = "dtpFilterTo"
        Me.dtpFilterTo.Size = New System.Drawing.Size(168, 20)
        Me.dtpFilterTo.TabIndex = 8
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 45)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "To"
        '
        'dtpFilterFrom
        '
        Me.dtpFilterFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFilterFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFilterFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFilterFrom.Location = New System.Drawing.Point(49, 13)
        Me.dtpFilterFrom.Name = "dtpFilterFrom"
        Me.dtpFilterFrom.Size = New System.Drawing.Size(168, 20)
        Me.dtpFilterFrom.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "From"
        '
        'chkFilterDate
        '
        Me.chkFilterDate.AutoSize = True
        Me.chkFilterDate.Location = New System.Drawing.Point(1, -1)
        Me.chkFilterDate.Name = "chkFilterDate"
        Me.chkFilterDate.Size = New System.Drawing.Size(49, 17)
        Me.chkFilterDate.TabIndex = 5
        Me.chkFilterDate.Text = "Date"
        Me.chkFilterDate.UseVisualStyleBackColor = True
        '
        'btnFilterFindMP
        '
        Me.btnFilterFindMP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilterFindMP.Location = New System.Drawing.Point(196, 283)
        Me.btnFilterFindMP.Name = "btnFilterFindMP"
        Me.btnFilterFindMP.Size = New System.Drawing.Size(34, 21)
        Me.btnFilterFindMP.TabIndex = 3
        Me.btnFilterFindMP.Text = "..."
        Me.btnFilterFindMP.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 286)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Doc No"
        '
        'cboFilterStatus
        '
        Me.cboFilterStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboFilterStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFilterStatus.FormattingEnabled = True
        Me.cboFilterStatus.Items.AddRange(New Object() {"UnAuthorized Only", "Authorized Only", "All"})
        Me.cboFilterStatus.Location = New System.Drawing.Point(56, 255)
        Me.cboFilterStatus.Name = "cboFilterStatus"
        Me.cboFilterStatus.Size = New System.Drawing.Size(174, 21)
        Me.cboFilterStatus.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 259)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Status"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Top
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(237, 249)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.dgvAuth)
        Me.TabPage1.Controls.Add(Me.LinkLabel2)
        Me.TabPage1.Controls.Add(Me.LinkLabel1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(229, 223)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Authorize"
        '
        'dgvAuth
        '
        Me.dgvAuth.AllowUserToAddRows = False
        Me.dgvAuth.AllowUserToDeleteRows = False
        Me.dgvAuth.AllowUserToResizeColumns = False
        Me.dgvAuth.AllowUserToResizeRows = False
        Me.dgvAuth.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAuth.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAuth.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvAuth.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.col1No, Me.col1DocNo, Me.col1Date, Me.col1Auth, Me.col1Primary})
        Me.dgvAuth.Location = New System.Drawing.Point(3, 3)
        Me.dgvAuth.MultiSelect = False
        Me.dgvAuth.Name = "dgvAuth"
        Me.dgvAuth.RowHeadersVisible = False
        Me.dgvAuth.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvAuth.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvAuth.Size = New System.Drawing.Size(223, 201)
        Me.dgvAuth.TabIndex = 0
        '
        'col1No
        '
        Me.col1No.Frozen = True
        Me.col1No.HeaderText = "No"
        Me.col1No.Name = "col1No"
        Me.col1No.ReadOnly = True
        Me.col1No.Width = 46
        '
        'col1DocNo
        '
        Me.col1DocNo.HeaderText = "Document No"
        Me.col1DocNo.Name = "col1DocNo"
        Me.col1DocNo.ReadOnly = True
        Me.col1DocNo.Width = 98
        '
        'col1Date
        '
        Me.col1Date.HeaderText = "Date"
        Me.col1Date.Name = "col1Date"
        Me.col1Date.ReadOnly = True
        Me.col1Date.Width = 55
        '
        'col1Auth
        '
        Me.col1Auth.HeaderText = "Authorize"
        Me.col1Auth.Name = "col1Auth"
        Me.col1Auth.Width = 57
        '
        'col1Primary
        '
        Me.col1Primary.HeaderText = "Primary"
        Me.col1Primary.Name = "col1Primary"
        Me.col1Primary.ReadOnly = True
        Me.col1Primary.Visible = False
        Me.col1Primary.Width = 66
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(158, 207)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel2.TabIndex = 1
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(100, 207)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel1.TabIndex = 0
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage2.Controls.Add(Me.dgvCancel)
        Me.TabPage2.Controls.Add(Me.LinkLabel3)
        Me.TabPage2.Controls.Add(Me.LinkLabel4)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(229, 223)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Cancel Authorize"
        '
        'dgvCancel
        '
        Me.dgvCancel.AllowUserToAddRows = False
        Me.dgvCancel.AllowUserToDeleteRows = False
        Me.dgvCancel.AllowUserToResizeColumns = False
        Me.dgvCancel.AllowUserToResizeRows = False
        Me.dgvCancel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCancel.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCancel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCancel.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.col2No, Me.col2DocumentNo, Me.col2Date, Me.col2CancelAuth, Me.col2Primary})
        Me.dgvCancel.Location = New System.Drawing.Point(3, 3)
        Me.dgvCancel.MultiSelect = False
        Me.dgvCancel.Name = "dgvCancel"
        Me.dgvCancel.RowHeadersVisible = False
        Me.dgvCancel.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvCancel.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvCancel.Size = New System.Drawing.Size(223, 201)
        Me.dgvCancel.TabIndex = 4
        '
        'col2No
        '
        Me.col2No.Frozen = True
        Me.col2No.HeaderText = "No"
        Me.col2No.Name = "col2No"
        Me.col2No.ReadOnly = True
        Me.col2No.Width = 46
        '
        'col2DocumentNo
        '
        Me.col2DocumentNo.HeaderText = "Document No"
        Me.col2DocumentNo.Name = "col2DocumentNo"
        Me.col2DocumentNo.ReadOnly = True
        Me.col2DocumentNo.Width = 98
        '
        'col2Date
        '
        Me.col2Date.HeaderText = "Date"
        Me.col2Date.Name = "col2Date"
        Me.col2Date.ReadOnly = True
        Me.col2Date.Width = 55
        '
        'col2CancelAuth
        '
        Me.col2CancelAuth.HeaderText = "Cancel Authorize"
        Me.col2CancelAuth.Name = "col2CancelAuth"
        Me.col2CancelAuth.Width = 93
        '
        'col2Primary
        '
        Me.col2Primary.HeaderText = "Primary"
        Me.col2Primary.Name = "col2Primary"
        Me.col2Primary.ReadOnly = True
        Me.col2Primary.Visible = False
        Me.col2Primary.Width = 66
        '
        'LinkLabel3
        '
        Me.LinkLabel3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel3.AutoSize = True
        Me.LinkLabel3.Location = New System.Drawing.Point(158, 207)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel3.TabIndex = 3
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "Uncheck All"
        '
        'LinkLabel4
        '
        Me.LinkLabel4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel4.AutoSize = True
        Me.LinkLabel4.Location = New System.Drawing.Point(100, 207)
        Me.LinkLabel4.Name = "LinkLabel4"
        Me.LinkLabel4.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel4.TabIndex = 2
        Me.LinkLabel4.TabStop = True
        Me.LinkLabel4.Text = "Check All"
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage3.Controls.Add(Me.dgvAll)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(229, 223)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "All"
        '
        'dgvAll
        '
        Me.dgvAll.AllowUserToAddRows = False
        Me.dgvAll.AllowUserToDeleteRows = False
        Me.dgvAll.AllowUserToResizeColumns = False
        Me.dgvAll.AllowUserToResizeRows = False
        Me.dgvAll.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAll.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvAll.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.col3No, Me.col3DocNo, Me.col3Date, Me.col3Status, Me.col3Primary})
        Me.dgvAll.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAll.Location = New System.Drawing.Point(3, 3)
        Me.dgvAll.MultiSelect = False
        Me.dgvAll.Name = "dgvAll"
        Me.dgvAll.RowHeadersVisible = False
        Me.dgvAll.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvAll.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvAll.Size = New System.Drawing.Size(223, 217)
        Me.dgvAll.TabIndex = 5
        '
        'col3No
        '
        Me.col3No.Frozen = True
        Me.col3No.HeaderText = "No"
        Me.col3No.Name = "col3No"
        Me.col3No.ReadOnly = True
        Me.col3No.Width = 46
        '
        'col3DocNo
        '
        Me.col3DocNo.HeaderText = "Document No"
        Me.col3DocNo.Name = "col3DocNo"
        Me.col3DocNo.ReadOnly = True
        Me.col3DocNo.Width = 98
        '
        'col3Date
        '
        Me.col3Date.HeaderText = "Date"
        Me.col3Date.Name = "col3Date"
        Me.col3Date.ReadOnly = True
        Me.col3Date.Width = 55
        '
        'col3Status
        '
        Me.col3Status.HeaderText = "Status"
        Me.col3Status.Name = "col3Status"
        Me.col3Status.ReadOnly = True
        Me.col3Status.Width = 62
        '
        'col3Primary
        '
        Me.col3Primary.HeaderText = "Primary"
        Me.col3Primary.Name = "col3Primary"
        Me.col3Primary.ReadOnly = True
        Me.col3Primary.Visible = False
        Me.col3Primary.Width = 66
        '
        'txtFilterPrimaryMP
        '
        Me.txtFilterPrimaryMP.Location = New System.Drawing.Point(56, 283)
        Me.txtFilterPrimaryMP.Name = "txtFilterPrimaryMP"
        Me.txtFilterPrimaryMP.Size = New System.Drawing.Size(27, 20)
        Me.txtFilterPrimaryMP.TabIndex = 6
        Me.txtFilterPrimaryMP.Text = "-1"
        Me.txtFilterPrimaryMP.Visible = False
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(12, 403)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'gboMP
        '
        Me.gboMP.Controls.Add(Me.dtpMP_Filter)
        Me.gboMP.Controls.Add(Me.SplitContainer5)
        Me.gboMP.Controls.Add(Me.chkMP_RequiredOnly)
        Me.gboMP.Controls.Add(Me.TabControl2)
        Me.gboMP.Controls.Add(Me.dgvMP)
        Me.gboMP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboMP.Location = New System.Drawing.Point(0, 0)
        Me.gboMP.Margin = New System.Windows.Forms.Padding(2)
        Me.gboMP.Name = "gboMP"
        Me.gboMP.Padding = New System.Windows.Forms.Padding(2)
        Me.gboMP.Size = New System.Drawing.Size(693, 464)
        Me.gboMP.TabIndex = 45
        Me.gboMP.TabStop = False
        '
        'dtpMP_Filter
        '
        Me.dtpMP_Filter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpMP_Filter.CustomFormat = "dd MMM yyyy"
        Me.dtpMP_Filter.Enabled = False
        Me.dtpMP_Filter.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpMP_Filter.Location = New System.Drawing.Point(86, 438)
        Me.dtpMP_Filter.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpMP_Filter.Name = "dtpMP_Filter"
        Me.dtpMP_Filter.Size = New System.Drawing.Size(100, 20)
        Me.dtpMP_Filter.TabIndex = 5
        '
        'SplitContainer5
        '
        Me.SplitContainer5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer5.Location = New System.Drawing.Point(2, 11)
        Me.SplitContainer5.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer5.Name = "SplitContainer5"
        '
        'SplitContainer5.Panel1
        '
        Me.SplitContainer5.Panel1.Controls.Add(Me.TxtMP_DaysPerMonth)
        Me.SplitContainer5.Panel1.Controls.Add(Me.Label27)
        Me.SplitContainer5.Panel1.Controls.Add(Me.Label29)
        Me.SplitContainer5.Panel1.Controls.Add(Me.dtpMP)
        Me.SplitContainer5.Panel1.Controls.Add(Me.dtpMP_RefDate)
        Me.SplitContainer5.Panel1.Controls.Add(Me.TxtMP_Primary)
        Me.SplitContainer5.Panel1.Controls.Add(Me.TxtMP_DocNo)
        Me.SplitContainer5.Panel1.Controls.Add(Me.Label26)
        '
        'SplitContainer5.Panel2
        '
        Me.SplitContainer5.Panel2.Controls.Add(Me.chkClose)
        Me.SplitContainer5.Panel2.Controls.Add(Me.TxtMP_Remarks)
        Me.SplitContainer5.Panel2.Controls.Add(Me.Label28)
        Me.SplitContainer5.Size = New System.Drawing.Size(686, 70)
        Me.SplitContainer5.SplitterDistance = 325
        Me.SplitContainer5.SplitterWidth = 3
        Me.SplitContainer5.TabIndex = 0
        Me.SplitContainer5.TabStop = False
        '
        'TxtMP_DaysPerMonth
        '
        Me.TxtMP_DaysPerMonth.Location = New System.Drawing.Point(70, 23)
        Me.TxtMP_DaysPerMonth.Name = "TxtMP_DaysPerMonth"
        Me.TxtMP_DaysPerMonth.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_DaysPerMonth.TabIndex = 73
        Me.TxtMP_DaysPerMonth.TabStop = False
        Me.TxtMP_DaysPerMonth.Text = "0"
        Me.TxtMP_DaysPerMonth.Visible = False
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(1, 27)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(30, 13)
        Me.Label27.TabIndex = 72
        Me.Label27.Text = "Date"
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(1, 51)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(53, 13)
        Me.Label29.TabIndex = 76
        Me.Label29.Text = "Ref. Date"
        '
        'dtpMP
        '
        Me.dtpMP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpMP.CustomFormat = "dd MMM yyyy"
        Me.dtpMP.Enabled = False
        Me.dtpMP.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpMP.Location = New System.Drawing.Point(96, 23)
        Me.dtpMP.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpMP.Name = "dtpMP"
        Me.dtpMP.Size = New System.Drawing.Size(227, 20)
        Me.dtpMP.TabIndex = 71
        '
        'dtpMP_RefDate
        '
        Me.dtpMP_RefDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpMP_RefDate.CustomFormat = "dd MMM yyyy"
        Me.dtpMP_RefDate.Enabled = False
        Me.dtpMP_RefDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpMP_RefDate.Location = New System.Drawing.Point(96, 47)
        Me.dtpMP_RefDate.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpMP_RefDate.Name = "dtpMP_RefDate"
        Me.dtpMP_RefDate.Size = New System.Drawing.Size(227, 20)
        Me.dtpMP_RefDate.TabIndex = 75
        '
        'TxtMP_Primary
        '
        Me.TxtMP_Primary.Location = New System.Drawing.Point(69, 0)
        Me.TxtMP_Primary.Name = "TxtMP_Primary"
        Me.TxtMP_Primary.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_Primary.TabIndex = 70
        Me.TxtMP_Primary.TabStop = False
        Me.TxtMP_Primary.Text = "-1"
        Me.TxtMP_Primary.Visible = False
        '
        'TxtMP_DocNo
        '
        Me.TxtMP_DocNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_DocNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMP_DocNo.Location = New System.Drawing.Point(96, 0)
        Me.TxtMP_DocNo.Name = "TxtMP_DocNo"
        Me.TxtMP_DocNo.ReadOnly = True
        Me.TxtMP_DocNo.Size = New System.Drawing.Size(226, 20)
        Me.TxtMP_DocNo.TabIndex = 52
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(1, 2)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(73, 13)
        Me.Label26.TabIndex = 51
        Me.Label26.Text = "Document No"
        '
        'chkClose
        '
        Me.chkClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkClose.AutoSize = True
        Me.chkClose.Enabled = False
        Me.chkClose.Location = New System.Drawing.Point(55, 50)
        Me.chkClose.Name = "chkClose"
        Me.chkClose.Size = New System.Drawing.Size(58, 17)
        Me.chkClose.TabIndex = 87
        Me.chkClose.Text = "Closed"
        Me.chkClose.UseVisualStyleBackColor = True
        '
        'TxtMP_Remarks
        '
        Me.TxtMP_Remarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_Remarks.Location = New System.Drawing.Point(55, 0)
        Me.TxtMP_Remarks.Multiline = True
        Me.TxtMP_Remarks.Name = "TxtMP_Remarks"
        Me.TxtMP_Remarks.ReadOnly = True
        Me.TxtMP_Remarks.Size = New System.Drawing.Size(300, 45)
        Me.TxtMP_Remarks.TabIndex = 74
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(1, 4)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(49, 13)
        Me.Label28.TabIndex = 73
        Me.Label28.Text = "Remarks"
        '
        'chkMP_RequiredOnly
        '
        Me.chkMP_RequiredOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkMP_RequiredOnly.AutoSize = True
        Me.chkMP_RequiredOnly.Checked = True
        Me.chkMP_RequiredOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMP_RequiredOnly.Enabled = False
        Me.chkMP_RequiredOnly.Location = New System.Drawing.Point(2, 440)
        Me.chkMP_RequiredOnly.Name = "chkMP_RequiredOnly"
        Me.chkMP_RequiredOnly.Size = New System.Drawing.Size(79, 17)
        Me.chkMP_RequiredOnly.TabIndex = 4
        Me.chkMP_RequiredOnly.Text = "Order Point"
        Me.chkMP_RequiredOnly.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Location = New System.Drawing.Point(2, 256)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(686, 179)
        Me.TabControl2.TabIndex = 2
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.SplitContainer2)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Size = New System.Drawing.Size(678, 153)
        Me.TabPage4.TabIndex = 0
        Me.TabPage4.Text = "Detail Qty"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(2, 2)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label33)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyRMReq)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label31)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyPrOB)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtySafetyStock)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label16)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtySOB)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyOptimumStock)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyPendingDO)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label17)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.BackColor = System.Drawing.SystemColors.Control
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyFree)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label37)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyMOB)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label36)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label30)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyRequired)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyOnHand)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label18)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyPOB)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyGIT)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_AvgPerMonth)
        Me.SplitContainer2.Panel2.Controls.Add(Me.cboMPOrderStrategy)
        Me.SplitContainer2.Size = New System.Drawing.Size(680, 158)
        Me.SplitContainer2.SplitterDistance = 317
        Me.SplitContainer2.SplitterWidth = 3
        Me.SplitContainer2.TabIndex = 71
        Me.SplitContainer2.TabStop = False
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(3, 81)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(95, 13)
        Me.Label33.TabIndex = 71
        Me.Label33.Text = "Raw Material Req."
        '
        'TxtMP_QtyRMReq
        '
        Me.TxtMP_QtyRMReq.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyRMReq.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyRMReq.Location = New System.Drawing.Point(101, 78)
        Me.TxtMP_QtyRMReq.Name = "TxtMP_QtyRMReq"
        Me.TxtMP_QtyRMReq.ReadOnly = True
        Me.TxtMP_QtyRMReq.Size = New System.Drawing.Size(213, 20)
        Me.TxtMP_QtyRMReq.TabIndex = 70
        Me.TxtMP_QtyRMReq.Text = "0"
        Me.TxtMP_QtyRMReq.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(3, 7)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(81, 13)
        Me.Label31.TabIndex = 69
        Me.Label31.Text = "Purch. Request"
        '
        'TxtMP_QtyPrOB
        '
        Me.TxtMP_QtyPrOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyPrOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyPrOB.Location = New System.Drawing.Point(101, 6)
        Me.TxtMP_QtyPrOB.Name = "TxtMP_QtyPrOB"
        Me.TxtMP_QtyPrOB.ReadOnly = True
        Me.TxtMP_QtyPrOB.Size = New System.Drawing.Size(213, 20)
        Me.TxtMP_QtyPrOB.TabIndex = 0
        Me.TxtMP_QtyPrOB.Text = "0"
        Me.TxtMP_QtyPrOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_QtySafetyStock
        '
        Me.TxtMP_QtySafetyStock.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtySafetyStock.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtySafetyStock.Location = New System.Drawing.Point(101, 128)
        Me.TxtMP_QtySafetyStock.Name = "TxtMP_QtySafetyStock"
        Me.TxtMP_QtySafetyStock.ReadOnly = True
        Me.TxtMP_QtySafetyStock.Size = New System.Drawing.Size(213, 20)
        Me.TxtMP_QtySafetyStock.TabIndex = 4
        Me.TxtMP_QtySafetyStock.Text = "0"
        Me.TxtMP_QtySafetyStock.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 32)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(75, 13)
        Me.Label11.TabIndex = 57
        Me.Label11.Text = "Sales Balance"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(3, 130)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(68, 13)
        Me.Label16.TabIndex = 67
        Me.Label16.Text = "Safety Stock"
        '
        'TxtMP_QtySOB
        '
        Me.TxtMP_QtySOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtySOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtySOB.Location = New System.Drawing.Point(101, 30)
        Me.TxtMP_QtySOB.Name = "TxtMP_QtySOB"
        Me.TxtMP_QtySOB.ReadOnly = True
        Me.TxtMP_QtySOB.Size = New System.Drawing.Size(213, 20)
        Me.TxtMP_QtySOB.TabIndex = 1
        Me.TxtMP_QtySOB.Text = "0"
        Me.TxtMP_QtySOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_QtyOptimumStock
        '
        Me.TxtMP_QtyOptimumStock.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyOptimumStock.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyOptimumStock.Location = New System.Drawing.Point(101, 103)
        Me.TxtMP_QtyOptimumStock.Name = "TxtMP_QtyOptimumStock"
        Me.TxtMP_QtyOptimumStock.ReadOnly = True
        Me.TxtMP_QtyOptimumStock.Size = New System.Drawing.Size(213, 20)
        Me.TxtMP_QtyOptimumStock.TabIndex = 3
        Me.TxtMP_QtyOptimumStock.Text = "0"
        Me.TxtMP_QtyOptimumStock.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(3, 57)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(65, 13)
        Me.Label14.TabIndex = 59
        Me.Label14.Text = "Pending DO"
        '
        'TxtMP_QtyPendingDO
        '
        Me.TxtMP_QtyPendingDO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyPendingDO.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyPendingDO.Location = New System.Drawing.Point(101, 54)
        Me.TxtMP_QtyPendingDO.Name = "TxtMP_QtyPendingDO"
        Me.TxtMP_QtyPendingDO.ReadOnly = True
        Me.TxtMP_QtyPendingDO.Size = New System.Drawing.Size(213, 20)
        Me.TxtMP_QtyPendingDO.TabIndex = 2
        Me.TxtMP_QtyPendingDO.Text = "0"
        Me.TxtMP_QtyPendingDO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(3, 106)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(79, 13)
        Me.Label17.TabIndex = 65
        Me.Label17.Text = "Optimum Stock"
        '
        'TxtMP_QtyFree
        '
        Me.TxtMP_QtyFree.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyFree.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyFree.Location = New System.Drawing.Point(270, 6)
        Me.TxtMP_QtyFree.Name = "TxtMP_QtyFree"
        Me.TxtMP_QtyFree.ReadOnly = True
        Me.TxtMP_QtyFree.Size = New System.Drawing.Size(80, 20)
        Me.TxtMP_QtyFree.TabIndex = 1
        Me.TxtMP_QtyFree.Text = "0"
        Me.TxtMP_QtyFree.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(217, 8)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(47, 13)
        Me.Label37.TabIndex = 77
        Me.Label37.Text = "Qty Free"
        '
        'TxtMP_QtyMOB
        '
        Me.TxtMP_QtyMOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyMOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyMOB.Location = New System.Drawing.Point(107, 77)
        Me.TxtMP_QtyMOB.Name = "TxtMP_QtyMOB"
        Me.TxtMP_QtyMOB.ReadOnly = True
        Me.TxtMP_QtyMOB.Size = New System.Drawing.Size(243, 20)
        Me.TxtMP_QtyMOB.TabIndex = 4
        Me.TxtMP_QtyMOB.Text = "0"
        Me.TxtMP_QtyMOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(3, 80)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(100, 13)
        Me.Label36.TabIndex = 75
        Me.Label36.Text = "Production Balance"
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(3, 104)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(81, 13)
        Me.Label30.TabIndex = 73
        Me.Label30.Text = "Avg. Per Month"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 104)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(75, 13)
        Me.Label6.TabIndex = 71
        Me.Label6.Text = "Order Strategy"
        Me.Label6.Visible = False
        '
        'TxtMP_QtyRequired
        '
        Me.TxtMP_QtyRequired.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyRequired.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyRequired.Location = New System.Drawing.Point(107, 128)
        Me.TxtMP_QtyRequired.Name = "TxtMP_QtyRequired"
        Me.TxtMP_QtyRequired.ReadOnly = True
        Me.TxtMP_QtyRequired.Size = New System.Drawing.Size(243, 20)
        Me.TxtMP_QtyRequired.TabIndex = 6
        Me.TxtMP_QtyRequired.Text = "0"
        Me.TxtMP_QtyRequired.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_QtyOnHand
        '
        Me.TxtMP_QtyOnHand.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyOnHand.Location = New System.Drawing.Point(107, 6)
        Me.TxtMP_QtyOnHand.Name = "TxtMP_QtyOnHand"
        Me.TxtMP_QtyOnHand.ReadOnly = True
        Me.TxtMP_QtyOnHand.Size = New System.Drawing.Size(104, 20)
        Me.TxtMP_QtyOnHand.TabIndex = 0
        Me.TxtMP_QtyOnHand.Text = "0"
        Me.TxtMP_QtyOnHand.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(3, 130)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(50, 13)
        Me.Label18.TabIndex = 69
        Me.Label18.Text = "Required"
        '
        'TxtMP_QtyPOB
        '
        Me.TxtMP_QtyPOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyPOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyPOB.Location = New System.Drawing.Point(107, 54)
        Me.TxtMP_QtyPOB.Name = "TxtMP_QtyPOB"
        Me.TxtMP_QtyPOB.ReadOnly = True
        Me.TxtMP_QtyPOB.Size = New System.Drawing.Size(243, 20)
        Me.TxtMP_QtyPOB.TabIndex = 3
        Me.TxtMP_QtyPOB.Text = "0"
        Me.TxtMP_QtyPOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 57)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(94, 13)
        Me.Label10.TabIndex = 53
        Me.Label10.Text = "Purchase Balance"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(3, 8)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(50, 13)
        Me.Label15.TabIndex = 63
        Me.Label15.Text = "On Hand"
        '
        'TxtMP_QtyGIT
        '
        Me.TxtMP_QtyGIT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyGIT.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyGIT.Location = New System.Drawing.Point(107, 30)
        Me.TxtMP_QtyGIT.Name = "TxtMP_QtyGIT"
        Me.TxtMP_QtyGIT.ReadOnly = True
        Me.TxtMP_QtyGIT.Size = New System.Drawing.Size(243, 20)
        Me.TxtMP_QtyGIT.TabIndex = 2
        Me.TxtMP_QtyGIT.Text = "0"
        Me.TxtMP_QtyGIT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 32)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(85, 13)
        Me.Label9.TabIndex = 51
        Me.Label9.Text = "Goods In Transit"
        '
        'TxtMP_AvgPerMonth
        '
        Me.TxtMP_AvgPerMonth.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_AvgPerMonth.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_AvgPerMonth.Location = New System.Drawing.Point(107, 100)
        Me.TxtMP_AvgPerMonth.Name = "TxtMP_AvgPerMonth"
        Me.TxtMP_AvgPerMonth.ReadOnly = True
        Me.TxtMP_AvgPerMonth.Size = New System.Drawing.Size(243, 20)
        Me.TxtMP_AvgPerMonth.TabIndex = 5
        Me.TxtMP_AvgPerMonth.Text = "0"
        Me.TxtMP_AvgPerMonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboMPOrderStrategy
        '
        Me.cboMPOrderStrategy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboMPOrderStrategy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMPOrderStrategy.FormattingEnabled = True
        Me.cboMPOrderStrategy.Items.AddRange(New Object() {"Order To Min Qty", "Order To Optimum Qty", "Order To Max Qty", "Order To Safety Stock", "Order By Min Order Qty", "Order By Opt Order Qty", "Order By Max Order Qty"})
        Me.cboMPOrderStrategy.Location = New System.Drawing.Point(107, 100)
        Me.cboMPOrderStrategy.Margin = New System.Windows.Forms.Padding(2)
        Me.cboMPOrderStrategy.Name = "cboMPOrderStrategy"
        Me.cboMPOrderStrategy.Size = New System.Drawing.Size(243, 21)
        Me.cboMPOrderStrategy.TabIndex = 5
        Me.cboMPOrderStrategy.Visible = False
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.SplitContainer3)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Size = New System.Drawing.Size(678, 153)
        Me.TabPage5.TabIndex = 1
        Me.TabPage5.Text = "Detail Purchase"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(2, 2)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.SplitContainer3.Panel1.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer3.Panel2Collapsed = True
        Me.SplitContainer3.Size = New System.Drawing.Size(674, 149)
        Me.SplitContainer3.SplitterDistance = 304
        Me.SplitContainer3.SplitterWidth = 3
        Me.SplitContainer3.TabIndex = 72
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_DiscAmount)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label32)
        Me.SplitContainer4.Panel1.Controls.Add(Me.LlbMP_UnitPrice)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_UnitPrice)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label22)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label20)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label19)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_NettPrice)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label13)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_PreferedSupplier)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_CurrencyCode)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_Discount)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_DiscFromPerct)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_PrimaryCurrency)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_PrimarySupplier)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtMP_PreferedSupplierDesc)
        Me.SplitContainer4.Panel2.Controls.Add(Me.llbMP_Ratio)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtMP_QtyRequired2)
        Me.SplitContainer4.Panel2.Controls.Add(Me.LblMP_Unit1)
        Me.SplitContainer4.Panel2.Controls.Add(Me.LblRatio)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label24)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label25)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtMP_SubTotal)
        Me.SplitContainer4.Panel2.Controls.Add(Me.cboMP_OrderUnit)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label23)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtMP_QtyOrder)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtMP_PrimaryOrderUnit)
        Me.SplitContainer4.Size = New System.Drawing.Size(680, 124)
        Me.SplitContainer4.SplitterDistance = 325
        Me.SplitContainer4.SplitterWidth = 3
        Me.SplitContainer4.TabIndex = 83
        Me.SplitContainer4.TabStop = False
        '
        'TxtMP_DiscAmount
        '
        Me.TxtMP_DiscAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_DiscAmount.Location = New System.Drawing.Point(165, 71)
        Me.TxtMP_DiscAmount.Name = "TxtMP_DiscAmount"
        Me.TxtMP_DiscAmount.ReadOnly = True
        Me.TxtMP_DiscAmount.Size = New System.Drawing.Size(153, 20)
        Me.TxtMP_DiscAmount.TabIndex = 8
        Me.TxtMP_DiscAmount.Text = "0"
        Me.TxtMP_DiscAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(142, 75)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(15, 13)
        Me.Label32.TabIndex = 73
        Me.Label32.Text = "%"
        '
        'LlbMP_UnitPrice
        '
        Me.LlbMP_UnitPrice.AutoSize = True
        Me.LlbMP_UnitPrice.Location = New System.Drawing.Point(3, 52)
        Me.LlbMP_UnitPrice.Name = "LlbMP_UnitPrice"
        Me.LlbMP_UnitPrice.Size = New System.Drawing.Size(53, 13)
        Me.LlbMP_UnitPrice.TabIndex = 5
        Me.LlbMP_UnitPrice.TabStop = True
        Me.LlbMP_UnitPrice.Text = "Unit Price"
        Me.LlbMP_UnitPrice.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtMP_UnitPrice
        '
        Me.TxtMP_UnitPrice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_UnitPrice.Location = New System.Drawing.Point(98, 48)
        Me.TxtMP_UnitPrice.Name = "TxtMP_UnitPrice"
        Me.TxtMP_UnitPrice.ReadOnly = True
        Me.TxtMP_UnitPrice.Size = New System.Drawing.Size(221, 20)
        Me.TxtMP_UnitPrice.TabIndex = 5
        Me.TxtMP_UnitPrice.Text = "0"
        Me.TxtMP_UnitPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(3, 75)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(49, 13)
        Me.Label22.TabIndex = 7
        Me.Label22.Text = "Discount"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(3, 102)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(54, 13)
        Me.Label20.TabIndex = 67
        Me.Label20.Text = "Nett Price"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(3, 29)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(49, 13)
        Me.Label19.TabIndex = 57
        Me.Label19.Text = "Currency"
        '
        'TxtMP_NettPrice
        '
        Me.TxtMP_NettPrice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_NettPrice.Location = New System.Drawing.Point(98, 98)
        Me.TxtMP_NettPrice.Name = "TxtMP_NettPrice"
        Me.TxtMP_NettPrice.ReadOnly = True
        Me.TxtMP_NettPrice.Size = New System.Drawing.Size(221, 20)
        Me.TxtMP_NettPrice.TabIndex = 9
        Me.TxtMP_NettPrice.TabStop = False
        Me.TxtMP_NettPrice.Text = "0"
        Me.TxtMP_NettPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(3, 6)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(88, 13)
        Me.Label13.TabIndex = 55
        Me.Label13.Text = "Prefered Supplier"
        '
        'TxtMP_PreferedSupplier
        '
        Me.TxtMP_PreferedSupplier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_PreferedSupplier.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMP_PreferedSupplier.Location = New System.Drawing.Point(98, 2)
        Me.TxtMP_PreferedSupplier.Name = "TxtMP_PreferedSupplier"
        Me.TxtMP_PreferedSupplier.ReadOnly = True
        Me.TxtMP_PreferedSupplier.Size = New System.Drawing.Size(221, 20)
        Me.TxtMP_PreferedSupplier.TabIndex = 1
        '
        'TxtMP_CurrencyCode
        '
        Me.TxtMP_CurrencyCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_CurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMP_CurrencyCode.Location = New System.Drawing.Point(98, 25)
        Me.TxtMP_CurrencyCode.Name = "TxtMP_CurrencyCode"
        Me.TxtMP_CurrencyCode.ReadOnly = True
        Me.TxtMP_CurrencyCode.Size = New System.Drawing.Size(221, 20)
        Me.TxtMP_CurrencyCode.TabIndex = 3
        '
        'TxtMP_Discount
        '
        Me.TxtMP_Discount.Location = New System.Drawing.Point(98, 71)
        Me.TxtMP_Discount.Name = "TxtMP_Discount"
        Me.TxtMP_Discount.ReadOnly = True
        Me.TxtMP_Discount.Size = New System.Drawing.Size(39, 20)
        Me.TxtMP_Discount.TabIndex = 6
        Me.TxtMP_Discount.Text = "0.00"
        Me.TxtMP_Discount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_DiscFromPerct
        '
        Me.TxtMP_DiscFromPerct.Location = New System.Drawing.Point(98, 71)
        Me.TxtMP_DiscFromPerct.Name = "TxtMP_DiscFromPerct"
        Me.TxtMP_DiscFromPerct.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_DiscFromPerct.TabIndex = 74
        Me.TxtMP_DiscFromPerct.TabStop = False
        Me.TxtMP_DiscFromPerct.Text = "1"
        Me.TxtMP_DiscFromPerct.Visible = False
        '
        'TxtMP_PrimaryCurrency
        '
        Me.TxtMP_PrimaryCurrency.Location = New System.Drawing.Point(98, 25)
        Me.TxtMP_PrimaryCurrency.Name = "TxtMP_PrimaryCurrency"
        Me.TxtMP_PrimaryCurrency.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_PrimaryCurrency.TabIndex = 72
        Me.TxtMP_PrimaryCurrency.TabStop = False
        Me.TxtMP_PrimaryCurrency.Text = "-1"
        Me.TxtMP_PrimaryCurrency.Visible = False
        '
        'TxtMP_PrimarySupplier
        '
        Me.TxtMP_PrimarySupplier.Location = New System.Drawing.Point(98, 2)
        Me.TxtMP_PrimarySupplier.Name = "TxtMP_PrimarySupplier"
        Me.TxtMP_PrimarySupplier.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_PrimarySupplier.TabIndex = 69
        Me.TxtMP_PrimarySupplier.TabStop = False
        Me.TxtMP_PrimarySupplier.Text = "-1"
        Me.TxtMP_PrimarySupplier.Visible = False
        '
        'TxtMP_PreferedSupplierDesc
        '
        Me.TxtMP_PreferedSupplierDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_PreferedSupplierDesc.Location = New System.Drawing.Point(3, 2)
        Me.TxtMP_PreferedSupplierDesc.Name = "TxtMP_PreferedSupplierDesc"
        Me.TxtMP_PreferedSupplierDesc.ReadOnly = True
        Me.TxtMP_PreferedSupplierDesc.Size = New System.Drawing.Size(341, 20)
        Me.TxtMP_PreferedSupplierDesc.TabIndex = 0
        Me.TxtMP_PreferedSupplierDesc.TabStop = False
        '
        'llbMP_Ratio
        '
        Me.llbMP_Ratio.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbMP_Ratio.AutoSize = True
        Me.llbMP_Ratio.Location = New System.Drawing.Point(77, 52)
        Me.llbMP_Ratio.Name = "llbMP_Ratio"
        Me.llbMP_Ratio.Size = New System.Drawing.Size(32, 13)
        Me.llbMP_Ratio.TabIndex = 3
        Me.llbMP_Ratio.TabStop = True
        Me.llbMP_Ratio.Text = "Ratio"
        Me.llbMP_Ratio.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtMP_QtyRequired2
        '
        Me.TxtMP_QtyRequired2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyRequired2.Location = New System.Drawing.Point(77, 25)
        Me.TxtMP_QtyRequired2.Name = "TxtMP_QtyRequired2"
        Me.TxtMP_QtyRequired2.ReadOnly = True
        Me.TxtMP_QtyRequired2.Size = New System.Drawing.Size(202, 20)
        Me.TxtMP_QtyRequired2.TabIndex = 1
        Me.TxtMP_QtyRequired2.TabStop = False
        Me.TxtMP_QtyRequired2.Text = "0"
        Me.TxtMP_QtyRequired2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblMP_Unit1
        '
        Me.LblMP_Unit1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblMP_Unit1.AutoSize = True
        Me.LblMP_Unit1.Location = New System.Drawing.Point(286, 29)
        Me.LblMP_Unit1.Name = "LblMP_Unit1"
        Me.LblMP_Unit1.Size = New System.Drawing.Size(26, 13)
        Me.LblMP_Unit1.TabIndex = 2
        Me.LblMP_Unit1.Text = "Unit"
        '
        'LblRatio
        '
        Me.LblRatio.AutoSize = True
        Me.LblRatio.Location = New System.Drawing.Point(3, 52)
        Me.LblRatio.Name = "LblRatio"
        Me.LblRatio.Size = New System.Drawing.Size(32, 13)
        Me.LblRatio.TabIndex = 81
        Me.LblRatio.Text = "Ratio"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(3, 102)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(53, 13)
        Me.Label24.TabIndex = 75
        Me.Label24.Text = "Sub Total"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(3, 29)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(69, 13)
        Me.Label25.TabIndex = 79
        Me.Label25.Text = "Required Qty"
        '
        'TxtMP_SubTotal
        '
        Me.TxtMP_SubTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_SubTotal.Location = New System.Drawing.Point(77, 98)
        Me.TxtMP_SubTotal.Name = "TxtMP_SubTotal"
        Me.TxtMP_SubTotal.ReadOnly = True
        Me.TxtMP_SubTotal.Size = New System.Drawing.Size(268, 20)
        Me.TxtMP_SubTotal.TabIndex = 6
        Me.TxtMP_SubTotal.TabStop = False
        Me.TxtMP_SubTotal.Text = "0"
        Me.TxtMP_SubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboMP_OrderUnit
        '
        Me.cboMP_OrderUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboMP_OrderUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMP_OrderUnit.Enabled = False
        Me.cboMP_OrderUnit.FormattingEnabled = True
        Me.cboMP_OrderUnit.Location = New System.Drawing.Point(276, 71)
        Me.cboMP_OrderUnit.Margin = New System.Windows.Forms.Padding(2)
        Me.cboMP_OrderUnit.Name = "cboMP_OrderUnit"
        Me.cboMP_OrderUnit.Size = New System.Drawing.Size(69, 21)
        Me.cboMP_OrderUnit.TabIndex = 5
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(3, 75)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(52, 13)
        Me.Label23.TabIndex = 73
        Me.Label23.Text = "Order Qty"
        '
        'TxtMP_QtyOrder
        '
        Me.TxtMP_QtyOrder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyOrder.Location = New System.Drawing.Point(77, 71)
        Me.TxtMP_QtyOrder.Name = "TxtMP_QtyOrder"
        Me.TxtMP_QtyOrder.ReadOnly = True
        Me.TxtMP_QtyOrder.Size = New System.Drawing.Size(194, 20)
        Me.TxtMP_QtyOrder.TabIndex = 4
        Me.TxtMP_QtyOrder.Text = "0"
        Me.TxtMP_QtyOrder.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_PrimaryOrderUnit
        '
        Me.TxtMP_PrimaryOrderUnit.Location = New System.Drawing.Point(77, 71)
        Me.TxtMP_PrimaryOrderUnit.Name = "TxtMP_PrimaryOrderUnit"
        Me.TxtMP_PrimaryOrderUnit.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_PrimaryOrderUnit.TabIndex = 78
        Me.TxtMP_PrimaryOrderUnit.TabStop = False
        Me.TxtMP_PrimaryOrderUnit.Text = "-1"
        Me.TxtMP_PrimaryOrderUnit.Visible = False
        '
        'dgvMP
        '
        Me.dgvMP.AllowUserToAddRows = False
        Me.dgvMP.AllowUserToDeleteRows = False
        Me.dgvMP.AllowUserToResizeColumns = False
        Me.dgvMP.AllowUserToResizeRows = False
        Me.dgvMP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMP.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvMP.ColumnHeadersHeight = 25
        Me.dgvMP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvMP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColMP_No, Me.ColMP_CodeNo, Me.ColMP_Desc, Me.ColMP_Unit, Me.ColMP_QtyOnHand, Me.ColMP_QtyPOB, Me.ColMP_QtyGIT, Me.ColMP_QtyInquiry, Me.ColMP_QtyPendingDO, Me.ColMP_QtySOB, Me.ColMP_QtyPrOB, Me.ColMP_QtySafetyStock, Me.ColMP_QtyLeadTime, Me.ColMP_QtyRequired, Me.ColMP_QtyOrderMaximum, Me.ColMP_QtyOrderMinimum, Me.ColMP_QtyOrderOptimum, Me.ColMP_QtyMax, Me.ColMP_QtyMin, Me.ColMP_QtyOpt, Me.ColMP_QtyOrder, Me.ColMP_OrderUnit, Me.ColMP_OrderPoint, Me.ColMP_NextOrderPoint, Me.ColMP_ProcurementType, Me.ColMP_ProcurementStrategy, Me.ColMP_OrderStrategy, Me.ColMP_FollowUp, Me.ColMP_PreferedSupplier, Me.ColMP_Currency, Me.ColMP_UnitPrice, Me.ColMP_Disc, Me.ColMP_NettPrice, Me.ColMP_SubTotal, Me.ColMP_PrimaryCurrency, Me.ColMP_PrimarySupplier, Me.ColMP_PrimaryCodeNo, Me.ColMP_DefPrecCurrency, Me.ColMP_DefPrecUnit, Me.ColMP_PrimaryOrderUnit, Me.ColMP_DefPrecOrderUnit, Me.ColMP_OrderPointDelay, Me.ColMP_NextOrderPointDelay, Me.ColMP_LeadTimeDays, Me.ColMP_QtyOrderMulti, Me.ColMP_PrimaryUnitMulti, Me.ColMP_DefPrecUnitMulti, Me.ColMP_RatioMulti, Me.ColMP_RatioDefault, Me.ColMP_DiscAmount, Me.ColMP_DiscFromPerct, Me.ColMP_DetailFollowUp})
        Me.dgvMP.Location = New System.Drawing.Point(2, 86)
        Me.dgvMP.MultiSelect = False
        Me.dgvMP.Name = "dgvMP"
        Me.dgvMP.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvMP.RowTemplate.Height = 24
        Me.dgvMP.Size = New System.Drawing.Size(686, 168)
        Me.dgvMP.TabIndex = 1
        '
        'ColMP_No
        '
        Me.ColMP_No.Frozen = True
        Me.ColMP_No.HeaderText = "No"
        Me.ColMP_No.Name = "ColMP_No"
        Me.ColMP_No.Width = 46
        '
        'ColMP_CodeNo
        '
        Me.ColMP_CodeNo.Frozen = True
        Me.ColMP_CodeNo.HeaderText = "Code No"
        Me.ColMP_CodeNo.Name = "ColMP_CodeNo"
        Me.ColMP_CodeNo.Width = 74
        '
        'ColMP_Desc
        '
        Me.ColMP_Desc.HeaderText = "Description"
        Me.ColMP_Desc.Name = "ColMP_Desc"
        Me.ColMP_Desc.Width = 85
        '
        'ColMP_Unit
        '
        Me.ColMP_Unit.HeaderText = "Unit"
        Me.ColMP_Unit.Name = "ColMP_Unit"
        Me.ColMP_Unit.Width = 51
        '
        'ColMP_QtyOnHand
        '
        Me.ColMP_QtyOnHand.HeaderText = "Qty On Hand"
        Me.ColMP_QtyOnHand.Name = "ColMP_QtyOnHand"
        Me.ColMP_QtyOnHand.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyOnHand.Visible = False
        Me.ColMP_QtyOnHand.Width = 94
        '
        'ColMP_QtyPOB
        '
        Me.ColMP_QtyPOB.HeaderText = "Qty Purchase Order Balance"
        Me.ColMP_QtyPOB.Name = "ColMP_QtyPOB"
        Me.ColMP_QtyPOB.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyPOB.Width = 167
        '
        'ColMP_QtyGIT
        '
        Me.ColMP_QtyGIT.HeaderText = "Qty Goods In Transit"
        Me.ColMP_QtyGIT.Name = "ColMP_QtyGIT"
        Me.ColMP_QtyGIT.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyGIT.Width = 129
        '
        'ColMP_QtyInquiry
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMP_QtyInquiry.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColMP_QtyInquiry.HeaderText = "Qty Inquiry"
        Me.ColMP_QtyInquiry.Name = "ColMP_QtyInquiry"
        Me.ColMP_QtyInquiry.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyInquiry.Width = 82
        '
        'ColMP_QtyPendingDO
        '
        Me.ColMP_QtyPendingDO.HeaderText = "Qty Pending DO"
        Me.ColMP_QtyPendingDO.Name = "ColMP_QtyPendingDO"
        Me.ColMP_QtyPendingDO.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyPendingDO.Visible = False
        Me.ColMP_QtyPendingDO.Width = 109
        '
        'ColMP_QtySOB
        '
        Me.ColMP_QtySOB.HeaderText = "Qty Sales Order Balance"
        Me.ColMP_QtySOB.Name = "ColMP_QtySOB"
        Me.ColMP_QtySOB.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtySOB.Width = 148
        '
        'ColMP_QtyPrOB
        '
        Me.ColMP_QtyPrOB.HeaderText = "Qty PrOB"
        Me.ColMP_QtyPrOB.Name = "ColMP_QtyPrOB"
        Me.ColMP_QtyPrOB.Visible = False
        Me.ColMP_QtyPrOB.Width = 76
        '
        'ColMP_QtySafetyStock
        '
        Me.ColMP_QtySafetyStock.HeaderText = "Qty Safety Stock"
        Me.ColMP_QtySafetyStock.Name = "ColMP_QtySafetyStock"
        Me.ColMP_QtySafetyStock.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtySafetyStock.Width = 112
        '
        'ColMP_QtyLeadTime
        '
        Me.ColMP_QtyLeadTime.HeaderText = "Qty Lead Time"
        Me.ColMP_QtyLeadTime.Name = "ColMP_QtyLeadTime"
        Me.ColMP_QtyLeadTime.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyLeadTime.Width = 101
        '
        'ColMP_QtyRequired
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMP_QtyRequired.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColMP_QtyRequired.HeaderText = "Qty Required"
        Me.ColMP_QtyRequired.Name = "ColMP_QtyRequired"
        Me.ColMP_QtyRequired.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyRequired.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColMP_QtyRequired.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColMP_QtyRequired.Width = 94
        '
        'ColMP_QtyOrderMaximum
        '
        Me.ColMP_QtyOrderMaximum.HeaderText = "Qty Order Maximum"
        Me.ColMP_QtyOrderMaximum.Name = "ColMP_QtyOrderMaximum"
        Me.ColMP_QtyOrderMaximum.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyOrderMaximum.Width = 124
        '
        'ColMP_QtyOrderMinimum
        '
        Me.ColMP_QtyOrderMinimum.HeaderText = "Qty Order Minimum"
        Me.ColMP_QtyOrderMinimum.Name = "ColMP_QtyOrderMinimum"
        Me.ColMP_QtyOrderMinimum.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyOrderMinimum.Width = 121
        '
        'ColMP_QtyOrderOptimum
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMP_QtyOrderOptimum.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColMP_QtyOrderOptimum.HeaderText = "Qty Order Optimum"
        Me.ColMP_QtyOrderOptimum.Name = "ColMP_QtyOrderOptimum"
        Me.ColMP_QtyOrderOptimum.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyOrderOptimum.Width = 121
        '
        'ColMP_QtyMax
        '
        Me.ColMP_QtyMax.HeaderText = "Qty Max"
        Me.ColMP_QtyMax.Name = "ColMP_QtyMax"
        Me.ColMP_QtyMax.Visible = False
        Me.ColMP_QtyMax.Width = 71
        '
        'ColMP_QtyMin
        '
        Me.ColMP_QtyMin.HeaderText = "Qty Min"
        Me.ColMP_QtyMin.Name = "ColMP_QtyMin"
        Me.ColMP_QtyMin.Visible = False
        Me.ColMP_QtyMin.Width = 68
        '
        'ColMP_QtyOpt
        '
        Me.ColMP_QtyOpt.HeaderText = "Qty Optimum"
        Me.ColMP_QtyOpt.Name = "ColMP_QtyOpt"
        Me.ColMP_QtyOpt.Visible = False
        Me.ColMP_QtyOpt.Width = 92
        '
        'ColMP_QtyOrder
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMP_QtyOrder.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColMP_QtyOrder.HeaderText = "Qty Order"
        Me.ColMP_QtyOrder.Name = "ColMP_QtyOrder"
        Me.ColMP_QtyOrder.Width = 77
        '
        'ColMP_OrderUnit
        '
        Me.ColMP_OrderUnit.HeaderText = "Unit"
        Me.ColMP_OrderUnit.Name = "ColMP_OrderUnit"
        Me.ColMP_OrderUnit.Visible = False
        Me.ColMP_OrderUnit.Width = 51
        '
        'ColMP_OrderPoint
        '
        Me.ColMP_OrderPoint.HeaderText = "Order Point Date"
        Me.ColMP_OrderPoint.Name = "ColMP_OrderPoint"
        Me.ColMP_OrderPoint.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_OrderPoint.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColMP_OrderPoint.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColMP_OrderPoint.Width = 111
        '
        'ColMP_NextOrderPoint
        '
        Me.ColMP_NextOrderPoint.HeaderText = "Next Order Point Date"
        Me.ColMP_NextOrderPoint.Name = "ColMP_NextOrderPoint"
        Me.ColMP_NextOrderPoint.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_NextOrderPoint.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColMP_NextOrderPoint.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColMP_NextOrderPoint.Width = 136
        '
        'ColMP_ProcurementType
        '
        Me.ColMP_ProcurementType.HeaderText = "Procurement"
        Me.ColMP_ProcurementType.Items.AddRange(New Object() {"Produce", "Buy"})
        Me.ColMP_ProcurementType.Name = "ColMP_ProcurementType"
        Me.ColMP_ProcurementType.Visible = False
        Me.ColMP_ProcurementType.Width = 73
        '
        'ColMP_ProcurementStrategy
        '
        Me.ColMP_ProcurementStrategy.HeaderText = "Procurement Strategy"
        Me.ColMP_ProcurementStrategy.Items.AddRange(New Object() {"Build To Order", "Build To Stock"})
        Me.ColMP_ProcurementStrategy.Name = "ColMP_ProcurementStrategy"
        Me.ColMP_ProcurementStrategy.Width = 115
        '
        'ColMP_OrderStrategy
        '
        Me.ColMP_OrderStrategy.HeaderText = "Order Strategy"
        Me.ColMP_OrderStrategy.Items.AddRange(New Object() {"Order To Optimum Qty", "Order By Optimum Order Qty"})
        Me.ColMP_OrderStrategy.Name = "ColMP_OrderStrategy"
        Me.ColMP_OrderStrategy.Width = 81
        '
        'ColMP_FollowUp
        '
        Me.ColMP_FollowUp.HeaderText = "Follow Up"
        Me.ColMP_FollowUp.Name = "ColMP_FollowUp"
        Me.ColMP_FollowUp.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColMP_FollowUp.Width = 79
        '
        'ColMP_PreferedSupplier
        '
        Me.ColMP_PreferedSupplier.HeaderText = "Prefered Supplier"
        Me.ColMP_PreferedSupplier.Name = "ColMP_PreferedSupplier"
        Me.ColMP_PreferedSupplier.Width = 113
        '
        'ColMP_Currency
        '
        Me.ColMP_Currency.HeaderText = "Currency"
        Me.ColMP_Currency.Name = "ColMP_Currency"
        Me.ColMP_Currency.Width = 74
        '
        'ColMP_UnitPrice
        '
        Me.ColMP_UnitPrice.HeaderText = "Unit Price"
        Me.ColMP_UnitPrice.Name = "ColMP_UnitPrice"
        Me.ColMP_UnitPrice.Width = 78
        '
        'ColMP_Disc
        '
        Me.ColMP_Disc.HeaderText = "Discount"
        Me.ColMP_Disc.Name = "ColMP_Disc"
        Me.ColMP_Disc.Width = 74
        '
        'ColMP_NettPrice
        '
        Me.ColMP_NettPrice.HeaderText = "Nett Price"
        Me.ColMP_NettPrice.Name = "ColMP_NettPrice"
        Me.ColMP_NettPrice.Width = 79
        '
        'ColMP_SubTotal
        '
        Me.ColMP_SubTotal.HeaderText = "Sub Total"
        Me.ColMP_SubTotal.Name = "ColMP_SubTotal"
        Me.ColMP_SubTotal.Width = 78
        '
        'ColMP_PrimaryCurrency
        '
        Me.ColMP_PrimaryCurrency.HeaderText = "Primary Curr"
        Me.ColMP_PrimaryCurrency.Name = "ColMP_PrimaryCurrency"
        Me.ColMP_PrimaryCurrency.Visible = False
        Me.ColMP_PrimaryCurrency.Width = 88
        '
        'ColMP_PrimarySupplier
        '
        Me.ColMP_PrimarySupplier.HeaderText = "Primary Supplier"
        Me.ColMP_PrimarySupplier.Name = "ColMP_PrimarySupplier"
        Me.ColMP_PrimarySupplier.Visible = False
        Me.ColMP_PrimarySupplier.Width = 107
        '
        'ColMP_PrimaryCodeNo
        '
        Me.ColMP_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColMP_PrimaryCodeNo.Name = "ColMP_PrimaryCodeNo"
        Me.ColMP_PrimaryCodeNo.Visible = False
        Me.ColMP_PrimaryCodeNo.Width = 111
        '
        'ColMP_DefPrecCurrency
        '
        Me.ColMP_DefPrecCurrency.HeaderText = "Def Prec Currency"
        Me.ColMP_DefPrecCurrency.Name = "ColMP_DefPrecCurrency"
        Me.ColMP_DefPrecCurrency.Visible = False
        Me.ColMP_DefPrecCurrency.Width = 119
        '
        'ColMP_DefPrecUnit
        '
        Me.ColMP_DefPrecUnit.HeaderText = "DefPrecProduct"
        Me.ColMP_DefPrecUnit.Name = "ColMP_DefPrecUnit"
        Me.ColMP_DefPrecUnit.Visible = False
        Me.ColMP_DefPrecUnit.Width = 108
        '
        'ColMP_PrimaryOrderUnit
        '
        Me.ColMP_PrimaryOrderUnit.HeaderText = "Primary Order Unit"
        Me.ColMP_PrimaryOrderUnit.Name = "ColMP_PrimaryOrderUnit"
        Me.ColMP_PrimaryOrderUnit.Visible = False
        Me.ColMP_PrimaryOrderUnit.Width = 117
        '
        'ColMP_DefPrecOrderUnit
        '
        Me.ColMP_DefPrecOrderUnit.HeaderText = "DefPrec Order Unit"
        Me.ColMP_DefPrecOrderUnit.Name = "ColMP_DefPrecOrderUnit"
        Me.ColMP_DefPrecOrderUnit.Visible = False
        Me.ColMP_DefPrecOrderUnit.Width = 122
        '
        'ColMP_OrderPointDelay
        '
        Me.ColMP_OrderPointDelay.HeaderText = "Order Point Detail"
        Me.ColMP_OrderPointDelay.Name = "ColMP_OrderPointDelay"
        Me.ColMP_OrderPointDelay.Visible = False
        Me.ColMP_OrderPointDelay.Width = 115
        '
        'ColMP_NextOrderPointDelay
        '
        Me.ColMP_NextOrderPointDelay.HeaderText = "Next Order Point Delay"
        Me.ColMP_NextOrderPointDelay.Name = "ColMP_NextOrderPointDelay"
        Me.ColMP_NextOrderPointDelay.Visible = False
        Me.ColMP_NextOrderPointDelay.Width = 140
        '
        'ColMP_LeadTimeDays
        '
        Me.ColMP_LeadTimeDays.HeaderText = "Lead Time (days)"
        Me.ColMP_LeadTimeDays.Name = "ColMP_LeadTimeDays"
        Me.ColMP_LeadTimeDays.Visible = False
        Me.ColMP_LeadTimeDays.Width = 113
        '
        'ColMP_QtyOrderMulti
        '
        Me.ColMP_QtyOrderMulti.HeaderText = "QtyOrderMulti"
        Me.ColMP_QtyOrderMulti.Name = "ColMP_QtyOrderMulti"
        Me.ColMP_QtyOrderMulti.Visible = False
        Me.ColMP_QtyOrderMulti.Width = 96
        '
        'ColMP_PrimaryUnitMulti
        '
        Me.ColMP_PrimaryUnitMulti.HeaderText = "Primary Unit Multi"
        Me.ColMP_PrimaryUnitMulti.Name = "ColMP_PrimaryUnitMulti"
        Me.ColMP_PrimaryUnitMulti.Visible = False
        Me.ColMP_PrimaryUnitMulti.Width = 113
        '
        'ColMP_DefPrecUnitMulti
        '
        Me.ColMP_DefPrecUnitMulti.HeaderText = "Def Prec Unit Multi"
        Me.ColMP_DefPrecUnitMulti.Name = "ColMP_DefPrecUnitMulti"
        Me.ColMP_DefPrecUnitMulti.Visible = False
        Me.ColMP_DefPrecUnitMulti.Width = 121
        '
        'ColMP_RatioMulti
        '
        Me.ColMP_RatioMulti.HeaderText = "Ratio Multi"
        Me.ColMP_RatioMulti.Name = "ColMP_RatioMulti"
        Me.ColMP_RatioMulti.Visible = False
        Me.ColMP_RatioMulti.Width = 82
        '
        'ColMP_RatioDefault
        '
        Me.ColMP_RatioDefault.HeaderText = "Ratio Default"
        Me.ColMP_RatioDefault.Name = "ColMP_RatioDefault"
        Me.ColMP_RatioDefault.Visible = False
        Me.ColMP_RatioDefault.Width = 94
        '
        'ColMP_DiscAmount
        '
        Me.ColMP_DiscAmount.HeaderText = "Disc Amount"
        Me.ColMP_DiscAmount.Name = "ColMP_DiscAmount"
        Me.ColMP_DiscAmount.Visible = False
        Me.ColMP_DiscAmount.Width = 92
        '
        'ColMP_DiscFromPerct
        '
        Me.ColMP_DiscFromPerct.HeaderText = "DiscFromPerct"
        Me.ColMP_DiscFromPerct.Name = "ColMP_DiscFromPerct"
        Me.ColMP_DiscFromPerct.Visible = False
        Me.ColMP_DiscFromPerct.Width = 101
        '
        'ColMP_DetailFollowUp
        '
        Me.ColMP_DetailFollowUp.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColMP_DetailFollowUp.HeaderText = "Detail FollowUp"
        Me.ColMP_DetailFollowUp.Name = "ColMP_DetailFollowUp"
        Me.ColMP_DetailFollowUp.ReadOnly = True
        Me.ColMP_DetailFollowUp.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColMP_DetailFollowUp.Width = 87
        '
        'BackgroundWorker1
        '
        '
        'FrmTrans_MaterialPlanningDashboard_Authorize
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(934, 511)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.MinimumSize = New System.Drawing.Size(950, 550)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_Authorize"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Material Planning Authorization"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.dgvAuth, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.dgvCancel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvAll, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboMP.ResumeLayout(False)
        Me.gboMP.PerformLayout()
        Me.SplitContainer5.Panel1.ResumeLayout(False)
        Me.SplitContainer5.Panel1.PerformLayout()
        Me.SplitContainer5.Panel2.ResumeLayout(False)
        Me.SplitContainer5.Panel2.PerformLayout()
        Me.SplitContainer5.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        Me.SplitContainer4.ResumeLayout(False)
        CType(Me.dgvMP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStripLabel2 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents tslMP_CreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslMP_CreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslMP_ModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslMP_ModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel3 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel4 As System.Windows.Forms.LinkLabel
    Friend WithEvents dgvAuth As System.Windows.Forms.DataGridView
    Friend WithEvents dgvCancel As System.Windows.Forms.DataGridView
    Friend WithEvents dgvAll As System.Windows.Forms.DataGridView
    Friend WithEvents cboFilterStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFilterDocNo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFilterTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpFilterFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkFilterDate As System.Windows.Forms.CheckBox
    Friend WithEvents btnFilterFindMP As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents txtFilterPrimaryMP As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents col4No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col4UserCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col4UserName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col4Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col1No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col1DocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col1Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col1Auth As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents col1Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col2No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col2DocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col2Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col2CancelAuth As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents col2Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col3No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col3DocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col3Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col3Status As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col3Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents gboMP As System.Windows.Forms.GroupBox
    Friend WithEvents dtpMP_Filter As System.Windows.Forms.DateTimePicker
    Friend WithEvents SplitContainer5 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtMP_DaysPerMonth As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents dtpMP As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpMP_RefDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtMP_Primary As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_DocNo As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents chkClose As System.Windows.Forms.CheckBox
    Friend WithEvents TxtMP_Remarks As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents chkMP_RequiredOnly As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyRMReq As System.Windows.Forms.TextBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyPrOB As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_QtySafetyStock As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtySOB As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_QtyOptimumStock As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyPendingDO As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyFree As System.Windows.Forms.TextBox
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyMOB As System.Windows.Forms.TextBox
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyRequired As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_QtyOnHand As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyPOB As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyGIT As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_AvgPerMonth As System.Windows.Forms.TextBox
    Friend WithEvents cboMPOrderStrategy As System.Windows.Forms.ComboBox
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtMP_DiscAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents LlbMP_UnitPrice As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtMP_CurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_UnitPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_PreferedSupplier As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_Discount As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_NettPrice As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_PrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_PrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_DiscFromPerct As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_PreferedSupplierDesc As System.Windows.Forms.TextBox
    Friend WithEvents llbMP_Ratio As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtMP_QtyRequired2 As System.Windows.Forms.TextBox
    Friend WithEvents LblMP_Unit1 As System.Windows.Forms.Label
    Friend WithEvents LblRatio As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_SubTotal As System.Windows.Forms.TextBox
    Friend WithEvents cboMP_OrderUnit As System.Windows.Forms.ComboBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyOrder As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_PrimaryOrderUnit As System.Windows.Forms.TextBox
    Friend WithEvents dgvMP As System.Windows.Forms.DataGridView
    Friend WithEvents ColMP_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOnHand As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyPOB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyGIT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyInquiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyPendingDO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtySOB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyPrOB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtySafetyStock As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyLeadTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyRequired As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColMP_QtyOrderMaximum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOrderMinimum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOrderOptimum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyMax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyMin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOpt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOrder As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_OrderUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_OrderPoint As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColMP_NextOrderPoint As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColMP_ProcurementType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColMP_ProcurementStrategy As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColMP_OrderStrategy As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColMP_FollowUp As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColMP_PreferedSupplier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_Currency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_UnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_Disc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_NettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_SubTotal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_PrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_PrimarySupplier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DefPrecCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DefPrecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_PrimaryOrderUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DefPrecOrderUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_OrderPointDelay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_NextOrderPointDelay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_LeadTimeDays As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOrderMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_PrimaryUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DefPrecUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_RatioMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_RatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DiscFromPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DetailFollowUp As System.Windows.Forms.DataGridViewLinkColumn
End Class
