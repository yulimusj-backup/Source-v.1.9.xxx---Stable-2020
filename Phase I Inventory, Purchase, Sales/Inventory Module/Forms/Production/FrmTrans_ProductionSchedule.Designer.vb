<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule
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
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.CboAturWaktu = New System.Windows.Forms.ComboBox
        Me.btnRefresh = New System.Windows.Forms.Button
        Me.ColTreeNode_DocumentNo = New AdvancedDataGridView.TreeGridColumn
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.ColTGV_DocNo = New AdvancedDataGridView.TreeGridColumn
        Me.ColTGV_Late = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColTGV_IncompleteRM = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColTGV_PrintPickList = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColTGV_Details = New System.Windows.Forms.DataGridViewButtonColumn
        Me.cboGroupBy = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.dtpFirmPlanned = New System.Windows.Forms.DateTimePicker
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Label7 = New System.Windows.Forms.Label
        Me.dtpToday = New System.Windows.Forms.DateTimePicker
        Me.chkShowProgress = New System.Windows.Forms.CheckBox
        Me.chkLate = New System.Windows.Forms.CheckBox
        Me.btnAddDownTime = New System.Windows.Forms.Button
        Me.chkIncompleteRM = New System.Windows.Forms.CheckBox
        Me.btnPrintPickList = New System.Windows.Forms.Button
        Me.chkShowPickList = New System.Windows.Forms.CheckBox
        Me.btnAllocateWorkshift = New System.Windows.Forms.Button
        Me.cboTimeframe = New System.Windows.Forms.ComboBox
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.btnPrintSchedule = New System.Windows.Forms.Button
        Me.btnException = New System.Windows.Forms.Button
        Me.btnAddDelay = New System.Windows.Forms.Button
        Me.btnAddSetupTime = New System.Windows.Forms.Button
        Me.Button10 = New System.Windows.Forms.Button
        Me.llbSetting = New System.Windows.Forms.LinkLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage_SchGanttChart = New System.Windows.Forms.TabPage
        Me.TabPage_SchTable = New System.Windows.Forms.TabPage
        Me.dgvSch = New System.Windows.Forms.DataGridView
        Me.ColSch_JONumber1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_WorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_StartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_StartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_EndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_EndTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_Type = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_WorkOrderNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_MONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_JONumber2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_Process = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_Progress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_PrimaryWorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_PrimaryWorkShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_PrimaryMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_a = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_r = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_g = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_b = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.llbExpandAll = New System.Windows.Forms.LinkLabel
        Me.llbCollapsAll = New System.Windows.Forms.LinkLabel
        Me.tgvSch = New AdvancedDataGridView.TreeGridView
        Me.ColLeft_Tree = New AdvancedDataGridView.TreeGridColumn
        Me.ColLeft_ProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_WONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_WorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_PrimaryLine = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_LineCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_ProductionRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_QtyProduce = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_TotalWH = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_DayOfWeekStart = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_StartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_StartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_DayOfWeekFinish = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_FinishDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_FinishTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_Progress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_PrimaryWorkUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_PrimaryWorkshift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_PrimaryMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_a = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_r = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_g = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_b = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_NodeLevel = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_NodeLv0Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_NodeLv1Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_PrimaryLocationInput = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_LocationInput = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_PrimaryLocationOutput = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_LocationOutput = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_ActualGood = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_ActualReject = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeft_Actualproduce = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.cms1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tms_ChangeColor = New System.Windows.Forms.ToolStripMenuItem
        Me.CreateLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.EditLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DeleteLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ViewLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.CloseJOToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SplitMergeScheduleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DeleteJobOrderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.FindScheduleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.OptimizeScheduleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ReorderScheduleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ShowAllMOScheduleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.llbColumnSettingSch = New System.Windows.Forms.LinkLabel
        Me.dgvBar = New System.Windows.Forms.DataGridView
        Me.dgvHeader0 = New System.Windows.Forms.DataGridView
        Me.dgvHeader1 = New System.Windows.Forms.DataGridView
        Me.Button8 = New System.Windows.Forms.Button
        Me.Button9 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.BtnClearGanttchart = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.BtnLoad = New System.Windows.Forms.Button
        Me.BtnRemove = New System.Windows.Forms.Button
        Me.BtnAdd = New System.Windows.Forms.Button
        Me.BtnSave = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnClear = New System.Windows.Forms.Button
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.Dgv = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTaskName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFromDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProgress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeBar = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_a = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_r = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_g = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_b = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVisible = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BtnSet2 = New System.Windows.Forms.Button
        Me.TxtTaks = New System.Windows.Forms.TextBox
        Me.CboTaskName = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.DtpFromDate = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.DtpToDate = New System.Windows.Forms.DateTimePicker
        Me.txtIdentifier = New System.Windows.Forms.TextBox
        Me.HScrollBar1 = New System.Windows.Forms.HScrollBar
        Me.dtpBarFrom = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.btnAddOT = New System.Windows.Forms.Button
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.btnAddSchedule = New System.Windows.Forms.Button
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.btnWorkshift = New System.Windows.Forms.Button
        Me.BackgroundWorker5 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker6 = New System.ComponentModel.BackgroundWorker
        Me.bgwFullReload = New System.ComponentModel.BackgroundWorker
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.LblProgress = New System.Windows.Forms.Label
        Me.prgBar1 = New System.Windows.Forms.ProgressBar
        Me.bgwFormLoad = New System.ComponentModel.BackgroundWorker
        Me.btnFindMO = New System.Windows.Forms.Button
        Me.btnAddRemoveHoliday = New System.Windows.Forms.Button
        Me.StatusStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage_SchTable.SuspendLayout()
        CType(Me.dgvSch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.tgvSch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cms1.SuspendLayout()
        CType(Me.dgvBar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvHeader0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvHeader1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.Dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'CboAturWaktu
        '
        Me.CboAturWaktu.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAturWaktu.FormattingEnabled = True
        Me.CboAturWaktu.Location = New System.Drawing.Point(305, 70)
        Me.CboAturWaktu.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CboAturWaktu.Name = "CboAturWaktu"
        Me.CboAturWaktu.Size = New System.Drawing.Size(197, 24)
        Me.CboAturWaktu.TabIndex = 70
        Me.CboAturWaktu.Visible = False
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(845, 546)
        Me.btnRefresh.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(131, 53)
        Me.btnRefresh.TabIndex = 61
        Me.btnRefresh.TabStop = False
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        Me.btnRefresh.Visible = False
        '
        'ColTreeNode_DocumentNo
        '
        Me.ColTreeNode_DocumentNo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ColTreeNode_DocumentNo.DefaultNodeImage = Nothing
        Me.ColTreeNode_DocumentNo.HeaderText = "Document No"
        Me.ColTreeNode_DocumentNo.Name = "ColTreeNode_DocumentNo"
        Me.ColTreeNode_DocumentNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColTreeNode_DocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 14)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(120, 17)
        Me.Label5.TabIndex = 60
        Me.Label5.Text = "Schedule Horizon"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(143, 11)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(184, 22)
        Me.dtpFrom.TabIndex = 59
        '
        'ColTGV_DocNo
        '
        Me.ColTGV_DocNo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColTGV_DocNo.DefaultNodeImage = Nothing
        Me.ColTGV_DocNo.HeaderText = "Document No"
        Me.ColTGV_DocNo.Name = "ColTGV_DocNo"
        Me.ColTGV_DocNo.ReadOnly = True
        Me.ColTGV_DocNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColTGV_DocNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColTGV_Late
        '
        Me.ColTGV_Late.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColTGV_Late.HeaderText = "Late"
        Me.ColTGV_Late.Name = "ColTGV_Late"
        Me.ColTGV_Late.ReadOnly = True
        Me.ColTGV_Late.Visible = False
        '
        'ColTGV_IncompleteRM
        '
        Me.ColTGV_IncompleteRM.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColTGV_IncompleteRM.HeaderText = "Incomplete Raw Material"
        Me.ColTGV_IncompleteRM.Name = "ColTGV_IncompleteRM"
        Me.ColTGV_IncompleteRM.ReadOnly = True
        Me.ColTGV_IncompleteRM.Visible = False
        '
        'ColTGV_PrintPickList
        '
        Me.ColTGV_PrintPickList.HeaderText = "Print Pick List"
        Me.ColTGV_PrintPickList.Name = "ColTGV_PrintPickList"
        Me.ColTGV_PrintPickList.ReadOnly = True
        Me.ColTGV_PrintPickList.Visible = False
        '
        'ColTGV_Details
        '
        Me.ColTGV_Details.HeaderText = "Details"
        Me.ColTGV_Details.Name = "ColTGV_Details"
        Me.ColTGV_Details.ReadOnly = True
        Me.ColTGV_Details.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'cboGroupBy
        '
        Me.cboGroupBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGroupBy.FormattingEnabled = True
        Me.cboGroupBy.Items.AddRange(New Object() {"Group By Job Order", "Group By Work Unit"})
        Me.cboGroupBy.Location = New System.Drawing.Point(339, 41)
        Me.cboGroupBy.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cboGroupBy.Name = "cboGroupBy"
        Me.cboGroupBy.Size = New System.Drawing.Size(213, 24)
        Me.cboGroupBy.TabIndex = 78
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(15, 46)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(91, 17)
        Me.Label6.TabIndex = 80
        Me.Label6.Text = "Firm Planned"
        '
        'dtpFirmPlanned
        '
        Me.dtpFirmPlanned.CustomFormat = "dd MMM yyyy"
        Me.dtpFirmPlanned.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFirmPlanned.Location = New System.Drawing.Point(143, 41)
        Me.dtpFirmPlanned.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpFirmPlanned.Name = "dtpFirmPlanned"
        Me.dtpFirmPlanned.Size = New System.Drawing.Size(184, 22)
        Me.dtpFirmPlanned.TabIndex = 79
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(15, 46)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(35, 17)
        Me.Label7.TabIndex = 86
        Me.Label7.Text = "Now"
        Me.Label7.Visible = False
        '
        'dtpToday
        '
        Me.dtpToday.CustomFormat = "dd MMM yyyy HH.mm"
        Me.dtpToday.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToday.Location = New System.Drawing.Point(143, 41)
        Me.dtpToday.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpToday.Name = "dtpToday"
        Me.dtpToday.Size = New System.Drawing.Size(184, 22)
        Me.dtpToday.TabIndex = 85
        Me.dtpToday.Visible = False
        '
        'chkShowProgress
        '
        Me.chkShowProgress.AutoSize = True
        Me.chkShowProgress.Checked = True
        Me.chkShowProgress.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowProgress.Location = New System.Drawing.Point(561, 43)
        Me.chkShowProgress.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkShowProgress.Name = "chkShowProgress"
        Me.chkShowProgress.Size = New System.Drawing.Size(125, 21)
        Me.chkShowProgress.TabIndex = 87
        Me.chkShowProgress.TabStop = False
        Me.chkShowProgress.Text = "Show Progress"
        Me.chkShowProgress.UseVisualStyleBackColor = True
        '
        'chkLate
        '
        Me.chkLate.AutoSize = True
        Me.chkLate.Location = New System.Drawing.Point(708, 36)
        Me.chkLate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkLate.Name = "chkLate"
        Me.chkLate.Size = New System.Drawing.Size(181, 21)
        Me.chkLate.TabIndex = 88
        Me.chkLate.TabStop = False
        Me.chkLate.Text = "Show Late Work Orders"
        Me.chkLate.UseVisualStyleBackColor = True
        Me.chkLate.Visible = False
        '
        'btnAddDownTime
        '
        Me.btnAddDownTime.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAddDownTime.BackColor = System.Drawing.Color.Transparent
        Me.btnAddDownTime.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnAddDownTime.ForeColor = System.Drawing.Color.Black
        Me.btnAddDownTime.Location = New System.Drawing.Point(1232, 538)
        Me.btnAddDownTime.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnAddDownTime.Name = "btnAddDownTime"
        Me.btnAddDownTime.Size = New System.Drawing.Size(131, 53)
        Me.btnAddDownTime.TabIndex = 89
        Me.btnAddDownTime.TabStop = False
        Me.btnAddDownTime.Text = "Sch. Down time (pending)"
        Me.btnAddDownTime.UseVisualStyleBackColor = True
        Me.btnAddDownTime.Visible = False
        '
        'chkIncompleteRM
        '
        Me.chkIncompleteRM.AutoSize = True
        Me.chkIncompleteRM.Location = New System.Drawing.Point(305, 74)
        Me.chkIncompleteRM.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkIncompleteRM.Name = "chkIncompleteRM"
        Me.chkIncompleteRM.Size = New System.Drawing.Size(396, 21)
        Me.chkIncompleteRM.TabIndex = 90
        Me.chkIncompleteRM.TabStop = False
        Me.chkIncompleteRM.Text = "Show Mark On Work Orders With Incomplete Raw Material"
        Me.chkIncompleteRM.UseVisualStyleBackColor = True
        Me.chkIncompleteRM.Visible = False
        '
        'btnPrintPickList
        '
        Me.btnPrintPickList.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPrintPickList.BackColor = System.Drawing.Color.Transparent
        Me.btnPrintPickList.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnPrintPickList.ForeColor = System.Drawing.Color.Black
        Me.btnPrintPickList.Location = New System.Drawing.Point(849, -1)
        Me.btnPrintPickList.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnPrintPickList.Name = "btnPrintPickList"
        Me.btnPrintPickList.Size = New System.Drawing.Size(131, 53)
        Me.btnPrintPickList.TabIndex = 91
        Me.btnPrintPickList.TabStop = False
        Me.btnPrintPickList.Text = "Print Pick List"
        Me.btnPrintPickList.UseVisualStyleBackColor = True
        Me.btnPrintPickList.Visible = False
        '
        'chkShowPickList
        '
        Me.chkShowPickList.AutoSize = True
        Me.chkShowPickList.Location = New System.Drawing.Point(333, 43)
        Me.chkShowPickList.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkShowPickList.Name = "chkShowPickList"
        Me.chkShowPickList.Size = New System.Drawing.Size(120, 21)
        Me.chkShowPickList.TabIndex = 92
        Me.chkShowPickList.TabStop = False
        Me.chkShowPickList.Text = "Show Pick List"
        Me.chkShowPickList.UseVisualStyleBackColor = True
        Me.chkShowPickList.Visible = False
        '
        'btnAllocateWorkshift
        '
        Me.btnAllocateWorkshift.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAllocateWorkshift.BackColor = System.Drawing.Color.Transparent
        Me.btnAllocateWorkshift.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnAllocateWorkshift.ForeColor = System.Drawing.Color.Black
        Me.btnAllocateWorkshift.Location = New System.Drawing.Point(428, 538)
        Me.btnAllocateWorkshift.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnAllocateWorkshift.Name = "btnAllocateWorkshift"
        Me.btnAllocateWorkshift.Size = New System.Drawing.Size(131, 53)
        Me.btnAllocateWorkshift.TabIndex = 94
        Me.btnAllocateWorkshift.TabStop = False
        Me.btnAllocateWorkshift.Text = "Allocate Workshift"
        Me.btnAllocateWorkshift.UseVisualStyleBackColor = True
        '
        'cboTimeframe
        '
        Me.cboTimeframe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTimeframe.FormattingEnabled = True
        Me.cboTimeframe.Items.AddRange(New Object() {"By Hour", "By Day", "By Weeks"})
        Me.cboTimeframe.Location = New System.Drawing.Point(305, 70)
        Me.cboTimeframe.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboTimeframe.Name = "cboTimeframe"
        Me.cboTimeframe.Size = New System.Drawing.Size(197, 24)
        Me.cboTimeframe.TabIndex = 95
        Me.cboTimeframe.Visible = False
        '
        'btnPrintSchedule
        '
        Me.btnPrintSchedule.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPrintSchedule.BackColor = System.Drawing.Color.Transparent
        Me.btnPrintSchedule.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnPrintSchedule.ForeColor = System.Drawing.Color.Black
        Me.btnPrintSchedule.Location = New System.Drawing.Point(567, 538)
        Me.btnPrintSchedule.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnPrintSchedule.Name = "btnPrintSchedule"
        Me.btnPrintSchedule.Size = New System.Drawing.Size(131, 53)
        Me.btnPrintSchedule.TabIndex = 96
        Me.btnPrintSchedule.TabStop = False
        Me.btnPrintSchedule.Text = "Print Schedule"
        Me.btnPrintSchedule.UseVisualStyleBackColor = True
        '
        'btnException
        '
        Me.btnException.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnException.BackColor = System.Drawing.Color.Transparent
        Me.btnException.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnException.ForeColor = System.Drawing.Color.Black
        Me.btnException.Location = New System.Drawing.Point(707, 538)
        Me.btnException.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnException.Name = "btnException"
        Me.btnException.Size = New System.Drawing.Size(131, 53)
        Me.btnException.TabIndex = 97
        Me.btnException.TabStop = False
        Me.btnException.Text = "Show Exception"
        Me.btnException.UseVisualStyleBackColor = True
        '
        'btnAddDelay
        '
        Me.btnAddDelay.Location = New System.Drawing.Point(560, 6)
        Me.btnAddDelay.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnAddDelay.Name = "btnAddDelay"
        Me.btnAddDelay.Size = New System.Drawing.Size(131, 32)
        Me.btnAddDelay.TabIndex = 98
        Me.btnAddDelay.TabStop = False
        Me.btnAddDelay.Text = "Add Delay"
        Me.btnAddDelay.UseVisualStyleBackColor = True
        Me.btnAddDelay.Visible = False
        '
        'btnAddSetupTime
        '
        Me.btnAddSetupTime.Location = New System.Drawing.Point(509, 68)
        Me.btnAddSetupTime.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnAddSetupTime.Name = "btnAddSetupTime"
        Me.btnAddSetupTime.Size = New System.Drawing.Size(131, 32)
        Me.btnAddSetupTime.TabIndex = 99
        Me.btnAddSetupTime.TabStop = False
        Me.btnAddSetupTime.Text = "Add Setup Time"
        Me.btnAddSetupTime.UseVisualStyleBackColor = True
        Me.btnAddSetupTime.Visible = False
        '
        'Button10
        '
        Me.Button10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button10.Location = New System.Drawing.Point(944, 538)
        Me.Button10.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(131, 53)
        Me.Button10.TabIndex = 100
        Me.Button10.TabStop = False
        Me.Button10.Text = "Load Balancing (pending)"
        Me.Button10.UseVisualStyleBackColor = True
        Me.Button10.Visible = False
        '
        'llbSetting
        '
        Me.llbSetting.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbSetting.AutoSize = True
        Me.llbSetting.Location = New System.Drawing.Point(9, 565)
        Me.llbSetting.Name = "llbSetting"
        Me.llbSetting.Size = New System.Drawing.Size(52, 17)
        Me.llbSetting.TabIndex = 101
        Me.llbSetting.TabStop = True
        Me.llbSetting.Text = "Setting"
        Me.llbSetting.Visible = False
        Me.llbSetting.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 614)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 13, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1035, 25)
        Me.StatusStrip1.TabIndex = 102
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(154, 20)
        Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage_SchGanttChart)
        Me.TabControl1.Controls.Add(Me.TabPage_SchTable)
        Me.TabControl1.Location = New System.Drawing.Point(1003, 16)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(904, 462)
        Me.TabControl1.TabIndex = 103
        Me.TabControl1.TabStop = False
        Me.TabControl1.Visible = False
        '
        'TabPage_SchGanttChart
        '
        Me.TabPage_SchGanttChart.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_SchGanttChart.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage_SchGanttChart.Name = "TabPage_SchGanttChart"
        Me.TabPage_SchGanttChart.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage_SchGanttChart.Size = New System.Drawing.Size(896, 433)
        Me.TabPage_SchGanttChart.TabIndex = 0
        Me.TabPage_SchGanttChart.Text = "Gantt Chart"
        Me.TabPage_SchGanttChart.UseVisualStyleBackColor = True
        '
        'TabPage_SchTable
        '
        Me.TabPage_SchTable.Controls.Add(Me.dgvSch)
        Me.TabPage_SchTable.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_SchTable.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage_SchTable.Name = "TabPage_SchTable"
        Me.TabPage_SchTable.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage_SchTable.Size = New System.Drawing.Size(896, 433)
        Me.TabPage_SchTable.TabIndex = 1
        Me.TabPage_SchTable.Text = "Schedule Table"
        Me.TabPage_SchTable.UseVisualStyleBackColor = True
        '
        'dgvSch
        '
        Me.dgvSch.AllowUserToAddRows = False
        Me.dgvSch.AllowUserToDeleteRows = False
        Me.dgvSch.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSch.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvSch.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSch_JONumber1, Me.ColSch_WorkUnit, Me.ColSch_StartDate, Me.ColSch_StartTime, Me.ColSch_EndDate, Me.ColSch_EndTime, Me.ColSch_Type, Me.ColSch_WorkOrderNo, Me.ColSch_MONumber, Me.ColSch_JONumber2, Me.ColSch_Process, Me.ColSch_Progress, Me.ColSch_PrimaryWorkUnit, Me.ColSch_PrimaryProcess, Me.ColSch_PrimaryWorkShift, Me.ColSch_PrimaryMO, Me.ColSch_IdentifierJO, Me.ColSch_IdentifierWO, Me.ColSch_a, Me.ColSch_r, Me.ColSch_g, Me.ColSch_b})
        Me.dgvSch.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSch.Location = New System.Drawing.Point(3, 2)
        Me.dgvSch.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvSch.Name = "dgvSch"
        Me.dgvSch.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvSch.RowTemplate.Height = 24
        Me.dgvSch.Size = New System.Drawing.Size(890, 429)
        Me.dgvSch.TabIndex = 0
        '
        'ColSch_JONumber1
        '
        Me.ColSch_JONumber1.HeaderText = "JO No"
        Me.ColSch_JONumber1.Name = "ColSch_JONumber1"
        Me.ColSch_JONumber1.Width = 73
        '
        'ColSch_WorkUnit
        '
        Me.ColSch_WorkUnit.HeaderText = "Work Unit"
        Me.ColSch_WorkUnit.Name = "ColSch_WorkUnit"
        Me.ColSch_WorkUnit.Width = 95
        '
        'ColSch_StartDate
        '
        Me.ColSch_StartDate.HeaderText = "Start"
        Me.ColSch_StartDate.Name = "ColSch_StartDate"
        Me.ColSch_StartDate.Width = 63
        '
        'ColSch_StartTime
        '
        Me.ColSch_StartTime.HeaderText = ""
        Me.ColSch_StartTime.Name = "ColSch_StartTime"
        Me.ColSch_StartTime.Width = 19
        '
        'ColSch_EndDate
        '
        Me.ColSch_EndDate.HeaderText = "End"
        Me.ColSch_EndDate.Name = "ColSch_EndDate"
        Me.ColSch_EndDate.Width = 58
        '
        'ColSch_EndTime
        '
        Me.ColSch_EndTime.HeaderText = ""
        Me.ColSch_EndTime.Name = "ColSch_EndTime"
        Me.ColSch_EndTime.Width = 19
        '
        'ColSch_Type
        '
        Me.ColSch_Type.HeaderText = "Type"
        Me.ColSch_Type.Name = "ColSch_Type"
        Me.ColSch_Type.Width = 65
        '
        'ColSch_WorkOrderNo
        '
        Me.ColSch_WorkOrderNo.HeaderText = "WO No"
        Me.ColSch_WorkOrderNo.Name = "ColSch_WorkOrderNo"
        Me.ColSch_WorkOrderNo.Width = 79
        '
        'ColSch_MONumber
        '
        Me.ColSch_MONumber.HeaderText = "MO No"
        Me.ColSch_MONumber.Name = "ColSch_MONumber"
        Me.ColSch_MONumber.Width = 77
        '
        'ColSch_JONumber2
        '
        Me.ColSch_JONumber2.HeaderText = "JO No"
        Me.ColSch_JONumber2.Name = "ColSch_JONumber2"
        Me.ColSch_JONumber2.Width = 73
        '
        'ColSch_Process
        '
        Me.ColSch_Process.HeaderText = "Process"
        Me.ColSch_Process.Name = "ColSch_Process"
        Me.ColSch_Process.Width = 84
        '
        'ColSch_Progress
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSch_Progress.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColSch_Progress.HeaderText = "Progress"
        Me.ColSch_Progress.Name = "ColSch_Progress"
        Me.ColSch_Progress.Width = 90
        '
        'ColSch_PrimaryWorkUnit
        '
        Me.ColSch_PrimaryWorkUnit.HeaderText = "Primary Work Unit"
        Me.ColSch_PrimaryWorkUnit.Name = "ColSch_PrimaryWorkUnit"
        Me.ColSch_PrimaryWorkUnit.Visible = False
        Me.ColSch_PrimaryWorkUnit.Width = 147
        '
        'ColSch_PrimaryProcess
        '
        Me.ColSch_PrimaryProcess.HeaderText = "Primary Process"
        Me.ColSch_PrimaryProcess.Name = "ColSch_PrimaryProcess"
        Me.ColSch_PrimaryProcess.Visible = False
        Me.ColSch_PrimaryProcess.Width = 136
        '
        'ColSch_PrimaryWorkShift
        '
        Me.ColSch_PrimaryWorkShift.HeaderText = "Primary Workshift"
        Me.ColSch_PrimaryWorkShift.Name = "ColSch_PrimaryWorkShift"
        Me.ColSch_PrimaryWorkShift.Visible = False
        Me.ColSch_PrimaryWorkShift.Width = 144
        '
        'ColSch_PrimaryMO
        '
        Me.ColSch_PrimaryMO.HeaderText = "Primary MO"
        Me.ColSch_PrimaryMO.Name = "ColSch_PrimaryMO"
        Me.ColSch_PrimaryMO.Visible = False
        Me.ColSch_PrimaryMO.Width = 107
        '
        'ColSch_IdentifierJO
        '
        Me.ColSch_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColSch_IdentifierJO.Name = "ColSch_IdentifierJO"
        Me.ColSch_IdentifierJO.Visible = False
        Me.ColSch_IdentifierJO.Width = 109
        '
        'ColSch_IdentifierWO
        '
        Me.ColSch_IdentifierWO.HeaderText = "Identifier WO"
        Me.ColSch_IdentifierWO.Name = "ColSch_IdentifierWO"
        Me.ColSch_IdentifierWO.Visible = False
        Me.ColSch_IdentifierWO.Width = 115
        '
        'ColSch_a
        '
        Me.ColSch_a.HeaderText = "a"
        Me.ColSch_a.Name = "ColSch_a"
        Me.ColSch_a.Visible = False
        Me.ColSch_a.Width = 41
        '
        'ColSch_r
        '
        Me.ColSch_r.HeaderText = "r"
        Me.ColSch_r.Name = "ColSch_r"
        Me.ColSch_r.Visible = False
        Me.ColSch_r.Width = 38
        '
        'ColSch_g
        '
        Me.ColSch_g.HeaderText = "g"
        Me.ColSch_g.Name = "ColSch_g"
        Me.ColSch_g.Visible = False
        Me.ColSch_g.Width = 41
        '
        'ColSch_b
        '
        Me.ColSch_b.HeaderText = "b"
        Me.ColSch_b.Name = "ColSch_b"
        Me.ColSch_b.Visible = False
        Me.ColSch_b.Width = 41
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(13, 69)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.llbExpandAll)
        Me.SplitContainer1.Panel1.Controls.Add(Me.llbCollapsAll)
        Me.SplitContainer1.Panel1.Controls.Add(Me.tgvSch)
        Me.SplitContainer1.Panel1.Controls.Add(Me.llbColumnSettingSch)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvBar)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvHeader0)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvHeader1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.HScrollBar1)
        Me.SplitContainer1.Size = New System.Drawing.Size(1011, 463)
        Me.SplitContainer1.SplitterDistance = 264
        Me.SplitContainer1.TabIndex = 74
        '
        'llbExpandAll
        '
        Me.llbExpandAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbExpandAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbExpandAll.AutoSize = True
        Me.llbExpandAll.Location = New System.Drawing.Point(108, 10)
        Me.llbExpandAll.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbExpandAll.Name = "llbExpandAll"
        Me.llbExpandAll.Size = New System.Drawing.Size(74, 17)
        Me.llbExpandAll.TabIndex = 109
        Me.llbExpandAll.TabStop = True
        Me.llbExpandAll.Text = "Expand All"
        Me.llbExpandAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbCollapsAll
        '
        Me.llbCollapsAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbCollapsAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCollapsAll.AutoSize = True
        Me.llbCollapsAll.Location = New System.Drawing.Point(192, 10)
        Me.llbCollapsAll.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbCollapsAll.Name = "llbCollapsAll"
        Me.llbCollapsAll.Size = New System.Drawing.Size(73, 17)
        Me.llbCollapsAll.TabIndex = 110
        Me.llbCollapsAll.TabStop = True
        Me.llbCollapsAll.Text = "Collaps All"
        Me.llbCollapsAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'tgvSch
        '
        Me.tgvSch.AllowUserToAddRows = False
        Me.tgvSch.AllowUserToDeleteRows = False
        Me.tgvSch.AllowUserToResizeRows = False
        Me.tgvSch.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.tgvSch.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.tgvSch.ColumnHeadersHeight = 25
        Me.tgvSch.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.tgvSch.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColLeft_Tree, Me.ColLeft_ProcessCode, Me.ColLeft_WONo, Me.ColLeft_WorkUnit, Me.ColLeft_PrimaryLine, Me.ColLeft_LineCode, Me.ColLeft_CodeNo, Me.ColLeft_Description, Me.ColLeft_ProductionRate, Me.ColLeft_QtyProduce, Me.ColLeft_Unit, Me.ColLeft_TotalWH, Me.ColLeft_DayOfWeekStart, Me.ColLeft_StartDate, Me.ColLeft_StartTime, Me.ColLeft_DayOfWeekFinish, Me.ColLeft_FinishDate, Me.ColLeft_FinishTime, Me.ColLeft_Progress, Me.ColLeft_PrimaryWorkUnit, Me.ColLeft_PrimaryProcess, Me.ColLeft_PrimaryWorkshift, Me.ColLeft_PrimaryMO, Me.ColLeft_IdentifierJO, Me.ColLeft_IdentifierWO, Me.ColLeft_a, Me.ColLeft_r, Me.ColLeft_g, Me.ColLeft_b, Me.ColLeft_NodeLevel, Me.ColLeft_NodeLv0Identifier, Me.ColLeft_NodeLv1Identifier, Me.ColLeft_PrimaryLocationInput, Me.ColLeft_LocationInput, Me.ColLeft_PrimaryLocationOutput, Me.ColLeft_LocationOutput, Me.ColLeft_ActualGood, Me.ColLeft_ActualReject, Me.ColLeft_Actualproduce})
        Me.tgvSch.ContextMenuStrip = Me.cms1
        Me.tgvSch.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.tgvSch.ImageList = Nothing
        Me.tgvSch.Location = New System.Drawing.Point(0, 33)
        Me.tgvSch.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.tgvSch.Name = "tgvSch"
        Me.tgvSch.RowHeadersVisible = False
        Me.tgvSch.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.tgvSch.Size = New System.Drawing.Size(264, 430)
        Me.tgvSch.TabIndex = 9
        '
        'ColLeft_Tree
        '
        Me.ColLeft_Tree.DefaultNodeImage = Global.Inventory_Module.My.Resources.Resources.archive_32
        Me.ColLeft_Tree.Frozen = True
        Me.ColLeft_Tree.HeaderText = "Tree"
        Me.ColLeft_Tree.Name = "ColLeft_Tree"
        Me.ColLeft_Tree.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLeft_Tree.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_Tree.Width = 35
        '
        'ColLeft_ProcessCode
        '
        Me.ColLeft_ProcessCode.HeaderText = "Process"
        Me.ColLeft_ProcessCode.Name = "ColLeft_ProcessCode"
        Me.ColLeft_ProcessCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_ProcessCode.Width = 51
        '
        'ColLeft_WONo
        '
        Me.ColLeft_WONo.HeaderText = "WO No"
        Me.ColLeft_WONo.Name = "ColLeft_WONo"
        Me.ColLeft_WONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_WONo.Width = 49
        '
        'ColLeft_WorkUnit
        '
        Me.ColLeft_WorkUnit.HeaderText = "Work Unit"
        Me.ColLeft_WorkUnit.Name = "ColLeft_WorkUnit"
        Me.ColLeft_WorkUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_WorkUnit.Width = 61
        '
        'ColLeft_PrimaryLine
        '
        Me.ColLeft_PrimaryLine.HeaderText = "Primary Line"
        Me.ColLeft_PrimaryLine.Name = "ColLeft_PrimaryLine"
        Me.ColLeft_PrimaryLine.ReadOnly = True
        Me.ColLeft_PrimaryLine.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_PrimaryLine.Visible = False
        Me.ColLeft_PrimaryLine.Width = 93
        '
        'ColLeft_LineCode
        '
        Me.ColLeft_LineCode.HeaderText = "Line"
        Me.ColLeft_LineCode.Name = "ColLeft_LineCode"
        Me.ColLeft_LineCode.ReadOnly = True
        Me.ColLeft_LineCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_LineCode.Width = 33
        '
        'ColLeft_CodeNo
        '
        Me.ColLeft_CodeNo.HeaderText = "Code No"
        Me.ColLeft_CodeNo.Name = "ColLeft_CodeNo"
        Me.ColLeft_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_CodeNo.Width = 55
        '
        'ColLeft_Description
        '
        Me.ColLeft_Description.HeaderText = "Description"
        Me.ColLeft_Description.Name = "ColLeft_Description"
        Me.ColLeft_Description.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_Description.Width = 66
        '
        'ColLeft_ProductionRate
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColLeft_ProductionRate.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColLeft_ProductionRate.HeaderText = "Rate (Unit/Hour)"
        Me.ColLeft_ProductionRate.Name = "ColLeft_ProductionRate"
        Me.ColLeft_ProductionRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_ProductionRate.Width = 92
        '
        'ColLeft_QtyProduce
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColLeft_QtyProduce.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColLeft_QtyProduce.HeaderText = "Qty"
        Me.ColLeft_QtyProduce.Name = "ColLeft_QtyProduce"
        Me.ColLeft_QtyProduce.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_QtyProduce.Width = 29
        '
        'ColLeft_Unit
        '
        Me.ColLeft_Unit.HeaderText = "Unit"
        Me.ColLeft_Unit.Name = "ColLeft_Unit"
        Me.ColLeft_Unit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_Unit.Width = 32
        '
        'ColLeft_TotalWH
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColLeft_TotalWH.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColLeft_TotalWH.HeaderText = "W/H"
        Me.ColLeft_TotalWH.Name = "ColLeft_TotalWH"
        Me.ColLeft_TotalWH.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_TotalWH.Width = 37
        '
        'ColLeft_DayOfWeekStart
        '
        Me.ColLeft_DayOfWeekStart.HeaderText = "Day of Week Start"
        Me.ColLeft_DayOfWeekStart.Name = "ColLeft_DayOfWeekStart"
        Me.ColLeft_DayOfWeekStart.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_DayOfWeekStart.Width = 101
        '
        'ColLeft_StartDate
        '
        Me.ColLeft_StartDate.HeaderText = "Start"
        Me.ColLeft_StartDate.Name = "ColLeft_StartDate"
        Me.ColLeft_StartDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_StartDate.Width = 35
        '
        'ColLeft_StartTime
        '
        Me.ColLeft_StartTime.HeaderText = " "
        Me.ColLeft_StartTime.Name = "ColLeft_StartTime"
        Me.ColLeft_StartTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_StartTime.Width = 16
        '
        'ColLeft_DayOfWeekFinish
        '
        Me.ColLeft_DayOfWeekFinish.HeaderText = "Day of Week Finish"
        Me.ColLeft_DayOfWeekFinish.Name = "ColLeft_DayOfWeekFinish"
        Me.ColLeft_DayOfWeekFinish.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_DayOfWeekFinish.Width = 106
        '
        'ColLeft_FinishDate
        '
        Me.ColLeft_FinishDate.HeaderText = "Finish"
        Me.ColLeft_FinishDate.Name = "ColLeft_FinishDate"
        Me.ColLeft_FinishDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_FinishDate.Width = 40
        '
        'ColLeft_FinishTime
        '
        Me.ColLeft_FinishTime.HeaderText = " "
        Me.ColLeft_FinishTime.Name = "ColLeft_FinishTime"
        Me.ColLeft_FinishTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_FinishTime.Width = 16
        '
        'ColLeft_Progress
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColLeft_Progress.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColLeft_Progress.HeaderText = "Progress (%)"
        Me.ColLeft_Progress.Name = "ColLeft_Progress"
        Me.ColLeft_Progress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_Progress.Width = 71
        '
        'ColLeft_PrimaryWorkUnit
        '
        Me.ColLeft_PrimaryWorkUnit.HeaderText = "Primary Work Unit"
        Me.ColLeft_PrimaryWorkUnit.Name = "ColLeft_PrimaryWorkUnit"
        Me.ColLeft_PrimaryWorkUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_PrimaryWorkUnit.Visible = False
        Me.ColLeft_PrimaryWorkUnit.Width = 128
        '
        'ColLeft_PrimaryProcess
        '
        Me.ColLeft_PrimaryProcess.HeaderText = "Primary Process"
        Me.ColLeft_PrimaryProcess.Name = "ColLeft_PrimaryProcess"
        Me.ColLeft_PrimaryProcess.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_PrimaryProcess.Visible = False
        Me.ColLeft_PrimaryProcess.Width = 117
        '
        'ColLeft_PrimaryWorkshift
        '
        Me.ColLeft_PrimaryWorkshift.HeaderText = "Primary Workshift"
        Me.ColLeft_PrimaryWorkshift.Name = "ColLeft_PrimaryWorkshift"
        Me.ColLeft_PrimaryWorkshift.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_PrimaryWorkshift.Visible = False
        Me.ColLeft_PrimaryWorkshift.Width = 125
        '
        'ColLeft_PrimaryMO
        '
        Me.ColLeft_PrimaryMO.HeaderText = "Primary MO"
        Me.ColLeft_PrimaryMO.Name = "ColLeft_PrimaryMO"
        Me.ColLeft_PrimaryMO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_PrimaryMO.Visible = False
        Me.ColLeft_PrimaryMO.Width = 88
        '
        'ColLeft_IdentifierJO
        '
        Me.ColLeft_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColLeft_IdentifierJO.Name = "ColLeft_IdentifierJO"
        Me.ColLeft_IdentifierJO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_IdentifierJO.Visible = False
        Me.ColLeft_IdentifierJO.Width = 90
        '
        'ColLeft_IdentifierWO
        '
        Me.ColLeft_IdentifierWO.HeaderText = "Identifier WO"
        Me.ColLeft_IdentifierWO.Name = "ColLeft_IdentifierWO"
        Me.ColLeft_IdentifierWO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_IdentifierWO.Visible = False
        Me.ColLeft_IdentifierWO.Width = 96
        '
        'ColLeft_a
        '
        Me.ColLeft_a.HeaderText = "a"
        Me.ColLeft_a.Name = "ColLeft_a"
        Me.ColLeft_a.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_a.Visible = False
        Me.ColLeft_a.Width = 22
        '
        'ColLeft_r
        '
        Me.ColLeft_r.HeaderText = "r"
        Me.ColLeft_r.Name = "ColLeft_r"
        Me.ColLeft_r.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_r.Visible = False
        Me.ColLeft_r.Width = 19
        '
        'ColLeft_g
        '
        Me.ColLeft_g.HeaderText = "g"
        Me.ColLeft_g.Name = "ColLeft_g"
        Me.ColLeft_g.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_g.Visible = False
        Me.ColLeft_g.Width = 22
        '
        'ColLeft_b
        '
        Me.ColLeft_b.HeaderText = "b"
        Me.ColLeft_b.Name = "ColLeft_b"
        Me.ColLeft_b.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_b.Visible = False
        Me.ColLeft_b.Width = 22
        '
        'ColLeft_NodeLevel
        '
        Me.ColLeft_NodeLevel.HeaderText = "Node Level"
        Me.ColLeft_NodeLevel.Name = "ColLeft_NodeLevel"
        Me.ColLeft_NodeLevel.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_NodeLevel.Visible = False
        Me.ColLeft_NodeLevel.Width = 86
        '
        'ColLeft_NodeLv0Identifier
        '
        Me.ColLeft_NodeLv0Identifier.HeaderText = "Node Identifier Lv 0"
        Me.ColLeft_NodeLv0Identifier.Name = "ColLeft_NodeLv0Identifier"
        Me.ColLeft_NodeLv0Identifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_NodeLv0Identifier.Visible = False
        Me.ColLeft_NodeLv0Identifier.Width = 137
        '
        'ColLeft_NodeLv1Identifier
        '
        Me.ColLeft_NodeLv1Identifier.HeaderText = "Node Identifier Lv 1"
        Me.ColLeft_NodeLv1Identifier.Name = "ColLeft_NodeLv1Identifier"
        Me.ColLeft_NodeLv1Identifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_NodeLv1Identifier.Visible = False
        Me.ColLeft_NodeLv1Identifier.Width = 137
        '
        'ColLeft_PrimaryLocationInput
        '
        Me.ColLeft_PrimaryLocationInput.HeaderText = "Primary Input Location"
        Me.ColLeft_PrimaryLocationInput.Name = "ColLeft_PrimaryLocationInput"
        Me.ColLeft_PrimaryLocationInput.ReadOnly = True
        Me.ColLeft_PrimaryLocationInput.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_PrimaryLocationInput.Visible = False
        Me.ColLeft_PrimaryLocationInput.Width = 155
        '
        'ColLeft_LocationInput
        '
        Me.ColLeft_LocationInput.HeaderText = "Input Location"
        Me.ColLeft_LocationInput.Name = "ColLeft_LocationInput"
        Me.ColLeft_LocationInput.ReadOnly = True
        Me.ColLeft_LocationInput.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_LocationInput.Width = 81
        '
        'ColLeft_PrimaryLocationOutput
        '
        Me.ColLeft_PrimaryLocationOutput.HeaderText = "Primary Output Location"
        Me.ColLeft_PrimaryLocationOutput.Name = "ColLeft_PrimaryLocationOutput"
        Me.ColLeft_PrimaryLocationOutput.ReadOnly = True
        Me.ColLeft_PrimaryLocationOutput.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_PrimaryLocationOutput.Visible = False
        Me.ColLeft_PrimaryLocationOutput.Width = 167
        '
        'ColLeft_LocationOutput
        '
        Me.ColLeft_LocationOutput.HeaderText = "Output Location"
        Me.ColLeft_LocationOutput.Name = "ColLeft_LocationOutput"
        Me.ColLeft_LocationOutput.ReadOnly = True
        Me.ColLeft_LocationOutput.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_LocationOutput.Width = 89
        '
        'ColLeft_ActualGood
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle7.NullValue = "0.00"
        Me.ColLeft_ActualGood.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColLeft_ActualGood.HeaderText = "Actual Good"
        Me.ColLeft_ActualGood.Name = "ColLeft_ActualGood"
        Me.ColLeft_ActualGood.ReadOnly = True
        Me.ColLeft_ActualGood.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_ActualGood.Width = 72
        '
        'ColLeft_ActualReject
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle8.NullValue = "0.00"
        Me.ColLeft_ActualReject.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColLeft_ActualReject.HeaderText = "Actual Reject"
        Me.ColLeft_ActualReject.Name = "ColLeft_ActualReject"
        Me.ColLeft_ActualReject.ReadOnly = True
        Me.ColLeft_ActualReject.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_ActualReject.Width = 77
        '
        'ColLeft_Actualproduce
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle9.NullValue = "0.00"
        Me.ColLeft_Actualproduce.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColLeft_Actualproduce.HeaderText = "Actual Produce"
        Me.ColLeft_Actualproduce.Name = "ColLeft_Actualproduce"
        Me.ColLeft_Actualproduce.ReadOnly = True
        Me.ColLeft_Actualproduce.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLeft_Actualproduce.Width = 86
        '
        'cms1
        '
        Me.cms1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tms_ChangeColor, Me.CreateLogToolStripMenuItem, Me.EditLogToolStripMenuItem, Me.DeleteLogToolStripMenuItem, Me.ViewLogToolStripMenuItem, Me.ToolStripSeparator1, Me.CloseJOToolStripMenuItem, Me.SplitMergeScheduleToolStripMenuItem, Me.DeleteJobOrderToolStripMenuItem, Me.FindScheduleToolStripMenuItem, Me.OptimizeScheduleToolStripMenuItem, Me.ReorderScheduleToolStripMenuItem, Me.ShowAllMOScheduleToolStripMenuItem})
        Me.cms1.Name = "cms1"
        Me.cms1.Size = New System.Drawing.Size(362, 298)
        '
        'tms_ChangeColor
        '
        Me.tms_ChangeColor.Name = "tms_ChangeColor"
        Me.tms_ChangeColor.Size = New System.Drawing.Size(361, 24)
        Me.tms_ChangeColor.Text = "Change Color"
        '
        'CreateLogToolStripMenuItem
        '
        Me.CreateLogToolStripMenuItem.Name = "CreateLogToolStripMenuItem"
        Me.CreateLogToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.CreateLogToolStripMenuItem.Text = "Create Log"
        '
        'EditLogToolStripMenuItem
        '
        Me.EditLogToolStripMenuItem.Name = "EditLogToolStripMenuItem"
        Me.EditLogToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.EditLogToolStripMenuItem.Text = "Edit Log"
        '
        'DeleteLogToolStripMenuItem
        '
        Me.DeleteLogToolStripMenuItem.Name = "DeleteLogToolStripMenuItem"
        Me.DeleteLogToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.DeleteLogToolStripMenuItem.Text = "Delete Log"
        '
        'ViewLogToolStripMenuItem
        '
        Me.ViewLogToolStripMenuItem.Name = "ViewLogToolStripMenuItem"
        Me.ViewLogToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.ViewLogToolStripMenuItem.Text = "View Log"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(358, 6)
        '
        'CloseJOToolStripMenuItem
        '
        Me.CloseJOToolStripMenuItem.Name = "CloseJOToolStripMenuItem"
        Me.CloseJOToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.CloseJOToolStripMenuItem.Text = "Close JO"
        '
        'SplitMergeScheduleToolStripMenuItem
        '
        Me.SplitMergeScheduleToolStripMenuItem.Name = "SplitMergeScheduleToolStripMenuItem"
        Me.SplitMergeScheduleToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.SplitMergeScheduleToolStripMenuItem.Text = "Change Workunit, Split, or Merge Schedule"
        '
        'DeleteJobOrderToolStripMenuItem
        '
        Me.DeleteJobOrderToolStripMenuItem.Name = "DeleteJobOrderToolStripMenuItem"
        Me.DeleteJobOrderToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.DeleteJobOrderToolStripMenuItem.Text = "Delete Schedule"
        '
        'FindScheduleToolStripMenuItem
        '
        Me.FindScheduleToolStripMenuItem.Name = "FindScheduleToolStripMenuItem"
        Me.FindScheduleToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.FindScheduleToolStripMenuItem.Text = "Find Schedule"
        '
        'OptimizeScheduleToolStripMenuItem
        '
        Me.OptimizeScheduleToolStripMenuItem.Name = "OptimizeScheduleToolStripMenuItem"
        Me.OptimizeScheduleToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.OptimizeScheduleToolStripMenuItem.Text = "Optimize Schedule"
        '
        'ReorderScheduleToolStripMenuItem
        '
        Me.ReorderScheduleToolStripMenuItem.Name = "ReorderScheduleToolStripMenuItem"
        Me.ReorderScheduleToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.ReorderScheduleToolStripMenuItem.Text = "Reorder Schedule"
        '
        'ShowAllMOScheduleToolStripMenuItem
        '
        Me.ShowAllMOScheduleToolStripMenuItem.Name = "ShowAllMOScheduleToolStripMenuItem"
        Me.ShowAllMOScheduleToolStripMenuItem.Size = New System.Drawing.Size(361, 24)
        Me.ShowAllMOScheduleToolStripMenuItem.Text = "Show All MO Schedule"
        '
        'llbColumnSettingSch
        '
        Me.llbColumnSettingSch.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbColumnSettingSch.AutoSize = True
        Me.llbColumnSettingSch.Location = New System.Drawing.Point(-1, 10)
        Me.llbColumnSettingSch.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbColumnSettingSch.Name = "llbColumnSettingSch"
        Me.llbColumnSettingSch.Size = New System.Drawing.Size(103, 17)
        Me.llbColumnSettingSch.TabIndex = 108
        Me.llbColumnSettingSch.TabStop = True
        Me.llbColumnSettingSch.Text = "Column Setting"
        Me.llbColumnSettingSch.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'dgvBar
        '
        Me.dgvBar.AllowUserToAddRows = False
        Me.dgvBar.AllowUserToDeleteRows = False
        Me.dgvBar.AllowUserToResizeColumns = False
        Me.dgvBar.AllowUserToResizeRows = False
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.Color.Transparent
        Me.dgvBar.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle10
        Me.dgvBar.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvBar.BackgroundColor = System.Drawing.Color.White
        Me.dgvBar.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        Me.dgvBar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvBar.ColumnHeadersVisible = False
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle11.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle11.ForeColor = System.Drawing.Color.Red
        DataGridViewCellStyle11.SelectionBackColor = System.Drawing.Color.Transparent
        DataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvBar.DefaultCellStyle = DataGridViewCellStyle11
        Me.dgvBar.Location = New System.Drawing.Point(0, 62)
        Me.dgvBar.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvBar.MultiSelect = False
        Me.dgvBar.Name = "dgvBar"
        Me.dgvBar.RowHeadersVisible = False
        Me.dgvBar.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvBar.RowTemplate.Height = 24
        Me.dgvBar.Size = New System.Drawing.Size(743, 401)
        Me.dgvBar.TabIndex = 100
        Me.dgvBar.TabStop = False
        '
        'dgvHeader0
        '
        Me.dgvHeader0.AllowUserToAddRows = False
        Me.dgvHeader0.AllowUserToDeleteRows = False
        Me.dgvHeader0.AllowUserToResizeColumns = False
        Me.dgvHeader0.AllowUserToResizeRows = False
        Me.dgvHeader0.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvHeader0.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle12.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvHeader0.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle12
        Me.dgvHeader0.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvHeader0.Location = New System.Drawing.Point(0, 33)
        Me.dgvHeader0.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvHeader0.Name = "dgvHeader0"
        Me.dgvHeader0.ReadOnly = True
        Me.dgvHeader0.RowHeadersVisible = False
        Me.dgvHeader0.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvHeader0.RowTemplate.Height = 24
        Me.dgvHeader0.Size = New System.Drawing.Size(743, 101)
        Me.dgvHeader0.TabIndex = 99
        '
        'dgvHeader1
        '
        Me.dgvHeader1.AllowUserToAddRows = False
        Me.dgvHeader1.AllowUserToDeleteRows = False
        Me.dgvHeader1.AllowUserToResizeColumns = False
        Me.dgvHeader1.AllowUserToResizeRows = False
        Me.dgvHeader1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvHeader1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle13.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle13.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.5!)
        DataGridViewCellStyle13.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvHeader1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle13
        Me.dgvHeader1.ColumnHeadersHeight = 29
        Me.dgvHeader1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvHeader1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.dgvHeader1.Location = New System.Drawing.Point(0, 0)
        Me.dgvHeader1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvHeader1.Name = "dgvHeader1"
        Me.dgvHeader1.ReadOnly = True
        Me.dgvHeader1.RowHeadersVisible = False
        Me.dgvHeader1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvHeader1.RowTemplate.Height = 24
        Me.dgvHeader1.Size = New System.Drawing.Size(743, 100)
        Me.dgvHeader1.TabIndex = 98
        '
        'Button8
        '
        Me.Button8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button8.Location = New System.Drawing.Point(57, 430)
        Me.Button8.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(33, 31)
        Me.Button8.TabIndex = 97
        Me.Button8.Text = "<"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button9.Location = New System.Drawing.Point(95, 430)
        Me.Button9.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(33, 31)
        Me.Button9.TabIndex = 96
        Me.Button9.Text = ">"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(15, 430)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(39, 31)
        Me.Button2.TabIndex = 72
        Me.Button2.Text = "<<"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(127, 430)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(39, 31)
        Me.Button1.TabIndex = 71
        Me.Button1.Text = ">>"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BtnClearGanttchart)
        Me.GroupBox2.Controls.Add(Me.Button7)
        Me.GroupBox2.Controls.Add(Me.Button5)
        Me.GroupBox2.Controls.Add(Me.Button6)
        Me.GroupBox2.Controls.Add(Me.Button4)
        Me.GroupBox2.Controls.Add(Me.BtnLoad)
        Me.GroupBox2.Controls.Add(Me.BtnRemove)
        Me.GroupBox2.Controls.Add(Me.BtnAdd)
        Me.GroupBox2.Controls.Add(Me.BtnSave)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Button3)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 186)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Size = New System.Drawing.Size(451, 100)
        Me.GroupBox2.TabIndex = 93
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "GroupBox2"
        Me.GroupBox2.Visible = False
        '
        'BtnClearGanttchart
        '
        Me.BtnClearGanttchart.Location = New System.Drawing.Point(39, 43)
        Me.BtnClearGanttchart.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnClearGanttchart.Name = "BtnClearGanttchart"
        Me.BtnClearGanttchart.Size = New System.Drawing.Size(144, 31)
        Me.BtnClearGanttchart.TabIndex = 69
        Me.BtnClearGanttchart.Text = "Clear Gannt Chart"
        Me.BtnClearGanttchart.UseVisualStyleBackColor = True
        Me.BtnClearGanttchart.Visible = False
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(103, 62)
        Me.Button7.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(88, 31)
        Me.Button7.TabIndex = 78
        Me.Button7.Text = "Orange"
        Me.Button7.UseVisualStyleBackColor = True
        Me.Button7.Visible = False
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(7, 22)
        Me.Button5.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(88, 31)
        Me.Button5.TabIndex = 77
        Me.Button5.Text = "Purple"
        Me.Button5.UseVisualStyleBackColor = True
        Me.Button5.Visible = False
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(103, 22)
        Me.Button6.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(88, 31)
        Me.Button6.TabIndex = 76
        Me.Button6.Text = "Green"
        Me.Button6.UseVisualStyleBackColor = True
        Me.Button6.Visible = False
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(7, 62)
        Me.Button4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(88, 31)
        Me.Button4.TabIndex = 75
        Me.Button4.Text = "Yellow"
        Me.Button4.UseVisualStyleBackColor = True
        Me.Button4.Visible = False
        '
        'BtnLoad
        '
        Me.BtnLoad.Location = New System.Drawing.Point(95, 39)
        Me.BtnLoad.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnLoad.Name = "BtnLoad"
        Me.BtnLoad.Size = New System.Drawing.Size(88, 31)
        Me.BtnLoad.TabIndex = 68
        Me.BtnLoad.Text = "Load"
        Me.BtnLoad.UseVisualStyleBackColor = True
        Me.BtnLoad.Visible = False
        '
        'BtnRemove
        '
        Me.BtnRemove.Location = New System.Drawing.Point(53, 39)
        Me.BtnRemove.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnRemove.Name = "BtnRemove"
        Me.BtnRemove.Size = New System.Drawing.Size(88, 31)
        Me.BtnRemove.TabIndex = 65
        Me.BtnRemove.Text = "Remove"
        Me.BtnRemove.UseVisualStyleBackColor = True
        Me.BtnRemove.Visible = False
        '
        'BtnAdd
        '
        Me.BtnAdd.Location = New System.Drawing.Point(29, 22)
        Me.BtnAdd.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(88, 31)
        Me.BtnAdd.TabIndex = 63
        Me.BtnAdd.Text = "Add"
        Me.BtnAdd.UseVisualStyleBackColor = True
        Me.BtnAdd.Visible = False
        '
        'BtnSave
        '
        Me.BtnSave.Location = New System.Drawing.Point(75, 22)
        Me.BtnSave.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(88, 31)
        Me.BtnSave.TabIndex = 62
        Me.BtnSave.Text = "Save"
        Me.BtnSave.UseVisualStyleBackColor = True
        Me.BtnSave.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(92, 39)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(25, 17)
        Me.Label1.TabIndex = 67
        Me.Label1.Text = "To"
        Me.Label1.Visible = False
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(39, 32)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(88, 31)
        Me.Button3.TabIndex = 74
        Me.Button3.Text = "Blue"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.BtnClear)
        Me.GroupBox1.Controls.Add(Me.txtPrimary)
        Me.GroupBox1.Controls.Add(Me.Dgv)
        Me.GroupBox1.Controls.Add(Me.BtnSet2)
        Me.GroupBox1.Controls.Add(Me.TxtTaks)
        Me.GroupBox1.Controls.Add(Me.CboTaskName)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.DtpFromDate)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.DtpToDate)
        Me.GroupBox1.Controls.Add(Me.txtIdentifier)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 270)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(975, 130)
        Me.GroupBox1.TabIndex = 64
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Bar Chart Info"
        Me.GroupBox1.Visible = False
        '
        'BtnClear
        '
        Me.BtnClear.Location = New System.Drawing.Point(407, 57)
        Me.BtnClear.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(88, 31)
        Me.BtnClear.TabIndex = 60
        Me.BtnClear.Text = "Clear"
        Me.BtnClear.UseVisualStyleBackColor = True
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(97, 23)
        Me.txtPrimary.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(36, 22)
        Me.txtPrimary.TabIndex = 58
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'Dgv
        '
        Me.Dgv.AllowUserToAddRows = False
        Me.Dgv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.Dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Dgv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPrimary, Me.ColTaskName, Me.ColFromDate, Me.ColToDate, Me.ColIdentifier, Me.ColProgress, Me.ColTipeBar, Me.Col_a, Me.Col_r, Me.Col_g, Me.Col_b, Me.ColVisible})
        Me.Dgv.Location = New System.Drawing.Point(503, 23)
        Me.Dgv.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Dgv.Name = "Dgv"
        Me.Dgv.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        Me.Dgv.RowHeadersVisible = False
        Me.Dgv.RowTemplate.Height = 24
        Me.Dgv.Size = New System.Drawing.Size(463, 98)
        Me.Dgv.TabIndex = 59
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.MinimumWidth = 35
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 51
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.Width = 81
        '
        'ColTaskName
        '
        Me.ColTaskName.HeaderText = "Task Name"
        Me.ColTaskName.Name = "ColTaskName"
        Me.ColTaskName.ReadOnly = True
        Me.ColTaskName.Width = 105
        '
        'ColFromDate
        '
        Me.ColFromDate.HeaderText = "From Date"
        Me.ColFromDate.Name = "ColFromDate"
        Me.ColFromDate.ReadOnly = True
        Me.ColFromDate.Width = 99
        '
        'ColToDate
        '
        Me.ColToDate.HeaderText = "To Date"
        Me.ColToDate.Name = "ColToDate"
        Me.ColToDate.ReadOnly = True
        Me.ColToDate.Width = 84
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.ReadOnly = True
        Me.ColIdentifier.Width = 87
        '
        'ColProgress
        '
        Me.ColProgress.HeaderText = "Progress"
        Me.ColProgress.Name = "ColProgress"
        Me.ColProgress.Width = 90
        '
        'ColTipeBar
        '
        Me.ColTipeBar.HeaderText = "Tipe Bar"
        Me.ColTipeBar.Name = "ColTipeBar"
        Me.ColTipeBar.Width = 87
        '
        'Col_a
        '
        Me.Col_a.HeaderText = "a"
        Me.Col_a.Name = "Col_a"
        Me.Col_a.Width = 41
        '
        'Col_r
        '
        Me.Col_r.HeaderText = "r"
        Me.Col_r.Name = "Col_r"
        Me.Col_r.Width = 38
        '
        'Col_g
        '
        Me.Col_g.HeaderText = "g"
        Me.Col_g.Name = "Col_g"
        Me.Col_g.Width = 41
        '
        'Col_b
        '
        Me.Col_b.HeaderText = "b"
        Me.Col_b.Name = "Col_b"
        Me.Col_b.Width = 41
        '
        'ColVisible
        '
        Me.ColVisible.HeaderText = "Visible"
        Me.ColVisible.Name = "ColVisible"
        Me.ColVisible.Width = 74
        '
        'BtnSet2
        '
        Me.BtnSet2.Location = New System.Drawing.Point(407, 23)
        Me.BtnSet2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnSet2.Name = "BtnSet2"
        Me.BtnSet2.Size = New System.Drawing.Size(88, 31)
        Me.BtnSet2.TabIndex = 50
        Me.BtnSet2.Text = "Set"
        Me.BtnSet2.UseVisualStyleBackColor = True
        '
        'TxtTaks
        '
        Me.TxtTaks.Location = New System.Drawing.Point(167, 23)
        Me.TxtTaks.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtTaks.Name = "TxtTaks"
        Me.TxtTaks.Size = New System.Drawing.Size(231, 22)
        Me.TxtTaks.TabIndex = 23
        '
        'CboTaskName
        '
        Me.CboTaskName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboTaskName.FormattingEnabled = True
        Me.CboTaskName.Location = New System.Drawing.Point(167, 23)
        Me.CboTaskName.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CboTaskName.Name = "CboTaskName"
        Me.CboTaskName.Size = New System.Drawing.Size(231, 24)
        Me.CboTaskName.TabIndex = 29
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 64)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 17)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "&Start :"
        '
        'DtpFromDate
        '
        Me.DtpFromDate.CustomFormat = "dd MMM yyyy HH:mm"
        Me.DtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFromDate.Location = New System.Drawing.Point(167, 57)
        Me.DtpFromDate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DtpFromDate.Name = "DtpFromDate"
        Me.DtpFromDate.Size = New System.Drawing.Size(231, 22)
        Me.DtpFromDate.TabIndex = 24
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 27)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 17)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "&Task Name :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 96)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(57, 17)
        Me.Label4.TabIndex = 28
        Me.Label4.Text = "&Finish : "
        '
        'DtpToDate
        '
        Me.DtpToDate.CustomFormat = "dd MMM yyyy HH:mm"
        Me.DtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpToDate.Location = New System.Drawing.Point(167, 89)
        Me.DtpToDate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DtpToDate.Name = "DtpToDate"
        Me.DtpToDate.Size = New System.Drawing.Size(231, 22)
        Me.DtpToDate.TabIndex = 25
        '
        'txtIdentifier
        '
        Me.txtIdentifier.Location = New System.Drawing.Point(65, 63)
        Me.txtIdentifier.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtIdentifier.Name = "txtIdentifier"
        Me.txtIdentifier.Size = New System.Drawing.Size(36, 22)
        Me.txtIdentifier.TabIndex = 59
        Me.txtIdentifier.Visible = False
        '
        'HScrollBar1
        '
        Me.HScrollBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.HScrollBar1.Location = New System.Drawing.Point(0, 446)
        Me.HScrollBar1.Name = "HScrollBar1"
        Me.HScrollBar1.Size = New System.Drawing.Size(743, 17)
        Me.HScrollBar1.TabIndex = 94
        '
        'dtpBarFrom
        '
        Me.dtpBarFrom.CustomFormat = "dd MMM yyyy HH.mm"
        Me.dtpBarFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpBarFrom.Location = New System.Drawing.Point(653, 34)
        Me.dtpBarFrom.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpBarFrom.Name = "dtpBarFrom"
        Me.dtpBarFrom.Size = New System.Drawing.Size(179, 22)
        Me.dtpBarFrom.TabIndex = 107
        Me.dtpBarFrom.TabStop = False
        Me.dtpBarFrom.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(335, 14)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(25, 17)
        Me.Label8.TabIndex = 105
        Me.Label8.Text = "To"
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(368, 11)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(184, 22)
        Me.dtpTo.TabIndex = 106
        '
        'btnAddOT
        '
        Me.btnAddOT.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAddOT.BackColor = System.Drawing.Color.Transparent
        Me.btnAddOT.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnAddOT.ForeColor = System.Drawing.Color.Black
        Me.btnAddOT.Location = New System.Drawing.Point(151, 538)
        Me.btnAddOT.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnAddOT.Name = "btnAddOT"
        Me.btnAddOT.Size = New System.Drawing.Size(131, 53)
        Me.btnAddOT.TabIndex = 109
        Me.btnAddOT.TabStop = False
        Me.btnAddOT.Text = "Add/Remove Overtime"
        Me.btnAddOT.UseVisualStyleBackColor = True
        '
        'btnAddSchedule
        '
        Me.btnAddSchedule.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAddSchedule.BackColor = System.Drawing.Color.Transparent
        Me.btnAddSchedule.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnAddSchedule.ForeColor = System.Drawing.Color.Black
        Me.btnAddSchedule.Location = New System.Drawing.Point(12, 538)
        Me.btnAddSchedule.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnAddSchedule.Name = "btnAddSchedule"
        Me.btnAddSchedule.Size = New System.Drawing.Size(131, 53)
        Me.btnAddSchedule.TabIndex = 110
        Me.btnAddSchedule.TabStop = False
        Me.btnAddSchedule.Text = "Add Schedule"
        Me.btnAddSchedule.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'btnWorkshift
        '
        Me.btnWorkshift.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnWorkshift.BackColor = System.Drawing.Color.Transparent
        Me.btnWorkshift.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnWorkshift.ForeColor = System.Drawing.Color.Black
        Me.btnWorkshift.Location = New System.Drawing.Point(291, 538)
        Me.btnWorkshift.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnWorkshift.Name = "btnWorkshift"
        Me.btnWorkshift.Size = New System.Drawing.Size(131, 53)
        Me.btnWorkshift.TabIndex = 111
        Me.btnWorkshift.TabStop = False
        Me.btnWorkshift.Text = "Workshift"
        Me.btnWorkshift.UseVisualStyleBackColor = True
        '
        'BackgroundWorker5
        '
        '
        'BackgroundWorker6
        '
        '
        'bgwFullReload
        '
        Me.bgwFullReload.WorkerReportsProgress = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GroupBox3.Controls.Add(Me.LblProgress)
        Me.GroupBox3.Controls.Add(Me.prgBar1)
        Me.GroupBox3.Location = New System.Drawing.Point(328, 249)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox3.Size = New System.Drawing.Size(369, 80)
        Me.GroupBox3.TabIndex = 112
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Visible = False
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(5, 48)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(51, 17)
        Me.LblProgress.TabIndex = 1
        Me.LblProgress.Text = "Label9"
        '
        'prgBar1
        '
        Me.prgBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBar1.Location = New System.Drawing.Point(5, 21)
        Me.prgBar1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prgBar1.Name = "prgBar1"
        Me.prgBar1.Size = New System.Drawing.Size(357, 23)
        Me.prgBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.prgBar1.TabIndex = 0
        '
        'bgwFormLoad
        '
        Me.bgwFormLoad.WorkerReportsProgress = True
        '
        'btnFindMO
        '
        Me.btnFindMO.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnFindMO.BackColor = System.Drawing.Color.Transparent
        Me.btnFindMO.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnFindMO.ForeColor = System.Drawing.Color.Black
        Me.btnFindMO.Location = New System.Drawing.Point(845, 538)
        Me.btnFindMO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFindMO.Name = "btnFindMO"
        Me.btnFindMO.Size = New System.Drawing.Size(131, 53)
        Me.btnFindMO.TabIndex = 113
        Me.btnFindMO.TabStop = False
        Me.btnFindMO.Text = "Find MO"
        Me.btnFindMO.UseVisualStyleBackColor = True
        '
        'btnAddRemoveHoliday
        '
        Me.btnAddRemoveHoliday.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAddRemoveHoliday.BackColor = System.Drawing.Color.Transparent
        Me.btnAddRemoveHoliday.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnAddRemoveHoliday.ForeColor = System.Drawing.Color.Black
        Me.btnAddRemoveHoliday.Location = New System.Drawing.Point(845, 538)
        Me.btnAddRemoveHoliday.Margin = New System.Windows.Forms.Padding(4)
        Me.btnAddRemoveHoliday.Name = "btnAddRemoveHoliday"
        Me.btnAddRemoveHoliday.Size = New System.Drawing.Size(131, 53)
        Me.btnAddRemoveHoliday.TabIndex = 115
        Me.btnAddRemoveHoliday.TabStop = False
        Me.btnAddRemoveHoliday.Text = "Add/Remove Holiday"
        Me.btnAddRemoveHoliday.UseVisualStyleBackColor = True
        '
        'FrmTrans_ProductionSchedule
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1035, 639)
        Me.Controls.Add(Me.btnAddRemoveHoliday)
        Me.Controls.Add(Me.btnFindMO)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.btnWorkshift)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.btnAddSchedule)
        Me.Controls.Add(Me.dtpBarFrom)
        Me.Controls.Add(Me.btnAddOT)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.llbSetting)
        Me.Controls.Add(Me.btnAddSetupTime)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.btnException)
        Me.Controls.Add(Me.btnPrintSchedule)
        Me.Controls.Add(Me.cboTimeframe)
        Me.Controls.Add(Me.btnAllocateWorkshift)
        Me.Controls.Add(Me.btnPrintPickList)
        Me.Controls.Add(Me.btnAddDownTime)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.dtpFirmPlanned)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.chkShowProgress)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.CboAturWaktu)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.chkIncompleteRM)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.chkLate)
        Me.Controls.Add(Me.dtpToday)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.chkShowPickList)
        Me.Controls.Add(Me.cboGroupBy)
        Me.Controls.Add(Me.btnAddDelay)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimumSize = New System.Drawing.Size(1043, 676)
        Me.Name = "FrmTrans_ProductionSchedule"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Production Schedule"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage_SchTable.ResumeLayout(False)
        CType(Me.dgvSch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.tgvSch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cms1.ResumeLayout(False)
        CType(Me.dgvBar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvHeader0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvHeader1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.Dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CboAturWaktu As System.Windows.Forms.ComboBox
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents ColTreeNode_DocumentNo As AdvancedDataGridView.TreeGridColumn
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboGroupBy As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpFirmPlanned As System.Windows.Forms.DateTimePicker
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents dtpToday As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkShowProgress As System.Windows.Forms.CheckBox
    Friend WithEvents chkLate As System.Windows.Forms.CheckBox
    Friend WithEvents btnAddDownTime As System.Windows.Forms.Button
    Friend WithEvents chkIncompleteRM As System.Windows.Forms.CheckBox
    Friend WithEvents btnPrintPickList As System.Windows.Forms.Button
    Friend WithEvents chkShowPickList As System.Windows.Forms.CheckBox
    Friend WithEvents btnAllocateWorkshift As System.Windows.Forms.Button
    Friend WithEvents cboTimeframe As System.Windows.Forms.ComboBox
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnPrintSchedule As System.Windows.Forms.Button
    Friend WithEvents btnException As System.Windows.Forms.Button
    Friend WithEvents ColTGV_DocNo As AdvancedDataGridView.TreeGridColumn
    Friend WithEvents ColTGV_Late As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColTGV_IncompleteRM As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColTGV_PrintPickList As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColTGV_Details As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents btnAddDelay As System.Windows.Forms.Button
    Friend WithEvents btnAddSetupTime As System.Windows.Forms.Button
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents llbSetting As System.Windows.Forms.LinkLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_SchGanttChart As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents tgvSch As AdvancedDataGridView.TreeGridView
    Friend WithEvents HScrollBar1 As System.Windows.Forms.HScrollBar
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnClearGanttchart As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents BtnLoad As System.Windows.Forms.Button
    Friend WithEvents BtnRemove As System.Windows.Forms.Button
    Friend WithEvents BtnAdd As System.Windows.Forms.Button
    Friend WithEvents BtnSave As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnClear As System.Windows.Forms.Button
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents Dgv As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTaskName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFromDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProgress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeBar As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_a As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_r As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_g As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_b As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVisible As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BtnSet2 As System.Windows.Forms.Button
    Friend WithEvents TxtTaks As System.Windows.Forms.TextBox
    Friend WithEvents CboTaskName As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtIdentifier As System.Windows.Forms.TextBox
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TabPage_SchTable As System.Windows.Forms.TabPage
    Friend WithEvents dgvSch As System.Windows.Forms.DataGridView
    Friend WithEvents DtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents ColSch_JONumber1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_WorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_StartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_StartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_EndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_EndTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_Type As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_WorkOrderNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_MONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_JONumber2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_Process As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_Progress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_PrimaryWorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_PrimaryWorkShift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_PrimaryMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_a As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_r As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_g As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_b As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvBar As System.Windows.Forms.DataGridView
    Friend WithEvents dgvHeader0 As System.Windows.Forms.DataGridView
    Friend WithEvents dgvHeader1 As System.Windows.Forms.DataGridView
    Friend WithEvents dtpBarFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents llbColumnSettingSch As System.Windows.Forms.LinkLabel
    Friend WithEvents llbExpandAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCollapsAll As System.Windows.Forms.LinkLabel
    Friend WithEvents btnAddOT As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnAddSchedule As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents btnWorkshift As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker5 As System.ComponentModel.BackgroundWorker
    Friend WithEvents cms1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tms_ChangeColor As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CreateLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteJobOrderToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReorderScheduleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptimizeScheduleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BackgroundWorker6 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ViewLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CloseJOToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ColLeft_Tree As AdvancedDataGridView.TreeGridColumn
    Friend WithEvents ColLeft_ProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_WONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_WorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_PrimaryLine As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_LineCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_ProductionRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_QtyProduce As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_TotalWH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_DayOfWeekStart As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_StartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_StartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_DayOfWeekFinish As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_FinishDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_FinishTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_Progress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_PrimaryWorkUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_PrimaryWorkshift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_PrimaryMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_a As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_r As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_g As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_b As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_NodeLevel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_NodeLv0Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_NodeLv1Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_PrimaryLocationInput As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_LocationInput As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_PrimaryLocationOutput As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_LocationOutput As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_ActualGood As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_ActualReject As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeft_Actualproduce As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SplitMergeScheduleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents bgwFullReload As System.ComponentModel.BackgroundWorker
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents prgBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents bgwFormLoad As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnFindMO As System.Windows.Forms.Button
    Friend WithEvents FindScheduleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ShowAllMOScheduleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnAddRemoveHoliday As System.Windows.Forms.Button
End Class
