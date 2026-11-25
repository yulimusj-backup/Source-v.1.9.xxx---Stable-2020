<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPLApproval
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColItem_DocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_StartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_EndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Status = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColItem_StatusNotes = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColItem_Modifiy = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColItem_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Budget = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColItem_Print = New System.Windows.Forms.DataGridViewButtonColumn
        Me.BtnWait = New System.Windows.Forms.Button
        Me.RboRejected = New System.Windows.Forms.RadioButton
        Me.RboApproved = New System.Windows.Forms.RadioButton
        Me.Label14 = New System.Windows.Forms.Label
        Me.BtnReject = New System.Windows.Forms.Button
        Me.RboWaitingForApproval = New System.Windows.Forms.RadioButton
        Me.BtnApprove = New System.Windows.Forms.Button
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.txtTotalDuration = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtTotalDurationPerEmp = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.TxtReason = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.txtDocumentNo = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TabControl3 = New System.Windows.Forms.TabControl
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.colEmpl_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEmpl_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEmpl_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEmpl_NIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEmpl_JobPosition = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEmpl_Department = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEmpl_TotalHours = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEmpl_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEmpl_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEmpl_PrimaryPos = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colEmpl_PrimaryDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage7 = New System.Windows.Forms.TabPage
        Me.dgvSchDetail = New System.Windows.Forms.DataGridView
        Me.ColSch_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_DayOfWeek = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_Shift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_NormalWH = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_OTtype = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_Start = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_Break = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_End = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_Duration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOverride = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColSch_Remark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSch_PrimaryShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage8 = New System.Windows.Forms.TabPage
        Me.dgvSchEmp = New System.Windows.Forms.DataGridView
        Me.btnClose = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.DgvEmpFilter = New System.Windows.Forms.DataGridView
        Me.ColFilterEmp_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterEmp_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterEmp_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterEmp_NIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterEmp_Dept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterEmp_JobPosition = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterEmp_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.DgvDeptFilter = New System.Windows.Forms.DataGridView
        Me.ColFilterDept_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterDept_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterDept_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterDept_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.DgvPosFilter = New System.Windows.Forms.DataGridView
        Me.ColFilterJobPos_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterJobPos_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterJobPos_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilterJobPos_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.llbClearFilter = New System.Windows.Forms.LinkLabel
        Me.DtpDocDateTo = New System.Windows.Forms.DateTimePicker
        Me.Label9 = New System.Windows.Forms.Label
        Me.DtpDocDateFrom = New System.Windows.Forms.DateTimePicker
        Me.ChkDocDate = New System.Windows.Forms.CheckBox
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.BgwDepartment = New System.ComponentModel.BackgroundWorker
        Me.BgwEmployee = New System.ComponentModel.BackgroundWorker
        Me.BgwJobPosition = New System.ComponentModel.BackgroundWorker
        Me.ColSchEmp_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_Emp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_DayOfWeek = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_Shift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_NormalWH = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_StartEarly = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_BreakEarly = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_EndEarly = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_EarlyOTDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_OverrideEarlyOT = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColSchEmp_StartOT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_BreakOT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_EndOT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_OTDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_OverrideOT = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColSchEmp_PrimaryShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchEmp_PrimaryEmp = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl3.SuspendLayout()
        Me.TabPage6.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        CType(Me.dgvSchDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage8.SuspendLayout()
        CType(Me.dgvSchEmp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.DgvEmpFilter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.DgvDeptFilter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.DgvPosFilter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_DocumentNo, Me.ColItem_Date, Me.ColItem_StartDate, Me.ColItem_EndDate, Me.ColItem_Status, Me.ColItem_StatusNotes, Me.ColItem_Modifiy, Me.ColItem_Primary, Me.ColItem_Budget, Me.ColItem_Print})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(2, 2)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(714, 149)
        Me.dgvItem.TabIndex = 48
        '
        'ColItem_DocumentNo
        '
        Me.ColItem_DocumentNo.DataPropertyName = "no transaksi"
        Me.ColItem_DocumentNo.HeaderText = "Document No"
        Me.ColItem_DocumentNo.Name = "ColItem_DocumentNo"
        Me.ColItem_DocumentNo.ReadOnly = True
        Me.ColItem_DocumentNo.Width = 98
        '
        'ColItem_Date
        '
        Me.ColItem_Date.DataPropertyName = "tanggal"
        Me.ColItem_Date.HeaderText = "Date"
        Me.ColItem_Date.Name = "ColItem_Date"
        Me.ColItem_Date.ReadOnly = True
        Me.ColItem_Date.Width = 55
        '
        'ColItem_StartDate
        '
        Me.ColItem_StartDate.HeaderText = "Start"
        Me.ColItem_StartDate.Name = "ColItem_StartDate"
        Me.ColItem_StartDate.ReadOnly = True
        Me.ColItem_StartDate.Width = 54
        '
        'ColItem_EndDate
        '
        Me.ColItem_EndDate.HeaderText = "End"
        Me.ColItem_EndDate.Name = "ColItem_EndDate"
        Me.ColItem_EndDate.ReadOnly = True
        Me.ColItem_EndDate.Width = 51
        '
        'ColItem_Status
        '
        Me.ColItem_Status.AutoComplete = False
        Me.ColItem_Status.HeaderText = "Status"
        Me.ColItem_Status.Items.AddRange(New Object() {"Waiting For Approval", "Approved", "Rejected"})
        Me.ColItem_Status.Name = "ColItem_Status"
        Me.ColItem_Status.ReadOnly = True
        Me.ColItem_Status.Width = 43
        '
        'ColItem_StatusNotes
        '
        Me.ColItem_StatusNotes.HeaderText = "Status Notes"
        Me.ColItem_StatusNotes.Name = "ColItem_StatusNotes"
        Me.ColItem_StatusNotes.ReadOnly = True
        Me.ColItem_StatusNotes.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItem_StatusNotes.Width = 74
        '
        'ColItem_Modifiy
        '
        Me.ColItem_Modifiy.HeaderText = "Modify"
        Me.ColItem_Modifiy.Name = "ColItem_Modifiy"
        Me.ColItem_Modifiy.Width = 44
        '
        'ColItem_Primary
        '
        Me.ColItem_Primary.DataPropertyName = "primary"
        Me.ColItem_Primary.HeaderText = "Primary"
        Me.ColItem_Primary.Name = "ColItem_Primary"
        Me.ColItem_Primary.ReadOnly = True
        Me.ColItem_Primary.Visible = False
        Me.ColItem_Primary.Width = 66
        '
        'ColItem_Budget
        '
        Me.ColItem_Budget.HeaderText = "Budget"
        Me.ColItem_Budget.Name = "ColItem_Budget"
        Me.ColItem_Budget.ReadOnly = True
        Me.ColItem_Budget.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColItem_Budget.Width = 47
        '
        'ColItem_Print
        '
        Me.ColItem_Print.HeaderText = "Print"
        Me.ColItem_Print.Name = "ColItem_Print"
        Me.ColItem_Print.ReadOnly = True
        Me.ColItem_Print.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItem_Print.Width = 34
        '
        'BtnWait
        '
        Me.BtnWait.Location = New System.Drawing.Point(511, 10)
        Me.BtnWait.Name = "BtnWait"
        Me.BtnWait.Size = New System.Drawing.Size(75, 23)
        Me.BtnWait.TabIndex = 5
        Me.BtnWait.Text = "Wait"
        Me.BtnWait.UseVisualStyleBackColor = True
        '
        'RboRejected
        '
        Me.RboRejected.AutoSize = True
        Me.RboRejected.Location = New System.Drawing.Point(272, 13)
        Me.RboRejected.Name = "RboRejected"
        Me.RboRejected.Size = New System.Drawing.Size(68, 17)
        Me.RboRejected.TabIndex = 2
        Me.RboRejected.Text = "Rejected"
        Me.RboRejected.UseVisualStyleBackColor = True
        '
        'RboApproved
        '
        Me.RboApproved.AutoSize = True
        Me.RboApproved.Location = New System.Drawing.Point(195, 13)
        Me.RboApproved.Name = "RboApproved"
        Me.RboApproved.Size = New System.Drawing.Size(71, 17)
        Me.RboApproved.TabIndex = 1
        Me.RboApproved.Text = "Approved"
        Me.RboApproved.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(16, 15)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(43, 13)
        Me.Label14.TabIndex = 50
        Me.Label14.Text = "Show : "
        '
        'BtnReject
        '
        Me.BtnReject.Location = New System.Drawing.Point(430, 10)
        Me.BtnReject.Name = "BtnReject"
        Me.BtnReject.Size = New System.Drawing.Size(75, 23)
        Me.BtnReject.TabIndex = 4
        Me.BtnReject.Text = "Reject"
        Me.BtnReject.UseVisualStyleBackColor = True
        '
        'RboWaitingForApproval
        '
        Me.RboWaitingForApproval.AutoSize = True
        Me.RboWaitingForApproval.Checked = True
        Me.RboWaitingForApproval.Location = New System.Drawing.Point(64, 14)
        Me.RboWaitingForApproval.Name = "RboWaitingForApproval"
        Me.RboWaitingForApproval.Size = New System.Drawing.Size(124, 17)
        Me.RboWaitingForApproval.TabIndex = 0
        Me.RboWaitingForApproval.TabStop = True
        Me.RboWaitingForApproval.Text = "Waiting For Approval"
        Me.RboWaitingForApproval.UseVisualStyleBackColor = True
        '
        'BtnApprove
        '
        Me.BtnApprove.Location = New System.Drawing.Point(349, 10)
        Me.BtnApprove.Name = "BtnApprove"
        Me.BtnApprove.Size = New System.Drawing.Size(75, 23)
        Me.BtnApprove.TabIndex = 3
        Me.BtnApprove.Text = "Approve"
        Me.BtnApprove.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(16, 247)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtTotalDuration)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtTotalDurationPerEmp)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpFrom)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtRemarks)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtReason)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtDocumentNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TabControl3)
        Me.SplitContainer1.Size = New System.Drawing.Size(718, 257)
        Me.SplitContainer1.SplitterDistance = 352
        Me.SplitContainer1.TabIndex = 49
        Me.SplitContainer1.TabStop = False
        '
        'txtTotalDuration
        '
        Me.txtTotalDuration.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTotalDuration.Location = New System.Drawing.Point(79, 126)
        Me.txtTotalDuration.MaxLength = 250
        Me.txtTotalDuration.Name = "txtTotalDuration"
        Me.txtTotalDuration.ReadOnly = True
        Me.txtTotalDuration.Size = New System.Drawing.Size(272, 20)
        Me.txtTotalDuration.TabIndex = 100000014
        Me.txtTotalDuration.TabStop = False
        Me.txtTotalDuration.Text = "00.00"
        Me.txtTotalDuration.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(281, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 13)
        Me.Label2.TabIndex = 100000015
        Me.Label2.Text = "per employee"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(4, 128)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(74, 13)
        Me.Label7.TabIndex = 100000013
        Me.Label7.Text = "Total Duration"
        '
        'Label10
        '
        Me.Label10.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(180, 128)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(12, 13)
        Me.Label10.TabIndex = 100000011
        Me.Label10.Text = "/"
        '
        'TxtTotalDurationPerEmp
        '
        Me.TxtTotalDurationPerEmp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalDurationPerEmp.Location = New System.Drawing.Point(195, 126)
        Me.TxtTotalDurationPerEmp.MaxLength = 250
        Me.TxtTotalDurationPerEmp.Name = "TxtTotalDurationPerEmp"
        Me.TxtTotalDurationPerEmp.ReadOnly = True
        Me.TxtTotalDurationPerEmp.Size = New System.Drawing.Size(81, 20)
        Me.TxtTotalDurationPerEmp.TabIndex = 100000012
        Me.TxtTotalDurationPerEmp.Text = "00.00"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 153)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 100000008
        Me.Label3.Text = "Remarks"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 106)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 100000009
        Me.Label4.Text = "From"
        '
        'DtpTo
        '
        Me.DtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Enabled = False
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(236, 102)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(116, 20)
        Me.DtpTo.TabIndex = 17
        Me.DtpTo.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(212, 106)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(20, 13)
        Me.Label5.TabIndex = 100000010
        Me.Label5.Text = "To"
        '
        'DtpFrom
        '
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Enabled = False
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(79, 102)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(128, 20)
        Me.DtpFrom.TabIndex = 16
        Me.DtpFrom.TabStop = False
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(79, 150)
        Me.txtRemarks.MaxLength = 250
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.ReadOnly = True
        Me.txtRemarks.Size = New System.Drawing.Size(274, 104)
        Me.txtRemarks.TabIndex = 100000006
        Me.txtRemarks.TabStop = False
        '
        'TxtReason
        '
        Me.TxtReason.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReason.BackColor = System.Drawing.SystemColors.Control
        Me.TxtReason.Location = New System.Drawing.Point(79, 28)
        Me.TxtReason.MaxLength = 250
        Me.TxtReason.Multiline = True
        Me.TxtReason.Name = "TxtReason"
        Me.TxtReason.ReadOnly = True
        Me.TxtReason.Size = New System.Drawing.Size(272, 68)
        Me.TxtReason.TabIndex = 15
        Me.TxtReason.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(2, 30)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 13)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "Reason"
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(9, 3)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(29, 20)
        Me.txtPrimary.TabIndex = 12
        Me.txtPrimary.TabStop = False
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Enabled = False
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(226, 3)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(126, 20)
        Me.dtpDate.TabIndex = 14
        Me.dtpDate.TabStop = False
        '
        'txtDocumentNo
        '
        Me.txtDocumentNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDocumentNo.BackColor = System.Drawing.SystemColors.Control
        Me.txtDocumentNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDocumentNo.Location = New System.Drawing.Point(79, 3)
        Me.txtDocumentNo.MaxLength = 250
        Me.txtDocumentNo.Name = "txtDocumentNo"
        Me.txtDocumentNo.ReadOnly = True
        Me.txtDocumentNo.Size = New System.Drawing.Size(142, 20)
        Me.txtDocumentNo.TabIndex = 13
        Me.txtDocumentNo.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Document No"
        '
        'TabControl3
        '
        Me.TabControl3.Controls.Add(Me.TabPage6)
        Me.TabControl3.Controls.Add(Me.TabPage7)
        Me.TabControl3.Controls.Add(Me.TabPage8)
        Me.TabControl3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl3.Location = New System.Drawing.Point(0, 0)
        Me.TabControl3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl3.Name = "TabControl3"
        Me.TabControl3.SelectedIndex = 0
        Me.TabControl3.Size = New System.Drawing.Size(362, 257)
        Me.TabControl3.TabIndex = 100000001
        Me.TabControl3.TabStop = False
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.dgvEmployee)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage6.Size = New System.Drawing.Size(354, 231)
        Me.TabPage6.TabIndex = 0
        Me.TabPage6.Text = "Employee List"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'dgvEmployee
        '
        Me.dgvEmployee.AllowUserToAddRows = False
        Me.dgvEmployee.AllowUserToResizeColumns = False
        Me.dgvEmployee.AllowUserToResizeRows = False
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvEmployee.ColumnHeadersHeight = 25
        Me.dgvEmployee.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colEmpl_No, Me.colEmpl_Code, Me.colEmpl_Name, Me.colEmpl_NIK, Me.colEmpl_JobPosition, Me.colEmpl_Department, Me.colEmpl_TotalHours, Me.colEmpl_Remarks, Me.colEmpl_Primary, Me.colEmpl_PrimaryPos, Me.colEmpl_PrimaryDept})
        Me.dgvEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEmployee.Location = New System.Drawing.Point(2, 2)
        Me.dgvEmployee.MultiSelect = False
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(350, 227)
        Me.dgvEmployee.TabIndex = 1
        '
        'colEmpl_No
        '
        Me.colEmpl_No.Frozen = True
        Me.colEmpl_No.HeaderText = "No"
        Me.colEmpl_No.Name = "colEmpl_No"
        Me.colEmpl_No.ReadOnly = True
        Me.colEmpl_No.Width = 46
        '
        'colEmpl_Code
        '
        Me.colEmpl_Code.Frozen = True
        Me.colEmpl_Code.HeaderText = "EmployeeCode"
        Me.colEmpl_Code.MaxInputLength = 250
        Me.colEmpl_Code.Name = "colEmpl_Code"
        Me.colEmpl_Code.Width = 103
        '
        'colEmpl_Name
        '
        Me.colEmpl_Name.HeaderText = "Name"
        Me.colEmpl_Name.Name = "colEmpl_Name"
        Me.colEmpl_Name.ReadOnly = True
        Me.colEmpl_Name.Width = 60
        '
        'colEmpl_NIK
        '
        Me.colEmpl_NIK.HeaderText = "NIK"
        Me.colEmpl_NIK.Name = "colEmpl_NIK"
        Me.colEmpl_NIK.Width = 50
        '
        'colEmpl_JobPosition
        '
        Me.colEmpl_JobPosition.HeaderText = "Job Position"
        Me.colEmpl_JobPosition.Name = "colEmpl_JobPosition"
        Me.colEmpl_JobPosition.Width = 89
        '
        'colEmpl_Department
        '
        Me.colEmpl_Department.HeaderText = "Department"
        Me.colEmpl_Department.Name = "colEmpl_Department"
        Me.colEmpl_Department.Width = 87
        '
        'colEmpl_TotalHours
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colEmpl_TotalHours.DefaultCellStyle = DataGridViewCellStyle1
        Me.colEmpl_TotalHours.HeaderText = "Total Hours"
        Me.colEmpl_TotalHours.Name = "colEmpl_TotalHours"
        Me.colEmpl_TotalHours.ReadOnly = True
        Me.colEmpl_TotalHours.Width = 87
        '
        'colEmpl_Remarks
        '
        Me.colEmpl_Remarks.HeaderText = "Remarks"
        Me.colEmpl_Remarks.Name = "colEmpl_Remarks"
        Me.colEmpl_Remarks.Width = 74
        '
        'colEmpl_Primary
        '
        Me.colEmpl_Primary.HeaderText = "Primary"
        Me.colEmpl_Primary.Name = "colEmpl_Primary"
        Me.colEmpl_Primary.ReadOnly = True
        Me.colEmpl_Primary.Visible = False
        Me.colEmpl_Primary.Width = 66
        '
        'colEmpl_PrimaryPos
        '
        Me.colEmpl_PrimaryPos.HeaderText = "Primary Position"
        Me.colEmpl_PrimaryPos.Name = "colEmpl_PrimaryPos"
        Me.colEmpl_PrimaryPos.ReadOnly = True
        Me.colEmpl_PrimaryPos.Width = 106
        '
        'colEmpl_PrimaryDept
        '
        Me.colEmpl_PrimaryDept.HeaderText = "Primary Department"
        Me.colEmpl_PrimaryDept.Name = "colEmpl_PrimaryDept"
        Me.colEmpl_PrimaryDept.ReadOnly = True
        Me.colEmpl_PrimaryDept.Width = 124
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.dgvSchDetail)
        Me.TabPage7.Location = New System.Drawing.Point(4, 22)
        Me.TabPage7.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage7.Size = New System.Drawing.Size(354, 231)
        Me.TabPage7.TabIndex = 1
        Me.TabPage7.Text = "Overtime Schedule"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'dgvSchDetail
        '
        Me.dgvSchDetail.AllowUserToAddRows = False
        Me.dgvSchDetail.AllowUserToDeleteRows = False
        Me.dgvSchDetail.AllowUserToResizeColumns = False
        Me.dgvSchDetail.AllowUserToResizeRows = False
        Me.dgvSchDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSchDetail.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSchDetail.ColumnHeadersHeight = 25
        Me.dgvSchDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvSchDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSch_No, Me.ColSch_DayOfWeek, Me.ColSch_Date, Me.ColSch_Shift, Me.ColSch_NormalWH, Me.ColSch_OTtype, Me.ColSch_Start, Me.ColSch_Break, Me.ColSch_End, Me.ColSch_Duration, Me.ColOverride, Me.ColSch_Remark, Me.ColSch_PrimaryShift})
        Me.dgvSchDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSchDetail.Location = New System.Drawing.Point(2, 2)
        Me.dgvSchDetail.MultiSelect = False
        Me.dgvSchDetail.Name = "dgvSchDetail"
        Me.dgvSchDetail.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvSchDetail.RowTemplate.Height = 24
        Me.dgvSchDetail.Size = New System.Drawing.Size(350, 227)
        Me.dgvSchDetail.TabIndex = 36
        '
        'ColSch_No
        '
        Me.ColSch_No.HeaderText = "No"
        Me.ColSch_No.Name = "ColSch_No"
        Me.ColSch_No.ReadOnly = True
        Me.ColSch_No.Width = 46
        '
        'ColSch_DayOfWeek
        '
        Me.ColSch_DayOfWeek.HeaderText = "Day Of Week"
        Me.ColSch_DayOfWeek.Name = "ColSch_DayOfWeek"
        Me.ColSch_DayOfWeek.Width = 97
        '
        'ColSch_Date
        '
        Me.ColSch_Date.HeaderText = "Date"
        Me.ColSch_Date.Name = "ColSch_Date"
        Me.ColSch_Date.Width = 55
        '
        'ColSch_Shift
        '
        Me.ColSch_Shift.HeaderText = "Shift"
        Me.ColSch_Shift.Name = "ColSch_Shift"
        Me.ColSch_Shift.Width = 53
        '
        'ColSch_NormalWH
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSch_NormalWH.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColSch_NormalWH.HeaderText = "Normal W/H"
        Me.ColSch_NormalWH.Name = "ColSch_NormalWH"
        Me.ColSch_NormalWH.Width = 92
        '
        'ColSch_OTtype
        '
        Me.ColSch_OTtype.HeaderText = "OT Type"
        Me.ColSch_OTtype.Name = "ColSch_OTtype"
        Me.ColSch_OTtype.Width = 74
        '
        'ColSch_Start
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSch_Start.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColSch_Start.HeaderText = "Start"
        Me.ColSch_Start.Name = "ColSch_Start"
        Me.ColSch_Start.Width = 54
        '
        'ColSch_Break
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSch_Break.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColSch_Break.HeaderText = "Break"
        Me.ColSch_Break.Name = "ColSch_Break"
        Me.ColSch_Break.Width = 60
        '
        'ColSch_End
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSch_End.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColSch_End.HeaderText = "End"
        Me.ColSch_End.Name = "ColSch_End"
        Me.ColSch_End.Width = 51
        '
        'ColSch_Duration
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSch_Duration.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColSch_Duration.HeaderText = "Duration"
        Me.ColSch_Duration.Name = "ColSch_Duration"
        Me.ColSch_Duration.Width = 72
        '
        'ColOverride
        '
        Me.ColOverride.HeaderText = "Override"
        Me.ColOverride.Name = "ColOverride"
        Me.ColOverride.Width = 53
        '
        'ColSch_Remark
        '
        Me.ColSch_Remark.HeaderText = "Remarks"
        Me.ColSch_Remark.Name = "ColSch_Remark"
        Me.ColSch_Remark.Width = 74
        '
        'ColSch_PrimaryShift
        '
        Me.ColSch_PrimaryShift.HeaderText = "Primary Shift"
        Me.ColSch_PrimaryShift.Name = "ColSch_PrimaryShift"
        Me.ColSch_PrimaryShift.ReadOnly = True
        Me.ColSch_PrimaryShift.Visible = False
        Me.ColSch_PrimaryShift.Width = 90
        '
        'TabPage8
        '
        Me.TabPage8.Controls.Add(Me.dgvSchEmp)
        Me.TabPage8.Location = New System.Drawing.Point(4, 22)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage8.Size = New System.Drawing.Size(354, 231)
        Me.TabPage8.TabIndex = 2
        Me.TabPage8.Text = "Employee Schedule"
        Me.TabPage8.UseVisualStyleBackColor = True
        '
        'dgvSchEmp
        '
        Me.dgvSchEmp.AllowUserToAddRows = False
        Me.dgvSchEmp.AllowUserToResizeColumns = False
        Me.dgvSchEmp.AllowUserToResizeRows = False
        Me.dgvSchEmp.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSchEmp.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSchEmp.ColumnHeadersHeight = 25
        Me.dgvSchEmp.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvSchEmp.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSchEmp_No, Me.ColSchEmp_Code, Me.ColSchEmp_Emp, Me.ColSchEmp_DayOfWeek, Me.ColSchEmp_Date, Me.ColSchEmp_Shift, Me.ColSchEmp_NormalWH, Me.ColSchEmp_StartEarly, Me.ColSchEmp_BreakEarly, Me.ColSchEmp_EndEarly, Me.ColSchEmp_EarlyOTDuration, Me.ColSchEmp_OverrideEarlyOT, Me.ColSchEmp_StartOT, Me.ColSchEmp_BreakOT, Me.ColSchEmp_EndOT, Me.ColSchEmp_OTDuration, Me.ColSchEmp_OverrideOT, Me.ColSchEmp_PrimaryShift, Me.ColSchEmp_PrimaryEmp})
        Me.dgvSchEmp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSchEmp.Location = New System.Drawing.Point(3, 3)
        Me.dgvSchEmp.Name = "dgvSchEmp"
        Me.dgvSchEmp.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvSchEmp.RowTemplate.Height = 24
        Me.dgvSchEmp.Size = New System.Drawing.Size(348, 225)
        Me.dgvSchEmp.TabIndex = 37
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(661, 3)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(54, 25)
        Me.btnClose.TabIndex = 51
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(9, 63)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(726, 179)
        Me.TabControl1.TabIndex = 11
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvItem)
        Me.TabPage1.Controls.Add(Me.btnClose)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Size = New System.Drawing.Size(718, 153)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "List"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.TabControl2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Size = New System.Drawing.Size(718, 153)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Filter"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.TabPage3)
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl2.Location = New System.Drawing.Point(2, 2)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(714, 149)
        Me.TabControl2.TabIndex = 0
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.DgvEmpFilter)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Size = New System.Drawing.Size(706, 123)
        Me.TabPage3.TabIndex = 0
        Me.TabPage3.Text = "Employee Filter"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'DgvEmpFilter
        '
        Me.DgvEmpFilter.AllowUserToResizeColumns = False
        Me.DgvEmpFilter.AllowUserToResizeRows = False
        Me.DgvEmpFilter.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvEmpFilter.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvEmpFilter.ColumnHeadersHeight = 25
        Me.DgvEmpFilter.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvEmpFilter.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFilterEmp_No, Me.ColFilterEmp_Code, Me.ColFilterEmp_Name, Me.ColFilterEmp_NIK, Me.ColFilterEmp_Dept, Me.ColFilterEmp_JobPosition, Me.ColFilterEmp_Primary})
        Me.DgvEmpFilter.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvEmpFilter.Location = New System.Drawing.Point(2, 2)
        Me.DgvEmpFilter.MultiSelect = False
        Me.DgvEmpFilter.Name = "DgvEmpFilter"
        Me.DgvEmpFilter.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvEmpFilter.RowTemplate.Height = 24
        Me.DgvEmpFilter.Size = New System.Drawing.Size(702, 119)
        Me.DgvEmpFilter.TabIndex = 49
        '
        'ColFilterEmp_No
        '
        Me.ColFilterEmp_No.HeaderText = "No"
        Me.ColFilterEmp_No.Name = "ColFilterEmp_No"
        Me.ColFilterEmp_No.Width = 46
        '
        'ColFilterEmp_Code
        '
        Me.ColFilterEmp_Code.HeaderText = "Code"
        Me.ColFilterEmp_Code.Name = "ColFilterEmp_Code"
        Me.ColFilterEmp_Code.Width = 57
        '
        'ColFilterEmp_Name
        '
        Me.ColFilterEmp_Name.HeaderText = "Name"
        Me.ColFilterEmp_Name.Name = "ColFilterEmp_Name"
        Me.ColFilterEmp_Name.Width = 60
        '
        'ColFilterEmp_NIK
        '
        Me.ColFilterEmp_NIK.HeaderText = "NIK"
        Me.ColFilterEmp_NIK.Name = "ColFilterEmp_NIK"
        Me.ColFilterEmp_NIK.Width = 50
        '
        'ColFilterEmp_Dept
        '
        Me.ColFilterEmp_Dept.HeaderText = "Department"
        Me.ColFilterEmp_Dept.Name = "ColFilterEmp_Dept"
        Me.ColFilterEmp_Dept.Width = 87
        '
        'ColFilterEmp_JobPosition
        '
        Me.ColFilterEmp_JobPosition.HeaderText = "Job Position"
        Me.ColFilterEmp_JobPosition.Name = "ColFilterEmp_JobPosition"
        Me.ColFilterEmp_JobPosition.Width = 89
        '
        'ColFilterEmp_Primary
        '
        Me.ColFilterEmp_Primary.HeaderText = "Primary"
        Me.ColFilterEmp_Primary.Name = "ColFilterEmp_Primary"
        Me.ColFilterEmp_Primary.Visible = False
        Me.ColFilterEmp_Primary.Width = 66
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.DgvDeptFilter)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Size = New System.Drawing.Size(706, 123)
        Me.TabPage4.TabIndex = 1
        Me.TabPage4.Text = "Department Filter"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'DgvDeptFilter
        '
        Me.DgvDeptFilter.AllowUserToResizeColumns = False
        Me.DgvDeptFilter.AllowUserToResizeRows = False
        Me.DgvDeptFilter.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvDeptFilter.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvDeptFilter.ColumnHeadersHeight = 25
        Me.DgvDeptFilter.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvDeptFilter.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFilterDept_No, Me.ColFilterDept_Primary, Me.ColFilterDept_Code, Me.ColFilterDept_Desc})
        Me.DgvDeptFilter.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvDeptFilter.Location = New System.Drawing.Point(2, 2)
        Me.DgvDeptFilter.MultiSelect = False
        Me.DgvDeptFilter.Name = "DgvDeptFilter"
        Me.DgvDeptFilter.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvDeptFilter.RowTemplate.Height = 24
        Me.DgvDeptFilter.Size = New System.Drawing.Size(702, 119)
        Me.DgvDeptFilter.TabIndex = 50
        '
        'ColFilterDept_No
        '
        Me.ColFilterDept_No.HeaderText = "No"
        Me.ColFilterDept_No.Name = "ColFilterDept_No"
        Me.ColFilterDept_No.Width = 46
        '
        'ColFilterDept_Primary
        '
        Me.ColFilterDept_Primary.HeaderText = "Primary"
        Me.ColFilterDept_Primary.Name = "ColFilterDept_Primary"
        Me.ColFilterDept_Primary.Visible = False
        Me.ColFilterDept_Primary.Width = 66
        '
        'ColFilterDept_Code
        '
        Me.ColFilterDept_Code.HeaderText = "Code"
        Me.ColFilterDept_Code.Name = "ColFilterDept_Code"
        Me.ColFilterDept_Code.Width = 57
        '
        'ColFilterDept_Desc
        '
        Me.ColFilterDept_Desc.HeaderText = "Description"
        Me.ColFilterDept_Desc.Name = "ColFilterDept_Desc"
        Me.ColFilterDept_Desc.Width = 85
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.DgvPosFilter)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Size = New System.Drawing.Size(706, 123)
        Me.TabPage5.TabIndex = 2
        Me.TabPage5.Text = "Job Position Filter"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'DgvPosFilter
        '
        Me.DgvPosFilter.AllowUserToResizeColumns = False
        Me.DgvPosFilter.AllowUserToResizeRows = False
        Me.DgvPosFilter.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvPosFilter.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvPosFilter.ColumnHeadersHeight = 25
        Me.DgvPosFilter.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvPosFilter.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFilterJobPos_No, Me.ColFilterJobPos_Primary, Me.ColFilterJobPos_Code, Me.ColFilterJobPos_Desc})
        Me.DgvPosFilter.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvPosFilter.Location = New System.Drawing.Point(2, 2)
        Me.DgvPosFilter.MultiSelect = False
        Me.DgvPosFilter.Name = "DgvPosFilter"
        Me.DgvPosFilter.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvPosFilter.RowTemplate.Height = 24
        Me.DgvPosFilter.Size = New System.Drawing.Size(702, 119)
        Me.DgvPosFilter.TabIndex = 51
        '
        'ColFilterJobPos_No
        '
        Me.ColFilterJobPos_No.HeaderText = "No"
        Me.ColFilterJobPos_No.Name = "ColFilterJobPos_No"
        Me.ColFilterJobPos_No.Width = 46
        '
        'ColFilterJobPos_Primary
        '
        Me.ColFilterJobPos_Primary.HeaderText = "Primary"
        Me.ColFilterJobPos_Primary.Name = "ColFilterJobPos_Primary"
        Me.ColFilterJobPos_Primary.Visible = False
        Me.ColFilterJobPos_Primary.Width = 66
        '
        'ColFilterJobPos_Code
        '
        Me.ColFilterJobPos_Code.HeaderText = "Code"
        Me.ColFilterJobPos_Code.Name = "ColFilterJobPos_Code"
        Me.ColFilterJobPos_Code.Width = 57
        '
        'ColFilterJobPos_Desc
        '
        Me.ColFilterJobPos_Desc.HeaderText = "Description"
        Me.ColFilterJobPos_Desc.Name = "ColFilterJobPos_Desc"
        Me.ColFilterJobPos_Desc.Width = 85
        '
        'llbClearFilter
        '
        Me.llbClearFilter.AutoSize = True
        Me.llbClearFilter.Location = New System.Drawing.Point(429, 41)
        Me.llbClearFilter.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbClearFilter.Name = "llbClearFilter"
        Me.llbClearFilter.Size = New System.Drawing.Size(56, 13)
        Me.llbClearFilter.TabIndex = 10
        Me.llbClearFilter.TabStop = True
        Me.llbClearFilter.Text = "Clear Filter"
        '
        'DtpDocDateTo
        '
        Me.DtpDocDateTo.CustomFormat = "dd MMM yyyy"
        Me.DtpDocDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDocDateTo.Location = New System.Drawing.Point(239, 39)
        Me.DtpDocDateTo.Name = "DtpDocDateTo"
        Me.DtpDocDateTo.Size = New System.Drawing.Size(102, 20)
        Me.DtpDocDateTo.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(214, 41)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(20, 13)
        Me.Label9.TabIndex = 57
        Me.Label9.Text = "To"
        '
        'DtpDocDateFrom
        '
        Me.DtpDocDateFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpDocDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDocDateFrom.Location = New System.Drawing.Point(108, 39)
        Me.DtpDocDateFrom.Name = "DtpDocDateFrom"
        Me.DtpDocDateFrom.Size = New System.Drawing.Size(102, 20)
        Me.DtpDocDateFrom.TabIndex = 7
        '
        'ChkDocDate
        '
        Me.ChkDocDate.AutoSize = True
        Me.ChkDocDate.Location = New System.Drawing.Point(9, 40)
        Me.ChkDocDate.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkDocDate.Name = "ChkDocDate"
        Me.ChkDocDate.Size = New System.Drawing.Size(101, 17)
        Me.ChkDocDate.TabIndex = 6
        Me.ChkDocDate.Text = "Document Date"
        Me.ChkDocDate.UseVisualStyleBackColor = True
        '
        'BtnFilter
        '
        Me.BtnFilter.Location = New System.Drawing.Point(349, 37)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(75, 23)
        Me.BtnFilter.TabIndex = 9
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'BgwDepartment
        '
        '
        'BgwEmployee
        '
        '
        'BgwJobPosition
        '
        '
        'ColSchEmp_No
        '
        Me.ColSchEmp_No.HeaderText = "No"
        Me.ColSchEmp_No.Name = "ColSchEmp_No"
        Me.ColSchEmp_No.ReadOnly = True
        Me.ColSchEmp_No.Width = 46
        '
        'ColSchEmp_Code
        '
        Me.ColSchEmp_Code.HeaderText = "Employee Code"
        Me.ColSchEmp_Code.Name = "ColSchEmp_Code"
        Me.ColSchEmp_Code.ReadOnly = True
        Me.ColSchEmp_Code.Width = 106
        '
        'ColSchEmp_Emp
        '
        Me.ColSchEmp_Emp.HeaderText = "Name"
        Me.ColSchEmp_Emp.Name = "ColSchEmp_Emp"
        Me.ColSchEmp_Emp.ReadOnly = True
        Me.ColSchEmp_Emp.Width = 60
        '
        'ColSchEmp_DayOfWeek
        '
        Me.ColSchEmp_DayOfWeek.HeaderText = "Day Of Week"
        Me.ColSchEmp_DayOfWeek.Name = "ColSchEmp_DayOfWeek"
        Me.ColSchEmp_DayOfWeek.Width = 97
        '
        'ColSchEmp_Date
        '
        Me.ColSchEmp_Date.HeaderText = "Date"
        Me.ColSchEmp_Date.Name = "ColSchEmp_Date"
        Me.ColSchEmp_Date.Width = 55
        '
        'ColSchEmp_Shift
        '
        Me.ColSchEmp_Shift.HeaderText = "Shift"
        Me.ColSchEmp_Shift.Name = "ColSchEmp_Shift"
        Me.ColSchEmp_Shift.Width = 53
        '
        'ColSchEmp_NormalWH
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchEmp_NormalWH.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColSchEmp_NormalWH.HeaderText = "Normal W/H"
        Me.ColSchEmp_NormalWH.Name = "ColSchEmp_NormalWH"
        Me.ColSchEmp_NormalWH.Visible = False
        Me.ColSchEmp_NormalWH.Width = 92
        '
        'ColSchEmp_StartEarly
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchEmp_StartEarly.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColSchEmp_StartEarly.HeaderText = "Start Early OT"
        Me.ColSchEmp_StartEarly.Name = "ColSchEmp_StartEarly"
        Me.ColSchEmp_StartEarly.Width = 98
        '
        'ColSchEmp_BreakEarly
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchEmp_BreakEarly.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColSchEmp_BreakEarly.HeaderText = "Break Early OT"
        Me.ColSchEmp_BreakEarly.Name = "ColSchEmp_BreakEarly"
        Me.ColSchEmp_BreakEarly.Width = 104
        '
        'ColSchEmp_EndEarly
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchEmp_EndEarly.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColSchEmp_EndEarly.HeaderText = "End Early OT"
        Me.ColSchEmp_EndEarly.Name = "ColSchEmp_EndEarly"
        Me.ColSchEmp_EndEarly.Width = 95
        '
        'ColSchEmp_EarlyOTDuration
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchEmp_EarlyOTDuration.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColSchEmp_EarlyOTDuration.HeaderText = "Early OT Duration"
        Me.ColSchEmp_EarlyOTDuration.Name = "ColSchEmp_EarlyOTDuration"
        Me.ColSchEmp_EarlyOTDuration.Width = 116
        '
        'ColSchEmp_OverrideEarlyOT
        '
        Me.ColSchEmp_OverrideEarlyOT.HeaderText = "Override Early OT"
        Me.ColSchEmp_OverrideEarlyOT.Name = "ColSchEmp_OverrideEarlyOT"
        Me.ColSchEmp_OverrideEarlyOT.Width = 97
        '
        'ColSchEmp_StartOT
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchEmp_StartOT.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColSchEmp_StartOT.HeaderText = "Start OT"
        Me.ColSchEmp_StartOT.Name = "ColSchEmp_StartOT"
        Me.ColSchEmp_StartOT.ReadOnly = True
        Me.ColSchEmp_StartOT.Width = 72
        '
        'ColSchEmp_BreakOT
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchEmp_BreakOT.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColSchEmp_BreakOT.HeaderText = "Break OT"
        Me.ColSchEmp_BreakOT.Name = "ColSchEmp_BreakOT"
        Me.ColSchEmp_BreakOT.ReadOnly = True
        Me.ColSchEmp_BreakOT.Width = 78
        '
        'ColSchEmp_EndOT
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchEmp_EndOT.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColSchEmp_EndOT.HeaderText = "End OT"
        Me.ColSchEmp_EndOT.Name = "ColSchEmp_EndOT"
        Me.ColSchEmp_EndOT.ReadOnly = True
        Me.ColSchEmp_EndOT.Width = 69
        '
        'ColSchEmp_OTDuration
        '
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchEmp_OTDuration.DefaultCellStyle = DataGridViewCellStyle15
        Me.ColSchEmp_OTDuration.HeaderText = "OT Duration"
        Me.ColSchEmp_OTDuration.Name = "ColSchEmp_OTDuration"
        Me.ColSchEmp_OTDuration.Width = 90
        '
        'ColSchEmp_OverrideOT
        '
        Me.ColSchEmp_OverrideOT.HeaderText = "Override OT"
        Me.ColSchEmp_OverrideOT.Name = "ColSchEmp_OverrideOT"
        Me.ColSchEmp_OverrideOT.Visible = False
        Me.ColSchEmp_OverrideOT.Width = 71
        '
        'ColSchEmp_PrimaryShift
        '
        Me.ColSchEmp_PrimaryShift.HeaderText = "Primary Shift"
        Me.ColSchEmp_PrimaryShift.Name = "ColSchEmp_PrimaryShift"
        Me.ColSchEmp_PrimaryShift.ReadOnly = True
        Me.ColSchEmp_PrimaryShift.Visible = False
        Me.ColSchEmp_PrimaryShift.Width = 90
        '
        'ColSchEmp_PrimaryEmp
        '
        Me.ColSchEmp_PrimaryEmp.HeaderText = "Primary Employee"
        Me.ColSchEmp_PrimaryEmp.Name = "ColSchEmp_PrimaryEmp"
        Me.ColSchEmp_PrimaryEmp.ReadOnly = True
        Me.ColSchEmp_PrimaryEmp.Visible = False
        Me.ColSchEmp_PrimaryEmp.Width = 115
        '
        'FrmTrans_SPLApproval
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(744, 514)
        Me.Controls.Add(Me.BtnFilter)
        Me.Controls.Add(Me.ChkDocDate)
        Me.Controls.Add(Me.DtpDocDateTo)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.DtpDocDateFrom)
        Me.Controls.Add(Me.llbClearFilter)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.BtnWait)
        Me.Controls.Add(Me.RboRejected)
        Me.Controls.Add(Me.RboApproved)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.BtnReject)
        Me.Controls.Add(Me.RboWaitingForApproval)
        Me.Controls.Add(Me.BtnApprove)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmTrans_SPLApproval"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Overtime Order Approval"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabControl3.ResumeLayout(False)
        Me.TabPage6.ResumeLayout(False)
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage7.ResumeLayout(False)
        CType(Me.dgvSchDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage8.ResumeLayout(False)
        CType(Me.dgvSchEmp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        CType(Me.DgvEmpFilter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.DgvDeptFilter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        CType(Me.DgvPosFilter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BtnWait As System.Windows.Forms.Button
    Friend WithEvents RboRejected As System.Windows.Forms.RadioButton
    Friend WithEvents RboApproved As System.Windows.Forms.RadioButton
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents BtnReject As System.Windows.Forms.Button
    Friend WithEvents RboWaitingForApproval As System.Windows.Forms.RadioButton
    Friend WithEvents BtnApprove As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents DgvEmpFilter As System.Windows.Forms.DataGridView
    Friend WithEvents ColFilterEmp_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterEmp_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterEmp_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterEmp_NIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterEmp_Dept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterEmp_JobPosition As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterEmp_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvDeptFilter As System.Windows.Forms.DataGridView
    Friend WithEvents ColFilterDept_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterDept_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterDept_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterDept_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvPosFilter As System.Windows.Forms.DataGridView
    Friend WithEvents ColFilterJobPos_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterJobPos_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterJobPos_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilterJobPos_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtReason As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtDocumentNo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents TabControl3 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents llbClearFilter As System.Windows.Forms.LinkLabel
    Friend WithEvents DtpDocDateTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents DtpDocDateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents ChkDocDate As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtTotalDuration As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalDurationPerEmp As System.Windows.Forms.TextBox
    Friend WithEvents ColItem_EnumStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_StatusNotesDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvSchDetail As System.Windows.Forms.DataGridView
    Friend WithEvents ColSch_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_DayOfWeek As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_Shift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_NormalWH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_OTtype As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_Start As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_Break As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_End As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_Duration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOverride As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColSch_Remark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSch_PrimaryShift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents BgwEmployee As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwJobPosition As System.ComponentModel.BackgroundWorker
    Private WithEvents BgwDepartment As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColItem_DocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_StartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_EndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Status As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColItem_StatusNotes As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColItem_Modifiy As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColItem_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Budget As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColItem_Print As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents colEmpl_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpl_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpl_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpl_NIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpl_JobPosition As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpl_Department As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpl_TotalHours As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpl_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpl_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpl_PrimaryPos As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpl_PrimaryDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
    Friend WithEvents dgvSchEmp As System.Windows.Forms.DataGridView
    Friend WithEvents ColSchEmp_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_Emp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_DayOfWeek As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_Shift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_NormalWH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_StartEarly As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_BreakEarly As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_EndEarly As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_EarlyOTDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_OverrideEarlyOT As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColSchEmp_StartOT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_BreakOT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_EndOT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_OTDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_OverrideOT As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColSchEmp_PrimaryShift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchEmp_PrimaryEmp As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
