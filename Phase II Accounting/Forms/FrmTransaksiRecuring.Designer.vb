<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiRecuring
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
        Me.btnViewJurnal = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDebet = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKredit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExchangeRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOriginalCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDebet_OriginalCurr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCredit_OriginalCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJob = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryJob = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAcc_IncrementStyle = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColAcc_IncrementValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvEntries = New System.Windows.Forms.DataGridView
        Me.ColTrans_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTrans_JournalNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTrans_JournalDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTrans_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTrans_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTrans_PrimaryJournal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTrans_Status = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTrans_FollowUp = New System.Windows.Forms.DataGridViewButtonColumn
        Me.cms_follow_up = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tms_Create = New System.Windows.Forms.ToolStripMenuItem
        Me.tms_Update = New System.Windows.Forms.ToolStripMenuItem
        Me.tms_Edit = New System.Windows.Forms.ToolStripMenuItem
        Me.tms_Delete = New System.Windows.Forms.ToolStripMenuItem
        Me.tms_Skip = New System.Windows.Forms.ToolStripMenuItem
        Me.tms_Lock = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.tms_UpdateAll = New System.Windows.Forms.ToolStripMenuItem
        Me.tms_DeleteAll = New System.Windows.Forms.ToolStripMenuItem
        Me.ColTrans_ScheduleDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTrans_NoUrut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvReminder = New System.Windows.Forms.DataGridView
        Me.ColReminder_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReminder_UserCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReminder_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReminder_Interval = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReminder_PrimaryUser = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.txtUnbalanceIncrement = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.chkIgnoreIncrementOnSkip = New System.Windows.Forms.CheckBox
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtPrimaryJournal = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtRefNo = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtUnbalance = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtLimit = New System.Windows.Forms.TextBox
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.cboEveryMonth = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.cboEveryWeek = New System.Windows.Forms.ComboBox
        Me.chkLimit = New System.Windows.Forms.CheckBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboFrequency = New System.Windows.Forms.ComboBox
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.gboDetail.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvEntries, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cms_follow_up.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvReminder, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnViewJurnal
        '
        Me.btnViewJurnal.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnViewJurnal.Location = New System.Drawing.Point(12, 544)
        Me.btnViewJurnal.Name = "btnViewJurnal"
        Me.btnViewJurnal.Size = New System.Drawing.Size(80, 26)
        Me.btnViewJurnal.TabIndex = 41
        Me.btnViewJurnal.TabStop = False
        Me.btnViewJurnal.Text = "View Journal"
        Me.btnViewJurnal.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(0, 3)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(892, 535)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(6, 96)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(880, 402)
        Me.TabControl1.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvItem)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(872, 376)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Account Entries"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColDescAccount, Me.ColDebet, Me.ColKredit, Me.ColExchangeRate, Me.ColOriginalCurrency, Me.ColDebet_OriginalCurr, Me.ColCredit_OriginalCurrency, Me.ColRemarks, Me.ColDept, Me.ColJob, Me.ColPrimaryDept, Me.ColPrimaryJob, Me.ColPrimaryAkun, Me.ColPrimaryCurrency, Me.ColAcc_IncrementStyle, Me.ColAcc_IncrementValue})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 3)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(866, 370)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 0
        '
        'ColAccount
        '
        Me.ColAccount.HeaderText = "Account"
        Me.ColAccount.MaxInputLength = 50
        Me.ColAccount.Name = "ColAccount"
        Me.ColAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccount.Width = 53
        '
        'ColDescAccount
        '
        Me.ColDescAccount.HeaderText = "Description"
        Me.ColDescAccount.MaxInputLength = 150
        Me.ColDescAccount.Name = "ColDescAccount"
        Me.ColDescAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDescAccount.Width = 66
        '
        'ColDebet
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.NullValue = "0"
        Me.ColDebet.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColDebet.HeaderText = "Debet"
        Me.ColDebet.MaxInputLength = 18
        Me.ColDebet.Name = "ColDebet"
        Me.ColDebet.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet.Width = 42
        '
        'ColKredit
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.NullValue = "0"
        Me.ColKredit.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColKredit.HeaderText = "Credit"
        Me.ColKredit.MaxInputLength = 18
        Me.ColKredit.Name = "ColKredit"
        Me.ColKredit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColKredit.Width = 40
        '
        'ColExchangeRate
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.NullValue = "1"
        Me.ColExchangeRate.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColExchangeRate.HeaderText = "Exchange Rate"
        Me.ColExchangeRate.Name = "ColExchangeRate"
        Me.ColExchangeRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExchangeRate.Visible = False
        Me.ColExchangeRate.Width = 87
        '
        'ColOriginalCurrency
        '
        Me.ColOriginalCurrency.HeaderText = "Currency"
        Me.ColOriginalCurrency.Name = "ColOriginalCurrency"
        Me.ColOriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOriginalCurrency.Visible = False
        Me.ColOriginalCurrency.Width = 55
        '
        'ColDebet_OriginalCurr
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.NullValue = "0"
        Me.ColDebet_OriginalCurr.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColDebet_OriginalCurr.HeaderText = "Debet"
        Me.ColDebet_OriginalCurr.Name = "ColDebet_OriginalCurr"
        Me.ColDebet_OriginalCurr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet_OriginalCurr.Visible = False
        Me.ColDebet_OriginalCurr.Width = 42
        '
        'ColCredit_OriginalCurrency
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.NullValue = "0"
        Me.ColCredit_OriginalCurrency.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColCredit_OriginalCurrency.HeaderText = "Credit"
        Me.ColCredit_OriginalCurrency.Name = "ColCredit_OriginalCurrency"
        Me.ColCredit_OriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCredit_OriginalCurrency.Visible = False
        Me.ColCredit_OriginalCurrency.Width = 40
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 150
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 55
        '
        'ColDept
        '
        Me.ColDept.HeaderText = "Departement"
        Me.ColDept.Name = "ColDept"
        Me.ColDept.Visible = False
        Me.ColDept.Width = 93
        '
        'ColJob
        '
        Me.ColJob.HeaderText = "Job"
        Me.ColJob.Name = "ColJob"
        Me.ColJob.Visible = False
        Me.ColJob.Width = 49
        '
        'ColPrimaryDept
        '
        Me.ColPrimaryDept.HeaderText = "Primary Dept"
        Me.ColPrimaryDept.Name = "ColPrimaryDept"
        Me.ColPrimaryDept.Visible = False
        Me.ColPrimaryDept.Width = 92
        '
        'ColPrimaryJob
        '
        Me.ColPrimaryJob.HeaderText = "Primary Job"
        Me.ColPrimaryJob.Name = "ColPrimaryJob"
        Me.ColPrimaryJob.Visible = False
        Me.ColPrimaryJob.Width = 86
        '
        'ColPrimaryAkun
        '
        Me.ColPrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColPrimaryAkun.Name = "ColPrimaryAkun"
        Me.ColPrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryAkun.Visible = False
        Me.ColPrimaryAkun.Width = 72
        '
        'ColPrimaryCurrency
        '
        Me.ColPrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.Visible = False
        Me.ColPrimaryCurrency.Width = 111
        '
        'ColAcc_IncrementStyle
        '
        Me.ColAcc_IncrementStyle.HeaderText = "Increment Style"
        Me.ColAcc_IncrementStyle.Items.AddRange(New Object() {"No Increment", "Fixed Increment", "Percentage Increment"})
        Me.ColAcc_IncrementStyle.Name = "ColAcc_IncrementStyle"
        Me.ColAcc_IncrementStyle.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAcc_IncrementStyle.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColAcc_IncrementStyle.Width = 105
        '
        'ColAcc_IncrementValue
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle7.NullValue = "0"
        Me.ColAcc_IncrementValue.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColAcc_IncrementValue.HeaderText = "Increment Value"
        Me.ColAcc_IncrementValue.Name = "ColAcc_IncrementValue"
        Me.ColAcc_IncrementValue.Width = 109
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvEntries)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(872, 376)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Transaction Entries"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvEntries
        '
        Me.dgvEntries.AllowUserToAddRows = False
        Me.dgvEntries.AllowUserToDeleteRows = False
        Me.dgvEntries.AllowUserToResizeColumns = False
        Me.dgvEntries.AllowUserToResizeRows = False
        DataGridViewCellStyle8.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvEntries.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle8
        Me.dgvEntries.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEntries.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvEntries.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColTrans_No, Me.ColTrans_JournalNumber, Me.ColTrans_JournalDate, Me.ColTrans_Description, Me.ColTrans_Remarks, Me.ColTrans_PrimaryJournal, Me.ColTrans_Status, Me.ColTrans_FollowUp, Me.ColTrans_ScheduleDate, Me.ColTrans_NoUrut})
        Me.dgvEntries.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEntries.Location = New System.Drawing.Point(3, 3)
        Me.dgvEntries.MultiSelect = False
        Me.dgvEntries.Name = "dgvEntries"
        Me.dgvEntries.RowTemplate.Height = 24
        Me.dgvEntries.Size = New System.Drawing.Size(866, 370)
        Me.dgvEntries.StandardTab = True
        Me.dgvEntries.TabIndex = 4
        '
        'ColTrans_No
        '
        Me.ColTrans_No.HeaderText = "No"
        Me.ColTrans_No.Name = "ColTrans_No"
        Me.ColTrans_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTrans_No.Width = 27
        '
        'ColTrans_JournalNumber
        '
        Me.ColTrans_JournalNumber.HeaderText = "Journal No"
        Me.ColTrans_JournalNumber.Name = "ColTrans_JournalNumber"
        Me.ColTrans_JournalNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTrans_JournalNumber.Width = 64
        '
        'ColTrans_JournalDate
        '
        Me.ColTrans_JournalDate.HeaderText = "Date"
        Me.ColTrans_JournalDate.Name = "ColTrans_JournalDate"
        Me.ColTrans_JournalDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTrans_JournalDate.Width = 36
        '
        'ColTrans_Description
        '
        Me.ColTrans_Description.HeaderText = "Description"
        Me.ColTrans_Description.Name = "ColTrans_Description"
        Me.ColTrans_Description.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTrans_Description.Width = 66
        '
        'ColTrans_Remarks
        '
        Me.ColTrans_Remarks.HeaderText = "Remarks"
        Me.ColTrans_Remarks.Name = "ColTrans_Remarks"
        Me.ColTrans_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTrans_Remarks.Width = 55
        '
        'ColTrans_PrimaryJournal
        '
        Me.ColTrans_PrimaryJournal.HeaderText = "Primary Journal"
        Me.ColTrans_PrimaryJournal.Name = "ColTrans_PrimaryJournal"
        Me.ColTrans_PrimaryJournal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTrans_PrimaryJournal.Visible = False
        Me.ColTrans_PrimaryJournal.Width = 84
        '
        'ColTrans_Status
        '
        Me.ColTrans_Status.HeaderText = "Status"
        Me.ColTrans_Status.Name = "ColTrans_Status"
        Me.ColTrans_Status.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTrans_Status.Width = 43
        '
        'ColTrans_FollowUp
        '
        Me.ColTrans_FollowUp.ContextMenuStrip = Me.cms_follow_up
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle9.NullValue = "Right Click"
        Me.ColTrans_FollowUp.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColTrans_FollowUp.HeaderText = "Follow Up"
        Me.ColTrans_FollowUp.Name = "ColTrans_FollowUp"
        Me.ColTrans_FollowUp.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColTrans_FollowUp.Text = "Right Click"
        Me.ColTrans_FollowUp.Width = 60
        '
        'cms_follow_up
        '
        Me.cms_follow_up.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tms_Create, Me.tms_Update, Me.tms_Edit, Me.tms_Delete, Me.tms_Skip, Me.tms_Lock, Me.ToolStripSeparator1, Me.tms_UpdateAll, Me.tms_DeleteAll})
        Me.cms_follow_up.Name = "ContextMenuStrip1"
        Me.cms_follow_up.Size = New System.Drawing.Size(243, 186)
        Me.cms_follow_up.Text = "Edit"
        '
        'tms_Create
        '
        Me.tms_Create.Name = "tms_Create"
        Me.tms_Create.Size = New System.Drawing.Size(242, 22)
        Me.tms_Create.Text = "Create"
        '
        'tms_Update
        '
        Me.tms_Update.Name = "tms_Update"
        Me.tms_Update.Size = New System.Drawing.Size(242, 22)
        Me.tms_Update.Text = "Update"
        '
        'tms_Edit
        '
        Me.tms_Edit.Name = "tms_Edit"
        Me.tms_Edit.Size = New System.Drawing.Size(242, 22)
        Me.tms_Edit.Text = "Edit"
        '
        'tms_Delete
        '
        Me.tms_Delete.Name = "tms_Delete"
        Me.tms_Delete.Size = New System.Drawing.Size(242, 22)
        Me.tms_Delete.Text = "Delete"
        '
        'tms_Skip
        '
        Me.tms_Skip.Name = "tms_Skip"
        Me.tms_Skip.Size = New System.Drawing.Size(242, 22)
        Me.tms_Skip.Text = "Skip"
        '
        'tms_Lock
        '
        Me.tms_Lock.Name = "tms_Lock"
        Me.tms_Lock.Size = New System.Drawing.Size(242, 22)
        Me.tms_Lock.Text = "Lock"
        Me.tms_Lock.Visible = False
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(239, 6)
        Me.ToolStripSeparator1.Visible = False
        '
        'tms_UpdateAll
        '
        Me.tms_UpdateAll.Name = "tms_UpdateAll"
        Me.tms_UpdateAll.Size = New System.Drawing.Size(242, 22)
        Me.tms_UpdateAll.Text = "Update All Entries Which Allows"
        Me.tms_UpdateAll.Visible = False
        '
        'tms_DeleteAll
        '
        Me.tms_DeleteAll.Name = "tms_DeleteAll"
        Me.tms_DeleteAll.Size = New System.Drawing.Size(242, 22)
        Me.tms_DeleteAll.Text = "Delete All Entries Which Allows"
        Me.tms_DeleteAll.Visible = False
        '
        'ColTrans_ScheduleDate
        '
        Me.ColTrans_ScheduleDate.HeaderText = "Schedule Date"
        Me.ColTrans_ScheduleDate.Name = "ColTrans_ScheduleDate"
        Me.ColTrans_ScheduleDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTrans_ScheduleDate.Visible = False
        Me.ColTrans_ScheduleDate.Width = 84
        '
        'ColTrans_NoUrut
        '
        Me.ColTrans_NoUrut.HeaderText = "NoUrut"
        Me.ColTrans_NoUrut.Name = "ColTrans_NoUrut"
        Me.ColTrans_NoUrut.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTrans_NoUrut.Visible = False
        Me.ColTrans_NoUrut.Width = 47
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvReminder)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Size = New System.Drawing.Size(872, 376)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Reminder"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'dgvReminder
        '
        Me.dgvReminder.AllowUserToResizeColumns = False
        Me.dgvReminder.AllowUserToResizeRows = False
        DataGridViewCellStyle10.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvReminder.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle10
        Me.dgvReminder.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvReminder.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvReminder.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColReminder_No, Me.ColReminder_UserCode, Me.ColReminder_Remarks, Me.ColReminder_Interval, Me.ColReminder_PrimaryUser})
        Me.dgvReminder.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvReminder.Location = New System.Drawing.Point(2, 2)
        Me.dgvReminder.MultiSelect = False
        Me.dgvReminder.Name = "dgvReminder"
        Me.dgvReminder.RowTemplate.Height = 24
        Me.dgvReminder.Size = New System.Drawing.Size(868, 372)
        Me.dgvReminder.StandardTab = True
        Me.dgvReminder.TabIndex = 4
        '
        'ColReminder_No
        '
        Me.ColReminder_No.HeaderText = "No"
        Me.ColReminder_No.Name = "ColReminder_No"
        Me.ColReminder_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColReminder_No.Width = 27
        '
        'ColReminder_UserCode
        '
        Me.ColReminder_UserCode.HeaderText = "User Code"
        Me.ColReminder_UserCode.Name = "ColReminder_UserCode"
        Me.ColReminder_UserCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColReminder_UserCode.Width = 63
        '
        'ColReminder_Remarks
        '
        Me.ColReminder_Remarks.HeaderText = "Remarks"
        Me.ColReminder_Remarks.Name = "ColReminder_Remarks"
        Me.ColReminder_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColReminder_Remarks.Width = 55
        '
        'ColReminder_Interval
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColReminder_Interval.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColReminder_Interval.HeaderText = "Start Alarm (days before)"
        Me.ColReminder_Interval.Name = "ColReminder_Interval"
        Me.ColReminder_Interval.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColReminder_Interval.Width = 128
        '
        'ColReminder_PrimaryUser
        '
        Me.ColReminder_PrimaryUser.HeaderText = "Primary User"
        Me.ColReminder_PrimaryUser.Name = "ColReminder_PrimaryUser"
        Me.ColReminder_PrimaryUser.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColReminder_PrimaryUser.Visible = False
        Me.ColReminder_PrimaryUser.Width = 72
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 8)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtUnbalanceIncrement)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkIgnoreIncrementOnSkip)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtReadOnly)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryJournal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRefNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtUnbalance)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtLimit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.chkActive)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboEveryMonth)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboEveryWeek)
        Me.SplitContainer1.Panel2.Controls.Add(Me.chkLimit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dtpEndDate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboFrequency)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dtpStartDate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Size = New System.Drawing.Size(875, 81)
        Me.SplitContainer1.SplitterDistance = 435
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'txtUnbalanceIncrement
        '
        Me.txtUnbalanceIncrement.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtUnbalanceIncrement.Location = New System.Drawing.Point(314, 55)
        Me.txtUnbalanceIncrement.MaxLength = 50
        Me.txtUnbalanceIncrement.Name = "txtUnbalanceIncrement"
        Me.txtUnbalanceIncrement.ReadOnly = True
        Me.txtUnbalanceIncrement.Size = New System.Drawing.Size(118, 20)
        Me.txtUnbalanceIncrement.TabIndex = 11
        Me.txtUnbalanceIncrement.TabStop = False
        Me.txtUnbalanceIncrement.Text = "0"
        Me.txtUnbalanceIncrement.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(199, 58)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(109, 13)
        Me.Label10.TabIndex = 10
        Me.Label10.Text = "Unbalance Increment"
        '
        'chkIgnoreIncrementOnSkip
        '
        Me.chkIgnoreIncrementOnSkip.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkIgnoreIncrementOnSkip.AutoSize = True
        Me.chkIgnoreIncrementOnSkip.Location = New System.Drawing.Point(286, 31)
        Me.chkIgnoreIncrementOnSkip.Margin = New System.Windows.Forms.Padding(2)
        Me.chkIgnoreIncrementOnSkip.Name = "chkIgnoreIncrementOnSkip"
        Me.chkIgnoreIncrementOnSkip.Size = New System.Drawing.Size(145, 17)
        Me.chkIgnoreIncrementOnSkip.TabIndex = 7
        Me.chkIgnoreIncrementOnSkip.Text = "Ignore Increment on Skip"
        Me.chkIgnoreIncrementOnSkip.UseVisualStyleBackColor = True
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(77, 29)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(194, 20)
        Me.TxtRemarks.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Recurring No"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Description"
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Location = New System.Drawing.Point(75, 3)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(23, 20)
        Me.txtReadOnly.TabIndex = 2
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'TxtPrimaryJournal
        '
        Me.TxtPrimaryJournal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryJournal.Location = New System.Drawing.Point(3, 3)
        Me.TxtPrimaryJournal.Name = "TxtPrimaryJournal"
        Me.TxtPrimaryJournal.Size = New System.Drawing.Size(45, 20)
        Me.TxtPrimaryJournal.TabIndex = 0
        Me.TxtPrimaryJournal.TabStop = False
        Me.TxtPrimaryJournal.Text = "-1"
        Me.TxtPrimaryJournal.Visible = False
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(278, 5)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Date"
        '
        'TxtRefNo
        '
        Me.TxtRefNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRefNo.Location = New System.Drawing.Point(77, 3)
        Me.TxtRefNo.MaxLength = 50
        Me.TxtRefNo.Name = "TxtRefNo"
        Me.TxtRefNo.ReadOnly = True
        Me.TxtRefNo.Size = New System.Drawing.Size(194, 20)
        Me.TxtRefNo.TabIndex = 0
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(314, 3)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(118, 20)
        Me.dtpDate.TabIndex = 4
        '
        'TxtUnbalance
        '
        Me.TxtUnbalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUnbalance.Location = New System.Drawing.Point(77, 55)
        Me.TxtUnbalance.MaxLength = 50
        Me.TxtUnbalance.Name = "TxtUnbalance"
        Me.TxtUnbalance.ReadOnly = True
        Me.TxtUnbalance.Size = New System.Drawing.Size(116, 20)
        Me.TxtUnbalance.TabIndex = 9
        Me.TxtUnbalance.TabStop = False
        Me.TxtUnbalance.Text = "0"
        Me.TxtUnbalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Unbalance"
        '
        'txtLimit
        '
        Me.txtLimit.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtLimit.Location = New System.Drawing.Point(82, 29)
        Me.txtLimit.MaxLength = 5
        Me.txtLimit.Name = "txtLimit"
        Me.txtLimit.ReadOnly = True
        Me.txtLimit.Size = New System.Drawing.Size(87, 20)
        Me.txtLimit.TabIndex = 6
        Me.txtLimit.TabStop = False
        Me.txtLimit.Text = "0"
        Me.txtLimit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'chkActive
        '
        Me.chkActive.AutoSize = True
        Me.chkActive.Checked = True
        Me.chkActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkActive.Location = New System.Drawing.Point(267, 29)
        Me.chkActive.Margin = New System.Windows.Forms.Padding(2)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.Size = New System.Drawing.Size(56, 17)
        Me.chkActive.TabIndex = 8
        Me.chkActive.Text = "Active"
        Me.chkActive.UseVisualStyleBackColor = True
        '
        'cboEveryMonth
        '
        Me.cboEveryMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEveryMonth.FormattingEnabled = True
        Me.cboEveryMonth.Items.AddRange(New Object() {"First Day Of The Month", "Last Day Of The Month", "First Monday", "First Tuesday", "First Wednesday", "First Thusday", "First Friday", "First Saturday", "First Sunday", "Last Monday", "Last Tuesday", "Last Wednesday", "Last Thusday", "Last Friday", "Last Saturday", "Last Sunday", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"})
        Me.cboEveryMonth.Location = New System.Drawing.Point(267, 2)
        Me.cboEveryMonth.Name = "cboEveryMonth"
        Me.cboEveryMonth.Size = New System.Drawing.Size(166, 21)
        Me.cboEveryMonth.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(227, 5)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(34, 13)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Every"
        '
        'cboEveryWeek
        '
        Me.cboEveryWeek.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEveryWeek.FormattingEnabled = True
        Me.cboEveryWeek.Items.AddRange(New Object() {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"})
        Me.cboEveryWeek.Location = New System.Drawing.Point(267, 2)
        Me.cboEveryWeek.Name = "cboEveryWeek"
        Me.cboEveryWeek.Size = New System.Drawing.Size(166, 21)
        Me.cboEveryWeek.TabIndex = 5
        '
        'chkLimit
        '
        Me.chkLimit.AutoSize = True
        Me.chkLimit.Location = New System.Drawing.Point(5, 29)
        Me.chkLimit.Margin = New System.Windows.Forms.Padding(2)
        Me.chkLimit.Name = "chkLimit"
        Me.chkLimit.Size = New System.Drawing.Size(47, 17)
        Me.chkLimit.TabIndex = 4
        Me.chkLimit.Text = "Limit"
        Me.chkLimit.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(209, 54)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(52, 13)
        Me.Label9.TabIndex = 11
        Me.Label9.Text = "End Date"
        '
        'dtpEndDate
        '
        Me.dtpEndDate.CustomFormat = "dd MMM yyyy"
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndDate.Location = New System.Drawing.Point(267, 52)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(166, 20)
        Me.dtpEndDate.TabIndex = 12
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(175, 32)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(31, 13)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "times"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 56)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Start Date"
        '
        'cboFrequency
        '
        Me.cboFrequency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFrequency.FormattingEnabled = True
        Me.cboFrequency.Items.AddRange(New Object() {"Daily", "Weekly", "Bi-Weekly", "Monthly", "Quarter Yearly", "Semester", "Yearly"})
        Me.cboFrequency.Location = New System.Drawing.Point(82, 2)
        Me.cboFrequency.Name = "cboFrequency"
        Me.cboFrequency.Size = New System.Drawing.Size(121, 21)
        Me.cboFrequency.TabIndex = 1
        '
        'dtpStartDate
        '
        Me.dtpStartDate.CustomFormat = "dd MMM yyyy"
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartDate.Location = New System.Drawing.Point(82, 54)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(121, 20)
        Me.dtpStartDate.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Frequency"
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.AccModule.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 503)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 2
        Me.btnUP.TabStop = False
        Me.btnUP.Text = "^"
        Me.btnUP.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.AccModule.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 503)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 3
        Me.btnDown.TabStop = False
        Me.btnDown.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(452, 503)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(336, 503)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 587)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(892, 22)
        Me.StatusStrip1.TabIndex = 48
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(219, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(219, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(219, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(219, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(622, 544)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 47
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(446, 544)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 45
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(358, 544)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 44
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(182, 544)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 42
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(270, 544)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 43
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(534, 544)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 46
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'FrmTransaksiRecuring
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(892, 609)
        Me.Controls.Add(Me.btnViewJurnal)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnPrint)
        Me.MinimumSize = New System.Drawing.Size(902, 597)
        Me.Name = "FrmTransaksiRecuring"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Recurring Journal"
        Me.gboDetail.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvEntries, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cms_follow_up.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvReminder, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnViewJurnal As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryJournal As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtRefNo As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtUnbalance As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents cboFrequency As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Public WithEvents dgvReminder As System.Windows.Forms.DataGridView
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cms_follow_up As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tms_Create As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tms_Edit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tms_Delete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tms_Skip As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tms_UpdateAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tms_DeleteAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cboEveryMonth As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cboEveryWeek As System.Windows.Forms.ComboBox
    Friend WithEvents chkLimit As System.Windows.Forms.CheckBox
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    Friend WithEvents tms_Update As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents dgvEntries As System.Windows.Forms.DataGridView
    Friend WithEvents chkIgnoreIncrementOnSkip As System.Windows.Forms.CheckBox
    Friend WithEvents txtLimit As System.Windows.Forms.TextBox
    Friend WithEvents tms_Lock As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents txtUnbalanceIncrement As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColReminder_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReminder_UserCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReminder_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReminder_Interval As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReminder_PrimaryUser As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTrans_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTrans_JournalNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTrans_JournalDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTrans_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTrans_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTrans_PrimaryJournal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTrans_Status As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTrans_FollowUp As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColTrans_ScheduleDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTrans_NoUrut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDebet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKredit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExchangeRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOriginalCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDebet_OriginalCurr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCredit_OriginalCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJob As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryJob As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAcc_IncrementStyle As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColAcc_IncrementValue As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
