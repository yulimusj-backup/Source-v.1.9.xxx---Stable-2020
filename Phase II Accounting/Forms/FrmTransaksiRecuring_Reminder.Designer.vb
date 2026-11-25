<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiRecuring_Reminder
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.dgvHeader = New System.Windows.Forms.DataGridView
        Me.ColRecurringTrans_RecurringNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRecurringTrans_ReccuringDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRecurringTrans_JournalNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRecurringTrans_JournalDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRecurringTrans_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRecurringTrans_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRecurringTrans_Status = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRecurringTrans_FollowUp = New System.Windows.Forms.DataGridViewButtonColumn
        Me.cmsRecurring_follow_up = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tms_Create = New System.Windows.Forms.ToolStripMenuItem
        Me.tms_CreateOpen = New System.Windows.Forms.ToolStripMenuItem
        Me.tms_Skip = New System.Windows.Forms.ToolStripMenuItem
        Me.GoToRecurringJournalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ColPrimaryRecurring = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colModify = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColRecurringTrans_PrimaryJournal = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.ColPrimaryJournal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryRecurringDet = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnRecurring_Show = New System.Windows.Forms.Button
        Me.btnCreate = New System.Windows.Forms.Button
        Me.btnCreateAll = New System.Windows.Forms.Button
        Me.btnSkip = New System.Windows.Forms.Button
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cmsRecurring_follow_up.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 6)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(301, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "The following recurring journal has matured, specify what to do"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(5, 47)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvHeader)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvItem)
        Me.SplitContainer1.Size = New System.Drawing.Size(873, 427)
        Me.SplitContainer1.SplitterDistance = 213
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 6
        '
        'dgvHeader
        '
        Me.dgvHeader.AllowUserToAddRows = False
        Me.dgvHeader.AllowUserToDeleteRows = False
        Me.dgvHeader.AllowUserToResizeColumns = False
        Me.dgvHeader.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvHeader.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvHeader.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvHeader.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvHeader.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColRecurringTrans_RecurringNo, Me.ColRecurringTrans_ReccuringDesc, Me.ColRecurringTrans_JournalNumber, Me.ColRecurringTrans_JournalDate, Me.ColRecurringTrans_Description, Me.ColRecurringTrans_Remarks, Me.ColRecurringTrans_Status, Me.ColRecurringTrans_FollowUp, Me.ColPrimaryRecurring, Me.colModify, Me.ColRecurringTrans_PrimaryJournal})
        Me.dgvHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvHeader.Location = New System.Drawing.Point(0, 0)
        Me.dgvHeader.MultiSelect = False
        Me.dgvHeader.Name = "dgvHeader"
        Me.dgvHeader.RowTemplate.Height = 24
        Me.dgvHeader.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvHeader.Size = New System.Drawing.Size(873, 213)
        Me.dgvHeader.StandardTab = True
        Me.dgvHeader.TabIndex = 6
        '
        'ColRecurringTrans_RecurringNo
        '
        Me.ColRecurringTrans_RecurringNo.Frozen = True
        Me.ColRecurringTrans_RecurringNo.HeaderText = "Recurring No"
        Me.ColRecurringTrans_RecurringNo.Name = "ColRecurringTrans_RecurringNo"
        Me.ColRecurringTrans_RecurringNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRecurringTrans_RecurringNo.Width = 76
        '
        'ColRecurringTrans_ReccuringDesc
        '
        Me.ColRecurringTrans_ReccuringDesc.Frozen = True
        Me.ColRecurringTrans_ReccuringDesc.HeaderText = "Description"
        Me.ColRecurringTrans_ReccuringDesc.Name = "ColRecurringTrans_ReccuringDesc"
        Me.ColRecurringTrans_ReccuringDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRecurringTrans_ReccuringDesc.Width = 66
        '
        'ColRecurringTrans_JournalNumber
        '
        Me.ColRecurringTrans_JournalNumber.Frozen = True
        Me.ColRecurringTrans_JournalNumber.HeaderText = "Journal No"
        Me.ColRecurringTrans_JournalNumber.Name = "ColRecurringTrans_JournalNumber"
        Me.ColRecurringTrans_JournalNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRecurringTrans_JournalNumber.Width = 64
        '
        'ColRecurringTrans_JournalDate
        '
        Me.ColRecurringTrans_JournalDate.Frozen = True
        Me.ColRecurringTrans_JournalDate.HeaderText = "Date"
        Me.ColRecurringTrans_JournalDate.Name = "ColRecurringTrans_JournalDate"
        Me.ColRecurringTrans_JournalDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRecurringTrans_JournalDate.Width = 36
        '
        'ColRecurringTrans_Description
        '
        Me.ColRecurringTrans_Description.Frozen = True
        Me.ColRecurringTrans_Description.HeaderText = "Description"
        Me.ColRecurringTrans_Description.Name = "ColRecurringTrans_Description"
        Me.ColRecurringTrans_Description.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRecurringTrans_Description.Width = 66
        '
        'ColRecurringTrans_Remarks
        '
        Me.ColRecurringTrans_Remarks.Frozen = True
        Me.ColRecurringTrans_Remarks.HeaderText = "Remarks"
        Me.ColRecurringTrans_Remarks.Name = "ColRecurringTrans_Remarks"
        Me.ColRecurringTrans_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRecurringTrans_Remarks.Width = 55
        '
        'ColRecurringTrans_Status
        '
        Me.ColRecurringTrans_Status.Frozen = True
        Me.ColRecurringTrans_Status.HeaderText = "Status"
        Me.ColRecurringTrans_Status.Name = "ColRecurringTrans_Status"
        Me.ColRecurringTrans_Status.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRecurringTrans_Status.Width = 43
        '
        'ColRecurringTrans_FollowUp
        '
        Me.ColRecurringTrans_FollowUp.ContextMenuStrip = Me.cmsRecurring_follow_up
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.NullValue = "Right Click"
        Me.ColRecurringTrans_FollowUp.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColRecurringTrans_FollowUp.Frozen = True
        Me.ColRecurringTrans_FollowUp.HeaderText = "Follow Up"
        Me.ColRecurringTrans_FollowUp.Name = "ColRecurringTrans_FollowUp"
        Me.ColRecurringTrans_FollowUp.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRecurringTrans_FollowUp.Text = "Right Click"
        Me.ColRecurringTrans_FollowUp.Width = 60
        '
        'cmsRecurring_follow_up
        '
        Me.cmsRecurring_follow_up.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tms_Create, Me.tms_CreateOpen, Me.tms_Skip, Me.GoToRecurringJournalToolStripMenuItem})
        Me.cmsRecurring_follow_up.Name = "ContextMenuStrip1"
        Me.cmsRecurring_follow_up.Size = New System.Drawing.Size(202, 92)
        Me.cmsRecurring_follow_up.Text = "Edit"
        '
        'tms_Create
        '
        Me.tms_Create.Name = "tms_Create"
        Me.tms_Create.Size = New System.Drawing.Size(201, 22)
        Me.tms_Create.Text = "Create"
        '
        'tms_CreateOpen
        '
        Me.tms_CreateOpen.Name = "tms_CreateOpen"
        Me.tms_CreateOpen.Size = New System.Drawing.Size(201, 22)
        Me.tms_CreateOpen.Text = "Create && Open"
        '
        'tms_Skip
        '
        Me.tms_Skip.Name = "tms_Skip"
        Me.tms_Skip.Size = New System.Drawing.Size(201, 22)
        Me.tms_Skip.Text = "Skip"
        '
        'GoToRecurringJournalToolStripMenuItem
        '
        Me.GoToRecurringJournalToolStripMenuItem.Name = "GoToRecurringJournalToolStripMenuItem"
        Me.GoToRecurringJournalToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.GoToRecurringJournalToolStripMenuItem.Text = "Go To Recurring Journal"
        '
        'ColPrimaryRecurring
        '
        Me.ColPrimaryRecurring.HeaderText = "Primary Recurring"
        Me.ColPrimaryRecurring.Name = "ColPrimaryRecurring"
        Me.ColPrimaryRecurring.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryRecurring.Visible = False
        Me.ColPrimaryRecurring.Width = 96
        '
        'colModify
        '
        Me.colModify.Frozen = True
        Me.colModify.HeaderText = "Modify"
        Me.colModify.Name = "colModify"
        Me.colModify.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colModify.Width = 44
        '
        'ColRecurringTrans_PrimaryJournal
        '
        Me.ColRecurringTrans_PrimaryJournal.HeaderText = "Primary Journal"
        Me.ColRecurringTrans_PrimaryJournal.Name = "ColRecurringTrans_PrimaryJournal"
        Me.ColRecurringTrans_PrimaryJournal.Visible = False
        Me.ColRecurringTrans_PrimaryJournal.Width = 103
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColDescAccount, Me.ColDebet, Me.ColKredit, Me.ColExchangeRate, Me.ColOriginalCurrency, Me.ColDebet_OriginalCurr, Me.ColCredit_OriginalCurrency, Me.ColRemarks, Me.ColDept, Me.ColJob, Me.ColPrimaryDept, Me.ColPrimaryJob, Me.ColPrimaryAkun, Me.ColPrimaryCurrency, Me.ColPrimaryJournal, Me.ColPrimaryRecurringDet})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItem.Size = New System.Drawing.Size(873, 211)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 2
        '
        'ColAccount
        '
        Me.ColAccount.HeaderText = "Account"
        Me.ColAccount.MaxInputLength = 50
        Me.ColAccount.Name = "ColAccount"
        Me.ColAccount.ReadOnly = True
        Me.ColAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccount.Width = 53
        '
        'ColDescAccount
        '
        Me.ColDescAccount.HeaderText = "Description"
        Me.ColDescAccount.MaxInputLength = 150
        Me.ColDescAccount.Name = "ColDescAccount"
        Me.ColDescAccount.ReadOnly = True
        Me.ColDescAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDescAccount.Width = 66
        '
        'ColDebet
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.NullValue = "0"
        Me.ColDebet.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColDebet.HeaderText = "Debet"
        Me.ColDebet.MaxInputLength = 25
        Me.ColDebet.Name = "ColDebet"
        Me.ColDebet.ReadOnly = True
        Me.ColDebet.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet.Width = 42
        '
        'ColKredit
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.NullValue = "0"
        Me.ColKredit.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColKredit.HeaderText = "Credit"
        Me.ColKredit.MaxInputLength = 25
        Me.ColKredit.Name = "ColKredit"
        Me.ColKredit.ReadOnly = True
        Me.ColKredit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColKredit.Width = 40
        '
        'ColExchangeRate
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.NullValue = "1"
        Me.ColExchangeRate.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColExchangeRate.HeaderText = "Exchange Rate"
        Me.ColExchangeRate.Name = "ColExchangeRate"
        Me.ColExchangeRate.ReadOnly = True
        Me.ColExchangeRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExchangeRate.Visible = False
        Me.ColExchangeRate.Width = 87
        '
        'ColOriginalCurrency
        '
        Me.ColOriginalCurrency.HeaderText = "Currency"
        Me.ColOriginalCurrency.Name = "ColOriginalCurrency"
        Me.ColOriginalCurrency.ReadOnly = True
        Me.ColOriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOriginalCurrency.Visible = False
        Me.ColOriginalCurrency.Width = 55
        '
        'ColDebet_OriginalCurr
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle7.NullValue = "0"
        Me.ColDebet_OriginalCurr.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColDebet_OriginalCurr.HeaderText = "Debet"
        Me.ColDebet_OriginalCurr.Name = "ColDebet_OriginalCurr"
        Me.ColDebet_OriginalCurr.ReadOnly = True
        Me.ColDebet_OriginalCurr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet_OriginalCurr.Visible = False
        Me.ColDebet_OriginalCurr.Width = 42
        '
        'ColCredit_OriginalCurrency
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle8.NullValue = "0"
        Me.ColCredit_OriginalCurrency.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColCredit_OriginalCurrency.HeaderText = "Credit"
        Me.ColCredit_OriginalCurrency.Name = "ColCredit_OriginalCurrency"
        Me.ColCredit_OriginalCurrency.ReadOnly = True
        Me.ColCredit_OriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCredit_OriginalCurrency.Visible = False
        Me.ColCredit_OriginalCurrency.Width = 40
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 150
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 55
        '
        'ColDept
        '
        Me.ColDept.HeaderText = "Departement"
        Me.ColDept.Name = "ColDept"
        Me.ColDept.ReadOnly = True
        Me.ColDept.Visible = False
        Me.ColDept.Width = 93
        '
        'ColJob
        '
        Me.ColJob.HeaderText = "Job"
        Me.ColJob.Name = "ColJob"
        Me.ColJob.ReadOnly = True
        Me.ColJob.Visible = False
        Me.ColJob.Width = 49
        '
        'ColPrimaryDept
        '
        Me.ColPrimaryDept.HeaderText = "Primary Dept"
        Me.ColPrimaryDept.Name = "ColPrimaryDept"
        Me.ColPrimaryDept.ReadOnly = True
        Me.ColPrimaryDept.Visible = False
        Me.ColPrimaryDept.Width = 92
        '
        'ColPrimaryJob
        '
        Me.ColPrimaryJob.HeaderText = "Primary Job"
        Me.ColPrimaryJob.Name = "ColPrimaryJob"
        Me.ColPrimaryJob.ReadOnly = True
        Me.ColPrimaryJob.Visible = False
        Me.ColPrimaryJob.Width = 86
        '
        'ColPrimaryAkun
        '
        Me.ColPrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColPrimaryAkun.Name = "ColPrimaryAkun"
        Me.ColPrimaryAkun.ReadOnly = True
        Me.ColPrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryAkun.Visible = False
        Me.ColPrimaryAkun.Width = 72
        '
        'ColPrimaryCurrency
        '
        Me.ColPrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.ReadOnly = True
        Me.ColPrimaryCurrency.Visible = False
        Me.ColPrimaryCurrency.Width = 111
        '
        'ColPrimaryJournal
        '
        Me.ColPrimaryJournal.HeaderText = "PrimaryJournal"
        Me.ColPrimaryJournal.Name = "ColPrimaryJournal"
        Me.ColPrimaryJournal.ReadOnly = True
        Me.ColPrimaryJournal.Visible = False
        '
        'ColPrimaryRecurringDet
        '
        Me.ColPrimaryRecurringDet.HeaderText = "Primary Recurring"
        Me.ColPrimaryRecurringDet.Name = "ColPrimaryRecurringDet"
        Me.ColPrimaryRecurringDet.ReadOnly = True
        Me.ColPrimaryRecurringDet.Visible = False
        Me.ColPrimaryRecurringDet.Width = 115
        '
        'btnRecurring_Show
        '
        Me.btnRecurring_Show.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRecurring_Show.Location = New System.Drawing.Point(5, 479)
        Me.btnRecurring_Show.Margin = New System.Windows.Forms.Padding(2)
        Me.btnRecurring_Show.Name = "btnRecurring_Show"
        Me.btnRecurring_Show.Size = New System.Drawing.Size(86, 19)
        Me.btnRecurring_Show.TabIndex = 7
        Me.btnRecurring_Show.Text = "Hide Detail"
        Me.btnRecurring_Show.UseVisualStyleBackColor = True
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(13, 22)
        Me.btnCreate.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(56, 22)
        Me.btnCreate.TabIndex = 8
        Me.btnCreate.Text = "Create"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'btnCreateAll
        '
        Me.btnCreateAll.Location = New System.Drawing.Point(314, 1)
        Me.btnCreateAll.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCreateAll.Name = "btnCreateAll"
        Me.btnCreateAll.Size = New System.Drawing.Size(76, 22)
        Me.btnCreateAll.TabIndex = 9
        Me.btnCreateAll.Text = "Create All"
        Me.btnCreateAll.UseVisualStyleBackColor = True
        Me.btnCreateAll.Visible = False
        '
        'btnSkip
        '
        Me.btnSkip.Location = New System.Drawing.Point(73, 21)
        Me.btnSkip.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSkip.Name = "btnSkip"
        Me.btnSkip.Size = New System.Drawing.Size(56, 22)
        Me.btnSkip.TabIndex = 10
        Me.btnSkip.Text = "Skip"
        Me.btnSkip.UseVisualStyleBackColor = True
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.BackColor = System.Drawing.SystemColors.Window
        Me.llbUncheckAll.Location = New System.Drawing.Point(813, 28)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAll.TabIndex = 14
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.BackColor = System.Drawing.SystemColors.Window
        Me.llbCheckAll.Location = New System.Drawing.Point(754, 28)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAll.TabIndex = 15
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(73, 21)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(56, 22)
        Me.btnExit.TabIndex = 16
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'FrmTransaksiRecuring_Reminder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(884, 508)
        Me.Controls.Add(Me.llbUncheckAll)
        Me.Controls.Add(Me.llbCheckAll)
        Me.Controls.Add(Me.btnRecurring_Show)
        Me.Controls.Add(Me.btnSkip)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnCreateAll)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(900, 546)
        Me.Name = "FrmTransaksiRecuring_Reminder"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Recurring Journal Reminder"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cmsRecurring_follow_up.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Public WithEvents dgvHeader As System.Windows.Forms.DataGridView
    Friend WithEvents btnRecurring_Show As System.Windows.Forms.Button
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents cmsRecurring_follow_up As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tms_Create As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tms_CreateOpen As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tms_Skip As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GoToRecurringJournalToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnCreate As System.Windows.Forms.Button
    Friend WithEvents btnCreateAll As System.Windows.Forms.Button
    Friend WithEvents btnSkip As System.Windows.Forms.Button
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
    Friend WithEvents ColPrimaryJournal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryRecurringDet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents ColRecurringTrans_RecurringNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRecurringTrans_ReccuringDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRecurringTrans_JournalNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRecurringTrans_JournalDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRecurringTrans_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRecurringTrans_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRecurringTrans_Status As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRecurringTrans_FollowUp As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColPrimaryRecurring As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colModify As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColRecurringTrans_PrimaryJournal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnExit As System.Windows.Forms.Button
End Class
