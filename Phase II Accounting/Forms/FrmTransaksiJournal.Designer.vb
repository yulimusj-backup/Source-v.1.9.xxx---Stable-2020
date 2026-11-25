<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiJournal
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
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.LblQuickJurnal = New System.Windows.Forms.Label
        Me.llbEditQuickJournal = New System.Windows.Forms.LinkLabel
        Me.cboQuickJournal = New System.Windows.Forms.ComboBox
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtPrimaryJournal = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtRefNo = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.TxtUnbalance = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
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
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnViewJurnal = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Location = New System.Drawing.Point(12, 6)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(854, 496)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblQuickJurnal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.llbEditQuickJournal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboQuickJournal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtReadOnly)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryJournal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRefNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtUnbalance)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Size = New System.Drawing.Size(837, 81)
        Me.SplitContainer1.SplitterDistance = 417
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'LblQuickJurnal
        '
        Me.LblQuickJurnal.AutoSize = True
        Me.LblQuickJurnal.Location = New System.Drawing.Point(3, 54)
        Me.LblQuickJurnal.Name = "LblQuickJurnal"
        Me.LblQuickJurnal.Size = New System.Drawing.Size(72, 13)
        Me.LblQuickJurnal.TabIndex = 214
        Me.LblQuickJurnal.Text = "Quick Journal"
        '
        'llbEditQuickJournal
        '
        Me.llbEditQuickJournal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbEditQuickJournal.AutoSize = True
        Me.llbEditQuickJournal.Location = New System.Drawing.Point(371, 54)
        Me.llbEditQuickJournal.Name = "llbEditQuickJournal"
        Me.llbEditQuickJournal.Size = New System.Drawing.Size(25, 13)
        Me.llbEditQuickJournal.TabIndex = 3
        Me.llbEditQuickJournal.TabStop = True
        Me.llbEditQuickJournal.Text = "Edit"
        '
        'cboQuickJournal
        '
        Me.cboQuickJournal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboQuickJournal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboQuickJournal.FormattingEnabled = True
        Me.cboQuickJournal.Location = New System.Drawing.Point(104, 51)
        Me.cboQuickJournal.Name = "cboQuickJournal"
        Me.cboQuickJournal.Size = New System.Drawing.Size(259, 21)
        Me.cboQuickJournal.TabIndex = 2
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Location = New System.Drawing.Point(75, 3)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(23, 20)
        Me.txtReadOnly.TabIndex = 211
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'TxtPrimaryJournal
        '
        Me.TxtPrimaryJournal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryJournal.Location = New System.Drawing.Point(69, 26)
        Me.TxtPrimaryJournal.Name = "TxtPrimaryJournal"
        Me.TxtPrimaryJournal.Size = New System.Drawing.Size(27, 20)
        Me.TxtPrimaryJournal.TabIndex = 163
        Me.TxtPrimaryJournal.TabStop = False
        Me.TxtPrimaryJournal.Text = "-1"
        Me.TxtPrimaryJournal.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "No Journal"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'TxtRefNo
        '
        Me.TxtRefNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRefNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRefNo.Location = New System.Drawing.Point(104, 3)
        Me.TxtRefNo.MaxLength = 50
        Me.TxtRefNo.Name = "TxtRefNo"
        Me.TxtRefNo.Size = New System.Drawing.Size(310, 20)
        Me.TxtRefNo.TabIndex = 0
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(104, 26)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(310, 20)
        Me.dtpDate.TabIndex = 1
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(106, 5)
        Me.TxtRemarks.MaxLength = 150
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(308, 43)
        Me.TxtRemarks.TabIndex = 0
        '
        'TxtUnbalance
        '
        Me.TxtUnbalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUnbalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUnbalance.Location = New System.Drawing.Point(106, 54)
        Me.TxtUnbalance.MaxLength = 50
        Me.TxtUnbalance.Name = "TxtUnbalance"
        Me.TxtUnbalance.ReadOnly = True
        Me.TxtUnbalance.Size = New System.Drawing.Size(308, 20)
        Me.TxtUnbalance.TabIndex = 1
        Me.TxtUnbalance.TabStop = False
        Me.TxtUnbalance.Text = "0"
        Me.TxtUnbalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Description"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 164
        Me.Label2.Text = "Unbalance"
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.AccModule.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 464)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 2
        Me.btnUP.TabStop = False
        Me.btnUP.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.AccModule.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 464)
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
        Me.btnCancel.Location = New System.Drawing.Point(433, 464)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(317, 464)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColDescAccount, Me.ColDebet, Me.ColKredit, Me.ColExchangeRate, Me.ColOriginalCurrency, Me.ColDebet_OriginalCurr, Me.ColCredit_OriginalCurrency, Me.ColRemarks, Me.ColDept, Me.ColJob, Me.ColPrimaryDept, Me.ColPrimaryJob, Me.ColPrimaryAkun, Me.ColPrimaryCurrency})
        Me.dgvItem.Location = New System.Drawing.Point(6, 95)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(837, 363)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 1
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
        Me.ColDebet.MaxInputLength = 25
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
        Me.ColKredit.MaxInputLength = 25
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
        Me.ColDept.Width = 93
        '
        'ColJob
        '
        Me.ColJob.HeaderText = "Job"
        Me.ColJob.Name = "ColJob"
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
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 537)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(886, 22)
        Me.StatusStrip1.TabIndex = 39
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(443, 508)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 5
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(619, 508)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(355, 508)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 4
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(179, 508)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 2
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(267, 508)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 3
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnViewJurnal
        '
        Me.btnViewJurnal.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnViewJurnal.Location = New System.Drawing.Point(12, 508)
        Me.btnViewJurnal.Name = "btnViewJurnal"
        Me.btnViewJurnal.Size = New System.Drawing.Size(80, 26)
        Me.btnViewJurnal.TabIndex = 1
        Me.btnViewJurnal.TabStop = False
        Me.btnViewJurnal.Text = "View Journal"
        Me.btnViewJurnal.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(531, 508)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 6
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'FrmTransaksiJournal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(886, 559)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnViewJurnal)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.StatusStrip1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(902, 597)
        Me.Name = "FrmTransaksiJournal"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Journal"
        Me.gboDetail.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtRefNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents TxtUnbalance As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnViewJurnal As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryJournal As System.Windows.Forms.TextBox
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents LblQuickJurnal As System.Windows.Forms.Label
    Friend WithEvents llbEditQuickJournal As System.Windows.Forms.LinkLabel
    Friend WithEvents cboQuickJournal As System.Windows.Forms.ComboBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
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
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
