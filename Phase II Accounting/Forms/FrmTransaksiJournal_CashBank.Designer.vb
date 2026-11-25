<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiJournal_CashBank
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
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle17 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle18 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle19 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle20 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnViewJurnal = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtReferenceParty = New System.Windows.Forms.TextBox
        Me.LblReferenceParty = New System.Windows.Forms.Label
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.TxtPrimaryAccount = New System.Windows.Forms.TextBox
        Me.cboType = New System.Windows.Forms.ComboBox
        Me.cboAccount = New System.Windows.Forms.ComboBox
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtPrimaryJournal = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtRefNo = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtAmount = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtExchangeRate = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.LblCurrency = New System.Windows.Forms.Label
        Me.TxtAmountOriginal = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvPayment = New System.Windows.Forms.DataGridView
        Me.COlItemPay_Account = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COlItemPay_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemPay_Curr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COlItemPay_AmountOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemPay_ExRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COlItemPay_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COlItemPay_DeptCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COlItemPay_JobCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemPay_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COlItemPay_PrimaryDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COlItemPay_PrimaryJob = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COlItemPay_PrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
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
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvPayment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnViewJurnal
        '
        Me.btnViewJurnal.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnViewJurnal.Location = New System.Drawing.Point(28, 502)
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
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Location = New System.Drawing.Point(0, 3)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(886, 480)
        Me.gboDetail.TabIndex = 6
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtReferenceParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblReferenceParty)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LinkLabel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryAccount)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboType)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboAccount)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtReadOnly)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryJournal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtRefNo)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtAmount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtExchangeRate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblCurrency)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtAmountOriginal)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Size = New System.Drawing.Size(869, 81)
        Me.SplitContainer1.SplitterDistance = 433
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'TxtReferenceParty
        '
        Me.TxtReferenceParty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReferenceParty.Location = New System.Drawing.Point(104, 29)
        Me.TxtReferenceParty.MaxLength = 150
        Me.TxtReferenceParty.Name = "TxtReferenceParty"
        Me.TxtReferenceParty.Size = New System.Drawing.Size(321, 20)
        Me.TxtReferenceParty.TabIndex = 2
        '
        'LblReferenceParty
        '
        Me.LblReferenceParty.AutoSize = True
        Me.LblReferenceParty.Location = New System.Drawing.Point(3, 32)
        Me.LblReferenceParty.Name = "LblReferenceParty"
        Me.LblReferenceParty.Size = New System.Drawing.Size(79, 13)
        Me.LblReferenceParty.TabIndex = 216
        Me.LblReferenceParty.Text = "Received From"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(342, 55)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(88, 13)
        Me.LinkLabel1.TabIndex = 4
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Cheque Payment"
        '
        'TxtPrimaryAccount
        '
        Me.TxtPrimaryAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryAccount.Location = New System.Drawing.Point(67, 52)
        Me.TxtPrimaryAccount.Name = "TxtPrimaryAccount"
        Me.TxtPrimaryAccount.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimaryAccount.TabIndex = 214
        Me.TxtPrimaryAccount.TabStop = False
        Me.TxtPrimaryAccount.Text = "-1"
        Me.TxtPrimaryAccount.Visible = False
        '
        'cboType
        '
        Me.cboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboType.FormattingEnabled = True
        Me.cboType.Items.AddRange(New Object() {"Cash In", "Cash Out"})
        Me.cboType.Location = New System.Drawing.Point(6, 52)
        Me.cboType.Name = "cboType"
        Me.cboType.Size = New System.Drawing.Size(92, 21)
        Me.cboType.TabIndex = 2
        '
        'cboAccount
        '
        Me.cboAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccount.FormattingEnabled = True
        Me.cboAccount.Location = New System.Drawing.Point(104, 52)
        Me.cboAccount.Name = "cboAccount"
        Me.cboAccount.Size = New System.Drawing.Size(232, 21)
        Me.cboAccount.TabIndex = 3
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
        Me.TxtPrimaryJournal.Size = New System.Drawing.Size(43, 20)
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
        'dtpDate
        '
        Me.dtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(300, 3)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(130, 20)
        Me.dtpDate.TabIndex = 1
        '
        'TxtRefNo
        '
        Me.TxtRefNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRefNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRefNo.Location = New System.Drawing.Point(104, 3)
        Me.TxtRefNo.MaxLength = 50
        Me.TxtRefNo.Name = "TxtRefNo"
        Me.TxtRefNo.Size = New System.Drawing.Size(190, 20)
        Me.TxtRefNo.TabIndex = 0
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(103, 62)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(26, 13)
        Me.Label5.TabIndex = 218
        Me.Label5.Text = "IDR"
        '
        'TxtAmount
        '
        Me.TxtAmount.Location = New System.Drawing.Point(135, 59)
        Me.TxtAmount.Name = "TxtAmount"
        Me.TxtAmount.ReadOnly = True
        Me.TxtAmount.Size = New System.Drawing.Size(97, 20)
        Me.TxtAmount.TabIndex = 3
        Me.TxtAmount.TabStop = False
        Me.TxtAmount.Text = "0"
        Me.TxtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 62)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(43, 13)
        Me.Label7.TabIndex = 216
        Me.Label7.Text = "Amount"
        '
        'TxtExchangeRate
        '
        Me.TxtExchangeRate.Location = New System.Drawing.Point(262, 31)
        Me.TxtExchangeRate.Name = "TxtExchangeRate"
        Me.TxtExchangeRate.Size = New System.Drawing.Size(153, 20)
        Me.TxtExchangeRate.TabIndex = 2
        Me.TxtExchangeRate.Text = "1.0000"
        Me.TxtExchangeRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(238, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(18, 13)
        Me.Label2.TabIndex = 214
        Me.Label2.Text = "@"
        '
        'LblCurrency
        '
        Me.LblCurrency.AutoSize = True
        Me.LblCurrency.Location = New System.Drawing.Point(103, 36)
        Me.LblCurrency.Name = "LblCurrency"
        Me.LblCurrency.Size = New System.Drawing.Size(26, 13)
        Me.LblCurrency.TabIndex = 213
        Me.LblCurrency.Text = "IDR"
        '
        'TxtAmountOriginal
        '
        Me.TxtAmountOriginal.Location = New System.Drawing.Point(135, 33)
        Me.TxtAmountOriginal.Name = "TxtAmountOriginal"
        Me.TxtAmountOriginal.ReadOnly = True
        Me.TxtAmountOriginal.Size = New System.Drawing.Size(97, 20)
        Me.TxtAmountOriginal.TabIndex = 1
        Me.TxtAmountOriginal.TabStop = False
        Me.TxtAmountOriginal.Text = "0"
        Me.TxtAmountOriginal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 144
        Me.Label1.Text = "Amount Original"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(106, 5)
        Me.TxtRemarks.MaxLength = 150
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(324, 20)
        Me.TxtRemarks.TabIndex = 0
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
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.AccModule.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 448)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 2
        Me.btnUP.TabStop = False
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
        Me.btnDown.Location = New System.Drawing.Point(40, 448)
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
        Me.btnCancel.Location = New System.Drawing.Point(449, 448)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(333, 448)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(6, 95)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(869, 347)
        Me.TabControl1.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvPayment)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(861, 321)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Payment Detail"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvPayment
        '
        Me.dgvPayment.AllowUserToResizeColumns = False
        Me.dgvPayment.AllowUserToResizeRows = False
        DataGridViewCellStyle11.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvPayment.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle11
        Me.dgvPayment.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvPayment.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.COlItemPay_Account, Me.COlItemPay_Desc, Me.ColItemPay_Curr, Me.COlItemPay_AmountOriginal, Me.ColItemPay_ExRate, Me.COlItemPay_Remarks, Me.COlItemPay_DeptCode, Me.COlItemPay_JobCode, Me.ColItemPay_Amount, Me.COlItemPay_PrimaryDept, Me.COlItemPay_PrimaryJob, Me.COlItemPay_PrimaryAkun})
        Me.dgvPayment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvPayment.Location = New System.Drawing.Point(3, 3)
        Me.dgvPayment.MultiSelect = False
        Me.dgvPayment.Name = "dgvPayment"
        Me.dgvPayment.RowTemplate.Height = 24
        Me.dgvPayment.Size = New System.Drawing.Size(855, 315)
        Me.dgvPayment.StandardTab = True
        Me.dgvPayment.TabIndex = 2
        '
        'COlItemPay_Account
        '
        Me.COlItemPay_Account.HeaderText = "Account"
        Me.COlItemPay_Account.MaxInputLength = 50
        Me.COlItemPay_Account.Name = "COlItemPay_Account"
        Me.COlItemPay_Account.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COlItemPay_Account.Width = 53
        '
        'COlItemPay_Desc
        '
        Me.COlItemPay_Desc.HeaderText = "Description"
        Me.COlItemPay_Desc.MaxInputLength = 150
        Me.COlItemPay_Desc.Name = "COlItemPay_Desc"
        Me.COlItemPay_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COlItemPay_Desc.Width = 66
        '
        'ColItemPay_Curr
        '
        Me.ColItemPay_Curr.HeaderText = "Currency"
        Me.ColItemPay_Curr.Name = "ColItemPay_Curr"
        Me.ColItemPay_Curr.Width = 74
        '
        'COlItemPay_AmountOriginal
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.COlItemPay_AmountOriginal.DefaultCellStyle = DataGridViewCellStyle12
        Me.COlItemPay_AmountOriginal.HeaderText = "Amount"
        Me.COlItemPay_AmountOriginal.Name = "COlItemPay_AmountOriginal"
        Me.COlItemPay_AmountOriginal.Width = 68
        '
        'ColItemPay_ExRate
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItemPay_ExRate.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColItemPay_ExRate.HeaderText = "Exchange Rate"
        Me.ColItemPay_ExRate.Name = "ColItemPay_ExRate"
        Me.ColItemPay_ExRate.Width = 106
        '
        'COlItemPay_Remarks
        '
        Me.COlItemPay_Remarks.HeaderText = "Remarks"
        Me.COlItemPay_Remarks.Name = "COlItemPay_Remarks"
        Me.COlItemPay_Remarks.Width = 74
        '
        'COlItemPay_DeptCode
        '
        Me.COlItemPay_DeptCode.HeaderText = "Departement"
        Me.COlItemPay_DeptCode.Name = "COlItemPay_DeptCode"
        Me.COlItemPay_DeptCode.Width = 93
        '
        'COlItemPay_JobCode
        '
        Me.COlItemPay_JobCode.HeaderText = "Job"
        Me.COlItemPay_JobCode.Name = "COlItemPay_JobCode"
        Me.COlItemPay_JobCode.Width = 49
        '
        'ColItemPay_Amount
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItemPay_Amount.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColItemPay_Amount.HeaderText = "Amount"
        Me.ColItemPay_Amount.Name = "ColItemPay_Amount"
        Me.ColItemPay_Amount.Visible = False
        Me.ColItemPay_Amount.Width = 68
        '
        'COlItemPay_PrimaryDept
        '
        Me.COlItemPay_PrimaryDept.HeaderText = "Primary Dept"
        Me.COlItemPay_PrimaryDept.Name = "COlItemPay_PrimaryDept"
        Me.COlItemPay_PrimaryDept.Visible = False
        Me.COlItemPay_PrimaryDept.Width = 92
        '
        'COlItemPay_PrimaryJob
        '
        Me.COlItemPay_PrimaryJob.HeaderText = "Primary Job"
        Me.COlItemPay_PrimaryJob.Name = "COlItemPay_PrimaryJob"
        Me.COlItemPay_PrimaryJob.Visible = False
        Me.COlItemPay_PrimaryJob.Width = 86
        '
        'COlItemPay_PrimaryAkun
        '
        Me.COlItemPay_PrimaryAkun.HeaderText = "PrimaryAkun"
        Me.COlItemPay_PrimaryAkun.Name = "COlItemPay_PrimaryAkun"
        Me.COlItemPay_PrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.COlItemPay_PrimaryAkun.Visible = False
        Me.COlItemPay_PrimaryAkun.Width = 72
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvItem)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(861, 321)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Account Detail"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle15.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle15
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColDescAccount, Me.ColDebet, Me.ColKredit, Me.ColExchangeRate, Me.ColOriginalCurrency, Me.ColDebet_OriginalCurr, Me.ColCredit_OriginalCurrency, Me.ColRemarks, Me.ColDept, Me.ColJob, Me.ColPrimaryDept, Me.ColPrimaryJob, Me.ColPrimaryAkun, Me.ColPrimaryCurrency})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 3)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.Size = New System.Drawing.Size(855, 315)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 1
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
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle16.NullValue = "0"
        Me.ColDebet.DefaultCellStyle = DataGridViewCellStyle16
        Me.ColDebet.HeaderText = "Debet"
        Me.ColDebet.MaxInputLength = 25
        Me.ColDebet.Name = "ColDebet"
        Me.ColDebet.ReadOnly = True
        Me.ColDebet.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet.Width = 42
        '
        'ColKredit
        '
        DataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle17.NullValue = "0"
        Me.ColKredit.DefaultCellStyle = DataGridViewCellStyle17
        Me.ColKredit.HeaderText = "Credit"
        Me.ColKredit.MaxInputLength = 25
        Me.ColKredit.Name = "ColKredit"
        Me.ColKredit.ReadOnly = True
        Me.ColKredit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColKredit.Width = 40
        '
        'ColExchangeRate
        '
        DataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle18.NullValue = "1"
        Me.ColExchangeRate.DefaultCellStyle = DataGridViewCellStyle18
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
        DataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle19.NullValue = "0"
        Me.ColDebet_OriginalCurr.DefaultCellStyle = DataGridViewCellStyle19
        Me.ColDebet_OriginalCurr.HeaderText = "Debet"
        Me.ColDebet_OriginalCurr.Name = "ColDebet_OriginalCurr"
        Me.ColDebet_OriginalCurr.ReadOnly = True
        Me.ColDebet_OriginalCurr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet_OriginalCurr.Visible = False
        Me.ColDebet_OriginalCurr.Width = 42
        '
        'ColCredit_OriginalCurrency
        '
        DataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle20.NullValue = "0"
        Me.ColCredit_OriginalCurrency.DefaultCellStyle = DataGridViewCellStyle20
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
        Me.ColDept.Width = 93
        '
        'ColJob
        '
        Me.ColJob.HeaderText = "Job"
        Me.ColJob.Name = "ColJob"
        Me.ColJob.ReadOnly = True
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
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 537)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(886, 22)
        Me.StatusStrip1.TabIndex = 48
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
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(638, 502)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(462, 502)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(374, 502)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(198, 502)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(286, 502)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(550, 502)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'FrmTransaksiJournal_CashBank
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(886, 559)
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
        Me.Name = "FrmTransaksiJournal_CashBank"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cash And Bank"
        Me.gboDetail.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvPayment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnViewJurnal As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryJournal As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtRefNo As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
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
    Friend WithEvents cboType As System.Windows.Forms.ComboBox
    Friend WithEvents cboAccount As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryAccount As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Public WithEvents dgvPayment As System.Windows.Forms.DataGridView
    Friend WithEvents TxtAmountOriginal As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblCurrency As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtExchangeRate As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
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
    Friend WithEvents COlItemPay_Account As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COlItemPay_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemPay_Curr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COlItemPay_AmountOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemPay_ExRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COlItemPay_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COlItemPay_DeptCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COlItemPay_JobCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemPay_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COlItemPay_PrimaryDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COlItemPay_PrimaryJob As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COlItemPay_PrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtReferenceParty As System.Windows.Forms.TextBox
    Friend WithEvents LblReferenceParty As System.Windows.Forms.Label
End Class
