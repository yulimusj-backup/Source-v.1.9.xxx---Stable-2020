<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiCreditNote
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
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.groupDetail = New System.Windows.Forms.GroupBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtReferred = New System.Windows.Forms.TextBox
        Me.txtSupplierInvoiceNo = New System.Windows.Forms.TextBox
        Me.txtBalanceReff = New System.Windows.Forms.TextBox
        Me.lblSuppInvoiceNo = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtExpenseAccDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryExpenseAcc = New System.Windows.Forms.TextBox
        Me.TxtReferedInvoice = New System.Windows.Forms.TextBox
        Me.LblRefered = New System.Windows.Forms.Label
        Me.FindExpenseAccount = New System.Windows.Forms.Button
        Me.BtnFindReferedInvoice = New System.Windows.Forms.Button
        Me.LblNote = New System.Windows.Forms.Label
        Me.cboCurrencyReff = New System.Windows.Forms.ComboBox
        Me.TxtExpenseAcc = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtCustomerDebitNote = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.txtPrimaryJournal = New System.Windows.Forms.TextBox
        Me.txtCustomerAddIdentifier = New System.Windows.Forms.TextBox
        Me.CboCustomerSuuplier = New System.Windows.Forms.ComboBox
        Me.LinkAddress = New System.Windows.Forms.LinkLabel
        Me.CboAddress = New System.Windows.Forms.ComboBox
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.DtpDueDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.TxtExchangeRate = New System.Windows.Forms.TextBox
        Me.LblExchangeRate = New System.Windows.Forms.Label
        Me.cboCurrency = New System.Windows.Forms.ComboBox
        Me.LblARAmount = New System.Windows.Forms.Label
        Me.TxtBalance = New System.Windows.Forms.TextBox
        Me.CboArAccount = New System.Windows.Forms.ComboBox
        Me.TxtPrimaryArAccount = New System.Windows.Forms.TextBox
        Me.LblAcc = New System.Windows.Forms.Label
        Me.TxtPrimaryReferedInvoice = New System.Windows.Forms.TextBox
        Me.TxtContactPerson = New System.Windows.Forms.TextBox
        Me.txtNama = New System.Windows.Forms.TextBox
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.txtPrimaryCustSupp = New System.Windows.Forms.TextBox
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.btnFindCustSupp = New System.Windows.Forms.Button
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtNoTransaksi = New System.Windows.Forms.TextBox
        Me.cboPrimaryReffAcc = New System.Windows.Forms.ComboBox
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tspBar1 = New System.Windows.Forms.ToolStripProgressBar
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtNoJurnal = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtDecJurnal = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.dtpTransaksi = New System.Windows.Forms.DateTimePicker
        Me.TxtUnbalance = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
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
        Me.ColDefPrec = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.groupDetail.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(386, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'groupDetail
        '
        Me.groupDetail.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.groupDetail.Controls.Add(Me.Label8)
        Me.groupDetail.Controls.Add(Me.txtReferred)
        Me.groupDetail.Controls.Add(Me.txtSupplierInvoiceNo)
        Me.groupDetail.Controls.Add(Me.txtBalanceReff)
        Me.groupDetail.Controls.Add(Me.lblSuppInvoiceNo)
        Me.groupDetail.Controls.Add(Me.Label5)
        Me.groupDetail.Controls.Add(Me.TxtExpenseAccDesc)
        Me.groupDetail.Controls.Add(Me.TxtPrimaryExpenseAcc)
        Me.groupDetail.Controls.Add(Me.TxtReferedInvoice)
        Me.groupDetail.Controls.Add(Me.LblRefered)
        Me.groupDetail.Controls.Add(Me.FindExpenseAccount)
        Me.groupDetail.Controls.Add(Me.BtnFindReferedInvoice)
        Me.groupDetail.Controls.Add(Me.LblNote)
        Me.groupDetail.Controls.Add(Me.cboCurrencyReff)
        Me.groupDetail.Controls.Add(Me.TxtExpenseAcc)
        Me.groupDetail.Controls.Add(Me.Label6)
        Me.groupDetail.Location = New System.Drawing.Point(99, 31)
        Me.groupDetail.Name = "groupDetail"
        Me.groupDetail.Size = New System.Drawing.Size(11, 18)
        Me.groupDetail.TabIndex = 0
        Me.groupDetail.TabStop = False
        Me.groupDetail.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(10, 34)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(87, 13)
        Me.Label8.TabIndex = 54
        Me.Label8.Text = "Referred Amount"
        '
        'txtReferred
        '
        Me.txtReferred.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReferred.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtReferred.Location = New System.Drawing.Point(6, 6)
        Me.txtReferred.MaxLength = 50
        Me.txtReferred.Name = "txtReferred"
        Me.txtReferred.Size = New System.Drawing.Size(1, 20)
        Me.txtReferred.TabIndex = 20
        '
        'txtSupplierInvoiceNo
        '
        Me.txtSupplierInvoiceNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSupplierInvoiceNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierInvoiceNo.Location = New System.Drawing.Point(91, 163)
        Me.txtSupplierInvoiceNo.MaxLength = 250
        Me.txtSupplierInvoiceNo.Name = "txtSupplierInvoiceNo"
        Me.txtSupplierInvoiceNo.ReadOnly = True
        Me.txtSupplierInvoiceNo.Size = New System.Drawing.Size(1, 20)
        Me.txtSupplierInvoiceNo.TabIndex = 45
        Me.txtSupplierInvoiceNo.Visible = False
        '
        'txtBalanceReff
        '
        Me.txtBalanceReff.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBalanceReff.Location = New System.Drawing.Point(91, 205)
        Me.txtBalanceReff.MaxLength = 21
        Me.txtBalanceReff.Name = "txtBalanceReff"
        Me.txtBalanceReff.Size = New System.Drawing.Size(1, 20)
        Me.txtBalanceReff.TabIndex = 36
        Me.txtBalanceReff.Text = "0"
        Me.txtBalanceReff.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblSuppInvoiceNo
        '
        Me.lblSuppInvoiceNo.AutoSize = True
        Me.lblSuppInvoiceNo.Location = New System.Drawing.Point(8, 162)
        Me.lblSuppInvoiceNo.Name = "lblSuppInvoiceNo"
        Me.lblSuppInvoiceNo.Size = New System.Drawing.Size(90, 13)
        Me.lblSuppInvoiceNo.TabIndex = 44
        Me.lblSuppInvoiceNo.Text = "Supp Invoice No."
        Me.lblSuppInvoiceNo.Visible = False
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(10, -150)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(87, 13)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "Referred Amount"
        '
        'TxtExpenseAccDesc
        '
        Me.TxtExpenseAccDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtExpenseAccDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtExpenseAccDesc.Location = New System.Drawing.Point(168, 74)
        Me.TxtExpenseAccDesc.MaxLength = 250
        Me.TxtExpenseAccDesc.Name = "TxtExpenseAccDesc"
        Me.TxtExpenseAccDesc.ReadOnly = True
        Me.TxtExpenseAccDesc.Size = New System.Drawing.Size(1, 20)
        Me.TxtExpenseAccDesc.TabIndex = 33
        Me.TxtExpenseAccDesc.TabStop = False
        '
        'TxtPrimaryExpenseAcc
        '
        Me.TxtPrimaryExpenseAcc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryExpenseAcc.Location = New System.Drawing.Point(82, -111)
        Me.TxtPrimaryExpenseAcc.Multiline = True
        Me.TxtPrimaryExpenseAcc.Name = "TxtPrimaryExpenseAcc"
        Me.TxtPrimaryExpenseAcc.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryExpenseAcc.TabIndex = 30
        Me.TxtPrimaryExpenseAcc.TabStop = False
        Me.TxtPrimaryExpenseAcc.Text = "-1"
        Me.TxtPrimaryExpenseAcc.Visible = False
        '
        'TxtReferedInvoice
        '
        Me.TxtReferedInvoice.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtReferedInvoice.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtReferedInvoice.Location = New System.Drawing.Point(71, -89)
        Me.TxtReferedInvoice.MaxLength = 250
        Me.TxtReferedInvoice.Name = "TxtReferedInvoice"
        Me.TxtReferedInvoice.Size = New System.Drawing.Size(92, 20)
        Me.TxtReferedInvoice.TabIndex = 18
        '
        'LblRefered
        '
        Me.LblRefered.AutoSize = True
        Me.LblRefered.Location = New System.Drawing.Point(22, 94)
        Me.LblRefered.Name = "LblRefered"
        Me.LblRefered.Size = New System.Drawing.Size(86, 13)
        Me.LblRefered.TabIndex = 16
        Me.LblRefered.Text = "Referred Invoice"
        '
        'FindExpenseAccount
        '
        Me.FindExpenseAccount.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FindExpenseAccount.Location = New System.Drawing.Point(-155, 184)
        Me.FindExpenseAccount.Name = "FindExpenseAccount"
        Me.FindExpenseAccount.Size = New System.Drawing.Size(39, 20)
        Me.FindExpenseAccount.TabIndex = 32
        Me.FindExpenseAccount.TabStop = False
        Me.FindExpenseAccount.Text = "...."
        Me.FindExpenseAccount.UseVisualStyleBackColor = True
        '
        'BtnFindReferedInvoice
        '
        Me.BtnFindReferedInvoice.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindReferedInvoice.Location = New System.Drawing.Point(-196, 140)
        Me.BtnFindReferedInvoice.Name = "BtnFindReferedInvoice"
        Me.BtnFindReferedInvoice.Size = New System.Drawing.Size(39, 20)
        Me.BtnFindReferedInvoice.TabIndex = 19
        Me.BtnFindReferedInvoice.TabStop = False
        Me.BtnFindReferedInvoice.Text = "...."
        Me.BtnFindReferedInvoice.UseVisualStyleBackColor = True
        '
        'LblNote
        '
        Me.LblNote.AutoSize = True
        Me.LblNote.Location = New System.Drawing.Point(111, 57)
        Me.LblNote.Name = "LblNote"
        Me.LblNote.Size = New System.Drawing.Size(112, 13)
        Me.LblNote.TabIndex = 14
        Me.LblNote.Text = "Customer's Debit Note"
        '
        'cboCurrencyReff
        '
        Me.cboCurrencyReff.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboCurrencyReff.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurrencyReff.FormattingEnabled = True
        Me.cboCurrencyReff.Location = New System.Drawing.Point(71, -63)
        Me.cboCurrencyReff.Name = "cboCurrencyReff"
        Me.cboCurrencyReff.Size = New System.Drawing.Size(57, 21)
        Me.cboCurrencyReff.TabIndex = 35
        '
        'TxtExpenseAcc
        '
        Me.TxtExpenseAcc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtExpenseAcc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtExpenseAcc.Location = New System.Drawing.Point(91, 185)
        Me.TxtExpenseAcc.MaxLength = 50
        Me.TxtExpenseAcc.Name = "TxtExpenseAcc"
        Me.TxtExpenseAcc.Size = New System.Drawing.Size(1, 20)
        Me.TxtExpenseAcc.TabIndex = 31
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(4, 189)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(91, 13)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "Referred Account"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 224)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 13)
        Me.Label1.TabIndex = 39
        Me.Label1.Text = "Reason for claim"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 132)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Contact Person"
        '
        'TxtCustomerDebitNote
        '
        Me.TxtCustomerDebitNote.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerDebitNote.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerDebitNote.Location = New System.Drawing.Point(127, 177)
        Me.TxtCustomerDebitNote.MaxLength = 100
        Me.TxtCustomerDebitNote.Name = "TxtCustomerDebitNote"
        Me.TxtCustomerDebitNote.Size = New System.Drawing.Size(259, 20)
        Me.TxtCustomerDebitNote.TabIndex = 11
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(383, 352)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 26)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(276, 352)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(88, 26)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'txtPrimaryJournal
        '
        Me.txtPrimaryJournal.Location = New System.Drawing.Point(7, 261)
        Me.txtPrimaryJournal.Name = "txtPrimaryJournal"
        Me.txtPrimaryJournal.Size = New System.Drawing.Size(37, 20)
        Me.txtPrimaryJournal.TabIndex = 41
        Me.txtPrimaryJournal.TabStop = False
        Me.txtPrimaryJournal.Text = "-1"
        Me.txtPrimaryJournal.Visible = False
        '
        'txtCustomerAddIdentifier
        '
        Me.txtCustomerAddIdentifier.Location = New System.Drawing.Point(6, 285)
        Me.txtCustomerAddIdentifier.Multiline = True
        Me.txtCustomerAddIdentifier.Name = "txtCustomerAddIdentifier"
        Me.txtCustomerAddIdentifier.Size = New System.Drawing.Size(38, 20)
        Me.txtCustomerAddIdentifier.TabIndex = 9
        Me.txtCustomerAddIdentifier.TabStop = False
        Me.txtCustomerAddIdentifier.Visible = False
        '
        'CboCustomerSuuplier
        '
        Me.CboCustomerSuuplier.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCustomerSuuplier.FormattingEnabled = True
        Me.CboCustomerSuuplier.Items.AddRange(New Object() {"Customer", "Supplier"})
        Me.CboCustomerSuuplier.Location = New System.Drawing.Point(3, 79)
        Me.CboCustomerSuuplier.Name = "CboCustomerSuuplier"
        Me.CboCustomerSuuplier.Size = New System.Drawing.Size(84, 21)
        Me.CboCustomerSuuplier.TabIndex = 4
        '
        'LinkAddress
        '
        Me.LinkAddress.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkAddress.AutoSize = True
        Me.LinkAddress.Location = New System.Drawing.Point(338, 154)
        Me.LinkAddress.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LinkAddress.Name = "LinkAddress"
        Me.LinkAddress.Size = New System.Drawing.Size(45, 13)
        Me.LinkAddress.TabIndex = 10
        Me.LinkAddress.TabStop = True
        Me.LinkAddress.Text = "Address"
        '
        'CboAddress
        '
        Me.CboAddress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboAddress.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAddress.FormattingEnabled = True
        Me.CboAddress.ItemHeight = 13
        Me.CboAddress.Location = New System.Drawing.Point(127, 152)
        Me.CboAddress.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.CboAddress.Name = "CboAddress"
        Me.CboAddress.Size = New System.Drawing.Size(186, 21)
        Me.CboAddress.TabIndex = 9
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.Location = New System.Drawing.Point(127, 222)
        Me.txtRemarks.MaxLength = 250
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(259, 88)
        Me.txtRemarks.TabIndex = 13
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(2, 203)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(53, 13)
        Me.Label10.TabIndex = 37
        Me.Label10.Text = "Due Date"
        '
        'DtpDueDate
        '
        Me.DtpDueDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDueDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDueDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDueDate.Location = New System.Drawing.Point(127, 199)
        Me.DtpDueDate.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DtpDueDate.Name = "DtpDueDate"
        Me.DtpDueDate.Size = New System.Drawing.Size(259, 20)
        Me.DtpDueDate.TabIndex = 12
        '
        'TxtPrimaryCurrency
        '
        Me.TxtPrimaryCurrency.Location = New System.Drawing.Point(50, 235)
        Me.TxtPrimaryCurrency.Name = "TxtPrimaryCurrency"
        Me.TxtPrimaryCurrency.Size = New System.Drawing.Size(37, 20)
        Me.TxtPrimaryCurrency.TabIndex = 24
        Me.TxtPrimaryCurrency.TabStop = False
        Me.TxtPrimaryCurrency.Text = "-1"
        Me.TxtPrimaryCurrency.Visible = False
        '
        'TxtExchangeRate
        '
        Me.TxtExchangeRate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtExchangeRate.Location = New System.Drawing.Point(118, 51)
        Me.TxtExchangeRate.MaxLength = 21
        Me.TxtExchangeRate.Name = "TxtExchangeRate"
        Me.TxtExchangeRate.Size = New System.Drawing.Size(247, 20)
        Me.TxtExchangeRate.TabIndex = 2
        Me.TxtExchangeRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblExchangeRate
        '
        Me.LblExchangeRate.AutoSize = True
        Me.LblExchangeRate.Location = New System.Drawing.Point(10, 54)
        Me.LblExchangeRate.Name = "LblExchangeRate"
        Me.LblExchangeRate.Size = New System.Drawing.Size(81, 13)
        Me.LblExchangeRate.TabIndex = 27
        Me.LblExchangeRate.Text = "Exchange Rate"
        '
        'cboCurrency
        '
        Me.cboCurrency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurrency.FormattingEnabled = True
        Me.cboCurrency.Location = New System.Drawing.Point(118, 27)
        Me.cboCurrency.Name = "cboCurrency"
        Me.cboCurrency.Size = New System.Drawing.Size(77, 21)
        Me.cboCurrency.TabIndex = 2
        '
        'LblARAmount
        '
        Me.LblARAmount.AutoSize = True
        Me.LblARAmount.Location = New System.Drawing.Point(10, 29)
        Me.LblARAmount.Name = "LblARAmount"
        Me.LblARAmount.Size = New System.Drawing.Size(61, 13)
        Me.LblARAmount.TabIndex = 23
        Me.LblARAmount.Text = "AR Amount"
        '
        'TxtBalance
        '
        Me.TxtBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBalance.Location = New System.Drawing.Point(200, 27)
        Me.TxtBalance.MaxLength = 21
        Me.TxtBalance.Name = "TxtBalance"
        Me.TxtBalance.Size = New System.Drawing.Size(165, 20)
        Me.TxtBalance.TabIndex = 1
        Me.TxtBalance.Text = "0"
        Me.TxtBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'CboArAccount
        '
        Me.CboArAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboArAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboArAccount.FormattingEnabled = True
        Me.CboArAccount.Location = New System.Drawing.Point(118, 3)
        Me.CboArAccount.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.CboArAccount.Name = "CboArAccount"
        Me.CboArAccount.Size = New System.Drawing.Size(247, 21)
        Me.CboArAccount.TabIndex = 0
        '
        'TxtPrimaryArAccount
        '
        Me.TxtPrimaryArAccount.Location = New System.Drawing.Point(6, 235)
        Me.TxtPrimaryArAccount.Multiline = True
        Me.TxtPrimaryArAccount.Name = "TxtPrimaryArAccount"
        Me.TxtPrimaryArAccount.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryArAccount.TabIndex = 21
        Me.TxtPrimaryArAccount.TabStop = False
        Me.TxtPrimaryArAccount.Text = "-1"
        Me.TxtPrimaryArAccount.Visible = False
        '
        'LblAcc
        '
        Me.LblAcc.AutoSize = True
        Me.LblAcc.Location = New System.Drawing.Point(10, 6)
        Me.LblAcc.Name = "LblAcc"
        Me.LblAcc.Size = New System.Drawing.Size(65, 13)
        Me.LblAcc.TabIndex = 20
        Me.LblAcc.Text = "AR Account"
        '
        'TxtPrimaryReferedInvoice
        '
        Me.TxtPrimaryReferedInvoice.Location = New System.Drawing.Point(6, 209)
        Me.TxtPrimaryReferedInvoice.Multiline = True
        Me.TxtPrimaryReferedInvoice.Name = "TxtPrimaryReferedInvoice"
        Me.TxtPrimaryReferedInvoice.Size = New System.Drawing.Size(38, 20)
        Me.TxtPrimaryReferedInvoice.TabIndex = 17
        Me.TxtPrimaryReferedInvoice.TabStop = False
        Me.TxtPrimaryReferedInvoice.Text = "-1"
        Me.TxtPrimaryReferedInvoice.Visible = False
        '
        'TxtContactPerson
        '
        Me.TxtContactPerson.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtContactPerson.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtContactPerson.Location = New System.Drawing.Point(127, 128)
        Me.TxtContactPerson.MaxLength = 100
        Me.TxtContactPerson.Name = "TxtContactPerson"
        Me.TxtContactPerson.Size = New System.Drawing.Size(259, 20)
        Me.TxtContactPerson.TabIndex = 8
        '
        'txtNama
        '
        Me.txtNama.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNama.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtNama.Location = New System.Drawing.Point(127, 104)
        Me.txtNama.MaxLength = 250
        Me.txtNama.Name = "txtNama"
        Me.txtNama.ReadOnly = True
        Me.txtNama.Size = New System.Drawing.Size(259, 20)
        Me.txtNama.TabIndex = 7
        Me.txtNama.TabStop = False
        '
        'DtpDate
        '
        Me.DtpDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(280, 3)
        Me.DtpDate.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.Size = New System.Drawing.Size(107, 20)
        Me.DtpDate.TabIndex = 1
        '
        'txtPrimaryCustSupp
        '
        Me.txtPrimaryCustSupp.Location = New System.Drawing.Point(49, 209)
        Me.txtPrimaryCustSupp.Multiline = True
        Me.txtPrimaryCustSupp.Name = "txtPrimaryCustSupp"
        Me.txtPrimaryCustSupp.Size = New System.Drawing.Size(38, 20)
        Me.txtPrimaryCustSupp.TabIndex = 11
        Me.txtPrimaryCustSupp.Text = "-1"
        Me.txtPrimaryCustSupp.Visible = False
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(52, 287)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(38, 20)
        Me.txtPrimary.TabIndex = 0
        Me.txtPrimary.TabStop = False
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'btnFindCustSupp
        '
        Me.btnFindCustSupp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCustSupp.Location = New System.Drawing.Point(348, 78)
        Me.btnFindCustSupp.Name = "btnFindCustSupp"
        Me.btnFindCustSupp.Size = New System.Drawing.Size(37, 20)
        Me.btnFindCustSupp.TabIndex = 6
        Me.btnFindCustSupp.TabStop = False
        Me.btnFindCustSupp.Text = "...."
        Me.btnFindCustSupp.UseVisualStyleBackColor = True
        '
        'txtCode
        '
        Me.txtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCode.Location = New System.Drawing.Point(127, 78)
        Me.txtCode.MaxLength = 150
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(215, 20)
        Me.txtCode.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Document No"
        '
        'txtNoTransaksi
        '
        Me.txtNoTransaksi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNoTransaksi.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtNoTransaksi.Location = New System.Drawing.Point(127, 3)
        Me.txtNoTransaksi.MaxLength = 100
        Me.txtNoTransaksi.Name = "txtNoTransaksi"
        Me.txtNoTransaksi.ReadOnly = True
        Me.txtNoTransaksi.Size = New System.Drawing.Size(114, 20)
        Me.txtNoTransaksi.TabIndex = 0
        Me.txtNoTransaksi.TabStop = False
        '
        'cboPrimaryReffAcc
        '
        Me.cboPrimaryReffAcc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryReffAcc.Enabled = False
        Me.cboPrimaryReffAcc.FormattingEnabled = True
        Me.cboPrimaryReffAcc.Location = New System.Drawing.Point(3, 79)
        Me.cboPrimaryReffAcc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboPrimaryReffAcc.Name = "cboPrimaryReffAcc"
        Me.cboPrimaryReffAcc.Size = New System.Drawing.Size(84, 21)
        Me.cboPrimaryReffAcc.TabIndex = 4
        Me.cboPrimaryReffAcc.TabStop = False
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 440)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(787, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 8
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(386, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(386, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(386, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate, Me.tspBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 462)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(787, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 9
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tspBar1
        '
        Me.tspBar1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tspBar1.Name = "tspBar1"
        Me.tspBar1.Size = New System.Drawing.Size(100, 16)
        Me.tspBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.tspBar1.Visible = False
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(313, 404)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(68, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(239, 404)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(68, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(165, 404)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(68, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(536, 404)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(68, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnDelete.Location = New System.Drawing.Point(388, 404)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(68, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(462, 404)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(68, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Location = New System.Drawing.Point(49, 259)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(22, 20)
        Me.txtReadOnly.TabIndex = 50
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(11, 5)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.gboDetail.Size = New System.Drawing.Size(765, 384)
        Me.gboDetail.TabIndex = 50
        Me.gboDetail.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(2, 15)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtNoTransaksi)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtNoJurnal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtDecJurnal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.groupDetail)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtReadOnly)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCustomerDebitNote)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryJournal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryCurrency)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindCustSupp)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryArAccount)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LinkAddress)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtRemarks)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCustomerAddIdentifier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpDueDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboAddress)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtNama)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryCustSupp)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryReferedInvoice)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtContactPerson)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboCustomerSuuplier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboPrimaryReffAcc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpTransaksi)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtUnbalance)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvItem)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtExchangeRate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblExchangeRate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.CboArAccount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblARAmount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboCurrency)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtBalance)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblAcc)
        Me.SplitContainer1.Size = New System.Drawing.Size(761, 318)
        Me.SplitContainer1.SplitterDistance = 390
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(246, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 59
        Me.Label4.Text = "Date"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(2, 54)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(92, 13)
        Me.Label12.TabIndex = 55
        Me.Label12.Text = "Decription Journal"
        '
        'txtNoJurnal
        '
        Me.txtNoJurnal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNoJurnal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtNoJurnal.Location = New System.Drawing.Point(127, 28)
        Me.txtNoJurnal.MaxLength = 100
        Me.txtNoJurnal.Name = "txtNoJurnal"
        Me.txtNoJurnal.Size = New System.Drawing.Size(259, 20)
        Me.txtNoJurnal.TabIndex = 2
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(2, 32)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(58, 13)
        Me.Label11.TabIndex = 53
        Me.Label11.Text = "Journal No"
        '
        'txtDecJurnal
        '
        Me.txtDecJurnal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDecJurnal.Location = New System.Drawing.Point(127, 51)
        Me.txtDecJurnal.MaxLength = 100
        Me.txtDecJurnal.Name = "txtDecJurnal"
        Me.txtDecJurnal.Size = New System.Drawing.Size(259, 20)
        Me.txtDecJurnal.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(2, 181)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(117, 13)
        Me.Label7.TabIndex = 51
        Me.Label7.Text = "Referred Document No"
        '
        'dtpTransaksi
        '
        Me.dtpTransaksi.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTransaksi.CustomFormat = "dd MMM yyyy"
        Me.dtpTransaksi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTransaksi.Location = New System.Drawing.Point(292, 27)
        Me.dtpTransaksi.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpTransaksi.Name = "dtpTransaksi"
        Me.dtpTransaksi.Size = New System.Drawing.Size(96, 20)
        Me.dtpTransaksi.TabIndex = 1
        Me.dtpTransaksi.TabStop = False
        Me.dtpTransaksi.Visible = False
        '
        'TxtUnbalance
        '
        Me.TxtUnbalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUnbalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUnbalance.Location = New System.Drawing.Point(118, 76)
        Me.TxtUnbalance.MaxLength = 50
        Me.TxtUnbalance.Name = "TxtUnbalance"
        Me.TxtUnbalance.ReadOnly = True
        Me.TxtUnbalance.Size = New System.Drawing.Size(247, 20)
        Me.TxtUnbalance.TabIndex = 3
        Me.TxtUnbalance.TabStop = False
        Me.TxtUnbalance.Text = "0"
        Me.TxtUnbalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(10, 78)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(59, 13)
        Me.Label9.TabIndex = 166
        Me.Label9.Text = "Unbalance"
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
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColDescAccount, Me.ColDebet, Me.ColKredit, Me.ColExchangeRate, Me.ColOriginalCurrency, Me.ColDebet_OriginalCurr, Me.ColCredit_OriginalCurrency, Me.ColRemarks, Me.ColDept, Me.ColJob, Me.ColPrimaryDept, Me.ColPrimaryJob, Me.ColPrimaryAkun, Me.ColPrimaryCurrency, Me.ColDefPrec})
        Me.dgvItem.Location = New System.Drawing.Point(11, 100)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(354, 216)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 4
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
        Me.ColExchangeRate.Width = 110
        '
        'ColOriginalCurrency
        '
        Me.ColOriginalCurrency.HeaderText = "Currency"
        Me.ColOriginalCurrency.Name = "ColOriginalCurrency"
        Me.ColOriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOriginalCurrency.Visible = False
        Me.ColOriginalCurrency.Width = 71
        '
        'ColDebet_OriginalCurr
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.NullValue = "0"
        Me.ColDebet_OriginalCurr.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColDebet_OriginalCurr.HeaderText = "Debet Original"
        Me.ColDebet_OriginalCurr.Name = "ColDebet_OriginalCurr"
        Me.ColDebet_OriginalCurr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDebet_OriginalCurr.Visible = False
        Me.ColDebet_OriginalCurr.Width = 105
        '
        'ColCredit_OriginalCurrency
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.NullValue = "0"
        Me.ColCredit_OriginalCurrency.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColCredit_OriginalCurrency.HeaderText = "Credit Original"
        Me.ColCredit_OriginalCurrency.Name = "ColCredit_OriginalCurrency"
        Me.ColCredit_OriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCredit_OriginalCurrency.Visible = False
        Me.ColCredit_OriginalCurrency.Width = 104
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
        Me.ColDept.Width = 115
        '
        'ColJob
        '
        Me.ColJob.HeaderText = "Job"
        Me.ColJob.Name = "ColJob"
        Me.ColJob.Visible = False
        Me.ColJob.Width = 56
        '
        'ColPrimaryDept
        '
        Me.ColPrimaryDept.HeaderText = "Primary Dept"
        Me.ColPrimaryDept.Name = "ColPrimaryDept"
        Me.ColPrimaryDept.Visible = False
        Me.ColPrimaryDept.Width = 115
        '
        'ColPrimaryJob
        '
        Me.ColPrimaryJob.HeaderText = "Primary Job"
        Me.ColPrimaryJob.Name = "ColPrimaryJob"
        Me.ColPrimaryJob.Visible = False
        Me.ColPrimaryJob.Width = 108
        '
        'ColPrimaryAkun
        '
        Me.ColPrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColPrimaryAkun.Name = "ColPrimaryAkun"
        Me.ColPrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryAkun.Visible = False
        Me.ColPrimaryAkun.Width = 94
        '
        'ColPrimaryCurrency
        '
        Me.ColPrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.Visible = False
        Me.ColPrimaryCurrency.Width = 142
        '
        'ColDefPrec
        '
        Me.ColDefPrec.HeaderText = "DefPrec"
        Me.ColDefPrec.Name = "ColDefPrec"
        Me.ColDefPrec.Visible = False
        Me.ColDefPrec.Width = 84
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'FrmTransaksiCreditNote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(787, 484)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimumSize = New System.Drawing.Size(800, 500)
        Me.Name = "FrmTransaksiCreditNote"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.groupDetail.ResumeLayout(False)
        Me.groupDetail.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents groupDetail As System.Windows.Forms.GroupBox
    Friend WithEvents txtPrimaryCustSupp As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCustSupp As System.Windows.Forms.Button
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNoTransaksi As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tspBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents txtNama As System.Windows.Forms.TextBox
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents CboArAccount As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryArAccount As System.Windows.Forms.TextBox
    Friend WithEvents LblAcc As System.Windows.Forms.Label
    Friend WithEvents TxtExpenseAccDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryExpenseAcc As System.Windows.Forms.TextBox
    Friend WithEvents FindExpenseAccount As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtExpenseAcc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryReferedInvoice As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindReferedInvoice As System.Windows.Forms.Button
    Friend WithEvents LblRefered As System.Windows.Forms.Label
    Friend WithEvents TxtReferedInvoice As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtContactPerson As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents TxtExchangeRate As System.Windows.Forms.TextBox
    Friend WithEvents LblExchangeRate As System.Windows.Forms.Label
    Friend WithEvents cboCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents LblARAmount As System.Windows.Forms.Label
    Friend WithEvents TxtBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents DtpDueDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblNote As System.Windows.Forms.Label
    Friend WithEvents TxtCustomerDebitNote As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents LinkAddress As System.Windows.Forms.LinkLabel
    Friend WithEvents CboAddress As System.Windows.Forms.ComboBox
    Friend WithEvents CboCustomerSuuplier As System.Windows.Forms.ComboBox
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerAddIdentifier As System.Windows.Forms.TextBox
    Friend WithEvents cboPrimaryReffAcc As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimaryJournal As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtBalanceReff As System.Windows.Forms.TextBox
    Friend WithEvents cboCurrencyReff As System.Windows.Forms.ComboBox
    Friend WithEvents lblSuppInvoiceNo As System.Windows.Forms.Label
    Friend WithEvents txtSupplierInvoiceNo As System.Windows.Forms.TextBox
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents txtReferred As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TxtUnbalance As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtDecJurnal As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtNoJurnal As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
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
    Friend WithEvents ColDefPrec As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dtpTransaksi As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
