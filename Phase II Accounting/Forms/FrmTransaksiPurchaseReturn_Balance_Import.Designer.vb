<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPurchaseReturn_Balance_Import
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
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPR_Number = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPR_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_PR_DueDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurchaseReturn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVAT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDutystamp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceived = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPR = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySupp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.DtpDueDate = New System.Windows.Forms.DateTimePicker
        Me.Label18 = New System.Windows.Forms.Label
        Me.TxtPrimaryInvoice = New System.Windows.Forms.TextBox
        Me.BtnFindINV = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtInvoiceNumber = New System.Windows.Forms.TextBox
        Me.TxtSuppName = New System.Windows.Forms.TextBox
        Me.dtpInvoiceDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtSuppCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.TxtReceived = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.TxtBalance = New System.Windows.Forms.TextBox
        Me.btnInput = New System.Windows.Forms.Button
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtDutystamp = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtVAT = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.TxtSales = New System.Windows.Forms.TextBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtNoTransaksi = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.TxtTotalSalesIDR = New System.Windows.Forms.TextBox
        Me.LblVATIdr = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.LblIDR = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtTotalVat = New System.Windows.Forms.TextBox
        Me.TxtTotalDutyStamp = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.TxtTotalBalance = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtGrandTotal = New System.Windows.Forms.TextBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.btnNew = New System.Windows.Forms.Button
        Me.gboDetail.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtPrimaryStockIn
        '
        Me.TxtPrimaryStockIn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryStockIn.Location = New System.Drawing.Point(54, 63)
        Me.TxtPrimaryStockIn.Name = "TxtPrimaryStockIn"
        Me.TxtPrimaryStockIn.Size = New System.Drawing.Size(20, 20)
        Me.TxtPrimaryStockIn.TabIndex = 47
        Me.TxtPrimaryStockIn.TabStop = False
        Me.TxtPrimaryStockIn.Text = "-1"
        Me.TxtPrimaryStockIn.Visible = False
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.GroupBox2)
        Me.gboDetail.Controls.Add(Me.GroupBox3)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(12, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(860, 584)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.dgvItem)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 324)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(848, 222)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Item"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPR_Number, Me.ColPR_Date, Me.Col_PR_DueDate, Me.ColSuppCode, Me.ColSuppName, Me.ColPurchaseReturn, Me.ColVAT, Me.ColDutystamp, Me.ColTotal, Me.ColReceived, Me.ColBalance, Me.ColRemarks, Me.ColPrimaryPR, Me.ColPrimarySupp})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 16)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.Size = New System.Drawing.Size(842, 203)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 0
        Me.dgvItem.TabStop = False
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColPR_Number
        '
        Me.ColPR_Number.HeaderText = "Purchase Return Number"
        Me.ColPR_Number.Name = "ColPR_Number"
        Me.ColPR_Number.Width = 152
        '
        'ColPR_Date
        '
        Me.ColPR_Date.HeaderText = "Purchase Return Date"
        Me.ColPR_Date.Name = "ColPR_Date"
        Me.ColPR_Date.Width = 138
        '
        'Col_PR_DueDate
        '
        Me.Col_PR_DueDate.HeaderText = "Purchase Return Due Date"
        Me.Col_PR_DueDate.Name = "Col_PR_DueDate"
        Me.Col_PR_DueDate.Width = 161
        '
        'ColSuppCode
        '
        Me.ColSuppCode.HeaderText = "Supplier Code"
        Me.ColSuppCode.Name = "ColSuppCode"
        Me.ColSuppCode.Width = 98
        '
        'ColSuppName
        '
        Me.ColSuppName.HeaderText = "Supplier Name"
        Me.ColSuppName.Name = "ColSuppName"
        Me.ColSuppName.Width = 101
        '
        'ColPurchaseReturn
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPurchaseReturn.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColPurchaseReturn.HeaderText = "Purchase Return"
        Me.ColPurchaseReturn.Name = "ColPurchaseReturn"
        Me.ColPurchaseReturn.Width = 112
        '
        'ColVAT
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColVAT.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColVAT.HeaderText = "VAT"
        Me.ColVAT.Name = "ColVAT"
        Me.ColVAT.Width = 53
        '
        'ColDutystamp
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDutystamp.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColDutystamp.HeaderText = "Dutystamp"
        Me.ColDutystamp.Name = "ColDutystamp"
        Me.ColDutystamp.Width = 82
        '
        'ColTotal
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotal.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColTotal.HeaderText = "Total"
        Me.ColTotal.Name = "ColTotal"
        Me.ColTotal.Width = 56
        '
        'ColReceived
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColReceived.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColReceived.HeaderText = "Received"
        Me.ColReceived.MaxInputLength = 18
        Me.ColReceived.Name = "ColReceived"
        Me.ColReceived.Width = 78
        '
        'ColBalance
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.Width = 71
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 150
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'ColPrimaryPR
        '
        Me.ColPrimaryPR.HeaderText = "Primary PR"
        Me.ColPrimaryPR.Name = "ColPrimaryPR"
        Me.ColPrimaryPR.Visible = False
        Me.ColPrimaryPR.Width = 84
        '
        'ColPrimarySupp
        '
        Me.ColPrimarySupp.HeaderText = "Primary Supp"
        Me.ColPrimarySupp.Name = "ColPrimarySupp"
        Me.ColPrimarySupp.Visible = False
        Me.ColPrimarySupp.Width = 94
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.SplitContainer3)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 119)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(848, 199)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.txtReadOnly)
        Me.SplitContainer3.Panel1.Controls.Add(Me.DtpDueDate)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label18)
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtPrimaryInvoice)
        Me.SplitContainer3.Panel1.Controls.Add(Me.BtnFindINV)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtInvoiceNumber)
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtSuppName)
        Me.SplitContainer3.Panel1.Controls.Add(Me.dtpInvoiceDate)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtSuppCode)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label4)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label19)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label22)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtReceived)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label23)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtBalance)
        Me.SplitContainer3.Panel2.Controls.Add(Me.btnInput)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label11)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label13)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label14)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtDutystamp)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtTotal)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label16)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtVAT)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Label17)
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtSales)
        Me.SplitContainer3.Size = New System.Drawing.Size(842, 180)
        Me.SplitContainer3.SplitterDistance = 415
        Me.SplitContainer3.TabIndex = 0
        Me.SplitContainer3.TabStop = False
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Location = New System.Drawing.Point(116, 109)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(23, 20)
        Me.txtReadOnly.TabIndex = 250
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'DtpDueDate
        '
        Me.DtpDueDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDueDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDueDate.Enabled = False
        Me.DtpDueDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDueDate.Location = New System.Drawing.Point(149, 52)
        Me.DtpDueDate.Name = "DtpDueDate"
        Me.DtpDueDate.Size = New System.Drawing.Size(263, 20)
        Me.DtpDueDate.TabIndex = 2
        Me.DtpDueDate.TabStop = False
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(3, 58)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(136, 13)
        Me.Label18.TabIndex = 171
        Me.Label18.Text = "Purchase Return Due Date"
        '
        'TxtPrimaryInvoice
        '
        Me.TxtPrimaryInvoice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryInvoice.Location = New System.Drawing.Point(72, 3)
        Me.TxtPrimaryInvoice.Name = "TxtPrimaryInvoice"
        Me.TxtPrimaryInvoice.Size = New System.Drawing.Size(20, 20)
        Me.TxtPrimaryInvoice.TabIndex = 48
        Me.TxtPrimaryInvoice.TabStop = False
        Me.TxtPrimaryInvoice.Text = "-1"
        Me.TxtPrimaryInvoice.Visible = False
        '
        'BtnFindINV
        '
        Me.BtnFindINV.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindINV.Location = New System.Drawing.Point(376, 1)
        Me.BtnFindINV.Name = "BtnFindINV"
        Me.BtnFindINV.Size = New System.Drawing.Size(36, 22)
        Me.BtnFindINV.TabIndex = 169
        Me.BtnFindINV.TabStop = False
        Me.BtnFindINV.Text = "Find"
        Me.BtnFindINV.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 113)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 13)
        Me.Label5.TabIndex = 51
        Me.Label5.Text = "Supplier Name"
        '
        'TxtInvoiceNumber
        '
        Me.TxtInvoiceNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtInvoiceNumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtInvoiceNumber.Location = New System.Drawing.Point(149, 3)
        Me.TxtInvoiceNumber.MaxLength = 50
        Me.TxtInvoiceNumber.Name = "TxtInvoiceNumber"
        Me.TxtInvoiceNumber.Size = New System.Drawing.Size(221, 20)
        Me.TxtInvoiceNumber.TabIndex = 0
        '
        'TxtSuppName
        '
        Me.TxtSuppName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSuppName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSuppName.Location = New System.Drawing.Point(149, 108)
        Me.TxtSuppName.MaxLength = 50
        Me.TxtSuppName.Name = "TxtSuppName"
        Me.TxtSuppName.ReadOnly = True
        Me.TxtSuppName.Size = New System.Drawing.Size(263, 20)
        Me.TxtSuppName.TabIndex = 4
        Me.TxtSuppName.TabStop = False
        '
        'dtpInvoiceDate
        '
        Me.dtpInvoiceDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpInvoiceDate.CustomFormat = "dd MMM yyyy"
        Me.dtpInvoiceDate.Enabled = False
        Me.dtpInvoiceDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpInvoiceDate.Location = New System.Drawing.Point(149, 26)
        Me.dtpInvoiceDate.Name = "dtpInvoiceDate"
        Me.dtpInvoiceDate.Size = New System.Drawing.Size(263, 20)
        Me.dtpInvoiceDate.TabIndex = 1
        Me.dtpInvoiceDate.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 87)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 49
        Me.Label1.Text = "Supplier Code"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(113, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Purchase Return Date"
        '
        'TxtSuppCode
        '
        Me.TxtSuppCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSuppCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSuppCode.Location = New System.Drawing.Point(149, 82)
        Me.TxtSuppCode.MaxLength = 50
        Me.TxtSuppCode.Name = "TxtSuppCode"
        Me.TxtSuppCode.ReadOnly = True
        Me.TxtSuppCode.Size = New System.Drawing.Size(263, 20)
        Me.TxtSuppCode.TabIndex = 3
        Me.TxtSuppCode.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(127, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Purchase Return Number"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(118, 135)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(26, 13)
        Me.Label9.TabIndex = 176
        Me.Label9.Text = "IDR"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(118, 109)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(26, 13)
        Me.Label19.TabIndex = 175
        Me.Label19.Text = "IDR"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(4, 111)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(53, 13)
        Me.Label22.TabIndex = 174
        Me.Label22.Text = "Received"
        '
        'TxtReceived
        '
        Me.TxtReceived.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReceived.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtReceived.Location = New System.Drawing.Point(150, 106)
        Me.TxtReceived.MaxLength = 18
        Me.TxtReceived.Name = "TxtReceived"
        Me.TxtReceived.Size = New System.Drawing.Size(268, 20)
        Me.TxtReceived.TabIndex = 4
        Me.TxtReceived.Text = "0"
        Me.TxtReceived.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(4, 137)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(46, 13)
        Me.Label23.TabIndex = 172
        Me.Label23.Text = "Balance"
        '
        'TxtBalance
        '
        Me.TxtBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBalance.Location = New System.Drawing.Point(150, 132)
        Me.TxtBalance.MaxLength = 50
        Me.TxtBalance.Name = "TxtBalance"
        Me.TxtBalance.ReadOnly = True
        Me.TxtBalance.Size = New System.Drawing.Size(268, 20)
        Me.TxtBalance.TabIndex = 5
        Me.TxtBalance.TabStop = False
        Me.TxtBalance.Text = "0"
        Me.TxtBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnInput
        '
        Me.btnInput.Location = New System.Drawing.Point(149, 158)
        Me.btnInput.Name = "btnInput"
        Me.btnInput.Size = New System.Drawing.Size(69, 22)
        Me.btnInput.TabIndex = 6
        Me.btnInput.Text = "Input"
        Me.btnInput.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(118, 83)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(26, 13)
        Me.Label10.TabIndex = 85
        Me.Label10.Text = "IDR"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(118, 57)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(26, 13)
        Me.Label11.TabIndex = 84
        Me.Label11.Text = "IDR"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(118, 32)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(26, 13)
        Me.Label12.TabIndex = 83
        Me.Label12.Text = "IDR"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(118, 6)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(26, 13)
        Me.Label13.TabIndex = 82
        Me.Label13.Text = "IDR"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(4, 59)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(62, 13)
        Me.Label14.TabIndex = 81
        Me.Label14.Text = "Duty Stamp"
        '
        'TxtDutystamp
        '
        Me.TxtDutystamp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDutystamp.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDutystamp.Location = New System.Drawing.Point(150, 54)
        Me.TxtDutystamp.MaxLength = 50
        Me.TxtDutystamp.Name = "TxtDutystamp"
        Me.TxtDutystamp.ReadOnly = True
        Me.TxtDutystamp.Size = New System.Drawing.Size(268, 20)
        Me.TxtDutystamp.TabIndex = 2
        Me.TxtDutystamp.TabStop = False
        Me.TxtDutystamp.Text = "0"
        Me.TxtDutystamp.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(4, 85)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(31, 13)
        Me.Label15.TabIndex = 79
        Me.Label15.Text = "Total"
        '
        'TxtTotal
        '
        Me.TxtTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotal.Location = New System.Drawing.Point(150, 80)
        Me.TxtTotal.MaxLength = 50
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(268, 20)
        Me.TxtTotal.TabIndex = 3
        Me.TxtTotal.TabStop = False
        Me.TxtTotal.Text = "0"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(4, 34)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(28, 13)
        Me.Label16.TabIndex = 77
        Me.Label16.Text = "VAT"
        '
        'TxtVAT
        '
        Me.TxtVAT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtVAT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtVAT.Location = New System.Drawing.Point(150, 29)
        Me.TxtVAT.MaxLength = 50
        Me.TxtVAT.Name = "TxtVAT"
        Me.TxtVAT.ReadOnly = True
        Me.TxtVAT.Size = New System.Drawing.Size(268, 20)
        Me.TxtVAT.TabIndex = 1
        Me.TxtVAT.TabStop = False
        Me.TxtVAT.Text = "0"
        Me.TxtVAT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(4, 8)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(87, 13)
        Me.Label17.TabIndex = 75
        Me.Label17.Text = "Purchase Return"
        '
        'TxtSales
        '
        Me.TxtSales.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSales.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSales.Location = New System.Drawing.Point(150, 3)
        Me.TxtSales.MaxLength = 50
        Me.TxtSales.Name = "TxtSales"
        Me.TxtSales.ReadOnly = True
        Me.TxtSales.Size = New System.Drawing.Size(268, 20)
        Me.TxtSales.TabIndex = 0
        Me.TxtSales.TabStop = False
        Me.TxtSales.Text = "0"
        Me.TxtSales.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtNoTransaksi)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label26)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryStockIn)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtTotalSalesIDR)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblVATIdr)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label20)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblIDR)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtTotalVat)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtTotalDutyStamp)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label27)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label24)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label25)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTotalBalance)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label21)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtGrandTotal)
        Me.SplitContainer1.Size = New System.Drawing.Size(843, 121)
        Me.SplitContainer1.SplitterDistance = 421
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'TxtNoTransaksi
        '
        Me.TxtNoTransaksi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNoTransaksi.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNoTransaksi.Location = New System.Drawing.Point(152, 11)
        Me.TxtNoTransaksi.MaxLength = 50
        Me.TxtNoTransaksi.Name = "TxtNoTransaksi"
        Me.TxtNoTransaksi.ReadOnly = True
        Me.TxtNoTransaksi.Size = New System.Drawing.Size(263, 20)
        Me.TxtNoTransaksi.TabIndex = 0
        Me.TxtNoTransaksi.TabStop = False
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(0, 16)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(64, 13)
        Me.Label26.TabIndex = 74
        Me.Label26.Text = "Ref Number"
        '
        'TxtTotalSalesIDR
        '
        Me.TxtTotalSalesIDR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalSalesIDR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotalSalesIDR.Location = New System.Drawing.Point(152, 35)
        Me.TxtTotalSalesIDR.MaxLength = 50
        Me.TxtTotalSalesIDR.Name = "TxtTotalSalesIDR"
        Me.TxtTotalSalesIDR.ReadOnly = True
        Me.TxtTotalSalesIDR.Size = New System.Drawing.Size(263, 20)
        Me.TxtTotalSalesIDR.TabIndex = 1
        Me.TxtTotalSalesIDR.TabStop = False
        Me.TxtTotalSalesIDR.Text = "0"
        Me.TxtTotalSalesIDR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblVATIdr
        '
        Me.LblVATIdr.AutoSize = True
        Me.LblVATIdr.Location = New System.Drawing.Point(120, 64)
        Me.LblVATIdr.Name = "LblVATIdr"
        Me.LblVATIdr.Size = New System.Drawing.Size(26, 13)
        Me.LblVATIdr.TabIndex = 71
        Me.LblVATIdr.Text = "IDR"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(0, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(114, 13)
        Me.Label6.TabIndex = 53
        Me.Label6.Text = "Total Purchase Return"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(120, 88)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(26, 13)
        Me.Label20.TabIndex = 72
        Me.Label20.Text = "IDR"
        '
        'LblIDR
        '
        Me.LblIDR.AutoSize = True
        Me.LblIDR.Location = New System.Drawing.Point(120, 38)
        Me.LblIDR.Name = "LblIDR"
        Me.LblIDR.Size = New System.Drawing.Size(26, 13)
        Me.LblIDR.TabIndex = 70
        Me.LblIDR.Text = "IDR"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(0, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(89, 13)
        Me.Label7.TabIndex = 59
        Me.Label7.Text = "Total Duty Stamp"
        '
        'TxtTotalVat
        '
        Me.TxtTotalVat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalVat.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotalVat.Location = New System.Drawing.Point(152, 61)
        Me.TxtTotalVat.MaxLength = 50
        Me.TxtTotalVat.Name = "TxtTotalVat"
        Me.TxtTotalVat.ReadOnly = True
        Me.TxtTotalVat.Size = New System.Drawing.Size(263, 20)
        Me.TxtTotalVat.TabIndex = 2
        Me.TxtTotalVat.TabStop = False
        Me.TxtTotalVat.Text = "0"
        Me.TxtTotalVat.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtTotalDutyStamp
        '
        Me.TxtTotalDutyStamp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalDutyStamp.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotalDutyStamp.Location = New System.Drawing.Point(152, 85)
        Me.TxtTotalDutyStamp.MaxLength = 50
        Me.TxtTotalDutyStamp.Name = "TxtTotalDutyStamp"
        Me.TxtTotalDutyStamp.ReadOnly = True
        Me.TxtTotalDutyStamp.Size = New System.Drawing.Size(263, 20)
        Me.TxtTotalDutyStamp.TabIndex = 3
        Me.TxtTotalDutyStamp.TabStop = False
        Me.TxtTotalDutyStamp.Text = "0"
        Me.TxtTotalDutyStamp.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(0, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 13)
        Me.Label2.TabIndex = 55
        Me.Label2.Text = "Total VAT"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(0, 68)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(49, 13)
        Me.Label27.TabIndex = 78
        Me.Label27.Text = "Remarks"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(147, 63)
        Me.TxtRemarks.MaxLength = 150
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(268, 20)
        Me.TxtRemarks.TabIndex = 2
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(115, 40)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(26, 13)
        Me.Label24.TabIndex = 76
        Me.Label24.Text = "IDR"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(0, 42)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(73, 13)
        Me.Label25.TabIndex = 75
        Me.Label25.Text = "Total Balance"
        '
        'TxtTotalBalance
        '
        Me.TxtTotalBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalBalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotalBalance.Location = New System.Drawing.Point(147, 37)
        Me.TxtTotalBalance.MaxLength = 50
        Me.TxtTotalBalance.Name = "TxtTotalBalance"
        Me.TxtTotalBalance.ReadOnly = True
        Me.TxtTotalBalance.Size = New System.Drawing.Size(268, 20)
        Me.TxtTotalBalance.TabIndex = 1
        Me.TxtTotalBalance.TabStop = False
        Me.TxtTotalBalance.Text = "0"
        Me.TxtTotalBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(115, 14)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(26, 13)
        Me.Label21.TabIndex = 73
        Me.Label21.Text = "IDR"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(0, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(63, 13)
        Me.Label8.TabIndex = 57
        Me.Label8.Text = "Grand Total"
        '
        'TxtGrandTotal
        '
        Me.TxtGrandTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtGrandTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtGrandTotal.Location = New System.Drawing.Point(147, 11)
        Me.TxtGrandTotal.MaxLength = 50
        Me.TxtGrandTotal.Name = "TxtGrandTotal"
        Me.TxtGrandTotal.ReadOnly = True
        Me.TxtGrandTotal.Size = New System.Drawing.Size(268, 20)
        Me.TxtGrandTotal.TabIndex = 0
        Me.TxtGrandTotal.TabStop = False
        Me.TxtGrandTotal.Text = "0"
        Me.TxtGrandTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.AccModule.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 552)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 13
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
        Me.btnDown.Location = New System.Drawing.Point(40, 552)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 14
        Me.btnDown.TabStop = False
        Me.btnDown.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(436, 552)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(320, 552)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
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
        Me.btnDelete.Location = New System.Drawing.Point(498, 591)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(589, 591)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(407, 591)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(225, 591)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 620)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(884, 22)
        Me.StatusStrip1.TabIndex = 62
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(316, 591)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'FrmTransaksiPurchaseReturn_Balance_Import
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(884, 642)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnNew)
        Me.Name = "FrmTransaksiPurchaseReturn_Balance_Import"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Historical Purchase Return Data Import"
        Me.gboDetail.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.Panel2.PerformLayout()
        Me.SplitContainer3.ResumeLayout(False)
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
    Friend WithEvents TxtPrimaryStockIn As System.Windows.Forms.TextBox
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents DtpDueDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryInvoice As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindINV As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtInvoiceNumber As System.Windows.Forms.TextBox
    Friend WithEvents TxtSuppName As System.Windows.Forms.TextBox
    Friend WithEvents dtpInvoiceDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtSuppCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TxtReceived As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents TxtBalance As System.Windows.Forms.TextBox
    Friend WithEvents btnInput As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtDutystamp As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtVAT As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TxtSales As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtNoTransaksi As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalSalesIDR As System.Windows.Forms.TextBox
    Friend WithEvents LblVATIdr As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents LblIDR As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalVat As System.Windows.Forms.TextBox
    Friend WithEvents TxtTotalDutyStamp As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtGrandTotal As System.Windows.Forms.TextBox
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPR_Number As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPR_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_PR_DueDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurchaseReturn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVAT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDutystamp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceived As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPR As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySupp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
End Class
