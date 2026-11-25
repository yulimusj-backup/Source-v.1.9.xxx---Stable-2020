<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSettingAOP
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
        Me.TxtVATUSD = New System.Windows.Forms.TextBox
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtUnbalance = New System.Windows.Forms.TextBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.LblDPCurrency = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtPurchaseDP = New System.Windows.Forms.TextBox
        Me.LblDP = New System.Windows.Forms.Label
        Me.TxtCustName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtCustCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtInvoiceNumber = New System.Windows.Forms.TextBox
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.LblTotalCost2 = New System.Windows.Forms.Label
        Me.LblTotalCost1 = New System.Windows.Forms.Label
        Me.TxtTotalCost = New System.Windows.Forms.TextBox
        Me.LblVATUSD = New System.Windows.Forms.Label
        Me.cboCurrency = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.LblUSD = New System.Windows.Forms.Label
        Me.TxtTaxRate = New System.Windows.Forms.TextBox
        Me.TxtPurchaseUSD = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.LblVATIdr = New System.Windows.Forms.Label
        Me.LblIDR = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtDutyStamp = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtVatIDR = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtPurchaseIDR = New System.Windows.Forms.TextBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccount_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurchase = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColVAT = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColDutyStamp = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColTotalCost = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColDownpayment = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.gboCost = New System.Windows.Forms.GroupBox
        Me.dgvCost = New System.Windows.Forms.DataGridView
        Me.ColNoCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrencyCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostNominal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExcRateCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarksCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostAccountNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostAccountDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrencyCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvAkun = New System.Windows.Forms.DataGridView
        Me.Col_Account_Show = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Account_Desc_Show = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDebet = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKredit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboCost.SuspendLayout()
        CType(Me.dgvCost, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvAkun, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtVATUSD
        '
        Me.TxtVATUSD.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtVATUSD.Location = New System.Drawing.Point(147, 55)
        Me.TxtVATUSD.MaxLength = 50
        Me.TxtVATUSD.Name = "TxtVATUSD"
        Me.TxtVATUSD.ReadOnly = True
        Me.TxtVATUSD.Size = New System.Drawing.Size(110, 20)
        Me.TxtVATUSD.TabIndex = 77
        Me.TxtVATUSD.TabStop = False
        Me.TxtVATUSD.Text = "112.57"
        Me.TxtVATUSD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.Label9)
        Me.gboDetail.Controls.Add(Me.TxtUnbalance)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.SplitContainer2)
        Me.gboDetail.Location = New System.Drawing.Point(12, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(860, 614)
        Me.gboDetail.TabIndex = 49
        Me.gboDetail.TabStop = False
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(9, 561)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(59, 13)
        Me.Label9.TabIndex = 168
        Me.Label9.Text = "Unbalance"
        '
        'TxtUnbalance
        '
        Me.TxtUnbalance.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtUnbalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUnbalance.Location = New System.Drawing.Point(110, 556)
        Me.TxtUnbalance.MaxLength = 50
        Me.TxtUnbalance.Name = "TxtUnbalance"
        Me.TxtUnbalance.ReadOnly = True
        Me.TxtUnbalance.Size = New System.Drawing.Size(314, 20)
        Me.TxtUnbalance.TabIndex = 167
        Me.TxtUnbalance.TabStop = False
        Me.TxtUnbalance.Text = "0"
        Me.TxtUnbalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblDPCurrency)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPurchaseDP)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblDP)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCustName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCustCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtInvoiceNumber)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryStockIn)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblTotalCost2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblTotalCost1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTotalCost)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblVATUSD)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtVATUSD)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboCurrency)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblUSD)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTaxRate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPurchaseUSD)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label21)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label20)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblVATIdr)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblIDR)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDutyStamp)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTotal)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtVatIDR)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPurchaseIDR)
        Me.SplitContainer1.Size = New System.Drawing.Size(843, 154)
        Me.SplitContainer1.SplitterDistance = 421
        Me.SplitContainer1.TabIndex = 165
        '
        'LblDPCurrency
        '
        Me.LblDPCurrency.AutoSize = True
        Me.LblDPCurrency.Location = New System.Drawing.Point(130, 107)
        Me.LblDPCurrency.Name = "LblDPCurrency"
        Me.LblDPCurrency.Size = New System.Drawing.Size(26, 13)
        Me.LblDPCurrency.TabIndex = 84
        Me.LblDPCurrency.Text = "IDR"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 83)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 13)
        Me.Label5.TabIndex = 51
        Me.Label5.Text = "Supplier Name"
        '
        'TxtPurchaseDP
        '
        Me.TxtPurchaseDP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPurchaseDP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPurchaseDP.Location = New System.Drawing.Point(162, 104)
        Me.TxtPurchaseDP.MaxLength = 50
        Me.TxtPurchaseDP.Name = "TxtPurchaseDP"
        Me.TxtPurchaseDP.ReadOnly = True
        Me.TxtPurchaseDP.Size = New System.Drawing.Size(256, 20)
        Me.TxtPurchaseDP.TabIndex = 83
        Me.TxtPurchaseDP.TabStop = False
        Me.TxtPurchaseDP.Text = "1,500,000"
        Me.TxtPurchaseDP.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblDP
        '
        Me.LblDP.AutoSize = True
        Me.LblDP.Location = New System.Drawing.Point(3, 107)
        Me.LblDP.Name = "LblDP"
        Me.LblDP.Size = New System.Drawing.Size(123, 13)
        Me.LblDP.TabIndex = 82
        Me.LblDP.Text = "Purchase Downpayment"
        '
        'TxtCustName
        '
        Me.TxtCustName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustName.Location = New System.Drawing.Point(133, 78)
        Me.TxtCustName.MaxLength = 50
        Me.TxtCustName.Name = "TxtCustName"
        Me.TxtCustName.ReadOnly = True
        Me.TxtCustName.Size = New System.Drawing.Size(285, 20)
        Me.TxtCustName.TabIndex = 50
        Me.TxtCustName.TabStop = False
        Me.TxtCustName.Text = "PT. EKA JAYA ANUGRAH NUSANTARA"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 57)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 49
        Me.Label1.Text = "Supplier Code"
        '
        'TxtCustCode
        '
        Me.TxtCustCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustCode.Location = New System.Drawing.Point(133, 52)
        Me.TxtCustCode.MaxLength = 50
        Me.TxtCustCode.Name = "TxtCustCode"
        Me.TxtCustCode.ReadOnly = True
        Me.TxtCustCode.Size = New System.Drawing.Size(285, 20)
        Me.TxtCustCode.TabIndex = 48
        Me.TxtCustCode.TabStop = False
        Me.TxtCustCode.Text = "LA001"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "AOP Number"
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
        'TxtInvoiceNumber
        '
        Me.TxtInvoiceNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtInvoiceNumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtInvoiceNumber.Location = New System.Drawing.Point(133, 3)
        Me.TxtInvoiceNumber.MaxLength = 50
        Me.TxtInvoiceNumber.Name = "TxtInvoiceNumber"
        Me.TxtInvoiceNumber.ReadOnly = True
        Me.TxtInvoiceNumber.Size = New System.Drawing.Size(285, 20)
        Me.TxtInvoiceNumber.TabIndex = 0
        Me.TxtInvoiceNumber.TabStop = False
        Me.TxtInvoiceNumber.Text = "0001/AOP/2010"
        '
        'TxtPrimaryStockIn
        '
        Me.TxtPrimaryStockIn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryStockIn.Location = New System.Drawing.Point(144, 3)
        Me.TxtPrimaryStockIn.Name = "TxtPrimaryStockIn"
        Me.TxtPrimaryStockIn.Size = New System.Drawing.Size(20, 20)
        Me.TxtPrimaryStockIn.TabIndex = 47
        Me.TxtPrimaryStockIn.TabStop = False
        Me.TxtPrimaryStockIn.Text = "-1"
        Me.TxtPrimaryStockIn.Visible = False
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Enabled = False
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(133, 26)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(285, 20)
        Me.dtpDate.TabIndex = 1
        '
        'LblTotalCost2
        '
        Me.LblTotalCost2.AutoSize = True
        Me.LblTotalCost2.Location = New System.Drawing.Point(115, 134)
        Me.LblTotalCost2.Name = "LblTotalCost2"
        Me.LblTotalCost2.Size = New System.Drawing.Size(26, 13)
        Me.LblTotalCost2.TabIndex = 81
        Me.LblTotalCost2.Text = "IDR"
        '
        'LblTotalCost1
        '
        Me.LblTotalCost1.AutoSize = True
        Me.LblTotalCost1.Location = New System.Drawing.Point(0, 136)
        Me.LblTotalCost1.Name = "LblTotalCost1"
        Me.LblTotalCost1.Size = New System.Drawing.Size(55, 13)
        Me.LblTotalCost1.TabIndex = 80
        Me.LblTotalCost1.Text = "Total Cost"
        '
        'TxtTotalCost
        '
        Me.TxtTotalCost.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalCost.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotalCost.Location = New System.Drawing.Point(147, 131)
        Me.TxtTotalCost.MaxLength = 50
        Me.TxtTotalCost.Name = "TxtTotalCost"
        Me.TxtTotalCost.ReadOnly = True
        Me.TxtTotalCost.Size = New System.Drawing.Size(268, 20)
        Me.TxtTotalCost.TabIndex = 79
        Me.TxtTotalCost.TabStop = False
        Me.TxtTotalCost.Text = "14,021,625"
        Me.TxtTotalCost.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblVATUSD
        '
        Me.LblVATUSD.AutoSize = True
        Me.LblVATUSD.Location = New System.Drawing.Point(115, 58)
        Me.LblVATUSD.Name = "LblVATUSD"
        Me.LblVATUSD.Size = New System.Drawing.Size(30, 13)
        Me.LblVATUSD.TabIndex = 78
        Me.LblVATUSD.Text = "USD"
        '
        'cboCurrency
        '
        Me.cboCurrency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurrency.FormattingEnabled = True
        Me.cboCurrency.Items.AddRange(New Object() {"IDR AOP", "Non-IDR AOP"})
        Me.cboCurrency.Location = New System.Drawing.Point(3, 5)
        Me.cboCurrency.Name = "cboCurrency"
        Me.cboCurrency.Size = New System.Drawing.Size(239, 21)
        Me.cboCurrency.TabIndex = 76
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(248, 8)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(51, 13)
        Me.Label12.TabIndex = 68
        Me.Label12.Text = "Tax Rate"
        '
        'LblUSD
        '
        Me.LblUSD.AutoSize = True
        Me.LblUSD.Location = New System.Drawing.Point(115, 32)
        Me.LblUSD.Name = "LblUSD"
        Me.LblUSD.Size = New System.Drawing.Size(30, 13)
        Me.LblUSD.TabIndex = 75
        Me.LblUSD.Text = "USD"
        '
        'TxtTaxRate
        '
        Me.TxtTaxRate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTaxRate.Location = New System.Drawing.Point(305, 5)
        Me.TxtTaxRate.Name = "TxtTaxRate"
        Me.TxtTaxRate.ReadOnly = True
        Me.TxtTaxRate.Size = New System.Drawing.Size(110, 20)
        Me.TxtTaxRate.TabIndex = 67
        Me.TxtTaxRate.TabStop = False
        Me.TxtTaxRate.Text = "9,025.65"
        Me.TxtTaxRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtPurchaseUSD
        '
        Me.TxtPurchaseUSD.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPurchaseUSD.Location = New System.Drawing.Point(147, 29)
        Me.TxtPurchaseUSD.MaxLength = 50
        Me.TxtPurchaseUSD.Name = "TxtPurchaseUSD"
        Me.TxtPurchaseUSD.ReadOnly = True
        Me.TxtPurchaseUSD.Size = New System.Drawing.Size(110, 20)
        Me.TxtPurchaseUSD.TabIndex = 74
        Me.TxtPurchaseUSD.TabStop = False
        Me.TxtPurchaseUSD.Text = "1,125.65"
        Me.TxtPurchaseUSD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(115, 109)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(26, 13)
        Me.Label21.TabIndex = 73
        Me.Label21.Text = "IDR"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(115, 83)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(26, 13)
        Me.Label20.TabIndex = 72
        Me.Label20.Text = "IDR"
        '
        'LblVATIdr
        '
        Me.LblVATIdr.AutoSize = True
        Me.LblVATIdr.Location = New System.Drawing.Point(273, 58)
        Me.LblVATIdr.Name = "LblVATIdr"
        Me.LblVATIdr.Size = New System.Drawing.Size(26, 13)
        Me.LblVATIdr.TabIndex = 71
        Me.LblVATIdr.Text = "IDR"
        '
        'LblIDR
        '
        Me.LblIDR.AutoSize = True
        Me.LblIDR.Location = New System.Drawing.Point(273, 32)
        Me.LblIDR.Name = "LblIDR"
        Me.LblIDR.Size = New System.Drawing.Size(26, 13)
        Me.LblIDR.TabIndex = 70
        Me.LblIDR.Text = "IDR"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(0, 85)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(62, 13)
        Me.Label7.TabIndex = 59
        Me.Label7.Text = "Duty Stamp"
        '
        'TxtDutyStamp
        '
        Me.TxtDutyStamp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDutyStamp.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDutyStamp.Location = New System.Drawing.Point(147, 80)
        Me.TxtDutyStamp.MaxLength = 50
        Me.TxtDutyStamp.Name = "TxtDutyStamp"
        Me.TxtDutyStamp.ReadOnly = True
        Me.TxtDutyStamp.Size = New System.Drawing.Size(268, 20)
        Me.TxtDutyStamp.TabIndex = 58
        Me.TxtDutyStamp.TabStop = False
        Me.TxtDutyStamp.Text = "6,000"
        Me.TxtDutyStamp.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(0, 111)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(31, 13)
        Me.Label8.TabIndex = 57
        Me.Label8.Text = "Total"
        '
        'TxtTotal
        '
        Me.TxtTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotal.Location = New System.Drawing.Point(147, 106)
        Me.TxtTotal.MaxLength = 50
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(268, 20)
        Me.TxtTotal.TabIndex = 56
        Me.TxtTotal.TabStop = False
        Me.TxtTotal.Text = "11,006,000"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(0, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(28, 13)
        Me.Label2.TabIndex = 55
        Me.Label2.Text = "VAT"
        '
        'TxtVatIDR
        '
        Me.TxtVatIDR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtVatIDR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtVatIDR.Location = New System.Drawing.Point(305, 55)
        Me.TxtVatIDR.MaxLength = 50
        Me.TxtVatIDR.Name = "TxtVatIDR"
        Me.TxtVatIDR.ReadOnly = True
        Me.TxtVatIDR.Size = New System.Drawing.Size(110, 20)
        Me.TxtVatIDR.TabIndex = 54
        Me.TxtVatIDR.TabStop = False
        Me.TxtVatIDR.Text = "1,000,000"
        Me.TxtVatIDR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(0, 34)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(95, 13)
        Me.Label6.TabIndex = 53
        Me.Label6.Text = "Purchase After DP"
        '
        'TxtPurchaseIDR
        '
        Me.TxtPurchaseIDR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPurchaseIDR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPurchaseIDR.Location = New System.Drawing.Point(305, 29)
        Me.TxtPurchaseIDR.MaxLength = 50
        Me.TxtPurchaseIDR.Name = "TxtPurchaseIDR"
        Me.TxtPurchaseIDR.ReadOnly = True
        Me.TxtPurchaseIDR.Size = New System.Drawing.Size(110, 20)
        Me.TxtPurchaseIDR.TabIndex = 52
        Me.TxtPurchaseIDR.TabStop = False
        Me.TxtPurchaseIDR.Text = "10,000,000"
        Me.TxtPurchaseIDR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.AccModule.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 582)
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
        Me.btnDown.Location = New System.Drawing.Point(40, 582)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 14
        Me.btnDown.TabStop = False
        Me.btnDown.Text = "v"
        Me.btnDown.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(436, 582)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 16
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(320, 582)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 15
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(6, 168)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.SplitContainer3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer2.Size = New System.Drawing.Size(848, 379)
        Me.SplitContainer2.SplitterDistance = 187
        Me.SplitContainer2.TabIndex = 166
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.gboCost)
        Me.SplitContainer3.Size = New System.Drawing.Size(848, 187)
        Me.SplitContainer3.SplitterDistance = 422
        Me.SplitContainer3.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvItem)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(422, 187)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Setting"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColAccount_Desc, Me.ColPurchase, Me.ColVAT, Me.ColDutyStamp, Me.ColTotalCost, Me.ColDownpayment, Me.ColRemarks, Me.ColPrimaryAkun})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 16)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.Size = New System.Drawing.Size(416, 168)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 12
        '
        'ColAccount
        '
        Me.ColAccount.HeaderText = "Account"
        Me.ColAccount.Name = "ColAccount"
        Me.ColAccount.Width = 72
        '
        'ColAccount_Desc
        '
        Me.ColAccount_Desc.HeaderText = "Description"
        Me.ColAccount_Desc.Name = "ColAccount_Desc"
        Me.ColAccount_Desc.Width = 85
        '
        'ColPurchase
        '
        Me.ColPurchase.HeaderText = "Purchase"
        Me.ColPurchase.Items.AddRange(New Object() {"-", "Debet", "Credit"})
        Me.ColPurchase.Name = "ColPurchase"
        Me.ColPurchase.Width = 58
        '
        'ColVAT
        '
        Me.ColVAT.HeaderText = "VAT"
        Me.ColVAT.Items.AddRange(New Object() {"-", "Debet", "Credit"})
        Me.ColVAT.Name = "ColVAT"
        Me.ColVAT.Width = 39
        '
        'ColDutyStamp
        '
        Me.ColDutyStamp.HeaderText = "Duty Stamp"
        Me.ColDutyStamp.Items.AddRange(New Object() {"-", "Debet", "Credit"})
        Me.ColDutyStamp.Name = "ColDutyStamp"
        Me.ColDutyStamp.Width = 68
        '
        'ColTotalCost
        '
        Me.ColTotalCost.HeaderText = "Total Cost"
        Me.ColTotalCost.Items.AddRange(New Object() {"-", "Debet", "Credit"})
        Me.ColTotalCost.Name = "ColTotalCost"
        Me.ColTotalCost.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColTotalCost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColTotalCost.Width = 80
        '
        'ColDownpayment
        '
        Me.ColDownpayment.HeaderText = "Down Payment"
        Me.ColDownpayment.Items.AddRange(New Object() {"-", "Debet", "Credit"})
        Me.ColDownpayment.Name = "ColDownpayment"
        Me.ColDownpayment.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDownpayment.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColDownpayment.Width = 104
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'ColPrimaryAkun
        '
        Me.ColPrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColPrimaryAkun.Name = "ColPrimaryAkun"
        Me.ColPrimaryAkun.Visible = False
        Me.ColPrimaryAkun.Width = 91
        '
        'gboCost
        '
        Me.gboCost.Controls.Add(Me.dgvCost)
        Me.gboCost.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboCost.Location = New System.Drawing.Point(0, 0)
        Me.gboCost.Name = "gboCost"
        Me.gboCost.Size = New System.Drawing.Size(422, 187)
        Me.gboCost.TabIndex = 1
        Me.gboCost.TabStop = False
        Me.gboCost.Text = "Cost Data"
        '
        'dgvCost
        '
        Me.dgvCost.AllowUserToAddRows = False
        Me.dgvCost.AllowUserToDeleteRows = False
        Me.dgvCost.AllowUserToResizeColumns = False
        Me.dgvCost.AllowUserToResizeRows = False
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvCost.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvCost.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCost.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCost.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNoCost, Me.ColDescCost, Me.ColCurrencyCost, Me.ColCostNominal, Me.ColExcRateCost, Me.ColCostAmount, Me.ColRemarksCost, Me.ColCostAccountNo, Me.ColCostAccountDesc, Me.ColPrimaryCurrencyCost})
        Me.dgvCost.Location = New System.Drawing.Point(3, 16)
        Me.dgvCost.MultiSelect = False
        Me.dgvCost.Name = "dgvCost"
        Me.dgvCost.ReadOnly = True
        Me.dgvCost.Size = New System.Drawing.Size(416, 168)
        Me.dgvCost.StandardTab = True
        Me.dgvCost.TabIndex = 0
        '
        'ColNoCost
        '
        Me.ColNoCost.Frozen = True
        Me.ColNoCost.HeaderText = "No"
        Me.ColNoCost.Name = "ColNoCost"
        Me.ColNoCost.ReadOnly = True
        Me.ColNoCost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNoCost.Width = 27
        '
        'ColDescCost
        '
        Me.ColDescCost.Frozen = True
        Me.ColDescCost.HeaderText = "Description"
        Me.ColDescCost.Name = "ColDescCost"
        Me.ColDescCost.ReadOnly = True
        Me.ColDescCost.Width = 85
        '
        'ColCurrencyCost
        '
        Me.ColCurrencyCost.HeaderText = "Currency"
        Me.ColCurrencyCost.Name = "ColCurrencyCost"
        Me.ColCurrencyCost.ReadOnly = True
        Me.ColCurrencyCost.Width = 74
        '
        'ColCostNominal
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCostNominal.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColCostNominal.HeaderText = "Cost"
        Me.ColCostNominal.Name = "ColCostNominal"
        Me.ColCostNominal.ReadOnly = True
        Me.ColCostNominal.Width = 53
        '
        'ColExcRateCost
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColExcRateCost.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColExcRateCost.HeaderText = "Exhange Rate"
        Me.ColExcRateCost.Name = "ColExcRateCost"
        Me.ColExcRateCost.ReadOnly = True
        '
        'ColCostAmount
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCostAmount.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColCostAmount.HeaderText = "Cost Amount (Curr. ID)"
        Me.ColCostAmount.Name = "ColCostAmount"
        Me.ColCostAmount.ReadOnly = True
        Me.ColCostAmount.Width = 137
        '
        'ColRemarksCost
        '
        Me.ColRemarksCost.HeaderText = "Remarks"
        Me.ColRemarksCost.Name = "ColRemarksCost"
        Me.ColRemarksCost.ReadOnly = True
        Me.ColRemarksCost.Width = 74
        '
        'ColCostAccountNo
        '
        Me.ColCostAccountNo.HeaderText = "Account No"
        Me.ColCostAccountNo.Name = "ColCostAccountNo"
        Me.ColCostAccountNo.ReadOnly = True
        Me.ColCostAccountNo.Width = 89
        '
        'ColCostAccountDesc
        '
        Me.ColCostAccountDesc.HeaderText = "Description"
        Me.ColCostAccountDesc.Name = "ColCostAccountDesc"
        Me.ColCostAccountDesc.ReadOnly = True
        Me.ColCostAccountDesc.Width = 85
        '
        'ColPrimaryCurrencyCost
        '
        Me.ColPrimaryCurrencyCost.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrencyCost.Name = "ColPrimaryCurrencyCost"
        Me.ColPrimaryCurrencyCost.ReadOnly = True
        Me.ColPrimaryCurrencyCost.Visible = False
        Me.ColPrimaryCurrencyCost.Width = 111
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dgvAkun)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(848, 188)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Preview"
        '
        'dgvAkun
        '
        Me.dgvAkun.AllowUserToAddRows = False
        Me.dgvAkun.AllowUserToDeleteRows = False
        Me.dgvAkun.AllowUserToResizeColumns = False
        Me.dgvAkun.AllowUserToResizeRows = False
        DataGridViewCellStyle6.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvAkun.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvAkun.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAkun.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvAkun.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col_Account_Show, Me.Col_Account_Desc_Show, Me.ColDebet, Me.ColKredit})
        Me.dgvAkun.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAkun.Location = New System.Drawing.Point(3, 16)
        Me.dgvAkun.MultiSelect = False
        Me.dgvAkun.Name = "dgvAkun"
        Me.dgvAkun.ReadOnly = True
        Me.dgvAkun.Size = New System.Drawing.Size(842, 169)
        Me.dgvAkun.StandardTab = True
        Me.dgvAkun.TabIndex = 13
        '
        'Col_Account_Show
        '
        Me.Col_Account_Show.HeaderText = "Account"
        Me.Col_Account_Show.Name = "Col_Account_Show"
        Me.Col_Account_Show.ReadOnly = True
        Me.Col_Account_Show.Width = 72
        '
        'Col_Account_Desc_Show
        '
        Me.Col_Account_Desc_Show.HeaderText = "Description"
        Me.Col_Account_Desc_Show.Name = "Col_Account_Desc_Show"
        Me.Col_Account_Desc_Show.ReadOnly = True
        Me.Col_Account_Desc_Show.Width = 85
        '
        'ColDebet
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDebet.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColDebet.HeaderText = "Debet"
        Me.ColDebet.Name = "ColDebet"
        Me.ColDebet.ReadOnly = True
        Me.ColDebet.Width = 61
        '
        'ColKredit
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColKredit.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColKredit.HeaderText = "Credit"
        Me.ColKredit.Name = "ColKredit"
        Me.ColKredit.ReadOnly = True
        Me.ColKredit.Width = 59
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 620)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(884, 22)
        Me.StatusStrip1.TabIndex = 50
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(434, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(434, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'FrmTransaksiSettingAOP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(884, 642)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.MinimumSize = New System.Drawing.Size(900, 678)
        Me.Name = "FrmTransaksiSettingAOP"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Setting AOP"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboCost.ResumeLayout(False)
        CType(Me.dgvCost, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dgvAkun, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtVATUSD As System.Windows.Forms.TextBox
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtUnbalance As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtCustName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtCustCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtInvoiceNumber As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryStockIn As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblVATUSD As System.Windows.Forms.Label
    Friend WithEvents cboCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents LblUSD As System.Windows.Forms.Label
    Friend WithEvents TxtTaxRate As System.Windows.Forms.TextBox
    Friend WithEvents TxtPurchaseUSD As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents LblVATIdr As System.Windows.Forms.Label
    Friend WithEvents LblIDR As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtDutyStamp As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtVatIDR As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtPurchaseIDR As System.Windows.Forms.TextBox
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents dgvAkun As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents gboCost As System.Windows.Forms.GroupBox
    Public WithEvents dgvCost As System.Windows.Forms.DataGridView
    Friend WithEvents LblTotalCost2 As System.Windows.Forms.Label
    Friend WithEvents LblTotalCost1 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalCost As System.Windows.Forms.TextBox
    Friend WithEvents Col_Account_Show As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Account_Desc_Show As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDebet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKredit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNoCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrencyCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostNominal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExcRateCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarksCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostAccountNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostAccountDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrencyCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LblDPCurrency As System.Windows.Forms.Label
    Friend WithEvents TxtPurchaseDP As System.Windows.Forms.TextBox
    Friend WithEvents LblDP As System.Windows.Forms.Label
    Friend WithEvents ColAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccount_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurchase As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColVAT As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColDutyStamp As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColTotalCost As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColDownpayment As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
