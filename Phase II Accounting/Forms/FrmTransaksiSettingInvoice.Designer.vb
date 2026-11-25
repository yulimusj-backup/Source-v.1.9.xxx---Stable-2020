<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSettingInvoice
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
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtUnbalance = New System.Windows.Forms.TextBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtCustName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtCustCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtInvoiceNumber = New System.Windows.Forms.TextBox
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.LblVATUSD = New System.Windows.Forms.Label
        Me.TxtVATUSD = New System.Windows.Forms.TextBox
        Me.cboCurrency = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.LblUSD = New System.Windows.Forms.Label
        Me.TxtTaxRate = New System.Windows.Forms.TextBox
        Me.TxtSalesUSD = New System.Windows.Forms.TextBox
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
        Me.TxtSalesIDR = New System.Windows.Forms.TextBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccount_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSales = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColVAT = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColDutyStamp = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvAkun = New System.Windows.Forms.DataGridView
        Me.ColAkun_Show = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAkun_Desc_Show = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDebet = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKredit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvAkun, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
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
        Me.gboDetail.Location = New System.Drawing.Point(12, 2)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(860, 614)
        Me.gboDetail.TabIndex = 47
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblVATUSD)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtVATUSD)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboCurrency)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblUSD)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTaxRate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtSalesUSD)
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtSalesIDR)
        Me.SplitContainer1.Size = New System.Drawing.Size(843, 134)
        Me.SplitContainer1.SplitterDistance = 421
        Me.SplitContainer1.TabIndex = 165
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 83)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(82, 13)
        Me.Label5.TabIndex = 51
        Me.Label5.Text = "Customer Name"
        '
        'TxtCustName
        '
        Me.TxtCustName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustName.Location = New System.Drawing.Point(104, 78)
        Me.TxtCustName.MaxLength = 50
        Me.TxtCustName.Name = "TxtCustName"
        Me.TxtCustName.ReadOnly = True
        Me.TxtCustName.Size = New System.Drawing.Size(314, 20)
        Me.TxtCustName.TabIndex = 50
        Me.TxtCustName.TabStop = False
        Me.TxtCustName.Text = "PT. AERO NUSANTARA"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 57)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 49
        Me.Label1.Text = "Customer Code"
        '
        'TxtCustCode
        '
        Me.TxtCustCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustCode.Location = New System.Drawing.Point(104, 52)
        Me.TxtCustCode.MaxLength = 50
        Me.TxtCustCode.Name = "TxtCustCode"
        Me.TxtCustCode.ReadOnly = True
        Me.TxtCustCode.Size = New System.Drawing.Size(314, 20)
        Me.TxtCustCode.TabIndex = 48
        Me.TxtCustCode.TabStop = False
        Me.TxtCustCode.Text = "EA001"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Invoice Number"
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
        Me.TxtInvoiceNumber.Location = New System.Drawing.Point(104, 3)
        Me.TxtInvoiceNumber.MaxLength = 50
        Me.TxtInvoiceNumber.Name = "TxtInvoiceNumber"
        Me.TxtInvoiceNumber.ReadOnly = True
        Me.TxtInvoiceNumber.Size = New System.Drawing.Size(314, 20)
        Me.TxtInvoiceNumber.TabIndex = 0
        Me.TxtInvoiceNumber.TabStop = False
        Me.TxtInvoiceNumber.Text = "0001/INV/2010"
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
        Me.dtpDate.Location = New System.Drawing.Point(104, 26)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(314, 20)
        Me.dtpDate.TabIndex = 1
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
        'cboCurrency
        '
        Me.cboCurrency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurrency.FormattingEnabled = True
        Me.cboCurrency.Items.AddRange(New Object() {"IDR Invoice", "Non-IDR Invoice"})
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
        'TxtSalesUSD
        '
        Me.TxtSalesUSD.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSalesUSD.Location = New System.Drawing.Point(147, 29)
        Me.TxtSalesUSD.MaxLength = 50
        Me.TxtSalesUSD.Name = "TxtSalesUSD"
        Me.TxtSalesUSD.ReadOnly = True
        Me.TxtSalesUSD.Size = New System.Drawing.Size(110, 20)
        Me.TxtSalesUSD.TabIndex = 74
        Me.TxtSalesUSD.TabStop = False
        Me.TxtSalesUSD.Text = "1,125.65"
        Me.TxtSalesUSD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
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
        Me.Label6.Size = New System.Drawing.Size(76, 13)
        Me.Label6.TabIndex = 53
        Me.Label6.Text = "Sales After DP"
        '
        'TxtSalesIDR
        '
        Me.TxtSalesIDR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalesIDR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSalesIDR.Location = New System.Drawing.Point(305, 29)
        Me.TxtSalesIDR.MaxLength = 50
        Me.TxtSalesIDR.Name = "TxtSalesIDR"
        Me.TxtSalesIDR.ReadOnly = True
        Me.TxtSalesIDR.Size = New System.Drawing.Size(110, 20)
        Me.TxtSalesIDR.TabIndex = 52
        Me.TxtSalesIDR.TabStop = False
        Me.TxtSalesIDR.Text = "10,000,000"
        Me.TxtSalesIDR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
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
        Me.SplitContainer2.Location = New System.Drawing.Point(6, 148)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer2.Size = New System.Drawing.Size(848, 399)
        Me.SplitContainer2.SplitterDistance = 198
        Me.SplitContainer2.TabIndex = 166
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvItem)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(848, 198)
        Me.GroupBox1.TabIndex = 13
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
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColAccount_Desc, Me.ColSales, Me.ColVAT, Me.ColDutyStamp, Me.ColRemarks, Me.ColPrimaryAkun})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 16)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.Size = New System.Drawing.Size(842, 179)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 12
        '
        'ColAccount
        '
        Me.ColAccount.HeaderText = "Account"
        Me.ColAccount.Name = "ColAccount"
        Me.ColAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccount.Width = 53
        '
        'ColAccount_Desc
        '
        Me.ColAccount_Desc.HeaderText = "Description"
        Me.ColAccount_Desc.Name = "ColAccount_Desc"
        Me.ColAccount_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccount_Desc.Width = 66
        '
        'ColSales
        '
        Me.ColSales.HeaderText = "Sales"
        Me.ColSales.Items.AddRange(New Object() {"-", "Debet", "Credit"})
        Me.ColSales.Name = "ColSales"
        Me.ColSales.Width = 39
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
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 55
        '
        'ColPrimaryAkun
        '
        Me.ColPrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColPrimaryAkun.Name = "ColPrimaryAkun"
        Me.ColPrimaryAkun.Visible = False
        Me.ColPrimaryAkun.Width = 91
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dgvAkun)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(848, 197)
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
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvAkun.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvAkun.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAkun.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvAkun.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAkun_Show, Me.ColAkun_Desc_Show, Me.ColDebet, Me.ColKredit})
        Me.dgvAkun.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAkun.Location = New System.Drawing.Point(3, 16)
        Me.dgvAkun.MultiSelect = False
        Me.dgvAkun.Name = "dgvAkun"
        Me.dgvAkun.ReadOnly = True
        Me.dgvAkun.Size = New System.Drawing.Size(842, 178)
        Me.dgvAkun.StandardTab = True
        Me.dgvAkun.TabIndex = 13
        '
        'ColAkun_Show
        '
        Me.ColAkun_Show.HeaderText = "Account"
        Me.ColAkun_Show.Name = "ColAkun_Show"
        Me.ColAkun_Show.ReadOnly = True
        Me.ColAkun_Show.Width = 72
        '
        'ColAkun_Desc_Show
        '
        Me.ColAkun_Desc_Show.HeaderText = "Description"
        Me.ColAkun_Desc_Show.Name = "ColAkun_Desc_Show"
        Me.ColAkun_Desc_Show.ReadOnly = True
        Me.ColAkun_Desc_Show.Width = 85
        '
        'ColDebet
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDebet.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColDebet.HeaderText = "Debet"
        Me.ColDebet.Name = "ColDebet"
        Me.ColDebet.ReadOnly = True
        Me.ColDebet.Width = 61
        '
        'ColKredit
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColKredit.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColKredit.HeaderText = "Credit"
        Me.ColKredit.Name = "ColKredit"
        Me.ColKredit.ReadOnly = True
        Me.ColKredit.Width = 59
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(434, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(434, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 620)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(884, 22)
        Me.StatusStrip1.TabIndex = 48
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'FrmTransaksiSettingInvoice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(884, 642)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.MinimumSize = New System.Drawing.Size(900, 678)
        Me.Name = "FrmTransaksiSettingInvoice"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Setting Invoice"
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
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dgvAkun, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
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
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtDutyStamp As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtVatIDR As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtSalesIDR As System.Windows.Forms.TextBox
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents dgvAkun As System.Windows.Forms.DataGridView
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents cboCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents LblUSD As System.Windows.Forms.Label
    Friend WithEvents TxtTaxRate As System.Windows.Forms.TextBox
    Friend WithEvents TxtSalesUSD As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents LblVATIdr As System.Windows.Forms.Label
    Friend WithEvents LblIDR As System.Windows.Forms.Label
    Friend WithEvents LblVATUSD As System.Windows.Forms.Label
    Friend WithEvents TxtVATUSD As System.Windows.Forms.TextBox
    Friend WithEvents ColAkun_Show As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAkun_Desc_Show As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDebet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKredit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccount_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSales As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColVAT As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColDutyStamp As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
