<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_SF_AddDO
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
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Label6 = New System.Windows.Forms.Label
        Me.dtpOverdue = New System.Windows.Forms.DateTimePicker
        Me.llbUncek = New System.Windows.Forms.LinkLabel
        Me.llbCek = New System.Windows.Forms.LinkLabel
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.dgvException = New System.Windows.Forms.DataGridView
        Me.ColExc_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExc_Customer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExc_OverdueInvoice = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColExc_AboveCreditLimit = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColExc_CreateDO = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColExc_PrimaryCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.LblCreditLimit = New System.Windows.Forms.Label
        Me.LblTotalARAbove = New System.Windows.Forms.Label
        Me.LblTotalOverDue = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.dgvSalesBalance = New System.Windows.Forms.DataGridView
        Me.ColInvoiceNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvoiceDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDueDateInvoice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalesValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVatSalesValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn31 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalSales_VAT_Value = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn32 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn33 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceivedOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceivementOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExhangeRateBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPayInv = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColBalanceSales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceived = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceivement = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySalesBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDocumenAcuan_Sales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeDocumentAcuan_Sales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn34 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrencyBalanceSales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn36 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn37 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalanceOriginal01 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn38 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn39 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label101 = New System.Windows.Forms.Label
        Me.Label100 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Lbl1 = New System.Windows.Forms.Label
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.chkCreateBySONumber = New System.Windows.Forms.CheckBox
        Me.chkDirectlyCreateDO = New System.Windows.Forms.CheckBox
        Me.chkDirectlyCreateStockOut = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.LblProgress = New System.Windows.Forms.Label
        Me.dgvDONumbering = New System.Windows.Forms.DataGridView
        Me.btnCancelProses = New System.Windows.Forms.Button
        Me.chkDirectlyPrint = New System.Windows.Forms.CheckBox
        Me.chkDirectlySave = New System.Windows.Forms.CheckBox
        Me.prgAll = New System.Windows.Forms.ProgressBar
        Me.btnNext = New System.Windows.Forms.Button
        Me.btnBack = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.ColDONo_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONo_Customer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONo_SoNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONo_DoNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONo_CurrencyCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONo_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONo_PrimaryCur = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONo_UseTax = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDONo_TarPercentage = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONo_PrimaryCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONo_PrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabControl1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvException, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvSalesBalance, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvDONumbering, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(9, 10)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(695, 466)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.dtpOverdue)
        Me.TabPage2.Controls.Add(Me.llbUncek)
        Me.TabPage2.Controls.Add(Me.llbCek)
        Me.TabPage2.Controls.Add(Me.SplitContainer1)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.Lbl1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Size = New System.Drawing.Size(687, 440)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Customer Exception"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(570, 2)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(10, 13)
        Me.Label6.TabIndex = 92
        Me.Label6.Text = ")"
        '
        'dtpOverdue
        '
        Me.dtpOverdue.CustomFormat = "dd MMM yyyy"
        Me.dtpOverdue.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOverdue.Location = New System.Drawing.Point(446, 0)
        Me.dtpOverdue.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpOverdue.Name = "dtpOverdue"
        Me.dtpOverdue.Size = New System.Drawing.Size(120, 20)
        Me.dtpOverdue.TabIndex = 91
        '
        'llbUncek
        '
        Me.llbUncek.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncek.AutoSize = True
        Me.llbUncek.Location = New System.Drawing.Point(621, 16)
        Me.llbUncek.Name = "llbUncek"
        Me.llbUncek.Size = New System.Drawing.Size(65, 13)
        Me.llbUncek.TabIndex = 90
        Me.llbUncek.TabStop = True
        Me.llbUncek.Text = "Uncheck All"
        Me.llbUncek.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbCek
        '
        Me.llbCek.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCek.AutoSize = True
        Me.llbCek.Location = New System.Drawing.Point(565, 16)
        Me.llbCek.Name = "llbCek"
        Me.llbCek.Size = New System.Drawing.Size(52, 13)
        Me.llbCek.TabIndex = 89
        Me.llbCek.TabStop = True
        Me.llbCek.Text = "Check All"
        Me.llbCek.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(4, 32)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvException)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Size = New System.Drawing.Size(679, 405)
        Me.SplitContainer1.SplitterDistance = 143
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 3
        Me.SplitContainer1.TabStop = False
        '
        'dgvException
        '
        Me.dgvException.AllowUserToAddRows = False
        Me.dgvException.AllowUserToDeleteRows = False
        Me.dgvException.AllowUserToResizeColumns = False
        Me.dgvException.AllowUserToResizeRows = False
        Me.dgvException.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvException.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColExc_No, Me.ColExc_Customer, Me.ColExc_OverdueInvoice, Me.ColExc_AboveCreditLimit, Me.ColExc_CreateDO, Me.ColExc_PrimaryCust})
        Me.dgvException.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvException.Location = New System.Drawing.Point(0, 0)
        Me.dgvException.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvException.MultiSelect = False
        Me.dgvException.Name = "dgvException"
        Me.dgvException.RowTemplate.Height = 24
        Me.dgvException.Size = New System.Drawing.Size(679, 143)
        Me.dgvException.TabIndex = 0
        '
        'ColExc_No
        '
        Me.ColExc_No.HeaderText = "No"
        Me.ColExc_No.Name = "ColExc_No"
        Me.ColExc_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExc_No.Width = 27
        '
        'ColExc_Customer
        '
        Me.ColExc_Customer.HeaderText = "Customer"
        Me.ColExc_Customer.Name = "ColExc_Customer"
        Me.ColExc_Customer.Width = 76
        '
        'ColExc_OverdueInvoice
        '
        Me.ColExc_OverdueInvoice.HeaderText = "Overdue Invoice"
        Me.ColExc_OverdueInvoice.Name = "ColExc_OverdueInvoice"
        Me.ColExc_OverdueInvoice.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColExc_OverdueInvoice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColExc_OverdueInvoice.Width = 111
        '
        'ColExc_AboveCreditLimit
        '
        Me.ColExc_AboveCreditLimit.HeaderText = "Above Credit Limit"
        Me.ColExc_AboveCreditLimit.Name = "ColExc_AboveCreditLimit"
        Me.ColExc_AboveCreditLimit.Width = 98
        '
        'ColExc_CreateDO
        '
        Me.ColExc_CreateDO.HeaderText = "Create DO"
        Me.ColExc_CreateDO.Name = "ColExc_CreateDO"
        Me.ColExc_CreateDO.Width = 63
        '
        'ColExc_PrimaryCust
        '
        Me.ColExc_PrimaryCust.HeaderText = "Primary Cust"
        Me.ColExc_PrimaryCust.Name = "ColExc_PrimaryCust"
        Me.ColExc_PrimaryCust.Visible = False
        Me.ColExc_PrimaryCust.Width = 113
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.LblCreditLimit)
        Me.GroupBox1.Controls.Add(Me.LblTotalARAbove)
        Me.GroupBox1.Controls.Add(Me.LblTotalOverDue)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.dgvSalesBalance)
        Me.GroupBox1.Controls.Add(Me.Label101)
        Me.GroupBox1.Controls.Add(Me.Label100)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(679, 259)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Detail"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(4, 238)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(249, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Note: The rows colored in red are overdue invoices"
        '
        'LblCreditLimit
        '
        Me.LblCreditLimit.AutoSize = True
        Me.LblCreditLimit.Location = New System.Drawing.Point(393, 31)
        Me.LblCreditLimit.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblCreditLimit.Name = "LblCreditLimit"
        Me.LblCreditLimit.Size = New System.Drawing.Size(10, 13)
        Me.LblCreditLimit.TabIndex = 6
        Me.LblCreditLimit.Text = ":"
        '
        'LblTotalARAbove
        '
        Me.LblTotalARAbove.AutoSize = True
        Me.LblTotalARAbove.Location = New System.Drawing.Point(393, 15)
        Me.LblTotalARAbove.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblTotalARAbove.Name = "LblTotalARAbove"
        Me.LblTotalARAbove.Size = New System.Drawing.Size(10, 13)
        Me.LblTotalARAbove.TabIndex = 5
        Me.LblTotalARAbove.Text = ":"
        '
        'LblTotalOverDue
        '
        Me.LblTotalOverDue.AutoSize = True
        Me.LblTotalOverDue.Location = New System.Drawing.Point(119, 15)
        Me.LblTotalOverDue.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblTotalOverDue.Name = "LblTotalOverDue"
        Me.LblTotalOverDue.Size = New System.Drawing.Size(10, 13)
        Me.LblTotalOverDue.TabIndex = 4
        Me.LblTotalOverDue.Text = ":"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(262, 31)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Credit Limit"
        '
        'dgvSalesBalance
        '
        Me.dgvSalesBalance.AllowUserToAddRows = False
        Me.dgvSalesBalance.AllowUserToDeleteRows = False
        Me.dgvSalesBalance.AllowUserToResizeRows = False
        Me.dgvSalesBalance.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSalesBalance.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSalesBalance.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColInvoiceNo, Me.ColInvoiceDate, Me.ColDueDateInvoice, Me.ColInvRemarks, Me.ColSalesValue, Me.ColVatSalesValue, Me.DataGridViewTextBoxColumn31, Me.ColTotalSales_VAT_Value, Me.DataGridViewTextBoxColumn32, Me.DataGridViewTextBoxColumn33, Me.ColReceivedOriginal, Me.ColReceivementOriginal, Me.ColExhangeRateBalance, Me.ColPayInv, Me.ColBalanceSales, Me.ColReceived, Me.ColReceivement, Me.ColPrimarySalesBalance, Me.ColPrimaryDocumenAcuan_Sales, Me.ColTipeDocumentAcuan_Sales, Me.DataGridViewTextBoxColumn34, Me.ColPrimaryCurrencyBalanceSales, Me.DataGridViewTextBoxColumn36, Me.DataGridViewTextBoxColumn37, Me.ColBalanceOriginal01, Me.DataGridViewTextBoxColumn38, Me.DataGridViewTextBoxColumn39})
        Me.dgvSalesBalance.Location = New System.Drawing.Point(5, 56)
        Me.dgvSalesBalance.MultiSelect = False
        Me.dgvSalesBalance.Name = "dgvSalesBalance"
        Me.dgvSalesBalance.ReadOnly = True
        Me.dgvSalesBalance.RowTemplate.Height = 24
        Me.dgvSalesBalance.Size = New System.Drawing.Size(669, 179)
        Me.dgvSalesBalance.TabIndex = 2
        '
        'ColInvoiceNo
        '
        Me.ColInvoiceNo.Frozen = True
        Me.ColInvoiceNo.HeaderText = "Invoice Number"
        Me.ColInvoiceNo.Name = "ColInvoiceNo"
        Me.ColInvoiceNo.ReadOnly = True
        Me.ColInvoiceNo.Width = 107
        '
        'ColInvoiceDate
        '
        Me.ColInvoiceDate.HeaderText = "Invoice Date"
        Me.ColInvoiceDate.Name = "ColInvoiceDate"
        Me.ColInvoiceDate.ReadOnly = True
        Me.ColInvoiceDate.Visible = False
        Me.ColInvoiceDate.Width = 111
        '
        'ColDueDateInvoice
        '
        Me.ColDueDateInvoice.HeaderText = "Due Date"
        Me.ColDueDateInvoice.Name = "ColDueDateInvoice"
        Me.ColDueDateInvoice.ReadOnly = True
        Me.ColDueDateInvoice.Width = 78
        '
        'ColInvRemarks
        '
        Me.ColInvRemarks.HeaderText = "Remarks"
        Me.ColInvRemarks.Name = "ColInvRemarks"
        Me.ColInvRemarks.ReadOnly = True
        Me.ColInvRemarks.Width = 74
        '
        'ColSalesValue
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSalesValue.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColSalesValue.HeaderText = "Sales"
        Me.ColSalesValue.Name = "ColSalesValue"
        Me.ColSalesValue.ReadOnly = True
        Me.ColSalesValue.Visible = False
        Me.ColSalesValue.Width = 68
        '
        'ColVatSalesValue
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColVatSalesValue.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColVatSalesValue.HeaderText = "VAT"
        Me.ColVatSalesValue.Name = "ColVatSalesValue"
        Me.ColVatSalesValue.ReadOnly = True
        Me.ColVatSalesValue.Visible = False
        Me.ColVatSalesValue.Width = 60
        '
        'DataGridViewTextBoxColumn31
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn31.DefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridViewTextBoxColumn31.HeaderText = "Duty stamp"
        Me.DataGridViewTextBoxColumn31.Name = "DataGridViewTextBoxColumn31"
        Me.DataGridViewTextBoxColumn31.ReadOnly = True
        Me.DataGridViewTextBoxColumn31.Visible = False
        Me.DataGridViewTextBoxColumn31.Width = 104
        '
        'ColTotalSales_VAT_Value
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalSales_VAT_Value.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColTotalSales_VAT_Value.HeaderText = "Sales + VAT + Duty stamp"
        Me.ColTotalSales_VAT_Value.Name = "ColTotalSales_VAT_Value"
        Me.ColTotalSales_VAT_Value.ReadOnly = True
        Me.ColTotalSales_VAT_Value.Visible = False
        Me.ColTotalSales_VAT_Value.Width = 198
        '
        'DataGridViewTextBoxColumn32
        '
        Me.DataGridViewTextBoxColumn32.HeaderText = "Currency"
        Me.DataGridViewTextBoxColumn32.Name = "DataGridViewTextBoxColumn32"
        Me.DataGridViewTextBoxColumn32.ReadOnly = True
        Me.DataGridViewTextBoxColumn32.Width = 74
        '
        'DataGridViewTextBoxColumn33
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn33.DefaultCellStyle = DataGridViewCellStyle5
        Me.DataGridViewTextBoxColumn33.HeaderText = "Balance"
        Me.DataGridViewTextBoxColumn33.Name = "DataGridViewTextBoxColumn33"
        Me.DataGridViewTextBoxColumn33.ReadOnly = True
        Me.DataGridViewTextBoxColumn33.Width = 71
        '
        'ColReceivedOriginal
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColReceivedOriginal.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColReceivedOriginal.HeaderText = "Received"
        Me.ColReceivedOriginal.Name = "ColReceivedOriginal"
        Me.ColReceivedOriginal.ReadOnly = True
        Me.ColReceivedOriginal.Visible = False
        Me.ColReceivedOriginal.Width = 92
        '
        'ColReceivementOriginal
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColReceivementOriginal.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColReceivementOriginal.HeaderText = "Pay"
        Me.ColReceivementOriginal.Name = "ColReceivementOriginal"
        Me.ColReceivementOriginal.ReadOnly = True
        Me.ColReceivementOriginal.Visible = False
        Me.ColReceivementOriginal.Width = 57
        '
        'ColExhangeRateBalance
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColExhangeRateBalance.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColExhangeRateBalance.HeaderText = "Exchange Rate"
        Me.ColExhangeRateBalance.Name = "ColExhangeRateBalance"
        Me.ColExhangeRateBalance.ReadOnly = True
        Me.ColExhangeRateBalance.Width = 106
        '
        'ColPayInv
        '
        Me.ColPayInv.HeaderText = ""
        Me.ColPayInv.Name = "ColPayInv"
        Me.ColPayInv.ReadOnly = True
        Me.ColPayInv.Text = "Pay"
        Me.ColPayInv.Visible = False
        Me.ColPayInv.Width = 5
        '
        'ColBalanceSales
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalanceSales.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColBalanceSales.HeaderText = "Balance IDR"
        Me.ColBalanceSales.Name = "ColBalanceSales"
        Me.ColBalanceSales.ReadOnly = True
        Me.ColBalanceSales.Width = 93
        '
        'ColReceived
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColReceived.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColReceived.HeaderText = "Received IDR"
        Me.ColReceived.Name = "ColReceived"
        Me.ColReceived.ReadOnly = True
        Me.ColReceived.Visible = False
        Me.ColReceived.Width = 119
        '
        'ColReceivement
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColReceivement.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColReceivement.HeaderText = "Receivement IDR"
        Me.ColReceivement.Name = "ColReceivement"
        Me.ColReceivement.ReadOnly = True
        Me.ColReceivement.Visible = False
        Me.ColReceivement.Width = 142
        '
        'ColPrimarySalesBalance
        '
        Me.ColPrimarySalesBalance.HeaderText = "PrimarySalesBalance"
        Me.ColPrimarySalesBalance.Name = "ColPrimarySalesBalance"
        Me.ColPrimarySalesBalance.ReadOnly = True
        Me.ColPrimarySalesBalance.Visible = False
        Me.ColPrimarySalesBalance.Width = 167
        '
        'ColPrimaryDocumenAcuan_Sales
        '
        Me.ColPrimaryDocumenAcuan_Sales.HeaderText = "Primary Documen Acuan"
        Me.ColPrimaryDocumenAcuan_Sales.Name = "ColPrimaryDocumenAcuan_Sales"
        Me.ColPrimaryDocumenAcuan_Sales.ReadOnly = True
        Me.ColPrimaryDocumenAcuan_Sales.Visible = False
        Me.ColPrimaryDocumenAcuan_Sales.Width = 189
        '
        'ColTipeDocumentAcuan_Sales
        '
        Me.ColTipeDocumentAcuan_Sales.HeaderText = "TipeDocumenAcuan"
        Me.ColTipeDocumentAcuan_Sales.Name = "ColTipeDocumentAcuan_Sales"
        Me.ColTipeDocumentAcuan_Sales.ReadOnly = True
        Me.ColTipeDocumentAcuan_Sales.Visible = False
        Me.ColTipeDocumentAcuan_Sales.Width = 161
        '
        'DataGridViewTextBoxColumn34
        '
        Me.DataGridViewTextBoxColumn34.HeaderText = "Primary Cust"
        Me.DataGridViewTextBoxColumn34.Name = "DataGridViewTextBoxColumn34"
        Me.DataGridViewTextBoxColumn34.ReadOnly = True
        Me.DataGridViewTextBoxColumn34.Visible = False
        Me.DataGridViewTextBoxColumn34.Width = 113
        '
        'ColPrimaryCurrencyBalanceSales
        '
        Me.ColPrimaryCurrencyBalanceSales.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrencyBalanceSales.Name = "ColPrimaryCurrencyBalanceSales"
        Me.ColPrimaryCurrencyBalanceSales.ReadOnly = True
        Me.ColPrimaryCurrencyBalanceSales.Visible = False
        Me.ColPrimaryCurrencyBalanceSales.Width = 142
        '
        'DataGridViewTextBoxColumn36
        '
        Me.DataGridViewTextBoxColumn36.HeaderText = "Default Prec Cur"
        Me.DataGridViewTextBoxColumn36.Name = "DataGridViewTextBoxColumn36"
        Me.DataGridViewTextBoxColumn36.ReadOnly = True
        Me.DataGridViewTextBoxColumn36.Visible = False
        Me.DataGridViewTextBoxColumn36.Width = 137
        '
        'DataGridViewTextBoxColumn37
        '
        Me.DataGridViewTextBoxColumn37.HeaderText = "Balance1"
        Me.DataGridViewTextBoxColumn37.Name = "DataGridViewTextBoxColumn37"
        Me.DataGridViewTextBoxColumn37.ReadOnly = True
        Me.DataGridViewTextBoxColumn37.Visible = False
        Me.DataGridViewTextBoxColumn37.Width = 92
        '
        'ColBalanceOriginal01
        '
        Me.ColBalanceOriginal01.HeaderText = "BalanceOriginal01"
        Me.ColBalanceOriginal01.Name = "ColBalanceOriginal01"
        Me.ColBalanceOriginal01.ReadOnly = True
        Me.ColBalanceOriginal01.Visible = False
        Me.ColBalanceOriginal01.Width = 149
        '
        'DataGridViewTextBoxColumn38
        '
        Me.DataGridViewTextBoxColumn38.HeaderText = "Balance Original 02"
        Me.DataGridViewTextBoxColumn38.Name = "DataGridViewTextBoxColumn38"
        Me.DataGridViewTextBoxColumn38.ReadOnly = True
        Me.DataGridViewTextBoxColumn38.Visible = False
        Me.DataGridViewTextBoxColumn38.Width = 157
        '
        'DataGridViewTextBoxColumn39
        '
        Me.DataGridViewTextBoxColumn39.HeaderText = "Primary Akun Balance"
        Me.DataGridViewTextBoxColumn39.Name = "DataGridViewTextBoxColumn39"
        Me.DataGridViewTextBoxColumn39.ReadOnly = True
        Me.DataGridViewTextBoxColumn39.Visible = False
        Me.DataGridViewTextBoxColumn39.Width = 172
        '
        'Label101
        '
        Me.Label101.AutoSize = True
        Me.Label101.Location = New System.Drawing.Point(262, 15)
        Me.Label101.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label101.Name = "Label101"
        Me.Label101.Size = New System.Drawing.Size(131, 13)
        Me.Label101.TabIndex = 1
        Me.Label101.Text = "Total Account Receivable"
        '
        'Label100
        '
        Me.Label100.AutoSize = True
        Me.Label100.Location = New System.Drawing.Point(4, 15)
        Me.Label100.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label100.Name = "Label100"
        Me.Label100.Size = New System.Drawing.Size(113, 13)
        Me.Label100.TabIndex = 0
        Me.Label100.Text = "Total Overdue Invoice"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 16)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(127, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "please specify what to do"
        '
        'Lbl1
        '
        Me.Lbl1.AutoSize = True
        Me.Lbl1.Location = New System.Drawing.Point(2, 2)
        Me.Lbl1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Lbl1.Name = "Lbl1"
        Me.Lbl1.Size = New System.Drawing.Size(442, 13)
        Me.Lbl1.TabIndex = 0
        Me.Lbl1.Text = "The following customer have account receivable above credit limit or overdue invo" & _
            "ice (as of "
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.chkCreateBySONumber)
        Me.TabPage3.Controls.Add(Me.chkDirectlyCreateDO)
        Me.TabPage3.Controls.Add(Me.chkDirectlyCreateStockOut)
        Me.TabPage3.Controls.Add(Me.Label3)
        Me.TabPage3.Controls.Add(Me.dtpDate)
        Me.TabPage3.Controls.Add(Me.LblProgress)
        Me.TabPage3.Controls.Add(Me.dgvDONumbering)
        Me.TabPage3.Controls.Add(Me.btnCancelProses)
        Me.TabPage3.Controls.Add(Me.chkDirectlyPrint)
        Me.TabPage3.Controls.Add(Me.chkDirectlySave)
        Me.TabPage3.Controls.Add(Me.prgAll)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage3.Size = New System.Drawing.Size(687, 440)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Progress"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'chkCreateBySONumber
        '
        Me.chkCreateBySONumber.AutoSize = True
        Me.chkCreateBySONumber.Location = New System.Drawing.Point(176, 50)
        Me.chkCreateBySONumber.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkCreateBySONumber.Name = "chkCreateBySONumber"
        Me.chkCreateBySONumber.Size = New System.Drawing.Size(127, 17)
        Me.chkCreateBySONumber.TabIndex = 54
        Me.chkCreateBySONumber.TabStop = False
        Me.chkCreateBySONumber.Text = "Create 1 DO for 1 SO"
        Me.chkCreateBySONumber.UseVisualStyleBackColor = True
        Me.chkCreateBySONumber.Visible = False
        '
        'chkDirectlyCreateDO
        '
        Me.chkDirectlyCreateDO.AutoSize = True
        Me.chkDirectlyCreateDO.Location = New System.Drawing.Point(331, 71)
        Me.chkDirectlyCreateDO.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkDirectlyCreateDO.Name = "chkDirectlyCreateDO"
        Me.chkDirectlyCreateDO.Size = New System.Drawing.Size(150, 17)
        Me.chkDirectlyCreateDO.TabIndex = 7
        Me.chkDirectlyCreateDO.TabStop = False
        Me.chkDirectlyCreateDO.Text = "Directly print delivery order"
        Me.chkDirectlyCreateDO.UseVisualStyleBackColor = True
        Me.chkDirectlyCreateDO.Visible = False
        '
        'chkDirectlyCreateStockOut
        '
        Me.chkDirectlyCreateStockOut.AutoSize = True
        Me.chkDirectlyCreateStockOut.Location = New System.Drawing.Point(331, 50)
        Me.chkDirectlyCreateStockOut.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkDirectlyCreateStockOut.Name = "chkDirectlyCreateStockOut"
        Me.chkDirectlyCreateStockOut.Size = New System.Drawing.Size(238, 17)
        Me.chkDirectlyCreateStockOut.TabIndex = 6
        Me.chkDirectlyCreateStockOut.TabStop = False
        Me.chkDirectlyCreateStockOut.Text = "Directly create stock out sales when possible"
        Me.chkDirectlyCreateStockOut.UseVisualStyleBackColor = True
        Me.chkDirectlyCreateStockOut.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(328, 28)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 53
        Me.Label3.Text = "DO Date"
        '
        'dtpDate
        '
        Me.dtpDate.Location = New System.Drawing.Point(380, 28)
        Me.dtpDate.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(151, 20)
        Me.dtpDate.TabIndex = 3
        '
        'LblProgress
        '
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(4, 28)
        Me.LblProgress.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(28, 13)
        Me.LblProgress.TabIndex = 2
        Me.LblProgress.Text = "......."
        '
        'dgvDONumbering
        '
        Me.dgvDONumbering.AllowUserToAddRows = False
        Me.dgvDONumbering.AllowUserToDeleteRows = False
        Me.dgvDONumbering.AllowUserToResizeColumns = False
        Me.dgvDONumbering.AllowUserToResizeRows = False
        Me.dgvDONumbering.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDONumbering.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDONumbering.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDONo_No, Me.ColDONo_Customer, Me.ColDONo_SoNumber, Me.ColDONo_DoNumber, Me.ColDONo_CurrencyCode, Me.ColDONo_Remarks, Me.ColDONo_PrimaryCur, Me.ColDONo_UseTax, Me.ColDONo_TarPercentage, Me.ColDONo_PrimaryCust, Me.ColDONo_PrimarySO})
        Me.dgvDONumbering.Location = New System.Drawing.Point(4, 91)
        Me.dgvDONumbering.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvDONumbering.Name = "dgvDONumbering"
        Me.dgvDONumbering.RowTemplate.Height = 24
        Me.dgvDONumbering.Size = New System.Drawing.Size(681, 347)
        Me.dgvDONumbering.TabIndex = 8
        '
        'btnCancelProses
        '
        Me.btnCancelProses.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancelProses.Enabled = False
        Me.btnCancelProses.Location = New System.Drawing.Point(543, 4)
        Me.btnCancelProses.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancelProses.Name = "btnCancelProses"
        Me.btnCancelProses.Size = New System.Drawing.Size(56, 20)
        Me.btnCancelProses.TabIndex = 1
        Me.btnCancelProses.TabStop = False
        Me.btnCancelProses.Text = "Cancel"
        Me.btnCancelProses.UseVisualStyleBackColor = True
        Me.btnCancelProses.Visible = False
        '
        'chkDirectlyPrint
        '
        Me.chkDirectlyPrint.AutoSize = True
        Me.chkDirectlyPrint.Location = New System.Drawing.Point(4, 71)
        Me.chkDirectlyPrint.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkDirectlyPrint.Name = "chkDirectlyPrint"
        Me.chkDirectlyPrint.Size = New System.Drawing.Size(140, 17)
        Me.chkDirectlyPrint.TabIndex = 5
        Me.chkDirectlyPrint.Text = "Directly print packing list"
        Me.chkDirectlyPrint.UseVisualStyleBackColor = True
        '
        'chkDirectlySave
        '
        Me.chkDirectlySave.AutoSize = True
        Me.chkDirectlySave.Location = New System.Drawing.Point(4, 50)
        Me.chkDirectlySave.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkDirectlySave.Name = "chkDirectlySave"
        Me.chkDirectlySave.Size = New System.Drawing.Size(146, 17)
        Me.chkDirectlySave.TabIndex = 4
        Me.chkDirectlySave.Text = "Directly save to database"
        Me.chkDirectlySave.UseVisualStyleBackColor = True
        '
        'prgAll
        '
        Me.prgAll.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgAll.Location = New System.Drawing.Point(4, 4)
        Me.prgAll.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.prgAll.Name = "prgAll"
        Me.prgAll.Size = New System.Drawing.Size(534, 19)
        Me.prgAll.TabIndex = 0
        '
        'btnNext
        '
        Me.btnNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNext.Location = New System.Drawing.Point(485, 481)
        Me.btnNext.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(69, 23)
        Me.btnNext.TabIndex = 2
        Me.btnNext.Text = "Next >>"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBack.Enabled = False
        Me.btnBack.Location = New System.Drawing.Point(411, 482)
        Me.btnBack.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(69, 23)
        Me.btnBack.TabIndex = 1
        Me.btnBack.Text = "<< Back"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Enabled = False
        Me.btnOK.Location = New System.Drawing.Point(558, 482)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(69, 23)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(631, 481)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(69, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        Me.BackgroundWorker2.WorkerReportsProgress = True
        '
        'ColDONo_No
        '
        Me.ColDONo_No.HeaderText = "No"
        Me.ColDONo_No.Name = "ColDONo_No"
        Me.ColDONo_No.Width = 46
        '
        'ColDONo_Customer
        '
        Me.ColDONo_Customer.HeaderText = "Customer"
        Me.ColDONo_Customer.Name = "ColDONo_Customer"
        Me.ColDONo_Customer.Width = 76
        '
        'ColDONo_SoNumber
        '
        Me.ColDONo_SoNumber.HeaderText = "SO Number"
        Me.ColDONo_SoNumber.Name = "ColDONo_SoNumber"
        Me.ColDONo_SoNumber.ReadOnly = True
        Me.ColDONo_SoNumber.Width = 87
        '
        'ColDONo_DoNumber
        '
        Me.ColDONo_DoNumber.HeaderText = "DO Number"
        Me.ColDONo_DoNumber.Name = "ColDONo_DoNumber"
        Me.ColDONo_DoNumber.Width = 88
        '
        'ColDONo_CurrencyCode
        '
        Me.ColDONo_CurrencyCode.HeaderText = "Currency"
        Me.ColDONo_CurrencyCode.Name = "ColDONo_CurrencyCode"
        Me.ColDONo_CurrencyCode.Width = 74
        '
        'ColDONo_Remarks
        '
        Me.ColDONo_Remarks.HeaderText = "Remarks"
        Me.ColDONo_Remarks.Name = "ColDONo_Remarks"
        Me.ColDONo_Remarks.Width = 74
        '
        'ColDONo_PrimaryCur
        '
        Me.ColDONo_PrimaryCur.HeaderText = "Primary Cur"
        Me.ColDONo_PrimaryCur.Name = "ColDONo_PrimaryCur"
        Me.ColDONo_PrimaryCur.Visible = False
        Me.ColDONo_PrimaryCur.Width = 85
        '
        'ColDONo_UseTax
        '
        Me.ColDONo_UseTax.HeaderText = "Use Tax"
        Me.ColDONo_UseTax.Name = "ColDONo_UseTax"
        Me.ColDONo_UseTax.Width = 53
        '
        'ColDONo_TarPercentage
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDONo_TarPercentage.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColDONo_TarPercentage.HeaderText = "Tax %"
        Me.ColDONo_TarPercentage.Name = "ColDONo_TarPercentage"
        Me.ColDONo_TarPercentage.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDONo_TarPercentage.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDONo_TarPercentage.Width = 42
        '
        'ColDONo_PrimaryCust
        '
        Me.ColDONo_PrimaryCust.HeaderText = "Primary Cust"
        Me.ColDONo_PrimaryCust.Name = "ColDONo_PrimaryCust"
        Me.ColDONo_PrimaryCust.Visible = False
        Me.ColDONo_PrimaryCust.Width = 90
        '
        'ColDONo_PrimarySO
        '
        Me.ColDONo_PrimarySO.HeaderText = "Primary SO"
        Me.ColDONo_PrimarySO.Name = "ColDONo_PrimarySO"
        Me.ColDONo_PrimarySO.ReadOnly = True
        Me.ColDONo_PrimarySO.Visible = False
        Me.ColDONo_PrimarySO.Width = 84
        '
        'FrmTrans_MaterialPlanningDashboard_SF_AddDO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(713, 515)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.TabControl1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(644, 552)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_SF_AddDO"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create DO"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvException, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvSalesBalance, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.dgvDONumbering, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvException As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label101 As System.Windows.Forms.Label
    Friend WithEvents Label100 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Public WithEvents dgvSalesBalance As System.Windows.Forms.DataGridView
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents dgvDONumbering As System.Windows.Forms.DataGridView
    Friend WithEvents btnCancelProses As System.Windows.Forms.Button
    Friend WithEvents chkDirectlyPrint As System.Windows.Forms.CheckBox
    Friend WithEvents chkDirectlySave As System.Windows.Forms.CheckBox
    Friend WithEvents prgAll As System.Windows.Forms.ProgressBar
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkDirectlyCreateStockOut As System.Windows.Forms.CheckBox
    Friend WithEvents chkDirectlyCreateDO As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LblCreditLimit As System.Windows.Forms.Label
    Friend WithEvents LblTotalARAbove As System.Windows.Forms.Label
    Friend WithEvents LblTotalOverDue As System.Windows.Forms.Label
    Friend WithEvents ColInvoiceNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvoiceDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDueDateInvoice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalesValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVatSalesValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn31 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalSales_VAT_Value As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn32 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn33 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceivedOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceivementOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExhangeRateBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPayInv As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColBalanceSales As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceived As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceivement As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySalesBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDocumenAcuan_Sales As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeDocumentAcuan_Sales As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn34 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrencyBalanceSales As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn36 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn37 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalanceOriginal01 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn38 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn39 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpOverdue As System.Windows.Forms.DateTimePicker
    Friend WithEvents llbUncek As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCek As System.Windows.Forms.LinkLabel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Lbl1 As System.Windows.Forms.Label
    Friend WithEvents ColExc_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExc_Customer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExc_OverdueInvoice As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColExc_AboveCreditLimit As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColExc_CreateDO As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColExc_PrimaryCust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents chkCreateBySONumber As System.Windows.Forms.CheckBox
    Friend WithEvents ColDONo_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONo_Customer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONo_SoNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONo_DoNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONo_CurrencyCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONo_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONo_PrimaryCur As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONo_UseTax As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDONo_TarPercentage As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONo_PrimaryCust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONo_PrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
