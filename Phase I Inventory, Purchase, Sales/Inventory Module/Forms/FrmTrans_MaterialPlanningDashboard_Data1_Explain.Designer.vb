<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_Data1_Explain
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
        Dim DataGridViewCellStyle21 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle22 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle23 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle24 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle25 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle26 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle27 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle28 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle29 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle30 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage_SalesDetail = New System.Windows.Forms.TabPage
        Me.TabPage_SalesReturn = New System.Windows.Forms.TabPage
        Me.TabPage_Bundling = New System.Windows.Forms.TabPage
        Me.TabPage_Adjustment = New System.Windows.Forms.TabPage
        Me.TxtAdjustmentQty = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtBundlingQty = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtSalesQty = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtSalesAmount = New System.Windows.Forms.TextBox
        Me.LblCalcZero_Qty = New System.Windows.Forms.Label
        Me.LblCalcZero_SalesAmount = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtTotalGrossMargin = New System.Windows.Forms.TextBox
        Me.LblCalcZero_GrossMargin = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtTotalQty = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.TxtUnitCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvItemInvoice = New System.Windows.Forms.DataGridView
        Me.ColSales_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSales_DocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSales_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSales_Customer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSales_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSales_SalesAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSales_Cost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSales_GrossMargin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvItemSR = New System.Windows.Forms.DataGridView
        Me.ColSalesReturn_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalesReturn_DocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalesReturn_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalesReturn_Customer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalesReturn_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalesReturn_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalesReturn_Cost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalesReturn_GrossMargin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvItemBundling = New System.Windows.Forms.DataGridView
        Me.ColBundling_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBundling_DocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBundling_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBundling_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvItemAdjustment = New System.Windows.Forms.DataGridView
        Me.ColAdjustment_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjustment_DocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjustment_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjustment_TipeAdjustment = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjustment_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAdjustment_PrimaryTipeAdjustment = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabControl1.SuspendLayout()
        Me.TabPage_SalesDetail.SuspendLayout()
        Me.TabPage_SalesReturn.SuspendLayout()
        Me.TabPage_Bundling.SuspendLayout()
        Me.TabPage_Adjustment.SuspendLayout()
        CType(Me.dgvItemInvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemSR, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemBundling, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemAdjustment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage_SalesDetail)
        Me.TabControl1.Controls.Add(Me.TabPage_SalesReturn)
        Me.TabControl1.Controls.Add(Me.TabPage_Bundling)
        Me.TabControl1.Controls.Add(Me.TabPage_Adjustment)
        Me.TabControl1.Location = New System.Drawing.Point(12, 124)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(747, 416)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage_SalesDetail
        '
        Me.TabPage_SalesDetail.Controls.Add(Me.dgvItemInvoice)
        Me.TabPage_SalesDetail.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_SalesDetail.Name = "TabPage_SalesDetail"
        Me.TabPage_SalesDetail.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_SalesDetail.Size = New System.Drawing.Size(739, 387)
        Me.TabPage_SalesDetail.TabIndex = 0
        Me.TabPage_SalesDetail.Text = "Sales Detail"
        Me.TabPage_SalesDetail.UseVisualStyleBackColor = True
        '
        'TabPage_SalesReturn
        '
        Me.TabPage_SalesReturn.Controls.Add(Me.dgvItemSR)
        Me.TabPage_SalesReturn.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_SalesReturn.Name = "TabPage_SalesReturn"
        Me.TabPage_SalesReturn.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_SalesReturn.Size = New System.Drawing.Size(739, 387)
        Me.TabPage_SalesReturn.TabIndex = 1
        Me.TabPage_SalesReturn.Text = "Sales Return"
        Me.TabPage_SalesReturn.UseVisualStyleBackColor = True
        '
        'TabPage_Bundling
        '
        Me.TabPage_Bundling.Controls.Add(Me.dgvItemBundling)
        Me.TabPage_Bundling.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_Bundling.Name = "TabPage_Bundling"
        Me.TabPage_Bundling.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Bundling.Size = New System.Drawing.Size(739, 387)
        Me.TabPage_Bundling.TabIndex = 2
        Me.TabPage_Bundling.Text = "Bundling"
        Me.TabPage_Bundling.UseVisualStyleBackColor = True
        '
        'TabPage_Adjustment
        '
        Me.TabPage_Adjustment.Controls.Add(Me.dgvItemAdjustment)
        Me.TabPage_Adjustment.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_Adjustment.Name = "TabPage_Adjustment"
        Me.TabPage_Adjustment.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Adjustment.Size = New System.Drawing.Size(739, 387)
        Me.TabPage_Adjustment.TabIndex = 3
        Me.TabPage_Adjustment.Text = "Adjustment"
        Me.TabPage_Adjustment.UseVisualStyleBackColor = True
        '
        'TxtAdjustmentQty
        '
        Me.TxtAdjustmentQty.Location = New System.Drawing.Point(162, 66)
        Me.TxtAdjustmentQty.Name = "TxtAdjustmentQty"
        Me.TxtAdjustmentQty.ReadOnly = True
        Me.TxtAdjustmentQty.Size = New System.Drawing.Size(163, 22)
        Me.TxtAdjustmentQty.TabIndex = 4
        Me.TxtAdjustmentQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 69)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(140, 17)
        Me.Label4.TabIndex = 156
        Me.Label4.Text = "Total Adjustment Qty"
        '
        'TxtBundlingQty
        '
        Me.TxtBundlingQty.Location = New System.Drawing.Point(162, 38)
        Me.TxtBundlingQty.Name = "TxtBundlingQty"
        Me.TxtBundlingQty.ReadOnly = True
        Me.TxtBundlingQty.Size = New System.Drawing.Size(163, 22)
        Me.TxtBundlingQty.TabIndex = 3
        Me.TxtBundlingQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(16, 41)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(125, 17)
        Me.Label5.TabIndex = 154
        Me.Label5.Text = "Total Bundling Qty"
        '
        'TxtSalesQty
        '
        Me.TxtSalesQty.Location = New System.Drawing.Point(477, 10)
        Me.TxtSalesQty.Name = "TxtSalesQty"
        Me.TxtSalesQty.ReadOnly = True
        Me.TxtSalesQty.Size = New System.Drawing.Size(146, 22)
        Me.TxtSalesQty.TabIndex = 5
        Me.TxtSalesQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(331, 13)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(105, 17)
        Me.Label6.TabIndex = 152
        Me.Label6.Text = "Total Sales Qty"
        '
        'TxtSalesAmount
        '
        Me.TxtSalesAmount.Location = New System.Drawing.Point(477, 38)
        Me.TxtSalesAmount.Name = "TxtSalesAmount"
        Me.TxtSalesAmount.ReadOnly = True
        Me.TxtSalesAmount.Size = New System.Drawing.Size(146, 22)
        Me.TxtSalesAmount.TabIndex = 6
        Me.TxtSalesAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblCalcZero_Qty
        '
        Me.LblCalcZero_Qty.AutoSize = True
        Me.LblCalcZero_Qty.ForeColor = System.Drawing.Color.Red
        Me.LblCalcZero_Qty.Location = New System.Drawing.Point(630, 99)
        Me.LblCalcZero_Qty.Name = "LblCalcZero_Qty"
        Me.LblCalcZero_Qty.Size = New System.Drawing.Size(125, 17)
        Me.LblCalcZero_Qty.TabIndex = 159
        Me.LblCalcZero_Qty.Text = "Calculated as zero"
        '
        'LblCalcZero_SalesAmount
        '
        Me.LblCalcZero_SalesAmount.AutoSize = True
        Me.LblCalcZero_SalesAmount.ForeColor = System.Drawing.Color.Red
        Me.LblCalcZero_SalesAmount.Location = New System.Drawing.Point(629, 41)
        Me.LblCalcZero_SalesAmount.Name = "LblCalcZero_SalesAmount"
        Me.LblCalcZero_SalesAmount.Size = New System.Drawing.Size(125, 17)
        Me.LblCalcZero_SalesAmount.TabIndex = 160
        Me.LblCalcZero_SalesAmount.Text = "Calculated as zero"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(331, 41)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(131, 17)
        Me.Label8.TabIndex = 161
        Me.Label8.Text = "Total Sales Amount"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(331, 69)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(129, 17)
        Me.Label3.TabIndex = 163
        Me.Label3.Text = "Total Gross Margin"
        '
        'TxtTotalGrossMargin
        '
        Me.TxtTotalGrossMargin.Location = New System.Drawing.Point(477, 66)
        Me.TxtTotalGrossMargin.Name = "TxtTotalGrossMargin"
        Me.TxtTotalGrossMargin.ReadOnly = True
        Me.TxtTotalGrossMargin.Size = New System.Drawing.Size(146, 22)
        Me.TxtTotalGrossMargin.TabIndex = 7
        Me.TxtTotalGrossMargin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblCalcZero_GrossMargin
        '
        Me.LblCalcZero_GrossMargin.AutoSize = True
        Me.LblCalcZero_GrossMargin.ForeColor = System.Drawing.Color.Red
        Me.LblCalcZero_GrossMargin.Location = New System.Drawing.Point(630, 69)
        Me.LblCalcZero_GrossMargin.Name = "LblCalcZero_GrossMargin"
        Me.LblCalcZero_GrossMargin.Size = New System.Drawing.Size(125, 17)
        Me.LblCalcZero_GrossMargin.TabIndex = 164
        Me.LblCalcZero_GrossMargin.Text = "Calculated as zero"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(331, 97)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(66, 17)
        Me.Label10.TabIndex = 166
        Me.Label10.Text = "Total Qty"
        '
        'TxtTotalQty
        '
        Me.TxtTotalQty.Location = New System.Drawing.Point(477, 94)
        Me.TxtTotalQty.Name = "TxtTotalQty"
        Me.TxtTotalQty.ReadOnly = True
        Me.TxtTotalQty.Size = New System.Drawing.Size(146, 22)
        Me.TxtTotalQty.TabIndex = 8
        Me.TxtTotalQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(660, 176)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 145
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TxtUnitCode
        '
        Me.TxtUnitCode.Location = New System.Drawing.Point(162, 10)
        Me.TxtUnitCode.Name = "TxtUnitCode"
        Me.TxtUnitCode.ReadOnly = True
        Me.TxtUnitCode.Size = New System.Drawing.Size(163, 22)
        Me.TxtUnitCode.TabIndex = 167
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 17)
        Me.Label1.TabIndex = 168
        Me.Label1.Text = "Unit Code"
        '
        'dgvItemInvoice
        '
        Me.dgvItemInvoice.AllowUserToAddRows = False
        Me.dgvItemInvoice.AllowUserToDeleteRows = False
        Me.dgvItemInvoice.AllowUserToResizeColumns = False
        Me.dgvItemInvoice.AllowUserToResizeRows = False
        Me.dgvItemInvoice.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemInvoice.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItemInvoice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemInvoice.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSales_No, Me.ColSales_DocumentNo, Me.ColSales_Date, Me.ColSales_Customer, Me.ColSales_Qty, Me.ColSales_SalesAmount, Me.ColSales_Cost, Me.ColSales_GrossMargin})
        Me.dgvItemInvoice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemInvoice.Location = New System.Drawing.Point(3, 3)
        Me.dgvItemInvoice.Name = "dgvItemInvoice"
        Me.dgvItemInvoice.ReadOnly = True
        Me.dgvItemInvoice.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemInvoice.RowTemplate.Height = 24
        Me.dgvItemInvoice.Size = New System.Drawing.Size(733, 381)
        Me.dgvItemInvoice.TabIndex = 0
        '
        'ColSales_No
        '
        Me.ColSales_No.HeaderText = "No"
        Me.ColSales_No.Name = "ColSales_No"
        Me.ColSales_No.ReadOnly = True
        Me.ColSales_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSales_No.Width = 32
        '
        'ColSales_DocumentNo
        '
        Me.ColSales_DocumentNo.HeaderText = "Invoice No"
        Me.ColSales_DocumentNo.Name = "ColSales_DocumentNo"
        Me.ColSales_DocumentNo.ReadOnly = True
        Me.ColSales_DocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSales_DocumentNo.Width = 80
        '
        'ColSales_Date
        '
        Me.ColSales_Date.HeaderText = "Date"
        Me.ColSales_Date.Name = "ColSales_Date"
        Me.ColSales_Date.ReadOnly = True
        Me.ColSales_Date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSales_Date.Width = 44
        '
        'ColSales_Customer
        '
        Me.ColSales_Customer.HeaderText = "Customer"
        Me.ColSales_Customer.Name = "ColSales_Customer"
        Me.ColSales_Customer.ReadOnly = True
        Me.ColSales_Customer.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSales_Customer.Width = 74
        '
        'ColSales_Qty
        '
        DataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSales_Qty.DefaultCellStyle = DataGridViewCellStyle21
        Me.ColSales_Qty.HeaderText = "Qty"
        Me.ColSales_Qty.Name = "ColSales_Qty"
        Me.ColSales_Qty.ReadOnly = True
        Me.ColSales_Qty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSales_Qty.Width = 36
        '
        'ColSales_SalesAmount
        '
        DataGridViewCellStyle22.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSales_SalesAmount.DefaultCellStyle = DataGridViewCellStyle22
        Me.ColSales_SalesAmount.HeaderText = "Sales Amount"
        Me.ColSales_SalesAmount.Name = "ColSales_SalesAmount"
        Me.ColSales_SalesAmount.ReadOnly = True
        Me.ColSales_SalesAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSales_SalesAmount.Width = 101
        '
        'ColSales_Cost
        '
        DataGridViewCellStyle23.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSales_Cost.DefaultCellStyle = DataGridViewCellStyle23
        Me.ColSales_Cost.HeaderText = "Cost"
        Me.ColSales_Cost.Name = "ColSales_Cost"
        Me.ColSales_Cost.ReadOnly = True
        Me.ColSales_Cost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSales_Cost.Width = 42
        '
        'ColSales_GrossMargin
        '
        DataGridViewCellStyle24.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSales_GrossMargin.DefaultCellStyle = DataGridViewCellStyle24
        Me.ColSales_GrossMargin.HeaderText = "Gross margin"
        Me.ColSales_GrossMargin.Name = "ColSales_GrossMargin"
        Me.ColSales_GrossMargin.ReadOnly = True
        Me.ColSales_GrossMargin.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSales_GrossMargin.Width = 99
        '
        'dgvItemSR
        '
        Me.dgvItemSR.AllowUserToAddRows = False
        Me.dgvItemSR.AllowUserToDeleteRows = False
        Me.dgvItemSR.AllowUserToResizeColumns = False
        Me.dgvItemSR.AllowUserToResizeRows = False
        Me.dgvItemSR.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemSR.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItemSR.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemSR.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSalesReturn_No, Me.ColSalesReturn_DocumentNo, Me.ColSalesReturn_Date, Me.ColSalesReturn_Customer, Me.ColSalesReturn_Qty, Me.ColSalesReturn_Amount, Me.ColSalesReturn_Cost, Me.ColSalesReturn_GrossMargin})
        Me.dgvItemSR.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemSR.Location = New System.Drawing.Point(3, 3)
        Me.dgvItemSR.Name = "dgvItemSR"
        Me.dgvItemSR.ReadOnly = True
        Me.dgvItemSR.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemSR.RowTemplate.Height = 24
        Me.dgvItemSR.Size = New System.Drawing.Size(733, 381)
        Me.dgvItemSR.TabIndex = 1
        '
        'ColSalesReturn_No
        '
        Me.ColSalesReturn_No.HeaderText = "No"
        Me.ColSalesReturn_No.Name = "ColSalesReturn_No"
        Me.ColSalesReturn_No.ReadOnly = True
        Me.ColSalesReturn_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSalesReturn_No.Width = 32
        '
        'ColSalesReturn_DocumentNo
        '
        Me.ColSalesReturn_DocumentNo.HeaderText = "Sales Return No"
        Me.ColSalesReturn_DocumentNo.Name = "ColSalesReturn_DocumentNo"
        Me.ColSalesReturn_DocumentNo.ReadOnly = True
        Me.ColSalesReturn_DocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSalesReturn_DocumentNo.Width = 118
        '
        'ColSalesReturn_Date
        '
        Me.ColSalesReturn_Date.HeaderText = "Date"
        Me.ColSalesReturn_Date.Name = "ColSalesReturn_Date"
        Me.ColSalesReturn_Date.ReadOnly = True
        Me.ColSalesReturn_Date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSalesReturn_Date.Width = 44
        '
        'ColSalesReturn_Customer
        '
        Me.ColSalesReturn_Customer.HeaderText = "Customer"
        Me.ColSalesReturn_Customer.Name = "ColSalesReturn_Customer"
        Me.ColSalesReturn_Customer.ReadOnly = True
        Me.ColSalesReturn_Customer.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSalesReturn_Customer.Width = 74
        '
        'ColSalesReturn_Qty
        '
        DataGridViewCellStyle25.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSalesReturn_Qty.DefaultCellStyle = DataGridViewCellStyle25
        Me.ColSalesReturn_Qty.HeaderText = "Qty"
        Me.ColSalesReturn_Qty.Name = "ColSalesReturn_Qty"
        Me.ColSalesReturn_Qty.ReadOnly = True
        Me.ColSalesReturn_Qty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSalesReturn_Qty.Width = 36
        '
        'ColSalesReturn_Amount
        '
        DataGridViewCellStyle26.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSalesReturn_Amount.DefaultCellStyle = DataGridViewCellStyle26
        Me.ColSalesReturn_Amount.HeaderText = "Sales Return Amount"
        Me.ColSalesReturn_Amount.Name = "ColSalesReturn_Amount"
        Me.ColSalesReturn_Amount.ReadOnly = True
        Me.ColSalesReturn_Amount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSalesReturn_Amount.Width = 148
        '
        'ColSalesReturn_Cost
        '
        DataGridViewCellStyle27.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSalesReturn_Cost.DefaultCellStyle = DataGridViewCellStyle27
        Me.ColSalesReturn_Cost.HeaderText = "Cost"
        Me.ColSalesReturn_Cost.Name = "ColSalesReturn_Cost"
        Me.ColSalesReturn_Cost.ReadOnly = True
        Me.ColSalesReturn_Cost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSalesReturn_Cost.Width = 42
        '
        'ColSalesReturn_GrossMargin
        '
        DataGridViewCellStyle28.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSalesReturn_GrossMargin.DefaultCellStyle = DataGridViewCellStyle28
        Me.ColSalesReturn_GrossMargin.HeaderText = "Gross Margin"
        Me.ColSalesReturn_GrossMargin.Name = "ColSalesReturn_GrossMargin"
        Me.ColSalesReturn_GrossMargin.ReadOnly = True
        Me.ColSalesReturn_GrossMargin.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSalesReturn_GrossMargin.Width = 99
        '
        'dgvItemBundling
        '
        Me.dgvItemBundling.AllowUserToAddRows = False
        Me.dgvItemBundling.AllowUserToDeleteRows = False
        Me.dgvItemBundling.AllowUserToResizeColumns = False
        Me.dgvItemBundling.AllowUserToResizeRows = False
        Me.dgvItemBundling.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemBundling.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItemBundling.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemBundling.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColBundling_No, Me.ColBundling_DocumentNo, Me.ColBundling_Date, Me.ColBundling_Qty})
        Me.dgvItemBundling.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemBundling.Location = New System.Drawing.Point(3, 3)
        Me.dgvItemBundling.Name = "dgvItemBundling"
        Me.dgvItemBundling.ReadOnly = True
        Me.dgvItemBundling.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemBundling.RowTemplate.Height = 24
        Me.dgvItemBundling.Size = New System.Drawing.Size(733, 381)
        Me.dgvItemBundling.TabIndex = 2
        '
        'ColBundling_No
        '
        Me.ColBundling_No.HeaderText = "No"
        Me.ColBundling_No.Name = "ColBundling_No"
        Me.ColBundling_No.ReadOnly = True
        Me.ColBundling_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBundling_No.Width = 32
        '
        'ColBundling_DocumentNo
        '
        Me.ColBundling_DocumentNo.HeaderText = "Document No"
        Me.ColBundling_DocumentNo.Name = "ColBundling_DocumentNo"
        Me.ColBundling_DocumentNo.ReadOnly = True
        Me.ColBundling_DocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColBundling_Date
        '
        Me.ColBundling_Date.HeaderText = "Date"
        Me.ColBundling_Date.Name = "ColBundling_Date"
        Me.ColBundling_Date.ReadOnly = True
        Me.ColBundling_Date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBundling_Date.Width = 44
        '
        'ColBundling_Qty
        '
        DataGridViewCellStyle29.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBundling_Qty.DefaultCellStyle = DataGridViewCellStyle29
        Me.ColBundling_Qty.HeaderText = "Qty"
        Me.ColBundling_Qty.Name = "ColBundling_Qty"
        Me.ColBundling_Qty.ReadOnly = True
        Me.ColBundling_Qty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBundling_Qty.Width = 36
        '
        'dgvItemAdjustment
        '
        Me.dgvItemAdjustment.AllowUserToAddRows = False
        Me.dgvItemAdjustment.AllowUserToDeleteRows = False
        Me.dgvItemAdjustment.AllowUserToResizeColumns = False
        Me.dgvItemAdjustment.AllowUserToResizeRows = False
        Me.dgvItemAdjustment.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemAdjustment.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItemAdjustment.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemAdjustment.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAdjustment_No, Me.ColAdjustment_DocumentNo, Me.ColAdjustment_Date, Me.ColAdjustment_TipeAdjustment, Me.ColAdjustment_Qty, Me.ColAdjustment_PrimaryTipeAdjustment})
        Me.dgvItemAdjustment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemAdjustment.Location = New System.Drawing.Point(3, 3)
        Me.dgvItemAdjustment.Name = "dgvItemAdjustment"
        Me.dgvItemAdjustment.ReadOnly = True
        Me.dgvItemAdjustment.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemAdjustment.RowTemplate.Height = 24
        Me.dgvItemAdjustment.Size = New System.Drawing.Size(733, 381)
        Me.dgvItemAdjustment.TabIndex = 3
        '
        'ColAdjustment_No
        '
        Me.ColAdjustment_No.HeaderText = "No"
        Me.ColAdjustment_No.Name = "ColAdjustment_No"
        Me.ColAdjustment_No.ReadOnly = True
        Me.ColAdjustment_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAdjustment_No.Width = 32
        '
        'ColAdjustment_DocumentNo
        '
        Me.ColAdjustment_DocumentNo.HeaderText = "Document No"
        Me.ColAdjustment_DocumentNo.Name = "ColAdjustment_DocumentNo"
        Me.ColAdjustment_DocumentNo.ReadOnly = True
        Me.ColAdjustment_DocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColAdjustment_Date
        '
        Me.ColAdjustment_Date.HeaderText = "Date"
        Me.ColAdjustment_Date.Name = "ColAdjustment_Date"
        Me.ColAdjustment_Date.ReadOnly = True
        Me.ColAdjustment_Date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAdjustment_Date.Width = 44
        '
        'ColAdjustment_TipeAdjustment
        '
        Me.ColAdjustment_TipeAdjustment.HeaderText = "Adjustment Type"
        Me.ColAdjustment_TipeAdjustment.Name = "ColAdjustment_TipeAdjustment"
        Me.ColAdjustment_TipeAdjustment.ReadOnly = True
        Me.ColAdjustment_TipeAdjustment.Width = 139
        '
        'ColAdjustment_Qty
        '
        DataGridViewCellStyle30.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAdjustment_Qty.DefaultCellStyle = DataGridViewCellStyle30
        Me.ColAdjustment_Qty.HeaderText = "Qty"
        Me.ColAdjustment_Qty.Name = "ColAdjustment_Qty"
        Me.ColAdjustment_Qty.ReadOnly = True
        Me.ColAdjustment_Qty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAdjustment_Qty.Width = 36
        '
        'ColAdjustment_PrimaryTipeAdjustment
        '
        Me.ColAdjustment_PrimaryTipeAdjustment.HeaderText = "Primary Tipe Adj"
        Me.ColAdjustment_PrimaryTipeAdjustment.Name = "ColAdjustment_PrimaryTipeAdjustment"
        Me.ColAdjustment_PrimaryTipeAdjustment.ReadOnly = True
        Me.ColAdjustment_PrimaryTipeAdjustment.Visible = False
        Me.ColAdjustment_PrimaryTipeAdjustment.Width = 137
        '
        'FrmTrans_MaterialPlanningDashboard_Data1_Explain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(771, 552)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.TxtUnitCode)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.TxtTotalQty)
        Me.Controls.Add(Me.LblCalcZero_GrossMargin)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TxtTotalGrossMargin)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.LblCalcZero_SalesAmount)
        Me.Controls.Add(Me.LblCalcZero_Qty)
        Me.Controls.Add(Me.TxtSalesAmount)
        Me.Controls.Add(Me.TxtAdjustmentQty)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtBundlingQty)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TxtSalesQty)
        Me.Controls.Add(Me.Label6)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(779, 589)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_Data1_Explain"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Calculation"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage_SalesDetail.ResumeLayout(False)
        Me.TabPage_SalesReturn.ResumeLayout(False)
        Me.TabPage_Bundling.ResumeLayout(False)
        Me.TabPage_Adjustment.ResumeLayout(False)
        CType(Me.dgvItemInvoice, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemSR, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemBundling, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemAdjustment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_SalesDetail As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_SalesReturn As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_Bundling As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_Adjustment As System.Windows.Forms.TabPage
    Friend WithEvents TxtAdjustmentQty As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtBundlingQty As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtSalesQty As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtSalesAmount As System.Windows.Forms.TextBox
    Friend WithEvents LblCalcZero_Qty As System.Windows.Forms.Label
    Friend WithEvents LblCalcZero_SalesAmount As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalGrossMargin As System.Windows.Forms.TextBox
    Friend WithEvents LblCalcZero_GrossMargin As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalQty As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TxtUnitCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvItemInvoice As System.Windows.Forms.DataGridView
    Friend WithEvents ColSales_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSales_DocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSales_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSales_Customer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSales_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSales_SalesAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSales_Cost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSales_GrossMargin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvItemSR As System.Windows.Forms.DataGridView
    Friend WithEvents ColSalesReturn_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalesReturn_DocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalesReturn_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalesReturn_Customer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalesReturn_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalesReturn_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalesReturn_Cost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalesReturn_GrossMargin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvItemBundling As System.Windows.Forms.DataGridView
    Friend WithEvents ColBundling_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBundling_DocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBundling_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBundling_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvItemAdjustment As System.Windows.Forms.DataGridView
    Friend WithEvents ColAdjustment_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjustment_DocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjustment_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjustment_TipeAdjustment As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjustment_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAdjustment_PrimaryTipeAdjustment As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
