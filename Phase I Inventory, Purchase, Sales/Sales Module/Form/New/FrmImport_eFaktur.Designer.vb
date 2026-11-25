<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmImport_eFaktur
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.txtFilePath = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnFindFilePath = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.btnCreateInvoice = New System.Windows.Forms.Button
        Me.dgvHeader = New System.Windows.Forms.DataGridView
        Me.LlbCheckAll = New System.Windows.Forms.LinkLabel
        Me.LlbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkShowInvoiceReg = New System.Windows.Forms.CheckBox
        Me.ChkShowInvalidOnly = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.colItem_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_ProductCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_ProductDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyOnHand = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_UnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_DiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_DPP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_PPN = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_InvalidReason = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.llbProgressBar = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colCreate = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColCustCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colInvoiceNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colTaxInvoiceNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colNPWP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDPP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPPN = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colInvalidReason = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(238, 50)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(25, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "To"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(91, 49)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(135, 22)
        Me.dtpFrom.TabIndex = 2
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(273, 49)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(135, 22)
        Me.dtpTo.TabIndex = 3
        '
        'txtFilePath
        '
        Me.txtFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFilePath.Location = New System.Drawing.Point(91, 21)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.ReadOnly = True
        Me.txtFilePath.Size = New System.Drawing.Size(765, 22)
        Me.txtFilePath.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 17)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "File Path"
        '
        'btnFindFilePath
        '
        Me.btnFindFilePath.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindFilePath.Location = New System.Drawing.Point(862, 20)
        Me.btnFindFilePath.Name = "btnFindFilePath"
        Me.btnFindFilePath.Size = New System.Drawing.Size(63, 23)
        Me.btnFindFilePath.TabIndex = 6
        Me.btnFindFilePath.Text = "Find"
        Me.btnFindFilePath.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShow.Location = New System.Drawing.Point(862, 48)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(63, 23)
        Me.btnShow.TabIndex = 7
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'btnCreateInvoice
        '
        Me.btnCreateInvoice.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCreateInvoice.Location = New System.Drawing.Point(756, 588)
        Me.btnCreateInvoice.Name = "btnCreateInvoice"
        Me.btnCreateInvoice.Size = New System.Drawing.Size(118, 26)
        Me.btnCreateInvoice.TabIndex = 8
        Me.btnCreateInvoice.Text = "Create Invoice"
        Me.btnCreateInvoice.UseVisualStyleBackColor = True
        '
        'dgvHeader
        '
        Me.dgvHeader.AllowUserToAddRows = False
        Me.dgvHeader.AllowUserToDeleteRows = False
        Me.dgvHeader.AllowUserToResizeColumns = False
        Me.dgvHeader.AllowUserToResizeRows = False
        Me.dgvHeader.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvHeader.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.colCreate, Me.ColCustCode, Me.ColCustName, Me.colCustomer, Me.colInvoiceNumber, Me.colTaxInvoiceNumber, Me.colDate, Me.colNPWP, Me.colDPP, Me.colPPN, Me.colInvalidReason, Me.colPrimary})
        Me.dgvHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvHeader.Location = New System.Drawing.Point(3, 18)
        Me.dgvHeader.Name = "dgvHeader"
        Me.dgvHeader.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvHeader.RowTemplate.Height = 24
        Me.dgvHeader.Size = New System.Drawing.Size(928, 291)
        Me.dgvHeader.TabIndex = 9
        '
        'LlbCheckAll
        '
        Me.LlbCheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll.AutoSize = True
        Me.LlbCheckAll.Location = New System.Drawing.Point(771, -2)
        Me.LlbCheckAll.Name = "LlbCheckAll"
        Me.LlbCheckAll.Size = New System.Drawing.Size(66, 17)
        Me.LlbCheckAll.TabIndex = 10
        Me.LlbCheckAll.TabStop = True
        Me.LlbCheckAll.Text = "Check All"
        Me.LlbCheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LlbUncheckAll
        '
        Me.LlbUncheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll.AutoSize = True
        Me.LlbUncheckAll.Location = New System.Drawing.Point(843, -2)
        Me.LlbUncheckAll.Name = "LlbUncheckAll"
        Me.LlbUncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.LlbUncheckAll.TabIndex = 11
        Me.LlbUncheckAll.TabStop = True
        Me.LlbUncheckAll.Text = "Uncheck All"
        Me.LlbUncheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.ChkShowInvoiceReg)
        Me.GroupBox1.Controls.Add(Me.ChkShowInvalidOnly)
        Me.GroupBox1.Controls.Add(Me.txtFilePath)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.btnShow)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Controls.Add(Me.dtpTo)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.btnFindFilePath)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(934, 84)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        '
        'ChkShowInvoiceReg
        '
        Me.ChkShowInvoiceReg.AutoSize = True
        Me.ChkShowInvoiceReg.Location = New System.Drawing.Point(563, 50)
        Me.ChkShowInvoiceReg.Name = "ChkShowInvoiceReg"
        Me.ChkShowInvoiceReg.Size = New System.Drawing.Size(234, 21)
        Me.ChkShowInvoiceReg.TabIndex = 13
        Me.ChkShowInvoiceReg.Text = "Show Invoice Already Registered"
        Me.ChkShowInvoiceReg.UseVisualStyleBackColor = True
        '
        'ChkShowInvalidOnly
        '
        Me.ChkShowInvalidOnly.AutoSize = True
        Me.ChkShowInvalidOnly.Location = New System.Drawing.Point(419, 50)
        Me.ChkShowInvalidOnly.Name = "ChkShowInvalidOnly"
        Me.ChkShowInvalidOnly.Size = New System.Drawing.Size(138, 21)
        Me.ChkShowInvalidOnly.TabIndex = 12
        Me.ChkShowInvalidOnly.Text = "Show Invalid Only"
        Me.ChkShowInvalidOnly.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.LlbUncheckAll)
        Me.GroupBox2.Controls.Add(Me.LlbCheckAll)
        Me.GroupBox2.Controls.Add(Me.dgvHeader)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(934, 312)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "e-Faktur Data"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(880, 588)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(63, 26)
        Me.btnClose.TabIndex = 14
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(15, 584)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(735, 20)
        Me.ProgressBar1.TabIndex = 15
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.dgvItem)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(934, 164)
        Me.GroupBox3.TabIndex = 14
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Item Data"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colItem_No, Me.colItem_ProductCode, Me.colItem_ProductDesc, Me.colItem_Qty, Me.colItem_Unit, Me.ColQtyOnHand, Me.colItem_UnitPrice, Me.colItem_DiscAmount, Me.colItem_Amount, Me.colItem_DPP, Me.colItem_PPN, Me.colItem_InvalidReason})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 18)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(928, 143)
        Me.dgvItem.TabIndex = 9
        '
        'colItem_No
        '
        Me.colItem_No.HeaderText = "No"
        Me.colItem_No.Name = "colItem_No"
        Me.colItem_No.ReadOnly = True
        '
        'colItem_ProductCode
        '
        Me.colItem_ProductCode.HeaderText = "Code No"
        Me.colItem_ProductCode.Name = "colItem_ProductCode"
        Me.colItem_ProductCode.ReadOnly = True
        '
        'colItem_ProductDesc
        '
        Me.colItem_ProductDesc.HeaderText = "Description"
        Me.colItem_ProductDesc.Name = "colItem_ProductDesc"
        Me.colItem_ProductDesc.ReadOnly = True
        '
        'colItem_Qty
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_Qty.DefaultCellStyle = DataGridViewCellStyle4
        Me.colItem_Qty.HeaderText = "Qty"
        Me.colItem_Qty.Name = "colItem_Qty"
        Me.colItem_Qty.ReadOnly = True
        '
        'colItem_Unit
        '
        Me.colItem_Unit.HeaderText = "Unit"
        Me.colItem_Unit.Name = "colItem_Unit"
        Me.colItem_Unit.ReadOnly = True
        '
        'ColQtyOnHand
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyOnHand.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColQtyOnHand.HeaderText = "Qty On Hand"
        Me.ColQtyOnHand.Name = "ColQtyOnHand"
        Me.ColQtyOnHand.ReadOnly = True
        '
        'colItem_UnitPrice
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_UnitPrice.DefaultCellStyle = DataGridViewCellStyle6
        Me.colItem_UnitPrice.HeaderText = "Unit Price"
        Me.colItem_UnitPrice.Name = "colItem_UnitPrice"
        Me.colItem_UnitPrice.ReadOnly = True
        '
        'colItem_DiscAmount
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_DiscAmount.DefaultCellStyle = DataGridViewCellStyle7
        Me.colItem_DiscAmount.HeaderText = "Discount Amount"
        Me.colItem_DiscAmount.Name = "colItem_DiscAmount"
        Me.colItem_DiscAmount.ReadOnly = True
        '
        'colItem_Amount
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_Amount.DefaultCellStyle = DataGridViewCellStyle8
        Me.colItem_Amount.HeaderText = "Amount"
        Me.colItem_Amount.Name = "colItem_Amount"
        Me.colItem_Amount.ReadOnly = True
        '
        'colItem_DPP
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_DPP.DefaultCellStyle = DataGridViewCellStyle9
        Me.colItem_DPP.HeaderText = "DPP"
        Me.colItem_DPP.Name = "colItem_DPP"
        Me.colItem_DPP.ReadOnly = True
        '
        'colItem_PPN
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_PPN.DefaultCellStyle = DataGridViewCellStyle10
        Me.colItem_PPN.HeaderText = "PPN"
        Me.colItem_PPN.Name = "colItem_PPN"
        Me.colItem_PPN.ReadOnly = True
        '
        'colItem_InvalidReason
        '
        Me.colItem_InvalidReason.HeaderText = "Invalid Reason"
        Me.colItem_InvalidReason.Name = "colItem_InvalidReason"
        Me.colItem_InvalidReason.ReadOnly = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 102)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox3)
        Me.SplitContainer1.Size = New System.Drawing.Size(934, 480)
        Me.SplitContainer1.SplitterDistance = 312
        Me.SplitContainer1.TabIndex = 12
        '
        'llbProgressBar
        '
        Me.llbProgressBar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbProgressBar.AutoSize = True
        Me.llbProgressBar.Location = New System.Drawing.Point(20, 604)
        Me.llbProgressBar.Name = "llbProgressBar"
        Me.llbProgressBar.Size = New System.Drawing.Size(40, 17)
        Me.llbProgressBar.TabIndex = 13
        Me.llbProgressBar.Text = "From"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'colNo
        '
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        '
        'colCreate
        '
        Me.colCreate.HeaderText = "Create"
        Me.colCreate.Name = "colCreate"
        Me.colCreate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colCreate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'ColCustCode
        '
        Me.ColCustCode.HeaderText = "Customer Code"
        Me.ColCustCode.Name = "ColCustCode"
        '
        'ColCustName
        '
        Me.ColCustName.HeaderText = "Customer Name"
        Me.ColCustName.Name = "ColCustName"
        '
        'colCustomer
        '
        Me.colCustomer.HeaderText = "Customer NPWP Name"
        Me.colCustomer.Name = "colCustomer"
        Me.colCustomer.ReadOnly = True
        '
        'colInvoiceNumber
        '
        Me.colInvoiceNumber.HeaderText = "Invoice Number"
        Me.colInvoiceNumber.Name = "colInvoiceNumber"
        Me.colInvoiceNumber.ReadOnly = True
        '
        'colTaxInvoiceNumber
        '
        Me.colTaxInvoiceNumber.HeaderText = "Tax Invoice Number"
        Me.colTaxInvoiceNumber.Name = "colTaxInvoiceNumber"
        Me.colTaxInvoiceNumber.ReadOnly = True
        '
        'colDate
        '
        DataGridViewCellStyle1.Format = "dd MMM yyyy"
        DataGridViewCellStyle1.NullValue = "01/01/1950"
        Me.colDate.DefaultCellStyle = DataGridViewCellStyle1
        Me.colDate.HeaderText = "Date"
        Me.colDate.Name = "colDate"
        Me.colDate.ReadOnly = True
        '
        'colNPWP
        '
        Me.colNPWP.HeaderText = "NPWP"
        Me.colNPWP.Name = "colNPWP"
        Me.colNPWP.ReadOnly = True
        '
        'colDPP
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colDPP.DefaultCellStyle = DataGridViewCellStyle2
        Me.colDPP.HeaderText = "DPP"
        Me.colDPP.Name = "colDPP"
        Me.colDPP.ReadOnly = True
        '
        'colPPN
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPPN.DefaultCellStyle = DataGridViewCellStyle3
        Me.colPPN.HeaderText = "PPN"
        Me.colPPN.Name = "colPPN"
        Me.colPPN.ReadOnly = True
        '
        'colInvalidReason
        '
        Me.colInvalidReason.HeaderText = "Invalid Reason"
        Me.colInvalidReason.Name = "colInvalidReason"
        Me.colInvalidReason.ReadOnly = True
        '
        'colPrimary
        '
        Me.colPrimary.HeaderText = "Primary"
        Me.colPrimary.Name = "colPrimary"
        Me.colPrimary.ReadOnly = True
        Me.colPrimary.Visible = False
        '
        'FrmImport_eFaktur
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(958, 623)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.llbProgressBar)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCreateInvoice)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "FrmImport_eFaktur"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form Import e-Faktur"
        CType(Me.dgvHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnFindFilePath As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents btnCreateInvoice As System.Windows.Forms.Button
    Friend WithEvents dgvHeader As System.Windows.Forms.DataGridView
    Friend WithEvents LlbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkShowInvalidOnly As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents llbProgressBar As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ChkShowInvoiceReg As System.Windows.Forms.CheckBox
    Friend WithEvents colItem_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_ProductCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_ProductDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyOnHand As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_UnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_DiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_DPP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_PPN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_InvalidReason As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colCreate As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColCustCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colInvoiceNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colTaxInvoiceNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colNPWP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDPP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPPN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colInvalidReason As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
