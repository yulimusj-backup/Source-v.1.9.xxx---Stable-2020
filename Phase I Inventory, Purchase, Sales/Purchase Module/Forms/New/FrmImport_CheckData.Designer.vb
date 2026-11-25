<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmImport_CheckData
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
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.txtFilePath = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnFindFilePath = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.btnCreateInvoice = New System.Windows.Forms.Button
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.LlbCheckAll = New System.Windows.Forms.LinkLabel
        Me.LlbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cboTipe = New System.Windows.Forms.ComboBox
        Me.ChkShowInvalidOnly = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.colItem_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_ProductCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_ProductDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_UnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_DiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_DPP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_PPN = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colItem_InvalidReason = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colNoAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colInvoiceNumberSMESS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colInvoiceDateSMESS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerCodeSMESS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerNameSMESS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colCustomerSMESS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colNPWPSMESS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDPPSMESS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPPNSMESS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colInvoiceNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colTaxInvoiceNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colInvoiceDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colNPWP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDPP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPPN = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colSelisihInvoice = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColSelisihNPWP = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colSelisihCustomer = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colSelisihDPP = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colSelisihPPN = New System.Windows.Forms.DataGridViewCheckBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.txtFilePath.Location = New System.Drawing.Point(273, 21)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.ReadOnly = True
        Me.txtFilePath.Size = New System.Drawing.Size(507, 22)
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
        Me.btnFindFilePath.Location = New System.Drawing.Point(786, 21)
        Me.btnFindFilePath.Name = "btnFindFilePath"
        Me.btnFindFilePath.Size = New System.Drawing.Size(63, 23)
        Me.btnFindFilePath.TabIndex = 6
        Me.btnFindFilePath.Text = "Find"
        Me.btnFindFilePath.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShow.Location = New System.Drawing.Point(786, 48)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(63, 23)
        Me.btnShow.TabIndex = 7
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'btnCreateInvoice
        '
        Me.btnCreateInvoice.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCreateInvoice.Location = New System.Drawing.Point(680, 588)
        Me.btnCreateInvoice.Name = "btnCreateInvoice"
        Me.btnCreateInvoice.Size = New System.Drawing.Size(118, 26)
        Me.btnCreateInvoice.TabIndex = 8
        Me.btnCreateInvoice.Text = "Create Invoice"
        Me.btnCreateInvoice.UseVisualStyleBackColor = True
        Me.btnCreateInvoice.Visible = False
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.colNoAcuan, Me.colInvoiceNumberSMESS, Me.colInvoiceDateSMESS, Me.ColCustomerCodeSMESS, Me.ColCustomerNameSMESS, Me.colCustomerSMESS, Me.colNPWPSMESS, Me.colDPPSMESS, Me.colPPNSMESS, Me.colInvoiceNumber, Me.colTaxInvoiceNo, Me.colInvoiceDate, Me.colCustomer, Me.colNPWP, Me.colDPP, Me.colPPN, Me.colSelisihInvoice, Me.ColSelisihNPWP, Me.colSelisihCustomer, Me.colSelisihDPP, Me.colSelisihPPN})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 18)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(852, 491)
        Me.dgvItem.TabIndex = 9
        '
        'LlbCheckAll
        '
        Me.LlbCheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll.AutoSize = True
        Me.LlbCheckAll.Location = New System.Drawing.Point(695, -2)
        Me.LlbCheckAll.Name = "LlbCheckAll"
        Me.LlbCheckAll.Size = New System.Drawing.Size(66, 17)
        Me.LlbCheckAll.TabIndex = 10
        Me.LlbCheckAll.TabStop = True
        Me.LlbCheckAll.Text = "Check All"
        Me.LlbCheckAll.Visible = False
        Me.LlbCheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LlbUncheckAll
        '
        Me.LlbUncheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll.AutoSize = True
        Me.LlbUncheckAll.Location = New System.Drawing.Point(767, -2)
        Me.LlbUncheckAll.Name = "LlbUncheckAll"
        Me.LlbUncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.LlbUncheckAll.TabIndex = 11
        Me.LlbUncheckAll.TabStop = True
        Me.LlbUncheckAll.Text = "Uncheck All"
        Me.LlbUncheckAll.Visible = False
        Me.LlbUncheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.cboTipe)
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
        Me.GroupBox1.Size = New System.Drawing.Size(858, 84)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        '
        'cboTipe
        '
        Me.cboTipe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTipe.FormattingEnabled = True
        Me.cboTipe.Items.AddRange(New Object() {"Purchase", "Sales", "Purchase Return", "Sales Return"})
        Me.cboTipe.Location = New System.Drawing.Point(91, 20)
        Me.cboTipe.Name = "cboTipe"
        Me.cboTipe.Size = New System.Drawing.Size(172, 24)
        Me.cboTipe.TabIndex = 12
        '
        'ChkShowInvalidOnly
        '
        Me.ChkShowInvalidOnly.AutoSize = True
        Me.ChkShowInvalidOnly.Location = New System.Drawing.Point(417, 50)
        Me.ChkShowInvalidOnly.Name = "ChkShowInvalidOnly"
        Me.ChkShowInvalidOnly.Size = New System.Drawing.Size(163, 21)
        Me.ChkShowInvalidOnly.TabIndex = 12
        Me.ChkShowInvalidOnly.Text = "Show Difference Only"
        Me.ChkShowInvalidOnly.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.LlbUncheckAll)
        Me.GroupBox2.Controls.Add(Me.LlbCheckAll)
        Me.GroupBox2.Controls.Add(Me.dgvItem)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(858, 512)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "e-Faktur/Art Data"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(804, 588)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(63, 26)
        Me.btnClose.TabIndex = 14
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        Me.btnClose.Visible = False
        '
        'BackgroundWorker1
        '
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(15, 589)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(659, 23)
        Me.ProgressBar1.TabIndex = 15
        Me.ProgressBar1.Visible = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.DataGridView1)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(150, 46)
        Me.GroupBox3.TabIndex = 14
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Item Data"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colItem_No, Me.colItem_ProductCode, Me.colItem_ProductDesc, Me.colItem_Qty, Me.colItem_Unit, Me.colItem_UnitPrice, Me.colItem_DiscAmount, Me.colItem_Amount, Me.colItem_DPP, Me.colItem_PPN, Me.colItem_InvalidReason})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 18)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(144, 25)
        Me.DataGridView1.TabIndex = 9
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
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_Qty.DefaultCellStyle = DataGridViewCellStyle7
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
        'colItem_UnitPrice
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_UnitPrice.DefaultCellStyle = DataGridViewCellStyle8
        Me.colItem_UnitPrice.HeaderText = "Unit Price"
        Me.colItem_UnitPrice.Name = "colItem_UnitPrice"
        Me.colItem_UnitPrice.ReadOnly = True
        '
        'colItem_DiscAmount
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_DiscAmount.DefaultCellStyle = DataGridViewCellStyle9
        Me.colItem_DiscAmount.HeaderText = "Discount Amount"
        Me.colItem_DiscAmount.Name = "colItem_DiscAmount"
        Me.colItem_DiscAmount.ReadOnly = True
        '
        'colItem_Amount
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_Amount.DefaultCellStyle = DataGridViewCellStyle10
        Me.colItem_Amount.HeaderText = "Amount"
        Me.colItem_Amount.Name = "colItem_Amount"
        Me.colItem_Amount.ReadOnly = True
        '
        'colItem_DPP
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_DPP.DefaultCellStyle = DataGridViewCellStyle11
        Me.colItem_DPP.HeaderText = "DPP"
        Me.colItem_DPP.Name = "colItem_DPP"
        Me.colItem_DPP.ReadOnly = True
        '
        'colItem_PPN
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colItem_PPN.DefaultCellStyle = DataGridViewCellStyle12
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
        Me.SplitContainer1.Panel2Collapsed = True
        Me.SplitContainer1.Size = New System.Drawing.Size(858, 512)
        Me.SplitContainer1.SplitterDistance = 312
        Me.SplitContainer1.TabIndex = 12
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
        Me.colNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colNoAcuan
        '
        Me.colNoAcuan.HeaderText = "No Acuan"
        Me.colNoAcuan.Name = "colNoAcuan"
        Me.colNoAcuan.ReadOnly = True
        Me.colNoAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colInvoiceNumberSMESS
        '
        Me.colInvoiceNumberSMESS.HeaderText = "Invoice Number (SMESS)"
        Me.colInvoiceNumberSMESS.Name = "colInvoiceNumberSMESS"
        Me.colInvoiceNumberSMESS.ReadOnly = True
        Me.colInvoiceNumberSMESS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colInvoiceDateSMESS
        '
        DataGridViewCellStyle1.Format = "dd MMM yyyy"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.colInvoiceDateSMESS.DefaultCellStyle = DataGridViewCellStyle1
        Me.colInvoiceDateSMESS.HeaderText = "Date (SMESS)"
        Me.colInvoiceDateSMESS.Name = "colInvoiceDateSMESS"
        Me.colInvoiceDateSMESS.ReadOnly = True
        Me.colInvoiceDateSMESS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColCustomerCodeSMESS
        '
        Me.ColCustomerCodeSMESS.HeaderText = "Customer Code (SMESS)"
        Me.ColCustomerCodeSMESS.Name = "ColCustomerCodeSMESS"
        Me.ColCustomerCodeSMESS.ReadOnly = True
        Me.ColCustomerCodeSMESS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColCustomerNameSMESS
        '
        Me.ColCustomerNameSMESS.HeaderText = "Customer Name (SMESS)"
        Me.ColCustomerNameSMESS.Name = "ColCustomerNameSMESS"
        Me.ColCustomerNameSMESS.ReadOnly = True
        Me.ColCustomerNameSMESS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colCustomerSMESS
        '
        Me.colCustomerSMESS.FillWeight = 175.0!
        Me.colCustomerSMESS.HeaderText = "Customer NPWP Name (SMESS)"
        Me.colCustomerSMESS.Name = "colCustomerSMESS"
        Me.colCustomerSMESS.ReadOnly = True
        Me.colCustomerSMESS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.colCustomerSMESS.Width = 175
        '
        'colNPWPSMESS
        '
        Me.colNPWPSMESS.HeaderText = "NPWP (SMESS)"
        Me.colNPWPSMESS.Name = "colNPWPSMESS"
        Me.colNPWPSMESS.ReadOnly = True
        Me.colNPWPSMESS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colDPPSMESS
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colDPPSMESS.DefaultCellStyle = DataGridViewCellStyle2
        Me.colDPPSMESS.HeaderText = "DPP (SMESS)"
        Me.colDPPSMESS.Name = "colDPPSMESS"
        Me.colDPPSMESS.ReadOnly = True
        Me.colDPPSMESS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colPPNSMESS
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPPNSMESS.DefaultCellStyle = DataGridViewCellStyle3
        Me.colPPNSMESS.HeaderText = "PPN (SMESS)"
        Me.colPPNSMESS.Name = "colPPNSMESS"
        Me.colPPNSMESS.ReadOnly = True
        Me.colPPNSMESS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colInvoiceNumber
        '
        Me.colInvoiceNumber.HeaderText = "Invoice Number"
        Me.colInvoiceNumber.Name = "colInvoiceNumber"
        Me.colInvoiceNumber.ReadOnly = True
        Me.colInvoiceNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colTaxInvoiceNo
        '
        Me.colTaxInvoiceNo.HeaderText = "Tax Invoice No"
        Me.colTaxInvoiceNo.Name = "colTaxInvoiceNo"
        Me.colTaxInvoiceNo.ReadOnly = True
        Me.colTaxInvoiceNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colInvoiceDate
        '
        DataGridViewCellStyle4.Format = "dd MMM yyyy"
        DataGridViewCellStyle4.NullValue = "1 Jan 1950"
        Me.colInvoiceDate.DefaultCellStyle = DataGridViewCellStyle4
        Me.colInvoiceDate.HeaderText = "Date"
        Me.colInvoiceDate.Name = "colInvoiceDate"
        Me.colInvoiceDate.ReadOnly = True
        Me.colInvoiceDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colCustomer
        '
        Me.colCustomer.HeaderText = "Customer"
        Me.colCustomer.Name = "colCustomer"
        Me.colCustomer.ReadOnly = True
        Me.colCustomer.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colNPWP
        '
        Me.colNPWP.HeaderText = "NPWP"
        Me.colNPWP.Name = "colNPWP"
        Me.colNPWP.ReadOnly = True
        Me.colNPWP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colDPP
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colDPP.DefaultCellStyle = DataGridViewCellStyle5
        Me.colDPP.HeaderText = "DPP"
        Me.colDPP.Name = "colDPP"
        Me.colDPP.ReadOnly = True
        Me.colDPP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colPPN
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPPN.DefaultCellStyle = DataGridViewCellStyle6
        Me.colPPN.HeaderText = "PPN"
        Me.colPPN.Name = "colPPN"
        Me.colPPN.ReadOnly = True
        Me.colPPN.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colSelisihInvoice
        '
        Me.colSelisihInvoice.HeaderText = "Difference Invoice"
        Me.colSelisihInvoice.Name = "colSelisihInvoice"
        Me.colSelisihInvoice.ReadOnly = True
        '
        'ColSelisihNPWP
        '
        Me.ColSelisihNPWP.HeaderText = "Difference NPWP"
        Me.ColSelisihNPWP.Name = "ColSelisihNPWP"
        '
        'colSelisihCustomer
        '
        Me.colSelisihCustomer.HeaderText = "Difference Customer"
        Me.colSelisihCustomer.Name = "colSelisihCustomer"
        Me.colSelisihCustomer.ReadOnly = True
        '
        'colSelisihDPP
        '
        Me.colSelisihDPP.HeaderText = "Difference DPP"
        Me.colSelisihDPP.Name = "colSelisihDPP"
        Me.colSelisihDPP.ReadOnly = True
        '
        'colSelisihPPN
        '
        Me.colSelisihPPN.HeaderText = "Difference PPN"
        Me.colSelisihPPN.Name = "colSelisihPPN"
        Me.colSelisihPPN.ReadOnly = True
        '
        'FrmImport_CheckData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(882, 623)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCreateInvoice)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "FrmImport_CheckData"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form Check Import Data"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

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
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents LlbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkShowInvalidOnly As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents colItem_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_ProductCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_ProductDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_UnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_DiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_DPP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_PPN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colItem_InvalidReason As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents cboTipe As System.Windows.Forms.ComboBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colNoAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colInvoiceNumberSMESS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colInvoiceDateSMESS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerCodeSMESS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerNameSMESS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colCustomerSMESS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colNPWPSMESS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDPPSMESS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPPNSMESS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colInvoiceNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colTaxInvoiceNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colInvoiceDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colNPWP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDPP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPPN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colSelisihInvoice As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColSelisihNPWP As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colSelisihCustomer As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colSelisihDPP As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colSelisihPPN As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
