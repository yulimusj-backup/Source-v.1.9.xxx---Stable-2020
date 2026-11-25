<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSOC_Quick
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
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustShipto = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSONo = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyCancel = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCancel = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColRemarksItem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPoduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaruUnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefPrec_QtyDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDecimalMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtSONumber = New System.Windows.Forms.TextBox
        Me.btnFindSO = New System.Windows.Forms.Button
        Me.TxtPrimarySO = New System.Windows.Forms.TextBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DgvUser = New System.Windows.Forms.DataGridView
        Me.ColCreatedBy_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvCustomer = New System.Windows.Forms.DataGridView
        Me.ColC_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColC_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColC_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColC_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvCustomerGroup = New System.Windows.Forms.DataGridView
        Me.ColSG_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSG_SalesGroup = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSG_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSG_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.DgvSalesArea = New System.Windows.Forms.DataGridView
        Me.ColSA_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_SalesArea = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSA_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.DgvMarketSegmentation = New System.Windows.Forms.DataGridView
        Me.ColMS_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_MarketSeg = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMS_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.DgvCustomerType = New System.Windows.Forms.DataGridView
        Me.ColCT_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_CustType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage7 = New System.Windows.Forms.TabPage
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage8 = New System.Windows.Forms.TabPage
        Me.DgvSalesPerson = New System.Windows.Forms.DataGridView
        Me.SalesPerson_ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SalesPerson_ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SalesPerson_ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SalesPerson_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage9 = New System.Windows.Forms.TabPage
        Me.DgvShipto = New System.Windows.Forms.DataGridView
        Me.ColShiptoNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColShiptoCustCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColShiptoCustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColShiptoCustPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LlbClearFilter = New System.Windows.Forms.LinkLabel
        Me.Label6 = New System.Windows.Forms.Label
        Me.chkDirectSave = New System.Windows.Forms.CheckBox
        Me.TxtCancellationReason = New System.Windows.Forms.TextBox
        Me.cboCustomerFormat = New System.Windows.Forms.ComboBox
        Me.ChkIgnoreLowTolerance = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.TxtProductName = New System.Windows.Forms.TextBox
        Me.TxtProductCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnFindProduct = New System.Windows.Forms.Button
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.TxtSupplierName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtSupplierCode = New System.Windows.Forms.TextBox
        Me.TxtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.prg01 = New System.Windows.Forms.ProgressBar
        Me.LblProgress = New System.Windows.Forms.Label
        Me.bgwCreatedBy = New System.ComponentModel.BackgroundWorker
        Me.bgwCustSegmen = New System.ComponentModel.BackgroundWorker
        Me.bgwSalesArea = New System.ComponentModel.BackgroundWorker
        Me.bgwCustType = New System.ComponentModel.BackgroundWorker
        Me.bgwCustGroup = New System.ComponentModel.BackgroundWorker
        Me.bgwCustomer = New System.ComponentModel.BackgroundWorker
        Me.bgwUser = New System.ComponentModel.BackgroundWorker
        Me.bgwSalesPerson = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.bgwShipto = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvCustomerGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.DgvSalesArea, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.DgvMarketSegmentation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.DgvCustomerType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage8.SuspendLayout()
        CType(Me.DgvSalesPerson, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage9.SuspendLayout()
        CType(Me.DgvShipto, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCustomer, Me.ColCustShipto, Me.ColSONo, Me.ColDate, Me.ColCodeNo, Me.ColDesc, Me.ColQty, Me.ColUnit, Me.ColCurrency, Me.ColUnitPrice, Me.ColDiscPerct, Me.ColDiscAmount, Me.ColAmount, Me.ColQtyCancel, Me.ColCancel, Me.ColRemarksItem, Me.ColPrimaryPoduct, Me.ColPrimaryUnitMulti, Me.ColPrimarySO, Me.ColQtyDefault, Me.ColPrimaruUnitDefault, Me.ColRatioDefault, Me.ColRatioReference, Me.ColDefPrec_QtyDefault, Me.ColUnitDefault, Me.ColDecimalMulti})
        Me.dgvItem.Location = New System.Drawing.Point(9, 175)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(728, 304)
        Me.dgvItem.TabIndex = 3
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColCustomer
        '
        Me.ColCustomer.HeaderText = "Customer"
        Me.ColCustomer.Name = "ColCustomer"
        Me.ColCustomer.Width = 76
        '
        'ColCustShipto
        '
        Me.ColCustShipto.HeaderText = "Customer Ship to "
        Me.ColCustShipto.Name = "ColCustShipto"
        Me.ColCustShipto.Width = 115
        '
        'ColSONo
        '
        Me.ColSONo.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColSONo.HeaderText = "SO Number"
        Me.ColSONo.Name = "ColSONo"
        Me.ColSONo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColSONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColSONo.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColSONo.Width = 87
        '
        'ColDate
        '
        DataGridViewCellStyle8.Format = "dd MMM yyyy"
        Me.ColDate.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.Width = 55
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 74
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.Width = 85
        '
        'ColQty
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Width = 48
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 51
        '
        'ColCurrency
        '
        Me.ColCurrency.HeaderText = "Currency"
        Me.ColCurrency.Name = "ColCurrency"
        Me.ColCurrency.Width = 74
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.Width = 78
        '
        'ColDiscPerct
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscPerct.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColDiscPerct.HeaderText = "Discount %"
        Me.ColDiscPerct.Name = "ColDiscPerct"
        Me.ColDiscPerct.Width = 85
        '
        'ColDiscAmount
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscAmount.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColDiscAmount.HeaderText = "Disc. Amount"
        Me.ColDiscAmount.Name = "ColDiscAmount"
        Me.ColDiscAmount.Width = 95
        '
        'ColAmount
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.Width = 68
        '
        'ColQtyCancel
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyCancel.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColQtyCancel.HeaderText = "Cancel Qty"
        Me.ColQtyCancel.Name = "ColQtyCancel"
        Me.ColQtyCancel.Width = 84
        '
        'ColCancel
        '
        Me.ColCancel.HeaderText = "Cancel"
        Me.ColCancel.Name = "ColCancel"
        Me.ColCancel.Width = 46
        '
        'ColRemarksItem
        '
        Me.ColRemarksItem.HeaderText = "Remarks"
        Me.ColRemarksItem.Name = "ColRemarksItem"
        Me.ColRemarksItem.Width = 74
        '
        'ColPrimaryPoduct
        '
        Me.ColPrimaryPoduct.HeaderText = "PrimaryProduct"
        Me.ColPrimaryPoduct.Name = "ColPrimaryPoduct"
        Me.ColPrimaryPoduct.Visible = False
        Me.ColPrimaryPoduct.Width = 130
        '
        'ColPrimaryUnitMulti
        '
        Me.ColPrimaryUnitMulti.HeaderText = "Primary Unit Multi"
        Me.ColPrimaryUnitMulti.Name = "ColPrimaryUnitMulti"
        Me.ColPrimaryUnitMulti.Visible = False
        Me.ColPrimaryUnitMulti.Width = 143
        '
        'ColPrimarySO
        '
        Me.ColPrimarySO.HeaderText = "PrimarySO"
        Me.ColPrimarySO.Name = "ColPrimarySO"
        Me.ColPrimarySO.Visible = False
        Me.ColPrimarySO.Width = 101
        '
        'ColQtyDefault
        '
        Me.ColQtyDefault.HeaderText = "Qty Default"
        Me.ColQtyDefault.Name = "ColQtyDefault"
        Me.ColQtyDefault.Visible = False
        Me.ColQtyDefault.Width = 104
        '
        'ColPrimaruUnitDefault
        '
        Me.ColPrimaruUnitDefault.HeaderText = "Primary Unit Default"
        Me.ColPrimaruUnitDefault.Name = "ColPrimaruUnitDefault"
        Me.ColPrimaruUnitDefault.Visible = False
        Me.ColPrimaruUnitDefault.Width = 159
        '
        'ColRatioDefault
        '
        Me.ColRatioDefault.HeaderText = "Ratio Default"
        Me.ColRatioDefault.Name = "ColRatioDefault"
        Me.ColRatioDefault.Visible = False
        Me.ColRatioDefault.Width = 115
        '
        'ColRatioReference
        '
        Me.ColRatioReference.HeaderText = "Ratio Reference"
        Me.ColRatioReference.Name = "ColRatioReference"
        Me.ColRatioReference.Visible = False
        Me.ColRatioReference.Width = 136
        '
        'ColDefPrec_QtyDefault
        '
        Me.ColDefPrec_QtyDefault.HeaderText = "DefPrecDefault"
        Me.ColDefPrec_QtyDefault.Name = "ColDefPrec_QtyDefault"
        Me.ColDefPrec_QtyDefault.Visible = False
        Me.ColDefPrec_QtyDefault.Width = 129
        '
        'ColUnitDefault
        '
        Me.ColUnitDefault.HeaderText = "Unit Default"
        Me.ColUnitDefault.Name = "ColUnitDefault"
        Me.ColUnitDefault.Visible = False
        Me.ColUnitDefault.Width = 107
        '
        'ColDecimalMulti
        '
        Me.ColDecimalMulti.HeaderText = "Decimal Multi"
        Me.ColDecimalMulti.Name = "ColDecimalMulti"
        Me.ColDecimalMulti.ReadOnly = True
        Me.ColDecimalMulti.Visible = False
        Me.ColDecimalMulti.Width = 116
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 68)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "SO No"
        '
        'TxtSONumber
        '
        Me.TxtSONumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSONumber.Location = New System.Drawing.Point(53, 66)
        Me.TxtSONumber.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtSONumber.Name = "TxtSONumber"
        Me.TxtSONumber.Size = New System.Drawing.Size(113, 20)
        Me.TxtSONumber.TabIndex = 0
        '
        'btnFindSO
        '
        Me.btnFindSO.Location = New System.Drawing.Point(170, 66)
        Me.btnFindSO.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindSO.Name = "btnFindSO"
        Me.btnFindSO.Size = New System.Drawing.Size(36, 19)
        Me.btnFindSO.TabIndex = 2
        Me.btnFindSO.TabStop = False
        Me.btnFindSO.Text = "Find"
        Me.btnFindSO.UseVisualStyleBackColor = True
        '
        'TxtPrimarySO
        '
        Me.TxtPrimarySO.Location = New System.Drawing.Point(75, 66)
        Me.TxtPrimarySO.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimarySO.Name = "TxtPrimarySO"
        Me.TxtPrimarySO.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimarySO.TabIndex = 4
        Me.TxtPrimarySO.Text = "-1"
        Me.TxtPrimarySO.Visible = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(9, 10)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TabControl1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.LlbClearFilter)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.chkDirectSave)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtCancellationReason)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboCustomerFormat)
        Me.SplitContainer1.Panel2.Controls.Add(Me.ChkIgnoreLowTolerance)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnCancel)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtSONumber)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimarySO)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindSO)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Size = New System.Drawing.Size(728, 160)
        Me.SplitContainer1.SplitterDistance = 264
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage6)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage7)
        Me.TabControl1.Controls.Add(Me.TabPage8)
        Me.TabControl1.Controls.Add(Me.TabPage9)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(264, 160)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvUser)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Size = New System.Drawing.Size(258, 134)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Created By"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DgvUser
        '
        Me.DgvUser.AllowUserToResizeColumns = False
        Me.DgvUser.AllowUserToResizeRows = False
        Me.DgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCreatedBy_No, Me.ColCreatedBy, Me.ColCreatedBy_Name, Me.ColCreatedBy_Primary})
        Me.DgvUser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvUser.Location = New System.Drawing.Point(2, 2)
        Me.DgvUser.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DgvUser.MultiSelect = False
        Me.DgvUser.Name = "DgvUser"
        Me.DgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvUser.RowTemplate.Height = 24
        Me.DgvUser.Size = New System.Drawing.Size(254, 130)
        Me.DgvUser.TabIndex = 0
        '
        'ColCreatedBy_No
        '
        Me.ColCreatedBy_No.FillWeight = 35.0!
        Me.ColCreatedBy_No.HeaderText = "No"
        Me.ColCreatedBy_No.MinimumWidth = 32
        Me.ColCreatedBy_No.Name = "ColCreatedBy_No"
        Me.ColCreatedBy_No.ReadOnly = True
        Me.ColCreatedBy_No.Width = 32
        '
        'ColCreatedBy
        '
        Me.ColCreatedBy.HeaderText = "User Code"
        Me.ColCreatedBy.Name = "ColCreatedBy"
        '
        'ColCreatedBy_Name
        '
        Me.ColCreatedBy_Name.HeaderText = "Name"
        Me.ColCreatedBy_Name.Name = "ColCreatedBy_Name"
        Me.ColCreatedBy_Name.ReadOnly = True
        Me.ColCreatedBy_Name.Width = 70
        '
        'ColCreatedBy_Primary
        '
        Me.ColCreatedBy_Primary.HeaderText = "Primary"
        Me.ColCreatedBy_Primary.Name = "ColCreatedBy_Primary"
        Me.ColCreatedBy_Primary.ReadOnly = True
        Me.ColCreatedBy_Primary.Visible = False
        Me.ColCreatedBy_Primary.Width = 81
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvCustomer)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Size = New System.Drawing.Size(258, 134)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Customer"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvCustomer
        '
        Me.DgvCustomer.AllowUserToResizeColumns = False
        Me.DgvCustomer.AllowUserToResizeRows = False
        Me.DgvCustomer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCustomer.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvCustomer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCustomer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColC_No, Me.ColC_Code, Me.ColC_Desc, Me.ColC_Primary})
        Me.DgvCustomer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvCustomer.Location = New System.Drawing.Point(2, 2)
        Me.DgvCustomer.MultiSelect = False
        Me.DgvCustomer.Name = "DgvCustomer"
        Me.DgvCustomer.RowHeadersWidth = 24
        Me.DgvCustomer.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCustomer.RowTemplate.Height = 24
        Me.DgvCustomer.Size = New System.Drawing.Size(254, 130)
        Me.DgvCustomer.TabIndex = 6
        '
        'ColC_No
        '
        Me.ColC_No.FillWeight = 32.0!
        Me.ColC_No.HeaderText = "No"
        Me.ColC_No.MinimumWidth = 32
        Me.ColC_No.Name = "ColC_No"
        Me.ColC_No.ReadOnly = True
        Me.ColC_No.Width = 46
        '
        'ColC_Code
        '
        Me.ColC_Code.HeaderText = "Customer Code"
        Me.ColC_Code.MinimumWidth = 85
        Me.ColC_Code.Name = "ColC_Code"
        Me.ColC_Code.Width = 104
        '
        'ColC_Desc
        '
        Me.ColC_Desc.HeaderText = "Description"
        Me.ColC_Desc.MinimumWidth = 85
        Me.ColC_Desc.Name = "ColC_Desc"
        Me.ColC_Desc.Width = 85
        '
        'ColC_Primary
        '
        Me.ColC_Primary.HeaderText = "Primary"
        Me.ColC_Primary.Name = "ColC_Primary"
        Me.ColC_Primary.ReadOnly = True
        Me.ColC_Primary.Visible = False
        Me.ColC_Primary.Width = 81
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvCustomerGroup)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(258, 134)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Cust Group"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'dgvCustomerGroup
        '
        Me.dgvCustomerGroup.AllowUserToResizeColumns = False
        Me.dgvCustomerGroup.AllowUserToResizeRows = False
        Me.dgvCustomerGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCustomerGroup.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvCustomerGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCustomerGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSG_No, Me.ColSG_SalesGroup, Me.ColSG_Desc, Me.ColSG_Primary})
        Me.dgvCustomerGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCustomerGroup.Location = New System.Drawing.Point(0, 0)
        Me.dgvCustomerGroup.MultiSelect = False
        Me.dgvCustomerGroup.Name = "dgvCustomerGroup"
        Me.dgvCustomerGroup.RowHeadersWidth = 24
        Me.dgvCustomerGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvCustomerGroup.RowTemplate.Height = 24
        Me.dgvCustomerGroup.Size = New System.Drawing.Size(258, 134)
        Me.dgvCustomerGroup.TabIndex = 6
        '
        'ColSG_No
        '
        Me.ColSG_No.HeaderText = "No"
        Me.ColSG_No.MinimumWidth = 32
        Me.ColSG_No.Name = "ColSG_No"
        Me.ColSG_No.ReadOnly = True
        Me.ColSG_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSG_No.Width = 32
        '
        'ColSG_SalesGroup
        '
        Me.ColSG_SalesGroup.HeaderText = "Group Code"
        Me.ColSG_SalesGroup.MinimumWidth = 85
        Me.ColSG_SalesGroup.Name = "ColSG_SalesGroup"
        Me.ColSG_SalesGroup.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSG_SalesGroup.Width = 85
        '
        'ColSG_Desc
        '
        Me.ColSG_Desc.HeaderText = "Description"
        Me.ColSG_Desc.MinimumWidth = 85
        Me.ColSG_Desc.Name = "ColSG_Desc"
        Me.ColSG_Desc.ReadOnly = True
        Me.ColSG_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSG_Desc.Width = 85
        '
        'ColSG_Primary
        '
        Me.ColSG_Primary.HeaderText = "Primary"
        Me.ColSG_Primary.Name = "ColSG_Primary"
        Me.ColSG_Primary.ReadOnly = True
        Me.ColSG_Primary.Visible = False
        Me.ColSG_Primary.Width = 81
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.DgvSalesArea)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(258, 134)
        Me.TabPage6.TabIndex = 5
        Me.TabPage6.Text = "Cust Sales Area"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'DgvSalesArea
        '
        Me.DgvSalesArea.AllowUserToResizeColumns = False
        Me.DgvSalesArea.AllowUserToResizeRows = False
        Me.DgvSalesArea.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvSalesArea.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvSalesArea.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvSalesArea.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSA_No, Me.ColSA_SalesArea, Me.ColSA_Desc, Me.ColSA_Primary})
        Me.DgvSalesArea.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvSalesArea.Location = New System.Drawing.Point(0, 0)
        Me.DgvSalesArea.MultiSelect = False
        Me.DgvSalesArea.Name = "DgvSalesArea"
        Me.DgvSalesArea.RowHeadersWidth = 24
        Me.DgvSalesArea.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvSalesArea.RowTemplate.Height = 24
        Me.DgvSalesArea.Size = New System.Drawing.Size(258, 134)
        Me.DgvSalesArea.TabIndex = 6
        '
        'ColSA_No
        '
        Me.ColSA_No.HeaderText = "No"
        Me.ColSA_No.Name = "ColSA_No"
        Me.ColSA_No.ReadOnly = True
        Me.ColSA_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSA_No.Width = 27
        '
        'ColSA_SalesArea
        '
        Me.ColSA_SalesArea.HeaderText = "Sales Area Code"
        Me.ColSA_SalesArea.Name = "ColSA_SalesArea"
        Me.ColSA_SalesArea.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSA_SalesArea.Width = 92
        '
        'ColSA_Desc
        '
        Me.ColSA_Desc.HeaderText = "Description"
        Me.ColSA_Desc.Name = "ColSA_Desc"
        Me.ColSA_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSA_Desc.Width = 66
        '
        'ColSA_Primary
        '
        Me.ColSA_Primary.HeaderText = "Primary"
        Me.ColSA_Primary.Name = "ColSA_Primary"
        Me.ColSA_Primary.ReadOnly = True
        Me.ColSA_Primary.Visible = False
        Me.ColSA_Primary.Width = 81
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.DgvMarketSegmentation)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Size = New System.Drawing.Size(258, 134)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Cust Segmentation"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'DgvMarketSegmentation
        '
        Me.DgvMarketSegmentation.AllowUserToResizeRows = False
        Me.DgvMarketSegmentation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvMarketSegmentation.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvMarketSegmentation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvMarketSegmentation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColMS_No, Me.ColMS_MarketSeg, Me.ColMS_Desc, Me.ColMS_Primary})
        Me.DgvMarketSegmentation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvMarketSegmentation.Location = New System.Drawing.Point(0, 0)
        Me.DgvMarketSegmentation.MultiSelect = False
        Me.DgvMarketSegmentation.Name = "DgvMarketSegmentation"
        Me.DgvMarketSegmentation.RowHeadersWidth = 24
        Me.DgvMarketSegmentation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvMarketSegmentation.RowTemplate.Height = 24
        Me.DgvMarketSegmentation.Size = New System.Drawing.Size(258, 134)
        Me.DgvMarketSegmentation.TabIndex = 6
        '
        'ColMS_No
        '
        Me.ColMS_No.HeaderText = "No"
        Me.ColMS_No.Name = "ColMS_No"
        Me.ColMS_No.ReadOnly = True
        Me.ColMS_No.Width = 46
        '
        'ColMS_MarketSeg
        '
        Me.ColMS_MarketSeg.HeaderText = "Segmentation Code"
        Me.ColMS_MarketSeg.Name = "ColMS_MarketSeg"
        Me.ColMS_MarketSeg.Width = 125
        '
        'ColMS_Desc
        '
        Me.ColMS_Desc.HeaderText = "Description"
        Me.ColMS_Desc.Name = "ColMS_Desc"
        Me.ColMS_Desc.Width = 85
        '
        'ColMS_Primary
        '
        Me.ColMS_Primary.HeaderText = "Primary"
        Me.ColMS_Primary.Name = "ColMS_Primary"
        Me.ColMS_Primary.ReadOnly = True
        Me.ColMS_Primary.Visible = False
        Me.ColMS_Primary.Width = 81
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.DgvCustomerType)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(258, 134)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Cust Type"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'DgvCustomerType
        '
        Me.DgvCustomerType.AllowUserToResizeColumns = False
        Me.DgvCustomerType.AllowUserToResizeRows = False
        Me.DgvCustomerType.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCustomerType.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvCustomerType.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCustomerType.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCT_No, Me.ColCT_CustType, Me.ColCT_Desc, Me.ColCT_Primary})
        Me.DgvCustomerType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvCustomerType.Location = New System.Drawing.Point(0, 0)
        Me.DgvCustomerType.MultiSelect = False
        Me.DgvCustomerType.Name = "DgvCustomerType"
        Me.DgvCustomerType.RowHeadersWidth = 24
        Me.DgvCustomerType.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCustomerType.RowTemplate.Height = 24
        Me.DgvCustomerType.Size = New System.Drawing.Size(258, 134)
        Me.DgvCustomerType.TabIndex = 6
        '
        'ColCT_No
        '
        Me.ColCT_No.HeaderText = "No"
        Me.ColCT_No.Name = "ColCT_No"
        Me.ColCT_No.ReadOnly = True
        Me.ColCT_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCT_No.Width = 27
        '
        'ColCT_CustType
        '
        Me.ColCT_CustType.HeaderText = "Type Code"
        Me.ColCT_CustType.Name = "ColCT_CustType"
        Me.ColCT_CustType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCT_CustType.Width = 65
        '
        'ColCT_Desc
        '
        Me.ColCT_Desc.HeaderText = "Description"
        Me.ColCT_Desc.Name = "ColCT_Desc"
        Me.ColCT_Desc.ReadOnly = True
        Me.ColCT_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCT_Desc.Width = 66
        '
        'ColCT_Primary
        '
        Me.ColCT_Primary.HeaderText = "Primary"
        Me.ColCT_Primary.Name = "ColCT_Primary"
        Me.ColCT_Primary.ReadOnly = True
        Me.ColCT_Primary.Visible = False
        Me.ColCT_Primary.Width = 81
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.dgvItemProductGroup)
        Me.TabPage7.Location = New System.Drawing.Point(4, 22)
        Me.TabPage7.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage7.Size = New System.Drawing.Size(258, 134)
        Me.TabPage7.TabIndex = 6
        Me.TabPage7.Text = "Product Group"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(2, 2)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidth = 24
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(254, 130)
        Me.dgvItemProductGroup.TabIndex = 6
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.ReadOnly = True
        Me.ColProductGroup_No.Width = 46
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 39
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.HeaderText = "Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.Width = 57
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.ReadOnly = True
        Me.ColProductGroup_Desc.Width = 85
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.ReadOnly = True
        Me.ColProductGroup_Remarks.Visible = False
        Me.ColProductGroup_Remarks.Width = 89
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.ReadOnly = True
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 81
        '
        'TabPage8
        '
        Me.TabPage8.Controls.Add(Me.DgvSalesPerson)
        Me.TabPage8.Location = New System.Drawing.Point(4, 22)
        Me.TabPage8.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage8.Size = New System.Drawing.Size(256, 134)
        Me.TabPage8.TabIndex = 7
        Me.TabPage8.Text = "Sales Person"
        Me.TabPage8.UseVisualStyleBackColor = True
        '
        'DgvSalesPerson
        '
        Me.DgvSalesPerson.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvSalesPerson.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvSalesPerson.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvSalesPerson.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.SalesPerson_ColNo, Me.SalesPerson_ColCode, Me.SalesPerson_ColName, Me.SalesPerson_Primary})
        Me.DgvSalesPerson.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvSalesPerson.Location = New System.Drawing.Point(2, 2)
        Me.DgvSalesPerson.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DgvSalesPerson.MultiSelect = False
        Me.DgvSalesPerson.Name = "DgvSalesPerson"
        Me.DgvSalesPerson.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvSalesPerson.RowTemplate.Height = 24
        Me.DgvSalesPerson.Size = New System.Drawing.Size(252, 130)
        Me.DgvSalesPerson.TabIndex = 2
        '
        'SalesPerson_ColNo
        '
        Me.SalesPerson_ColNo.FillWeight = 50.0!
        Me.SalesPerson_ColNo.HeaderText = "No"
        Me.SalesPerson_ColNo.Name = "SalesPerson_ColNo"
        Me.SalesPerson_ColNo.Width = 46
        '
        'SalesPerson_ColCode
        '
        Me.SalesPerson_ColCode.HeaderText = "Sales Code"
        Me.SalesPerson_ColCode.Name = "SalesPerson_ColCode"
        Me.SalesPerson_ColCode.Width = 86
        '
        'SalesPerson_ColName
        '
        Me.SalesPerson_ColName.HeaderText = "Name"
        Me.SalesPerson_ColName.Name = "SalesPerson_ColName"
        Me.SalesPerson_ColName.Width = 60
        '
        'SalesPerson_Primary
        '
        Me.SalesPerson_Primary.HeaderText = "Primary"
        Me.SalesPerson_Primary.Name = "SalesPerson_Primary"
        Me.SalesPerson_Primary.Visible = False
        Me.SalesPerson_Primary.Width = 81
        '
        'TabPage9
        '
        Me.TabPage9.Controls.Add(Me.DgvShipto)
        Me.TabPage9.Location = New System.Drawing.Point(4, 22)
        Me.TabPage9.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage9.Name = "TabPage9"
        Me.TabPage9.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage9.Size = New System.Drawing.Size(252, 134)
        Me.TabPage9.TabIndex = 8
        Me.TabPage9.Text = "Cust Ship to"
        Me.TabPage9.UseVisualStyleBackColor = True
        '
        'DgvShipto
        '
        Me.DgvShipto.AllowUserToResizeColumns = False
        Me.DgvShipto.AllowUserToResizeRows = False
        Me.DgvShipto.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvShipto.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvShipto.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvShipto.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColShiptoNo, Me.ColShiptoCustCode, Me.ColShiptoCustName, Me.ColShiptoCustPrimary})
        Me.DgvShipto.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvShipto.Location = New System.Drawing.Point(2, 2)
        Me.DgvShipto.MultiSelect = False
        Me.DgvShipto.Name = "DgvShipto"
        Me.DgvShipto.RowHeadersWidth = 24
        Me.DgvShipto.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvShipto.RowTemplate.Height = 24
        Me.DgvShipto.Size = New System.Drawing.Size(248, 130)
        Me.DgvShipto.TabIndex = 7
        '
        'ColShiptoNo
        '
        Me.ColShiptoNo.FillWeight = 32.0!
        Me.ColShiptoNo.HeaderText = "No"
        Me.ColShiptoNo.MinimumWidth = 32
        Me.ColShiptoNo.Name = "ColShiptoNo"
        Me.ColShiptoNo.ReadOnly = True
        Me.ColShiptoNo.Width = 46
        '
        'ColShiptoCustCode
        '
        Me.ColShiptoCustCode.HeaderText = "Ship to Customer Code"
        Me.ColShiptoCustCode.MinimumWidth = 85
        Me.ColShiptoCustCode.Name = "ColShiptoCustCode"
        Me.ColShiptoCustCode.Width = 140
        '
        'ColShiptoCustName
        '
        Me.ColShiptoCustName.HeaderText = "Ship to Customer Name"
        Me.ColShiptoCustName.MinimumWidth = 85
        Me.ColShiptoCustName.Name = "ColShiptoCustName"
        Me.ColShiptoCustName.Width = 143
        '
        'ColShiptoCustPrimary
        '
        Me.ColShiptoCustPrimary.HeaderText = "Primary"
        Me.ColShiptoCustPrimary.Name = "ColShiptoCustPrimary"
        Me.ColShiptoCustPrimary.ReadOnly = True
        Me.ColShiptoCustPrimary.Visible = False
        Me.ColShiptoCustPrimary.Width = 81
        '
        'LlbClearFilter
        '
        Me.LlbClearFilter.AutoSize = True
        Me.LlbClearFilter.Location = New System.Drawing.Point(6, 141)
        Me.LlbClearFilter.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LlbClearFilter.Name = "LlbClearFilter"
        Me.LlbClearFilter.Size = New System.Drawing.Size(56, 13)
        Me.LlbClearFilter.TabIndex = 3
        Me.LlbClearFilter.TabStop = True
        Me.LlbClearFilter.Text = "Clear Filter"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(216, 19)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 13)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Reason"
        '
        'chkDirectSave
        '
        Me.chkDirectSave.AutoSize = True
        Me.chkDirectSave.Location = New System.Drawing.Point(272, 114)
        Me.chkDirectSave.Name = "chkDirectSave"
        Me.chkDirectSave.Size = New System.Drawing.Size(89, 17)
        Me.chkDirectSave.TabIndex = 7
        Me.chkDirectSave.Text = "Directly Save"
        Me.chkDirectSave.UseVisualStyleBackColor = True
        '
        'TxtCancellationReason
        '
        Me.TxtCancellationReason.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCancellationReason.Location = New System.Drawing.Point(272, 19)
        Me.TxtCancellationReason.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtCancellationReason.Multiline = True
        Me.TxtCancellationReason.Name = "TxtCancellationReason"
        Me.TxtCancellationReason.Size = New System.Drawing.Size(183, 88)
        Me.TxtCancellationReason.TabIndex = 6
        '
        'cboCustomerFormat
        '
        Me.cboCustomerFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustomerFormat.FormattingEnabled = True
        Me.cboCustomerFormat.Items.AddRange(New Object() {"Show Customer Code", "Show Customer Name", "Show Customer Code & Name"})
        Me.cboCustomerFormat.Location = New System.Drawing.Point(8, 112)
        Me.cboCustomerFormat.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboCustomerFormat.Name = "cboCustomerFormat"
        Me.cboCustomerFormat.Size = New System.Drawing.Size(198, 21)
        Me.cboCustomerFormat.TabIndex = 2
        '
        'ChkIgnoreLowTolerance
        '
        Me.ChkIgnoreLowTolerance.AutoSize = True
        Me.ChkIgnoreLowTolerance.Location = New System.Drawing.Point(53, 89)
        Me.ChkIgnoreLowTolerance.Name = "ChkIgnoreLowTolerance"
        Me.ChkIgnoreLowTolerance.Size = New System.Drawing.Size(139, 17)
        Me.ChkIgnoreLowTolerance.TabIndex = 1
        Me.ChkIgnoreLowTolerance.Text = "Ignore Lower Tolerance"
        Me.ChkIgnoreLowTolerance.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.dtpTo)
        Me.GroupBox1.Location = New System.Drawing.Point(2, 2)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(202, 61)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Sales Order Period"
        '
        'dtpFrom
        '
        Me.dtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(51, 17)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(145, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 39)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "To"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 18)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "From"
        '
        'dtpTo
        '
        Me.dtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(51, 38)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(145, 20)
        Me.dtpTo.TabIndex = 1
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(344, 136)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(115, 19)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Start Cancellation"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(154, 136)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(52, 19)
        Me.btnFilter.TabIndex = 5
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(98, 135)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(52, 19)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'TxtProductName
        '
        Me.TxtProductName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtProductName.Location = New System.Drawing.Point(362, 267)
        Me.TxtProductName.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtProductName.Name = "TxtProductName"
        Me.TxtProductName.ReadOnly = True
        Me.TxtProductName.Size = New System.Drawing.Size(202, 20)
        Me.TxtProductName.TabIndex = 14
        Me.TxtProductName.TabStop = False
        Me.TxtProductName.Visible = False
        '
        'TxtProductCode
        '
        Me.TxtProductCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtProductCode.Location = New System.Drawing.Point(179, 267)
        Me.TxtProductCode.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtProductCode.Name = "TxtProductCode"
        Me.TxtProductCode.Size = New System.Drawing.Size(127, 20)
        Me.TxtProductCode.TabIndex = 5
        Me.TxtProductCode.Visible = False
        '
        'TxtPrimaryProduct
        '
        Me.TxtPrimaryProduct.Location = New System.Drawing.Point(164, 268)
        Me.TxtPrimaryProduct.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryProduct.Name = "TxtPrimaryProduct"
        Me.TxtPrimaryProduct.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryProduct.TabIndex = 13
        Me.TxtPrimaryProduct.Text = "-1"
        Me.TxtPrimaryProduct.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(113, 270)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Code No"
        Me.Label5.Visible = False
        '
        'btnFindProduct
        '
        Me.btnFindProduct.Location = New System.Drawing.Point(310, 267)
        Me.btnFindProduct.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindProduct.Name = "btnFindProduct"
        Me.btnFindProduct.Size = New System.Drawing.Size(48, 19)
        Me.btnFindProduct.TabIndex = 6
        Me.btnFindProduct.TabStop = False
        Me.btnFindProduct.Text = "Find"
        Me.btnFindProduct.UseVisualStyleBackColor = True
        Me.btnFindProduct.Visible = False
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Location = New System.Drawing.Point(310, 245)
        Me.btnFindSupplier.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(48, 19)
        Me.btnFindSupplier.TabIndex = 4
        Me.btnFindSupplier.TabStop = False
        Me.btnFindSupplier.Text = "Find"
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        Me.btnFindSupplier.Visible = False
        '
        'TxtSupplierName
        '
        Me.TxtSupplierName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSupplierName.Location = New System.Drawing.Point(362, 245)
        Me.TxtSupplierName.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtSupplierName.Name = "TxtSupplierName"
        Me.TxtSupplierName.ReadOnly = True
        Me.TxtSupplierName.Size = New System.Drawing.Size(202, 20)
        Me.TxtSupplierName.TabIndex = 9
        Me.TxtSupplierName.TabStop = False
        Me.TxtSupplierName.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(113, 248)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Supplier"
        Me.Label2.Visible = False
        '
        'TxtSupplierCode
        '
        Me.TxtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSupplierCode.Location = New System.Drawing.Point(179, 245)
        Me.TxtSupplierCode.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtSupplierCode.Name = "TxtSupplierCode"
        Me.TxtSupplierCode.Size = New System.Drawing.Size(127, 20)
        Me.TxtSupplierCode.TabIndex = 3
        Me.TxtSupplierCode.Visible = False
        '
        'TxtPrimarySupplier
        '
        Me.TxtPrimarySupplier.Location = New System.Drawing.Point(164, 246)
        Me.TxtPrimarySupplier.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimarySupplier.Name = "TxtPrimarySupplier"
        Me.TxtPrimarySupplier.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimarySupplier.TabIndex = 8
        Me.TxtPrimarySupplier.Text = "-1"
        Me.TxtPrimarySupplier.Visible = False
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(9, 483)
        Me.llbCheckAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAll.TabIndex = 4
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(68, 483)
        Me.llbUncheckAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAll.TabIndex = 5
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'prg01
        '
        Me.prg01.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.prg01.Location = New System.Drawing.Point(134, 482)
        Me.prg01.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.prg01.Name = "prg01"
        Me.prg01.Size = New System.Drawing.Size(266, 19)
        Me.prg01.TabIndex = 16
        Me.prg01.Visible = False
        '
        'LblProgress
        '
        Me.LblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(405, 484)
        Me.LblProgress.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(39, 13)
        Me.LblProgress.TabIndex = 17
        Me.LblProgress.Text = "Label6"
        Me.LblProgress.Visible = False
        '
        'bgwCustSegmen
        '
        '
        'bgwSalesArea
        '
        '
        'bgwCustType
        '
        '
        'bgwCustGroup
        '
        '
        'bgwCustomer
        '
        '
        'bgwUser
        '
        '
        'bgwSalesPerson
        '
        '
        'BackgroundWorker3
        '
        '
        'bgwShipto
        '
        '
        'FrmTransaksiSOC_Quick
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(742, 514)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.prg01)
        Me.Controls.Add(Me.llbUncheckAll)
        Me.Controls.Add(Me.llbCheckAll)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.TxtProductName)
        Me.Controls.Add(Me.TxtPrimarySupplier)
        Me.Controls.Add(Me.TxtSupplierCode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtProductCode)
        Me.Controls.Add(Me.TxtSupplierName)
        Me.Controls.Add(Me.btnFindSupplier)
        Me.Controls.Add(Me.TxtPrimaryProduct)
        Me.Controls.Add(Me.btnFindProduct)
        Me.Controls.Add(Me.Label5)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimumSize = New System.Drawing.Size(758, 552)
        Me.Name = "FrmTransaksiSOC_Quick"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quick Sales Order Cancellation"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvCustomerGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        CType(Me.DgvSalesArea, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        CType(Me.DgvMarketSegmentation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.DgvCustomerType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage7.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage8.ResumeLayout(False)
        CType(Me.DgvSalesPerson, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage9.ResumeLayout(False)
        CType(Me.DgvShipto, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtSONumber As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSO As System.Windows.Forms.Button
    Friend WithEvents TxtPrimarySO As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents TxtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtProductName As System.Windows.Forms.TextBox
    Friend WithEvents TxtProductCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryProduct As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnFindProduct As System.Windows.Forms.Button
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents prg01 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents ChkIgnoreLowTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents cboCustomerFormat As System.Windows.Forms.ComboBox
    Friend WithEvents bgwCreatedBy As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtCancellationReason As System.Windows.Forms.TextBox
    Friend WithEvents chkDirectSave As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
    Friend WithEvents DgvSalesPerson As System.Windows.Forms.DataGridView
    Friend WithEvents LlbClearFilter As System.Windows.Forms.LinkLabel
    Friend WithEvents DgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents DgvCustomer As System.Windows.Forms.DataGridView
    Friend WithEvents DgvSalesArea As System.Windows.Forms.DataGridView
    Friend WithEvents ColSA_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_SalesArea As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSA_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvMarketSegmentation As System.Windows.Forms.DataGridView
    Friend WithEvents DgvCustomerType As System.Windows.Forms.DataGridView
    Friend WithEvents ColCT_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_CustType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents bgwCustSegmen As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwSalesArea As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwCustType As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwCustGroup As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwCustomer As System.ComponentModel.BackgroundWorker
    Friend WithEvents dgvCustomerGroup As System.Windows.Forms.DataGridView
    Friend WithEvents bgwUser As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgwSalesPerson As System.ComponentModel.BackgroundWorker
    Friend WithEvents SalesPerson_ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SalesPerson_ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SalesPerson_ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SalesPerson_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSG_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSG_SalesGroup As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSG_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSG_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_MarketSeg As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMS_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColC_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabPage9 As System.Windows.Forms.TabPage
    Friend WithEvents DgvShipto As System.Windows.Forms.DataGridView
    Friend WithEvents bgwShipto As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustShipto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSONo As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyCancel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCancel As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColRemarksItem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPoduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaruUnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefPrec_QtyDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDecimalMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShiptoNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShiptoCustCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShiptoCustName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShiptoCustPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
