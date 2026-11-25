<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPOC_QuickDel
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
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle17 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle18 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle19 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle20 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle21 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupplier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONo = New System.Windows.Forms.DataGridViewLinkColumn
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
        Me.ColPrimaryPO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaruUnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefPrec_QtyDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDecimalMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtPONumber = New System.Windows.Forms.TextBox
        Me.btnFindPO = New System.Windows.Forms.Button
        Me.TxtPrimaryPO = New System.Windows.Forms.TextBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvCreatedBy = New System.Windows.Forms.DataGridView
        Me.ColUserNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUser_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvSupplier = New System.Windows.Forms.DataGridView
        Me.ColSupp_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupp_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupp_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupp_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupp_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LlbClearFilter = New System.Windows.Forms.LinkLabel
        Me.chkDirectSave = New System.Windows.Forms.CheckBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtCancellationReason = New System.Windows.Forms.TextBox
        Me.cboSupplierFormat = New System.Windows.Forms.ComboBox
        Me.ChkIgnoreLowTolerance = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.prg01 = New System.Windows.Forms.ProgressBar
        Me.LblProgress = New System.Windows.Forms.Label
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BgwSupplier = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvCreatedBy, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvSupplier, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColSupplier, Me.ColPONo, Me.ColDate, Me.ColCodeNo, Me.ColDesc, Me.ColQty, Me.ColUnit, Me.ColCurrency, Me.ColUnitPrice, Me.ColDiscPerct, Me.ColDiscAmount, Me.ColAmount, Me.ColQtyCancel, Me.ColCancel, Me.ColRemarksItem, Me.ColPrimaryPoduct, Me.ColPrimaryUnitMulti, Me.ColPrimaryPO, Me.ColQtyDefault, Me.ColPrimaruUnitDefault, Me.ColRatioDefault, Me.ColRatioReference, Me.ColDefPrec_QtyDefault, Me.ColUnitDefault, Me.ColDecimalMulti})
        Me.dgvItem.Location = New System.Drawing.Point(9, 175)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(727, 315)
        Me.dgvItem.TabIndex = 3
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColSupplier
        '
        Me.ColSupplier.HeaderText = "Supplier"
        Me.ColSupplier.Name = "ColSupplier"
        Me.ColSupplier.Width = 70
        '
        'ColPONo
        '
        Me.ColPONo.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColPONo.HeaderText = "PO Number"
        Me.ColPONo.Name = "ColPONo"
        Me.ColPONo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColPONo.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColPONo.Width = 87
        '
        'ColDate
        '
        DataGridViewCellStyle15.Format = "dd MMM yyyy"
        Me.ColDate.DefaultCellStyle = DataGridViewCellStyle15
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
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle16
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
        DataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle17
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.Width = 78
        '
        'ColDiscPerct
        '
        DataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscPerct.DefaultCellStyle = DataGridViewCellStyle18
        Me.ColDiscPerct.HeaderText = "Discount %"
        Me.ColDiscPerct.Name = "ColDiscPerct"
        Me.ColDiscPerct.Width = 85
        '
        'ColDiscAmount
        '
        DataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscAmount.DefaultCellStyle = DataGridViewCellStyle19
        Me.ColDiscAmount.HeaderText = "Disc. Amount"
        Me.ColDiscAmount.Name = "ColDiscAmount"
        Me.ColDiscAmount.Width = 95
        '
        'ColAmount
        '
        DataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle20
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.Width = 68
        '
        'ColQtyCancel
        '
        DataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyCancel.DefaultCellStyle = DataGridViewCellStyle21
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
        'ColPrimaryPO
        '
        Me.ColPrimaryPO.HeaderText = "PrimaryPO"
        Me.ColPrimaryPO.Name = "ColPrimaryPO"
        Me.ColPrimaryPO.Visible = False
        Me.ColPrimaryPO.Width = 101
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
        Me.Label1.Location = New System.Drawing.Point(14, 64)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "PO Number"
        '
        'TxtPONumber
        '
        Me.TxtPONumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPONumber.Location = New System.Drawing.Point(80, 66)
        Me.TxtPONumber.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPONumber.Name = "TxtPONumber"
        Me.TxtPONumber.Size = New System.Drawing.Size(124, 20)
        Me.TxtPONumber.TabIndex = 0
        '
        'btnFindPO
        '
        Me.btnFindPO.Location = New System.Drawing.Point(207, 65)
        Me.btnFindPO.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindPO.Name = "btnFindPO"
        Me.btnFindPO.Size = New System.Drawing.Size(36, 19)
        Me.btnFindPO.TabIndex = 2
        Me.btnFindPO.TabStop = False
        Me.btnFindPO.Text = "Find"
        Me.btnFindPO.UseVisualStyleBackColor = True
        '
        'TxtPrimaryPO
        '
        Me.TxtPrimaryPO.Location = New System.Drawing.Point(73, 66)
        Me.TxtPrimaryPO.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryPO.Name = "TxtPrimaryPO"
        Me.TxtPrimaryPO.Size = New System.Drawing.Size(26, 20)
        Me.TxtPrimaryPO.TabIndex = 4
        Me.TxtPrimaryPO.Text = "-1"
        Me.TxtPrimaryPO.Visible = False
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.chkDirectSave)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtCancellationReason)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboSupplierFormat)
        Me.SplitContainer1.Panel2.Controls.Add(Me.ChkIgnoreLowTolerance)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnCancel)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPONumber)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindPO)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryPO)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Size = New System.Drawing.Size(727, 160)
        Me.SplitContainer1.SplitterDistance = 237
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(237, 160)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvCreatedBy)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Size = New System.Drawing.Size(229, 134)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Created By"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvCreatedBy
        '
        Me.dgvCreatedBy.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCreatedBy.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCreatedBy.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColUserNo, Me.ColUserCode, Me.ColUserName, Me.ColUser_Primary})
        Me.dgvCreatedBy.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCreatedBy.Location = New System.Drawing.Point(2, 2)
        Me.dgvCreatedBy.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvCreatedBy.MultiSelect = False
        Me.dgvCreatedBy.Name = "dgvCreatedBy"
        Me.dgvCreatedBy.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvCreatedBy.RowTemplate.Height = 24
        Me.dgvCreatedBy.Size = New System.Drawing.Size(225, 130)
        Me.dgvCreatedBy.TabIndex = 0
        '
        'ColUserNo
        '
        Me.ColUserNo.HeaderText = "No"
        Me.ColUserNo.Name = "ColUserNo"
        Me.ColUserNo.Width = 46
        '
        'ColUserCode
        '
        Me.ColUserCode.HeaderText = "User Code"
        Me.ColUserCode.Name = "ColUserCode"
        Me.ColUserCode.Width = 82
        '
        'ColUserName
        '
        Me.ColUserName.HeaderText = "Name"
        Me.ColUserName.Name = "ColUserName"
        Me.ColUserName.Width = 60
        '
        'ColUser_Primary
        '
        Me.ColUser_Primary.HeaderText = "Primary"
        Me.ColUser_Primary.Name = "ColUser_Primary"
        Me.ColUser_Primary.Visible = False
        Me.ColUser_Primary.Width = 81
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvSupplier)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Size = New System.Drawing.Size(229, 134)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Supplier"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvSupplier
        '
        Me.DgvSupplier.AllowUserToResizeColumns = False
        Me.DgvSupplier.AllowUserToResizeRows = False
        Me.DgvSupplier.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvSupplier.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvSupplier.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSupp_No, Me.ColSupp_Code, Me.ColSupp_Name, Me.ColSupp_Remarks, Me.ColSupp_Primary})
        Me.DgvSupplier.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvSupplier.Location = New System.Drawing.Point(2, 2)
        Me.DgvSupplier.MultiSelect = False
        Me.DgvSupplier.Name = "DgvSupplier"
        Me.DgvSupplier.RowHeadersWidth = 24
        Me.DgvSupplier.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvSupplier.RowTemplate.Height = 24
        Me.DgvSupplier.Size = New System.Drawing.Size(225, 130)
        Me.DgvSupplier.TabIndex = 6
        '
        'ColSupp_No
        '
        Me.ColSupp_No.HeaderText = "No"
        Me.ColSupp_No.Name = "ColSupp_No"
        Me.ColSupp_No.ReadOnly = True
        Me.ColSupp_No.Width = 46
        '
        'ColSupp_Code
        '
        Me.ColSupp_Code.HeaderText = "Supplier Code"
        Me.ColSupp_Code.Name = "ColSupp_Code"
        Me.ColSupp_Code.Width = 98
        '
        'ColSupp_Name
        '
        Me.ColSupp_Name.HeaderText = "Supplier Name"
        Me.ColSupp_Name.Name = "ColSupp_Name"
        Me.ColSupp_Name.Width = 101
        '
        'ColSupp_Remarks
        '
        Me.ColSupp_Remarks.HeaderText = "Remarks"
        Me.ColSupp_Remarks.Name = "ColSupp_Remarks"
        Me.ColSupp_Remarks.ReadOnly = True
        Me.ColSupp_Remarks.Visible = False
        Me.ColSupp_Remarks.Width = 89
        '
        'ColSupp_Primary
        '
        Me.ColSupp_Primary.HeaderText = "Primary"
        Me.ColSupp_Primary.Name = "ColSupp_Primary"
        Me.ColSupp_Primary.Visible = False
        Me.ColSupp_Primary.Width = 81
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvItemProductGroup)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage3.Size = New System.Drawing.Size(229, 134)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Product"
        Me.TabPage3.UseVisualStyleBackColor = True
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
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(225, 130)
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
        'LlbClearFilter
        '
        Me.LlbClearFilter.AutoSize = True
        Me.LlbClearFilter.Location = New System.Drawing.Point(14, 136)
        Me.LlbClearFilter.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LlbClearFilter.Name = "LlbClearFilter"
        Me.LlbClearFilter.Size = New System.Drawing.Size(56, 13)
        Me.LlbClearFilter.TabIndex = 3
        Me.LlbClearFilter.TabStop = True
        Me.LlbClearFilter.Text = "Clear Filter"
        '
        'chkDirectSave
        '
        Me.chkDirectSave.AutoSize = True
        Me.chkDirectSave.Location = New System.Drawing.Point(302, 111)
        Me.chkDirectSave.Name = "chkDirectSave"
        Me.chkDirectSave.Size = New System.Drawing.Size(89, 17)
        Me.chkDirectSave.TabIndex = 7
        Me.chkDirectSave.Text = "Directly Save"
        Me.chkDirectSave.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(254, 11)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 13)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Reason"
        '
        'TxtCancellationReason
        '
        Me.TxtCancellationReason.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCancellationReason.Location = New System.Drawing.Point(302, 11)
        Me.TxtCancellationReason.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtCancellationReason.MaxLength = 250
        Me.TxtCancellationReason.Multiline = True
        Me.TxtCancellationReason.Name = "TxtCancellationReason"
        Me.TxtCancellationReason.Size = New System.Drawing.Size(183, 88)
        Me.TxtCancellationReason.TabIndex = 6
        '
        'cboSupplierFormat
        '
        Me.cboSupplierFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSupplierFormat.FormattingEnabled = True
        Me.cboSupplierFormat.Items.AddRange(New Object() {"Show Supplier Code", "Show Supplier Name", "Show Supplier Code & Name"})
        Me.cboSupplierFormat.Location = New System.Drawing.Point(16, 109)
        Me.cboSupplierFormat.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboSupplierFormat.Name = "cboSupplierFormat"
        Me.cboSupplierFormat.Size = New System.Drawing.Size(228, 21)
        Me.cboSupplierFormat.TabIndex = 2
        '
        'ChkIgnoreLowTolerance
        '
        Me.ChkIgnoreLowTolerance.AutoSize = True
        Me.ChkIgnoreLowTolerance.Location = New System.Drawing.Point(80, 89)
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
        Me.GroupBox1.Location = New System.Drawing.Point(15, 2)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(228, 61)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Purchase Order Period"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(39, 17)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(183, 20)
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
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(39, 38)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(183, 20)
        Me.dtpTo.TabIndex = 1
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(191, 133)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(52, 19)
        Me.btnFilter.TabIndex = 5
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(369, 133)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(115, 19)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Start Cancellation"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(135, 133)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(52, 19)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'llbCheckAll
        '
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(9, 490)
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
        Me.llbUncheckAll.Location = New System.Drawing.Point(68, 490)
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
        Me.prg01.Location = New System.Drawing.Point(134, 493)
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
        Me.LblProgress.Location = New System.Drawing.Point(405, 495)
        Me.LblProgress.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(39, 13)
        Me.LblProgress.TabIndex = 17
        Me.LblProgress.Text = "Label6"
        Me.LblProgress.Visible = False
        '
        'BackgroundWorker1
        '
        '
        'BgwSupplier
        '
        '
        'BackgroundWorker3
        '
        '
        'FrmTransaksiPOC_QuickDel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(740, 514)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.prg01)
        Me.Controls.Add(Me.llbUncheckAll)
        Me.Controls.Add(Me.llbCheckAll)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.dgvItem)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimumSize = New System.Drawing.Size(756, 552)
        Me.Name = "FrmTransaksiPOC_QuickDel"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quick Purchase Order Cancellation"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvCreatedBy, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvSupplier, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPONumber As System.Windows.Forms.TextBox
    Friend WithEvents btnFindPO As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryPO As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents prg01 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents dgvCreatedBy As System.Windows.Forms.DataGridView
    Friend WithEvents ChkIgnoreLowTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents cboSupplierFormat As System.Windows.Forms.ComboBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtCancellationReason As System.Windows.Forms.TextBox
    Friend WithEvents chkDirectSave As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents LlbClearFilter As System.Windows.Forms.LinkLabel
    Friend WithEvents DgvSupplier As System.Windows.Forms.DataGridView
    Friend WithEvents ColSupp_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupp_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupp_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupp_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupp_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwSupplier As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColUserNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUser_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupplier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONo As System.Windows.Forms.DataGridViewLinkColumn
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
    Friend WithEvents ColPrimaryPO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaruUnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefPrec_QtyDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDecimalMulti As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
