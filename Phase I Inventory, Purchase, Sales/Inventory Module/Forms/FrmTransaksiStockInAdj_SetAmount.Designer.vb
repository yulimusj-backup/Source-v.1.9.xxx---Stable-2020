<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockInAdj_SetAmount
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
        Me.dgvProd = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroupCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroupDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBrandCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBrandDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSubBrandCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSubBrandDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCategoryCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCategoryDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColModifiy = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimaryProd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtUnitPriceMin = New System.Windows.Forms.TextBox
        Me.dgvGroup = New System.Windows.Forms.DataGridView
        Me.colProduct_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_GroupLevel = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.colProduct_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BgwProduct = New System.ComponentModel.BackgroundWorker
        Me.btnReset = New System.Windows.Forms.Button
        Me.bgwClose = New System.ComponentModel.BackgroundWorker
        Me.prg1 = New System.Windows.Forms.ProgressBar
        Me.lblProgess = New System.Windows.Forms.Label
        Me.bgwReset = New System.ComponentModel.BackgroundWorker
        Me.chkProductGroup = New System.Windows.Forms.CheckBox
        Me.chkShowBrand = New System.Windows.Forms.CheckBox
        Me.chkShowCategory = New System.Windows.Forms.CheckBox
        Me.chkShowSubBrand = New System.Windows.Forms.CheckBox
        Me.TxtUnitPriceMax = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtRef = New System.Windows.Forms.TextBox
        Me.btnSet = New System.Windows.Forms.Button
        Me.LlbCheckAll = New System.Windows.Forms.LinkLabel
        Me.LlbUncheckAll = New System.Windows.Forms.LinkLabel
        CType(Me.dgvProd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvProd
        '
        Me.dgvProd.AllowUserToAddRows = False
        Me.dgvProd.AllowUserToDeleteRows = False
        Me.dgvProd.AllowUserToResizeColumns = False
        Me.dgvProd.AllowUserToResizeRows = False
        Me.dgvProd.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvProd.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvProd.ColumnHeadersHeight = 25
        Me.dgvProd.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvProd.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColProductGroupCode, Me.ColProductGroupDesc, Me.ColBrandCode, Me.ColBrandDesc, Me.ColSubBrandCode, Me.ColSubBrandDesc, Me.ColCategoryCode, Me.ColCategoryDesc, Me.ColCodeNo, Me.ColDesc, Me.ColQty, Me.ColUnit, Me.ColUnitPrice, Me.ColAmount, Me.ColDetail, Me.ColModifiy, Me.ColPrimaryProd})
        Me.dgvProd.Location = New System.Drawing.Point(9, 163)
        Me.dgvProd.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvProd.Name = "dgvProd"
        Me.dgvProd.RowHeadersWidth = 24
        Me.dgvProd.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvProd.RowTemplate.Height = 24
        Me.dgvProd.Size = New System.Drawing.Size(726, 316)
        Me.dgvProd.TabIndex = 10
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColProductGroupCode
        '
        Me.ColProductGroupCode.HeaderText = "Product Group Code"
        Me.ColProductGroupCode.Name = "ColProductGroupCode"
        Me.ColProductGroupCode.Width = 129
        '
        'ColProductGroupDesc
        '
        Me.ColProductGroupDesc.HeaderText = "Product Group Desc"
        Me.ColProductGroupDesc.Name = "ColProductGroupDesc"
        Me.ColProductGroupDesc.Width = 129
        '
        'ColBrandCode
        '
        Me.ColBrandCode.HeaderText = "Brand Code"
        Me.ColBrandCode.Name = "ColBrandCode"
        Me.ColBrandCode.Width = 88
        '
        'ColBrandDesc
        '
        Me.ColBrandDesc.HeaderText = "Brand Desc"
        Me.ColBrandDesc.Name = "ColBrandDesc"
        Me.ColBrandDesc.Width = 88
        '
        'ColSubBrandCode
        '
        Me.ColSubBrandCode.HeaderText = "Sub Brand Code"
        Me.ColSubBrandCode.Name = "ColSubBrandCode"
        Me.ColSubBrandCode.Width = 110
        '
        'ColSubBrandDesc
        '
        Me.ColSubBrandDesc.HeaderText = "Sub Brand Desc"
        Me.ColSubBrandDesc.Name = "ColSubBrandDesc"
        Me.ColSubBrandDesc.Width = 110
        '
        'ColCategoryCode
        '
        Me.ColCategoryCode.HeaderText = "Category Code"
        Me.ColCategoryCode.Name = "ColCategoryCode"
        Me.ColCategoryCode.Width = 102
        '
        'ColCategoryDesc
        '
        Me.ColCategoryDesc.HeaderText = "Category Desc"
        Me.ColCategoryDesc.Name = "ColCategoryDesc"
        Me.ColCategoryDesc.Width = 102
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
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 29
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 32
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitPrice.Width = 59
        '
        'ColAmount
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAmount.Width = 49
        '
        'ColDetail
        '
        Me.ColDetail.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColDetail.HeaderText = "Detail"
        Me.ColDetail.Name = "ColDetail"
        Me.ColDetail.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColDetail.Width = 40
        '
        'ColModifiy
        '
        Me.ColModifiy.HeaderText = "Modify"
        Me.ColModifiy.Name = "ColModifiy"
        Me.ColModifiy.Width = 44
        '
        'ColPrimaryProd
        '
        Me.ColPrimaryProd.HeaderText = "Primary Prod"
        Me.ColPrimaryProd.Name = "ColPrimaryProd"
        Me.ColPrimaryProd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryProd.Visible = False
        Me.ColPrimaryProd.Width = 72
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(620, 10)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(88, 21)
        Me.btnFilter.TabIndex = 9
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(512, 253)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(56, 19)
        Me.btnClose.TabIndex = 12
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(413, 13)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Minimum Unit Price"
        '
        'TxtUnitPriceMin
        '
        Me.TxtUnitPriceMin.Location = New System.Drawing.Point(514, 11)
        Me.TxtUnitPriceMin.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtUnitPriceMin.Name = "TxtUnitPriceMin"
        Me.TxtUnitPriceMin.Size = New System.Drawing.Size(103, 20)
        Me.TxtUnitPriceMin.TabIndex = 2
        Me.TxtUnitPriceMin.Text = "0"
        Me.TxtUnitPriceMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dgvGroup
        '
        Me.dgvGroup.AllowUserToResizeColumns = False
        Me.dgvGroup.AllowUserToResizeRows = False
        Me.dgvGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvGroup.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvGroup.ColumnHeadersHeight = 25
        Me.dgvGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colProduct_No, Me.colProduct_GroupLevel, Me.colProduct_Code, Me.colProduct_Desc, Me.colProduct_Primary})
        Me.dgvGroup.Location = New System.Drawing.Point(9, 10)
        Me.dgvGroup.MultiSelect = False
        Me.dgvGroup.Name = "dgvGroup"
        Me.dgvGroup.RowHeadersWidth = 24
        Me.dgvGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvGroup.RowTemplate.Height = 24
        Me.dgvGroup.Size = New System.Drawing.Size(399, 148)
        Me.dgvGroup.TabIndex = 0
        '
        'colProduct_No
        '
        Me.colProduct_No.HeaderText = "No"
        Me.colProduct_No.Name = "colProduct_No"
        Me.colProduct_No.Width = 46
        '
        'colProduct_GroupLevel
        '
        Me.colProduct_GroupLevel.HeaderText = "Group Level"
        Me.colProduct_GroupLevel.Items.AddRange(New Object() {"Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.colProduct_GroupLevel.Name = "colProduct_GroupLevel"
        Me.colProduct_GroupLevel.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colProduct_GroupLevel.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.colProduct_GroupLevel.Width = 90
        '
        'colProduct_Code
        '
        Me.colProduct_Code.HeaderText = "Code"
        Me.colProduct_Code.Name = "colProduct_Code"
        Me.colProduct_Code.Width = 57
        '
        'colProduct_Desc
        '
        Me.colProduct_Desc.HeaderText = "Desc"
        Me.colProduct_Desc.Name = "colProduct_Desc"
        Me.colProduct_Desc.ReadOnly = True
        Me.colProduct_Desc.Width = 57
        '
        'colProduct_Primary
        '
        Me.colProduct_Primary.HeaderText = "Primary"
        Me.colProduct_Primary.Name = "colProduct_Primary"
        Me.colProduct_Primary.Visible = False
        Me.colProduct_Primary.Width = 66
        '
        'BgwProduct
        '
        '
        'btnReset
        '
        Me.btnReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnReset.Location = New System.Drawing.Point(9, 484)
        Me.btnReset.Margin = New System.Windows.Forms.Padding(2)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(53, 21)
        Me.btnReset.TabIndex = 11
        Me.btnReset.Text = "Reset"
        Me.btnReset.UseVisualStyleBackColor = True
        '
        'bgwClose
        '
        '
        'prg1
        '
        Me.prg1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.prg1.Location = New System.Drawing.Point(67, 485)
        Me.prg1.Margin = New System.Windows.Forms.Padding(2)
        Me.prg1.Name = "prg1"
        Me.prg1.Size = New System.Drawing.Size(341, 20)
        Me.prg1.TabIndex = 13
        '
        'lblProgess
        '
        Me.lblProgess.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblProgess.AutoSize = True
        Me.lblProgess.Location = New System.Drawing.Point(412, 489)
        Me.lblProgess.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblProgess.Name = "lblProgess"
        Me.lblProgess.Size = New System.Drawing.Size(39, 13)
        Me.lblProgess.TabIndex = 14
        Me.lblProgess.Text = "Label2"
        '
        'bgwReset
        '
        '
        'chkProductGroup
        '
        Me.chkProductGroup.AutoSize = True
        Me.chkProductGroup.Checked = True
        Me.chkProductGroup.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkProductGroup.Location = New System.Drawing.Point(416, 56)
        Me.chkProductGroup.Margin = New System.Windows.Forms.Padding(2)
        Me.chkProductGroup.Name = "chkProductGroup"
        Me.chkProductGroup.Size = New System.Drawing.Size(125, 17)
        Me.chkProductGroup.TabIndex = 5
        Me.chkProductGroup.Text = "Show Product Group"
        Me.chkProductGroup.UseVisualStyleBackColor = True
        '
        'chkShowBrand
        '
        Me.chkShowBrand.AutoSize = True
        Me.chkShowBrand.Checked = True
        Me.chkShowBrand.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowBrand.Location = New System.Drawing.Point(416, 78)
        Me.chkShowBrand.Margin = New System.Windows.Forms.Padding(2)
        Me.chkShowBrand.Name = "chkShowBrand"
        Me.chkShowBrand.Size = New System.Drawing.Size(84, 17)
        Me.chkShowBrand.TabIndex = 6
        Me.chkShowBrand.Text = "Show Brand"
        Me.chkShowBrand.UseVisualStyleBackColor = True
        '
        'chkShowCategory
        '
        Me.chkShowCategory.AutoSize = True
        Me.chkShowCategory.Checked = True
        Me.chkShowCategory.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowCategory.Location = New System.Drawing.Point(416, 122)
        Me.chkShowCategory.Margin = New System.Windows.Forms.Padding(2)
        Me.chkShowCategory.Name = "chkShowCategory"
        Me.chkShowCategory.Size = New System.Drawing.Size(98, 17)
        Me.chkShowCategory.TabIndex = 8
        Me.chkShowCategory.Text = "Show Category"
        Me.chkShowCategory.UseVisualStyleBackColor = True
        '
        'chkShowSubBrand
        '
        Me.chkShowSubBrand.AutoSize = True
        Me.chkShowSubBrand.Checked = True
        Me.chkShowSubBrand.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowSubBrand.Location = New System.Drawing.Point(416, 100)
        Me.chkShowSubBrand.Margin = New System.Windows.Forms.Padding(2)
        Me.chkShowSubBrand.Name = "chkShowSubBrand"
        Me.chkShowSubBrand.Size = New System.Drawing.Size(106, 17)
        Me.chkShowSubBrand.TabIndex = 7
        Me.chkShowSubBrand.Text = "Show Sub Brand"
        Me.chkShowSubBrand.UseVisualStyleBackColor = True
        '
        'TxtUnitPriceMax
        '
        Me.TxtUnitPriceMax.Location = New System.Drawing.Point(514, 33)
        Me.TxtUnitPriceMax.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtUnitPriceMax.Name = "TxtUnitPriceMax"
        Me.TxtUnitPriceMax.Size = New System.Drawing.Size(103, 20)
        Me.TxtUnitPriceMax.TabIndex = 4
        Me.TxtUnitPriceMax.Text = "0"
        Me.TxtUnitPriceMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(413, 36)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Maximum Unit Price"
        '
        'TxtRef
        '
        Me.TxtRef.Location = New System.Drawing.Point(572, 242)
        Me.TxtRef.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtRef.Name = "TxtRef"
        Me.TxtRef.Size = New System.Drawing.Size(52, 20)
        Me.TxtRef.TabIndex = 15
        Me.TxtRef.TabStop = False
        Me.TxtRef.Visible = False
        '
        'btnSet
        '
        Me.btnSet.Location = New System.Drawing.Point(620, 33)
        Me.btnSet.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSet.Name = "btnSet"
        Me.btnSet.Size = New System.Drawing.Size(88, 21)
        Me.btnSet.TabIndex = 16
        Me.btnSet.Text = "Set Unit Price"
        Me.btnSet.UseVisualStyleBackColor = True
        '
        'LlbCheckAll
        '
        Me.LlbCheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll.AutoSize = True
        Me.LlbCheckAll.Location = New System.Drawing.Point(617, 145)
        Me.LlbCheckAll.Name = "LlbCheckAll"
        Me.LlbCheckAll.Size = New System.Drawing.Size(52, 13)
        Me.LlbCheckAll.TabIndex = 17
        Me.LlbCheckAll.TabStop = True
        Me.LlbCheckAll.Text = "Check All"
        Me.LlbCheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LlbUncheckAll
        '
        Me.LlbUncheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll.AutoSize = True
        Me.LlbUncheckAll.Location = New System.Drawing.Point(670, 145)
        Me.LlbUncheckAll.Name = "LlbUncheckAll"
        Me.LlbUncheckAll.Size = New System.Drawing.Size(65, 13)
        Me.LlbUncheckAll.TabIndex = 18
        Me.LlbUncheckAll.TabStop = True
        Me.LlbUncheckAll.Text = "Uncheck All"
        Me.LlbUncheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'FrmTransaksiStockInAdj_SetAmount
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(744, 514)
        Me.Controls.Add(Me.LlbUncheckAll)
        Me.Controls.Add(Me.LlbCheckAll)
        Me.Controls.Add(Me.btnSet)
        Me.Controls.Add(Me.TxtUnitPriceMax)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.chkShowCategory)
        Me.Controls.Add(Me.chkShowSubBrand)
        Me.Controls.Add(Me.chkShowBrand)
        Me.Controls.Add(Me.chkProductGroup)
        Me.Controls.Add(Me.lblProgess)
        Me.Controls.Add(Me.prg1)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.dgvGroup)
        Me.Controls.Add(Me.TxtUnitPriceMin)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvProd)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.TxtRef)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmTransaksiStockInAdj_SetAmount"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set Begining Balance Amount"
        CType(Me.dgvProd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvProd As System.Windows.Forms.DataGridView
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtUnitPriceMin As System.Windows.Forms.TextBox
    Friend WithEvents dgvGroup As System.Windows.Forms.DataGridView
    Friend WithEvents colProduct_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_GroupLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colProduct_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwProduct As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents bgwClose As System.ComponentModel.BackgroundWorker
    Friend WithEvents prg1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblProgess As System.Windows.Forms.Label
    Friend WithEvents bgwReset As System.ComponentModel.BackgroundWorker
    Friend WithEvents chkProductGroup As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowBrand As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowCategory As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowSubBrand As System.Windows.Forms.CheckBox
    Friend WithEvents TxtUnitPriceMax As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtRef As System.Windows.Forms.TextBox
    Friend WithEvents btnSet As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroupCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroupDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBrandCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBrandDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSubBrandCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSubBrandDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCategoryCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCategoryDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColModifiy As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimaryProd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LlbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbUncheckAll As System.Windows.Forms.LinkLabel
End Class
