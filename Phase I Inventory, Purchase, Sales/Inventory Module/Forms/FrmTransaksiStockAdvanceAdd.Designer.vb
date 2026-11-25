<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockAdvanceAdd
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.cboRefDoc = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtRefDoc = New System.Windows.Forms.TextBox
        Me.btnFindRefDoc = New System.Windows.Forms.Button
        Me.TxtPrimaryRefDoc = New System.Windows.Forms.TextBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyRef = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyOnHand = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyAdd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecondaryQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecondaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWareHouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecondaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDecimalDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocationControl = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsService = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BtnFindLocationSource = New System.Windows.Forms.Button
        Me.TxtLocationSource = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtLocationSourceDesc = New System.Windows.Forms.TextBox
        Me.TxtWarehouseSourceDesc = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtWarehouseSource = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.llbRemoveAll = New System.Windows.Forms.LinkLabel
        Me.llbAddAll = New System.Windows.Forms.LinkLabel
        Me.rboAdd = New System.Windows.Forms.RadioButton
        Me.rboReplace = New System.Windows.Forms.RadioButton
        Me.TxtWarehouseTargetDesc = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtWarehouseTarget = New System.Windows.Forms.TextBox
        Me.TxtLocationTargetDesc = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.BtnFindLocationTarget = New System.Windows.Forms.Button
        Me.TxtLocationTarget = New System.Windows.Forms.TextBox
        Me.TxtPrimaryWarehouseTarget = New System.Windows.Forms.TextBox
        Me.TxtPrimaryLocationTarget = New System.Windows.Forms.TextBox
        Me.TxtPrimaryWarehouseSource = New System.Windows.Forms.TextBox
        Me.TxtPrimaryLocationSource = New System.Windows.Forms.TextBox
        Me.llbAddByQtyOnHand = New System.Windows.Forms.LinkLabel
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboRefDoc
        '
        Me.cboRefDoc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRefDoc.FormattingEnabled = True
        Me.cboRefDoc.Items.AddRange(New Object() {"Stock Bundling Item After", "Stock Bundling Item Before", "Stock In Adjust", "Stock In Purchase", "Stock In Sales Return", "Stock Out Adjust", "Stock Out Purchase Return", "Stock Out Sales", "Location Transfer"})
        Me.cboRefDoc.Location = New System.Drawing.Point(120, 10)
        Me.cboRefDoc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboRefDoc.Name = "cboRefDoc"
        Me.cboRefDoc.Size = New System.Drawing.Size(146, 21)
        Me.cboRefDoc.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 13)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(109, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Reference Document"
        '
        'TxtRefDoc
        '
        Me.TxtRefDoc.Location = New System.Drawing.Point(270, 11)
        Me.TxtRefDoc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtRefDoc.Name = "TxtRefDoc"
        Me.TxtRefDoc.Size = New System.Drawing.Size(167, 20)
        Me.TxtRefDoc.TabIndex = 1
        '
        'btnFindRefDoc
        '
        Me.btnFindRefDoc.Location = New System.Drawing.Point(440, 10)
        Me.btnFindRefDoc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFindRefDoc.Name = "btnFindRefDoc"
        Me.btnFindRefDoc.Size = New System.Drawing.Size(56, 20)
        Me.btnFindRefDoc.TabIndex = 4
        Me.btnFindRefDoc.TabStop = False
        Me.btnFindRefDoc.Text = "Find"
        Me.btnFindRefDoc.UseVisualStyleBackColor = True
        '
        'TxtPrimaryRefDoc
        '
        Me.TxtPrimaryRefDoc.Location = New System.Drawing.Point(414, 11)
        Me.TxtPrimaryRefDoc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryRefDoc.Name = "TxtPrimaryRefDoc"
        Me.TxtPrimaryRefDoc.Size = New System.Drawing.Size(23, 20)
        Me.TxtPrimaryRefDoc.TabIndex = 3
        Me.TxtPrimaryRefDoc.TabStop = False
        Me.TxtPrimaryRefDoc.Text = "-1"
        Me.TxtPrimaryRefDoc.Visible = False
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle6.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColQtyRef, Me.ColQtyOnHand, Me.ColQtyAdd, Me.ColUnit, Me.ColSecondaryQty, Me.ColSecondaryUnit, Me.ColBatch, Me.ColExpiredDate, Me.ColLocation, Me.ColWareHouse, Me.ColRemark, Me.ColPrimaryCodeNo, Me.ColPrimaryLocation, Me.ColPrimaryUnits, Me.ColPrimarySecondaryUnits, Me.ColDecimalDefault, Me.ColLocationControl, Me.ColIsService})
        Me.dgvItem.Location = New System.Drawing.Point(10, 126)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersVisible = False
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(742, 377)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 27
        Me.dgvItem.TabStop = False
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 74
        '
        'ColDescription
        '
        Me.ColDescription.Frozen = True
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.Width = 85
        '
        'ColQtyRef
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyRef.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColQtyRef.HeaderText = "Qty Reference"
        Me.ColQtyRef.Name = "ColQtyRef"
        Me.ColQtyRef.Width = 101
        '
        'ColQtyOnHand
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyOnHand.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColQtyOnHand.HeaderText = "Qty On Hand"
        Me.ColQtyOnHand.Name = "ColQtyOnHand"
        Me.ColQtyOnHand.Width = 94
        '
        'ColQtyAdd
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyAdd.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColQtyAdd.HeaderText = "Qty"
        Me.ColQtyAdd.Name = "ColQtyAdd"
        Me.ColQtyAdd.Width = 48
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 56
        '
        'ColSecondaryQty
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecondaryQty.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColSecondaryQty.HeaderText = "@"
        Me.ColSecondaryQty.Name = "ColSecondaryQty"
        Me.ColSecondaryQty.Visible = False
        Me.ColSecondaryQty.Width = 43
        '
        'ColSecondaryUnit
        '
        Me.ColSecondaryUnit.HeaderText = "Units"
        Me.ColSecondaryUnit.Name = "ColSecondaryUnit"
        Me.ColSecondaryUnit.Visible = False
        Me.ColSecondaryUnit.Width = 56
        '
        'ColBatch
        '
        Me.ColBatch.HeaderText = "Batch No"
        Me.ColBatch.Name = "ColBatch"
        Me.ColBatch.Visible = False
        Me.ColBatch.Width = 77
        '
        'ColExpiredDate
        '
        Me.ColExpiredDate.HeaderText = "Expiry Date"
        Me.ColExpiredDate.Name = "ColExpiredDate"
        Me.ColExpiredDate.Visible = False
        Me.ColExpiredDate.Width = 86
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.Visible = False
        Me.ColLocation.Width = 73
        '
        'ColWareHouse
        '
        Me.ColWareHouse.HeaderText = "WareHouse"
        Me.ColWareHouse.Name = "ColWareHouse"
        Me.ColWareHouse.Visible = False
        Me.ColWareHouse.Width = 89
        '
        'ColRemark
        '
        Me.ColRemark.HeaderText = "Remarks"
        Me.ColRemark.Name = "ColRemark"
        Me.ColRemark.Width = 74
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 105
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 110
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 90
        '
        'ColPrimarySecondaryUnits
        '
        Me.ColPrimarySecondaryUnits.HeaderText = "PrimarySecondaryUnits"
        Me.ColPrimarySecondaryUnits.Name = "ColPrimarySecondaryUnits"
        Me.ColPrimarySecondaryUnits.Visible = False
        Me.ColPrimarySecondaryUnits.Width = 141
        '
        'ColDecimalDefault
        '
        Me.ColDecimalDefault.HeaderText = "Decimal"
        Me.ColDecimalDefault.Name = "ColDecimalDefault"
        Me.ColDecimalDefault.Visible = False
        Me.ColDecimalDefault.Width = 70
        '
        'ColLocationControl
        '
        Me.ColLocationControl.HeaderText = "Location Control"
        Me.ColLocationControl.Name = "ColLocationControl"
        Me.ColLocationControl.ReadOnly = True
        Me.ColLocationControl.Visible = False
        Me.ColLocationControl.Width = 109
        '
        'ColIsService
        '
        Me.ColIsService.HeaderText = "Is Service"
        Me.ColIsService.Name = "ColIsService"
        Me.ColIsService.Visible = False
        Me.ColIsService.Width = 79
        '
        'BtnFindLocationSource
        '
        Me.BtnFindLocationSource.Location = New System.Drawing.Point(209, 33)
        Me.BtnFindLocationSource.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnFindLocationSource.Name = "BtnFindLocationSource"
        Me.BtnFindLocationSource.Size = New System.Drawing.Size(56, 20)
        Me.BtnFindLocationSource.TabIndex = 3
        Me.BtnFindLocationSource.TabStop = False
        Me.BtnFindLocationSource.Text = "Find"
        Me.BtnFindLocationSource.UseVisualStyleBackColor = True
        '
        'TxtLocationSource
        '
        Me.TxtLocationSource.Location = New System.Drawing.Point(120, 34)
        Me.TxtLocationSource.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtLocationSource.Name = "TxtLocationSource"
        Me.TxtLocationSource.Size = New System.Drawing.Size(86, 20)
        Me.TxtLocationSource.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 37)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Location"
        '
        'TxtLocationSourceDesc
        '
        Me.TxtLocationSourceDesc.AcceptsReturn = True
        Me.TxtLocationSourceDesc.Location = New System.Drawing.Point(270, 34)
        Me.TxtLocationSourceDesc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtLocationSourceDesc.Name = "TxtLocationSourceDesc"
        Me.TxtLocationSourceDesc.ReadOnly = True
        Me.TxtLocationSourceDesc.Size = New System.Drawing.Size(228, 20)
        Me.TxtLocationSourceDesc.TabIndex = 8
        Me.TxtLocationSourceDesc.TabStop = False
        '
        'TxtWarehouseSourceDesc
        '
        Me.TxtWarehouseSourceDesc.Location = New System.Drawing.Point(270, 57)
        Me.TxtWarehouseSourceDesc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtWarehouseSourceDesc.Name = "TxtWarehouseSourceDesc"
        Me.TxtWarehouseSourceDesc.ReadOnly = True
        Me.TxtWarehouseSourceDesc.Size = New System.Drawing.Size(228, 20)
        Me.TxtWarehouseSourceDesc.TabIndex = 9
        Me.TxtWarehouseSourceDesc.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 59)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Warehouse"
        '
        'TxtWarehouseSource
        '
        Me.TxtWarehouseSource.Location = New System.Drawing.Point(120, 57)
        Me.TxtWarehouseSource.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtWarehouseSource.Name = "TxtWarehouseSource"
        Me.TxtWarehouseSource.ReadOnly = True
        Me.TxtWarehouseSource.Size = New System.Drawing.Size(146, 20)
        Me.TxtWarehouseSource.TabIndex = 4
        Me.TxtWarehouseSource.TabStop = False
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(633, 10)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 20)
        Me.btnOK.TabIndex = 14
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(694, 10)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 20)
        Me.btnCancel.TabIndex = 15
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'llbRemoveAll
        '
        Me.llbRemoveAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbRemoveAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbRemoveAll.AutoSize = True
        Me.llbRemoveAll.Location = New System.Drawing.Point(693, 109)
        Me.llbRemoveAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbRemoveAll.Name = "llbRemoveAll"
        Me.llbRemoveAll.Size = New System.Drawing.Size(61, 13)
        Me.llbRemoveAll.TabIndex = 18
        Me.llbRemoveAll.TabStop = True
        Me.llbRemoveAll.Text = "Remove All"
        Me.llbRemoveAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbAddAll
        '
        Me.llbAddAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbAddAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAddAll.AutoSize = True
        Me.llbAddAll.Location = New System.Drawing.Point(651, 109)
        Me.llbAddAll.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbAddAll.Name = "llbAddAll"
        Me.llbAddAll.Size = New System.Drawing.Size(40, 13)
        Me.llbAddAll.TabIndex = 17
        Me.llbAddAll.TabStop = True
        Me.llbAddAll.Text = "Add All"
        Me.llbAddAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'rboAdd
        '
        Me.rboAdd.AutoSize = True
        Me.rboAdd.Checked = True
        Me.rboAdd.Location = New System.Drawing.Point(501, 11)
        Me.rboAdd.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rboAdd.Name = "rboAdd"
        Me.rboAdd.Size = New System.Drawing.Size(44, 17)
        Me.rboAdd.TabIndex = 12
        Me.rboAdd.TabStop = True
        Me.rboAdd.Text = "Add"
        Me.rboAdd.UseVisualStyleBackColor = True
        '
        'rboReplace
        '
        Me.rboReplace.AutoSize = True
        Me.rboReplace.Location = New System.Drawing.Point(546, 11)
        Me.rboReplace.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rboReplace.Name = "rboReplace"
        Me.rboReplace.Size = New System.Drawing.Size(65, 17)
        Me.rboReplace.TabIndex = 13
        Me.rboReplace.Text = "Replace"
        Me.rboReplace.UseVisualStyleBackColor = True
        '
        'TxtWarehouseTargetDesc
        '
        Me.TxtWarehouseTargetDesc.Location = New System.Drawing.Point(270, 102)
        Me.TxtWarehouseTargetDesc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtWarehouseTargetDesc.Name = "TxtWarehouseTargetDesc"
        Me.TxtWarehouseTargetDesc.ReadOnly = True
        Me.TxtWarehouseTargetDesc.Size = New System.Drawing.Size(228, 20)
        Me.TxtWarehouseTargetDesc.TabIndex = 11
        Me.TxtWarehouseTargetDesc.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 105)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 13)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Target Warehouse"
        '
        'TxtWarehouseTarget
        '
        Me.TxtWarehouseTarget.AcceptsReturn = True
        Me.TxtWarehouseTarget.Location = New System.Drawing.Point(120, 102)
        Me.TxtWarehouseTarget.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtWarehouseTarget.Name = "TxtWarehouseTarget"
        Me.TxtWarehouseTarget.ReadOnly = True
        Me.TxtWarehouseTarget.Size = New System.Drawing.Size(146, 20)
        Me.TxtWarehouseTarget.TabIndex = 7
        Me.TxtWarehouseTarget.TabStop = False
        '
        'TxtLocationTargetDesc
        '
        Me.TxtLocationTargetDesc.AcceptsReturn = True
        Me.TxtLocationTargetDesc.Location = New System.Drawing.Point(270, 80)
        Me.TxtLocationTargetDesc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtLocationTargetDesc.Name = "TxtLocationTargetDesc"
        Me.TxtLocationTargetDesc.ReadOnly = True
        Me.TxtLocationTargetDesc.Size = New System.Drawing.Size(228, 20)
        Me.TxtLocationTargetDesc.TabIndex = 10
        Me.TxtLocationTargetDesc.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 82)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(82, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Target Location"
        '
        'BtnFindLocationTarget
        '
        Me.BtnFindLocationTarget.Location = New System.Drawing.Point(209, 79)
        Me.BtnFindLocationTarget.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnFindLocationTarget.Name = "BtnFindLocationTarget"
        Me.BtnFindLocationTarget.Size = New System.Drawing.Size(56, 20)
        Me.BtnFindLocationTarget.TabIndex = 6
        Me.BtnFindLocationTarget.TabStop = False
        Me.BtnFindLocationTarget.Text = "Find"
        Me.BtnFindLocationTarget.UseVisualStyleBackColor = True
        '
        'TxtLocationTarget
        '
        Me.TxtLocationTarget.Location = New System.Drawing.Point(120, 80)
        Me.TxtLocationTarget.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtLocationTarget.Name = "TxtLocationTarget"
        Me.TxtLocationTarget.Size = New System.Drawing.Size(86, 20)
        Me.TxtLocationTarget.TabIndex = 5
        '
        'TxtPrimaryWarehouseTarget
        '
        Me.TxtPrimaryWarehouseTarget.Location = New System.Drawing.Point(502, 102)
        Me.TxtPrimaryWarehouseTarget.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryWarehouseTarget.Name = "TxtPrimaryWarehouseTarget"
        Me.TxtPrimaryWarehouseTarget.Size = New System.Drawing.Size(32, 20)
        Me.TxtPrimaryWarehouseTarget.TabIndex = 41
        Me.TxtPrimaryWarehouseTarget.TabStop = False
        Me.TxtPrimaryWarehouseTarget.Text = "-1"
        Me.TxtPrimaryWarehouseTarget.Visible = False
        '
        'TxtPrimaryLocationTarget
        '
        Me.TxtPrimaryLocationTarget.Location = New System.Drawing.Point(502, 80)
        Me.TxtPrimaryLocationTarget.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryLocationTarget.Name = "TxtPrimaryLocationTarget"
        Me.TxtPrimaryLocationTarget.Size = New System.Drawing.Size(32, 20)
        Me.TxtPrimaryLocationTarget.TabIndex = 40
        Me.TxtPrimaryLocationTarget.TabStop = False
        Me.TxtPrimaryLocationTarget.Text = "-1"
        Me.TxtPrimaryLocationTarget.Visible = False
        '
        'TxtPrimaryWarehouseSource
        '
        Me.TxtPrimaryWarehouseSource.Location = New System.Drawing.Point(502, 58)
        Me.TxtPrimaryWarehouseSource.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryWarehouseSource.Name = "TxtPrimaryWarehouseSource"
        Me.TxtPrimaryWarehouseSource.Size = New System.Drawing.Size(32, 20)
        Me.TxtPrimaryWarehouseSource.TabIndex = 39
        Me.TxtPrimaryWarehouseSource.TabStop = False
        Me.TxtPrimaryWarehouseSource.Text = "-1"
        Me.TxtPrimaryWarehouseSource.Visible = False
        '
        'TxtPrimaryLocationSource
        '
        Me.TxtPrimaryLocationSource.Location = New System.Drawing.Point(502, 36)
        Me.TxtPrimaryLocationSource.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtPrimaryLocationSource.Name = "TxtPrimaryLocationSource"
        Me.TxtPrimaryLocationSource.Size = New System.Drawing.Size(32, 20)
        Me.TxtPrimaryLocationSource.TabIndex = 38
        Me.TxtPrimaryLocationSource.TabStop = False
        Me.TxtPrimaryLocationSource.Text = "-1"
        Me.TxtPrimaryLocationSource.Visible = False
        '
        'llbAddByQtyOnHand
        '
        Me.llbAddByQtyOnHand.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbAddByQtyOnHand.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAddByQtyOnHand.AutoSize = True
        Me.llbAddByQtyOnHand.Location = New System.Drawing.Point(543, 109)
        Me.llbAddByQtyOnHand.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbAddByQtyOnHand.Name = "llbAddByQtyOnHand"
        Me.llbAddByQtyOnHand.Size = New System.Drawing.Size(106, 13)
        Me.llbAddByQtyOnHand.TabIndex = 16
        Me.llbAddByQtyOnHand.TabStop = True
        Me.llbAddByQtyOnHand.Text = "Add By Qty On Hand"
        Me.llbAddByQtyOnHand.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'FrmTransaksiStockAdvanceAdd
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(764, 515)
        Me.Controls.Add(Me.llbAddByQtyOnHand)
        Me.Controls.Add(Me.TxtPrimaryWarehouseTarget)
        Me.Controls.Add(Me.TxtPrimaryLocationTarget)
        Me.Controls.Add(Me.TxtPrimaryWarehouseSource)
        Me.Controls.Add(Me.TxtPrimaryLocationSource)
        Me.Controls.Add(Me.TxtWarehouseTargetDesc)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtWarehouseTarget)
        Me.Controls.Add(Me.TxtLocationTargetDesc)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.BtnFindLocationTarget)
        Me.Controls.Add(Me.TxtLocationTarget)
        Me.Controls.Add(Me.rboReplace)
        Me.Controls.Add(Me.rboAdd)
        Me.Controls.Add(Me.llbAddAll)
        Me.Controls.Add(Me.llbRemoveAll)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.TxtWarehouseSourceDesc)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TxtWarehouseSource)
        Me.Controls.Add(Me.TxtLocationSourceDesc)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.BtnFindLocationSource)
        Me.Controls.Add(Me.TxtLocationSource)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.TxtPrimaryRefDoc)
        Me.Controls.Add(Me.btnFindRefDoc)
        Me.Controls.Add(Me.TxtRefDoc)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboRefDoc)
        Me.Controls.Add(Me.btnOK)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(758, 552)
        Me.Name = "FrmTransaksiStockAdvanceAdd"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add By Other Document"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboRefDoc As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtRefDoc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindRefDoc As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryRefDoc As System.Windows.Forms.TextBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BtnFindLocationSource As System.Windows.Forms.Button
    Friend WithEvents TxtLocationSource As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtLocationSourceDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtWarehouseSourceDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouseSource As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents llbRemoveAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbAddAll As System.Windows.Forms.LinkLabel
    Friend WithEvents rboAdd As System.Windows.Forms.RadioButton
    Friend WithEvents rboReplace As System.Windows.Forms.RadioButton
    Friend WithEvents TxtWarehouseTargetDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouseTarget As System.Windows.Forms.TextBox
    Friend WithEvents TxtLocationTargetDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents BtnFindLocationTarget As System.Windows.Forms.Button
    Friend WithEvents TxtLocationTarget As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryWarehouseTarget As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryLocationTarget As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryWarehouseSource As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryLocationSource As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyRef As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyOnHand As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyAdd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecondaryQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecondaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWareHouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecondaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDecimalDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocationControl As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsService As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbAddByQtyOnHand As System.Windows.Forms.LinkLabel
End Class
