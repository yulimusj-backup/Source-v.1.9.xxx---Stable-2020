<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPackingList
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
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtPrimaryPO = New System.Windows.Forms.TextBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitCodeMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyPerMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMasterUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnNew = New System.Windows.Forms.Button
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtPrimaryUnitsDefault = New System.Windows.Forms.TextBox
        Me.lblUnitsMulti = New System.Windows.Forms.Label
        Me.TxtQtyperBatchMulti = New System.Windows.Forms.TextBox
        Me.dtpETD = New System.Windows.Forms.DateTimePicker
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.TxtSupplierName = New System.Windows.Forms.TextBox
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TxtQtyperBatchDefault = New System.Windows.Forms.TextBox
        Me.lblUnitsDefault = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtPackingListNo = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimaryPackingList = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtSupplier_GITNo = New System.Windows.Forms.TextBox
        Me.dtpReceivedDate = New System.Windows.Forms.DateTimePicker
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPrimaryUnitsMulti = New System.Windows.Forms.TextBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.TxtRatioDefault = New System.Windows.Forms.TextBox
        Me.TxtRatioReference = New System.Windows.Forms.TextBox
        Me.cboPurchaseOrder = New System.Windows.Forms.ComboBox
        Me.cboCodeNo = New System.Windows.Forms.ComboBox
        Me.txtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.TxtToleranceUp = New System.Windows.Forms.TextBox
        Me.ChkIgnoreTolerance = New System.Windows.Forms.CheckBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.LlbQty = New System.Windows.Forms.LinkLabel
        Me.LblUnits2 = New System.Windows.Forms.Label
        Me.BtnInput = New System.Windows.Forms.Button
        Me.TxtOrderBalance = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtRemarkItem = New System.Windows.Forms.TextBox
        Me.cboSupplierCodeNo = New System.Windows.Forms.ComboBox
        Me.TxtSupplierDescription = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.gboItem.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(6, 61)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(111, 13)
        Me.Label26.TabIndex = 181
        Me.Label26.Text = "Estimated Arrival Date"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(5, 10)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(98, 13)
        Me.Label19.TabIndex = 177
        Me.Label19.Text = "Purchase Order No"
        '
        'txtPrimaryPO
        '
        Me.txtPrimaryPO.Location = New System.Drawing.Point(142, 6)
        Me.txtPrimaryPO.Name = "txtPrimaryPO"
        Me.txtPrimaryPO.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryPO.TabIndex = 178
        Me.txtPrimaryPO.TabStop = False
        Me.txtPrimaryPO.Visible = False
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPONo, Me.ColCodeNo, Me.ColDescription, Me.ColQtyMulti, Me.ColUnitCodeMulti, Me.ColRatioDesc, Me.ColTotalQty, Me.ColUnit, Me.ColRemark, Me.ColPrimaryCodeNo, Me.ColPrimaryUnits, Me.ColPrimaryMasterUnits, Me.ColPrimaryPO, Me.ColRatioReference, Me.ColRatioDefault, Me.ColPrimaryUnitMulti, Me.ColQty, Me.ColSuppCodeNo, Me.ColSuppDesc, Me.ColQtyMaster, Me.ColQtyPerMaster, Me.ColMasterUnit})
        Me.dgvItem.Location = New System.Drawing.Point(7, 246)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.Size = New System.Drawing.Size(856, 169)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 10
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColPONo
        '
        Me.ColPONo.Frozen = True
        Me.ColPONo.HeaderText = "Purchase Order No"
        Me.ColPONo.Name = "ColPONo"
        Me.ColPONo.Width = 123
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
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.Width = 85
        '
        'ColQtyMulti
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMulti.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyMulti.HeaderText = "Qty"
        Me.ColQtyMulti.Name = "ColQtyMulti"
        Me.ColQtyMulti.Width = 48
        '
        'ColUnitCodeMulti
        '
        Me.ColUnitCodeMulti.HeaderText = "Units"
        Me.ColUnitCodeMulti.Name = "ColUnitCodeMulti"
        Me.ColUnitCodeMulti.Width = 56
        '
        'ColRatioDesc
        '
        Me.ColRatioDesc.HeaderText = "Ratio"
        Me.ColRatioDesc.Name = "ColRatioDesc"
        Me.ColRatioDesc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRatioDesc.Width = 57
        '
        'ColTotalQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColTotalQty.HeaderText = "Qty Default"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.Width = 85
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units Default"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 93
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
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 90
        '
        'ColPrimaryMasterUnits
        '
        Me.ColPrimaryMasterUnits.HeaderText = "Primary Master Units"
        Me.ColPrimaryMasterUnits.Name = "ColPrimaryMasterUnits"
        Me.ColPrimaryMasterUnits.Visible = False
        Me.ColPrimaryMasterUnits.Width = 128
        '
        'ColPrimaryPO
        '
        Me.ColPrimaryPO.HeaderText = "PrimaryPO"
        Me.ColPrimaryPO.Name = "ColPrimaryPO"
        Me.ColPrimaryPO.Visible = False
        Me.ColPrimaryPO.Width = 81
        '
        'ColRatioReference
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRatioReference.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColRatioReference.HeaderText = "Ratio Reference"
        Me.ColRatioReference.Name = "ColRatioReference"
        Me.ColRatioReference.Visible = False
        Me.ColRatioReference.Width = 110
        '
        'ColRatioDefault
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRatioDefault.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColRatioDefault.HeaderText = "Ratio Default"
        Me.ColRatioDefault.Name = "ColRatioDefault"
        Me.ColRatioDefault.Visible = False
        Me.ColRatioDefault.Width = 94
        '
        'ColPrimaryUnitMulti
        '
        Me.ColPrimaryUnitMulti.HeaderText = "Primary Unit Multi"
        Me.ColPrimaryUnitMulti.Name = "ColPrimaryUnitMulti"
        Me.ColPrimaryUnitMulti.Visible = False
        Me.ColPrimaryUnitMulti.Width = 113
        '
        'ColQty
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Visible = False
        Me.ColQty.Width = 48
        '
        'ColSuppCodeNo
        '
        Me.ColSuppCodeNo.HeaderText = "Supplier's Code No"
        Me.ColSuppCodeNo.Name = "ColSuppCodeNo"
        Me.ColSuppCodeNo.Visible = False
        Me.ColSuppCodeNo.Width = 122
        '
        'ColSuppDesc
        '
        Me.ColSuppDesc.HeaderText = "Supplier's Description"
        Me.ColSuppDesc.Name = "ColSuppDesc"
        Me.ColSuppDesc.Visible = False
        Me.ColSuppDesc.Width = 133
        '
        'ColQtyMaster
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMaster.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColQtyMaster.HeaderText = "Qty Master"
        Me.ColQtyMaster.Name = "ColQtyMaster"
        Me.ColQtyMaster.Visible = False
        Me.ColQtyMaster.Width = 83
        '
        'ColQtyPerMaster
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyPerMaster.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColQtyPerMaster.HeaderText = "Qty Per Master"
        Me.ColQtyPerMaster.Name = "ColQtyPerMaster"
        Me.ColQtyPerMaster.Visible = False
        Me.ColQtyPerMaster.Width = 102
        '
        'ColMasterUnit
        '
        Me.ColMasterUnit.HeaderText = "Master Units"
        Me.ColMasterUnit.Name = "ColMasterUnit"
        Me.ColMasterUnit.Visible = False
        Me.ColMasterUnit.Width = 91
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(528, 459)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(442, 459)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(356, 459)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 13)
        Me.Label2.TabIndex = 134
        Me.Label2.Text = "Supplier Name"
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindSupplier.Location = New System.Drawing.Point(390, 56)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(32, 20)
        Me.btnFindSupplier.TabIndex = 3
        Me.btnFindSupplier.TabStop = False
        Me.btnFindSupplier.Text = "...."
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(614, 459)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(184, 459)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 490)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(886, 22)
        Me.StatusStrip1.TabIndex = 47
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(270, 459)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(138, 85)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(279, 20)
        Me.TxtRemark.TabIndex = 8
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 89)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Remarks"
        '
        'txtDesc
        '
        Me.txtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDesc.Location = New System.Drawing.Point(119, 58)
        Me.txtDesc.MaxLength = 6
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(299, 20)
        Me.txtDesc.TabIndex = 2
        Me.txtDesc.TabStop = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(5, 62)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 13)
        Me.Label11.TabIndex = 109
        Me.Label11.Text = "Description"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(5, 36)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(49, 13)
        Me.Label10.TabIndex = 107
        Me.Label10.Text = "Code No"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 36)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(39, 13)
        Me.Label9.TabIndex = 106
        Me.Label9.Text = "PO No"
        Me.Label9.Visible = False
        '
        'txtPrimaryUnitsDefault
        '
        Me.txtPrimaryUnitsDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryUnitsDefault.Location = New System.Drawing.Point(9, 60)
        Me.txtPrimaryUnitsDefault.Name = "txtPrimaryUnitsDefault"
        Me.txtPrimaryUnitsDefault.Size = New System.Drawing.Size(24, 20)
        Me.txtPrimaryUnitsDefault.TabIndex = 87
        Me.txtPrimaryUnitsDefault.TabStop = False
        Me.txtPrimaryUnitsDefault.Visible = False
        '
        'lblUnitsMulti
        '
        Me.lblUnitsMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsMulti.AutoSize = True
        Me.lblUnitsMulti.Location = New System.Drawing.Point(379, 52)
        Me.lblUnitsMulti.Name = "lblUnitsMulti"
        Me.lblUnitsMulti.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsMulti.TabIndex = 11
        Me.lblUnitsMulti.Text = "Units"
        '
        'TxtQtyperBatchMulti
        '
        Me.TxtQtyperBatchMulti.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchMulti.Location = New System.Drawing.Point(136, 48)
        Me.TxtQtyperBatchMulti.MaxLength = 50
        Me.TxtQtyperBatchMulti.Name = "TxtQtyperBatchMulti"
        Me.TxtQtyperBatchMulti.Size = New System.Drawing.Size(238, 20)
        Me.TxtQtyperBatchMulti.TabIndex = 3
        Me.TxtQtyperBatchMulti.Text = "0"
        Me.TxtQtyperBatchMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dtpETD
        '
        Me.dtpETD.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpETD.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpETD.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpETD.Location = New System.Drawing.Point(138, 57)
        Me.dtpETD.Name = "dtpETD"
        Me.dtpETD.Size = New System.Drawing.Size(279, 20)
        Me.dtpETD.TabIndex = 7
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(143, 56)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimarySupplier.TabIndex = 133
        Me.txtPrimarySupplier.TabStop = False
        Me.txtPrimarySupplier.Visible = False
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierCode.Location = New System.Drawing.Point(107, 56)
        Me.txtSupplierCode.MaxLength = 50
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.Size = New System.Drawing.Size(277, 20)
        Me.txtSupplierCode.TabIndex = 2
        '
        'TxtSupplierName
        '
        Me.TxtSupplierName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSupplierName.Location = New System.Drawing.Point(107, 84)
        Me.TxtSupplierName.MaxLength = 30
        Me.TxtSupplierName.Name = "TxtSupplierName"
        Me.TxtSupplierName.ReadOnly = True
        Me.TxtSupplierName.Size = New System.Drawing.Size(315, 20)
        Me.TxtSupplierName.TabIndex = 4
        Me.TxtSupplierName.TabStop = False
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.TxtQtyperBatchDefault)
        Me.gboDetail.Controls.Add(Me.lblUnitsDefault)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.gboItem)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(4, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(871, 453)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'TxtQtyperBatchDefault
        '
        Me.TxtQtyperBatchDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchDefault.Location = New System.Drawing.Point(740, 253)
        Me.TxtQtyperBatchDefault.Name = "TxtQtyperBatchDefault"
        Me.TxtQtyperBatchDefault.Size = New System.Drawing.Size(75, 20)
        Me.TxtQtyperBatchDefault.TabIndex = 270
        Me.TxtQtyperBatchDefault.Visible = False
        '
        'lblUnitsDefault
        '
        Me.lblUnitsDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsDefault.AutoSize = True
        Me.lblUnitsDefault.Location = New System.Drawing.Point(821, 257)
        Me.lblUnitsDefault.Name = "lblUnitsDefault"
        Me.lblUnitsDefault.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsDefault.TabIndex = 269
        Me.lblUnitsDefault.Text = "Units"
        Me.lblUnitsDefault.Visible = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 12)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPackingListNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryPackingList)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtSupplierName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtSupplierCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimarySupplier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindSupplier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtSupplier_GITNo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dtpReceivedDate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label26)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dtpETD)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPrimaryUnitsDefault)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPrimaryUnitsMulti)
        Me.SplitContainer1.Size = New System.Drawing.Size(859, 114)
        Me.SplitContainer1.SplitterDistance = 429
        Me.SplitContainer1.TabIndex = 268
        Me.SplitContainer1.TabStop = False
        '
        'TxtPackingListNo
        '
        Me.TxtPackingListNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPackingListNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPackingListNo.Location = New System.Drawing.Point(107, 3)
        Me.TxtPackingListNo.MaxLength = 50
        Me.TxtPackingListNo.Name = "TxtPackingListNo"
        Me.TxtPackingListNo.ReadOnly = True
        Me.TxtPackingListNo.Size = New System.Drawing.Size(315, 20)
        Me.TxtPackingListNo.TabIndex = 0
        Me.TxtPackingListNo.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(42, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "GIT No"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 33)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "GIT Date"
        '
        'TxtPrimaryPackingList
        '
        Me.TxtPrimaryPackingList.AcceptsReturn = True
        Me.TxtPrimaryPackingList.Location = New System.Drawing.Point(147, 3)
        Me.TxtPrimaryPackingList.Name = "TxtPrimaryPackingList"
        Me.TxtPrimaryPackingList.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryPackingList.TabIndex = 47
        Me.TxtPrimaryPackingList.TabStop = False
        Me.TxtPrimaryPackingList.Visible = False
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(107, 29)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(315, 20)
        Me.dtpDate.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 60)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 132
        Me.Label1.Text = "Supplier Code"
        '
        'TxtSupplier_GITNo
        '
        Me.TxtSupplier_GITNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSupplier_GITNo.Location = New System.Drawing.Point(138, 4)
        Me.TxtSupplier_GITNo.MaxLength = 50
        Me.TxtSupplier_GITNo.Name = "TxtSupplier_GITNo"
        Me.TxtSupplier_GITNo.Size = New System.Drawing.Size(279, 20)
        Me.TxtSupplier_GITNo.TabIndex = 5
        '
        'dtpReceivedDate
        '
        Me.dtpReceivedDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpReceivedDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpReceivedDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpReceivedDate.Location = New System.Drawing.Point(138, 30)
        Me.dtpReceivedDate.Name = "dtpReceivedDate"
        Me.dtpReceivedDate.Size = New System.Drawing.Size(279, 20)
        Me.dtpReceivedDate.TabIndex = 6
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 34)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(79, 13)
        Me.Label7.TabIndex = 190
        Me.Label7.Text = "Received Date"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(123, 13)
        Me.Label5.TabIndex = 184
        Me.Label5.Text = "Supplier Packing List No"
        '
        'txtPrimaryUnitsMulti
        '
        Me.txtPrimaryUnitsMulti.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryUnitsMulti.Location = New System.Drawing.Point(9, 85)
        Me.txtPrimaryUnitsMulti.Name = "txtPrimaryUnitsMulti"
        Me.txtPrimaryUnitsMulti.Size = New System.Drawing.Size(24, 20)
        Me.txtPrimaryUnitsMulti.TabIndex = 191
        Me.txtPrimaryUnitsMulti.TabStop = False
        Me.txtPrimaryUnitsMulti.Visible = False
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = Global.Purchase_Module.My.Resources.Resources.arrow_up
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 421)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 13
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = Global.Purchase_Module.My.Resources.Resources.arrow_down
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 421)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 14
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.SplitContainer2)
        Me.gboItem.Controls.Add(Me.cboSupplierCodeNo)
        Me.gboItem.Controls.Add(Me.TxtSupplierDescription)
        Me.gboItem.Controls.Add(Me.Label20)
        Me.gboItem.Controls.Add(Me.Label21)
        Me.gboItem.Controls.Add(Me.Label23)
        Me.gboItem.Location = New System.Drawing.Point(6, 126)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Size = New System.Drawing.Size(859, 118)
        Me.gboItem.TabIndex = 9
        Me.gboItem.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(9, 10)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtDesc)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtRatioDefault)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtRatioReference)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboPurchaseOrder)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtPrimaryPO)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label19)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cboCodeNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtPrimaryCodeNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label9)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtToleranceUp)
        Me.SplitContainer2.Panel2.Controls.Add(Me.ChkIgnoreTolerance)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer2.Panel2.Controls.Add(Me.LlbQty)
        Me.SplitContainer2.Panel2.Controls.Add(Me.LblUnits2)
        Me.SplitContainer2.Panel2.Controls.Add(Me.BtnInput)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtOrderBalance)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label16)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtQtyperBatchMulti)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtRemarkItem)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblUnitsMulti)
        Me.SplitContainer2.Size = New System.Drawing.Size(841, 100)
        Me.SplitContainer2.SplitterDistance = 422
        Me.SplitContainer2.TabIndex = 269
        Me.SplitContainer2.TabStop = False
        '
        'TxtRatioDefault
        '
        Me.TxtRatioDefault.Location = New System.Drawing.Point(207, 58)
        Me.TxtRatioDefault.Name = "TxtRatioDefault"
        Me.TxtRatioDefault.Size = New System.Drawing.Size(68, 20)
        Me.TxtRatioDefault.TabIndex = 190
        Me.TxtRatioDefault.Visible = False
        '
        'TxtRatioReference
        '
        Me.TxtRatioReference.Location = New System.Drawing.Point(133, 58)
        Me.TxtRatioReference.Name = "TxtRatioReference"
        Me.TxtRatioReference.Size = New System.Drawing.Size(68, 20)
        Me.TxtRatioReference.TabIndex = 189
        Me.TxtRatioReference.Visible = False
        '
        'cboPurchaseOrder
        '
        Me.cboPurchaseOrder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboPurchaseOrder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPurchaseOrder.FormattingEnabled = True
        Me.cboPurchaseOrder.Location = New System.Drawing.Point(119, 6)
        Me.cboPurchaseOrder.Name = "cboPurchaseOrder"
        Me.cboPurchaseOrder.Size = New System.Drawing.Size(300, 21)
        Me.cboPurchaseOrder.TabIndex = 0
        '
        'cboCodeNo
        '
        Me.cboCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboCodeNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCodeNo.FormattingEnabled = True
        Me.cboCodeNo.Location = New System.Drawing.Point(119, 32)
        Me.cboCodeNo.Name = "cboCodeNo"
        Me.cboCodeNo.Size = New System.Drawing.Size(300, 21)
        Me.cboCodeNo.TabIndex = 1
        '
        'txtPrimaryCodeNo
        '
        Me.txtPrimaryCodeNo.Location = New System.Drawing.Point(143, 32)
        Me.txtPrimaryCodeNo.Name = "txtPrimaryCodeNo"
        Me.txtPrimaryCodeNo.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCodeNo.TabIndex = 188
        Me.txtPrimaryCodeNo.TabStop = False
        Me.txtPrimaryCodeNo.Visible = False
        '
        'TxtToleranceUp
        '
        Me.TxtToleranceUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtToleranceUp.Location = New System.Drawing.Point(299, 6)
        Me.TxtToleranceUp.MaxLength = 6
        Me.TxtToleranceUp.Name = "TxtToleranceUp"
        Me.TxtToleranceUp.ReadOnly = True
        Me.TxtToleranceUp.Size = New System.Drawing.Size(75, 20)
        Me.TxtToleranceUp.TabIndex = 269
        Me.TxtToleranceUp.TabStop = False
        Me.TxtToleranceUp.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ChkIgnoreTolerance
        '
        Me.ChkIgnoreTolerance.AutoSize = True
        Me.ChkIgnoreTolerance.Location = New System.Drawing.Point(136, 29)
        Me.ChkIgnoreTolerance.Name = "ChkIgnoreTolerance"
        Me.ChkIgnoreTolerance.Size = New System.Drawing.Size(139, 17)
        Me.ChkIgnoreTolerance.TabIndex = 271
        Me.ChkIgnoreTolerance.Text = "Igoner Lower Tolerance"
        Me.ChkIgnoreTolerance.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(286, 10)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(13, 13)
        Me.Label12.TabIndex = 270
        Me.Label12.Text = "+"
        '
        'LlbQty
        '
        Me.LlbQty.ActiveLinkColor = System.Drawing.Color.Black
        Me.LlbQty.AutoSize = True
        Me.LlbQty.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.LlbQty.LinkColor = System.Drawing.Color.Black
        Me.LlbQty.Location = New System.Drawing.Point(10, 52)
        Me.LlbQty.Name = "LlbQty"
        Me.LlbQty.Size = New System.Drawing.Size(23, 13)
        Me.LlbQty.TabIndex = 268
        Me.LlbQty.TabStop = True
        Me.LlbQty.Text = "Qty"
        Me.LlbQty.VisitedLinkColor = System.Drawing.Color.Black
        '
        'LblUnits2
        '
        Me.LblUnits2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnits2.AutoSize = True
        Me.LblUnits2.Location = New System.Drawing.Point(379, 10)
        Me.LblUnits2.Name = "LblUnits2"
        Me.LblUnits2.Size = New System.Drawing.Size(31, 13)
        Me.LblUnits2.TabIndex = 267
        Me.LblUnits2.Text = "Units"
        '
        'BtnInput
        '
        Me.BtnInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInput.Location = New System.Drawing.Point(355, 71)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(52, 23)
        Me.BtnInput.TabIndex = 5
        Me.BtnInput.Text = "&Input"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'TxtOrderBalance
        '
        Me.TxtOrderBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtOrderBalance.Location = New System.Drawing.Point(136, 6)
        Me.TxtOrderBalance.MaxLength = 6
        Me.TxtOrderBalance.Name = "TxtOrderBalance"
        Me.TxtOrderBalance.ReadOnly = True
        Me.TxtOrderBalance.Size = New System.Drawing.Size(149, 20)
        Me.TxtOrderBalance.TabIndex = 265
        Me.TxtOrderBalance.TabStop = False
        Me.TxtOrderBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(10, 76)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(49, 13)
        Me.Label16.TabIndex = 63
        Me.Label16.Text = "Remarks"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(10, 10)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(75, 13)
        Me.Label8.TabIndex = 266
        Me.Label8.Text = "Order Balance"
        '
        'TxtRemarkItem
        '
        Me.TxtRemarkItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItem.Location = New System.Drawing.Point(136, 72)
        Me.TxtRemarkItem.MaxLength = 150
        Me.TxtRemarkItem.Name = "TxtRemarkItem"
        Me.TxtRemarkItem.Size = New System.Drawing.Size(213, 20)
        Me.TxtRemarkItem.TabIndex = 4
        '
        'cboSupplierCodeNo
        '
        Me.cboSupplierCodeNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSupplierCodeNo.FormattingEnabled = True
        Me.cboSupplierCodeNo.Location = New System.Drawing.Point(117, 120)
        Me.cboSupplierCodeNo.Name = "cboSupplierCodeNo"
        Me.cboSupplierCodeNo.Size = New System.Drawing.Size(304, 21)
        Me.cboSupplierCodeNo.TabIndex = 260
        '
        'TxtSupplierDescription
        '
        Me.TxtSupplierDescription.Location = New System.Drawing.Point(117, 149)
        Me.TxtSupplierDescription.MaxLength = 6
        Me.TxtSupplierDescription.Name = "TxtSupplierDescription"
        Me.TxtSupplierDescription.ReadOnly = True
        Me.TxtSupplierDescription.Size = New System.Drawing.Size(303, 20)
        Me.TxtSupplierDescription.TabIndex = 261
        Me.TxtSupplierDescription.TabStop = False
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(3, 152)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(108, 13)
        Me.Label20.TabIndex = 264
        Me.Label20.Text = "Supplier's Description"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(3, 126)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(97, 13)
        Me.Label21.TabIndex = 263
        Me.Label21.Text = "Supplier's Code No"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(3, 126)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(39, 13)
        Me.Label23.TabIndex = 262
        Me.Label23.Text = "PO No"
        Me.Label23.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(441, 421)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 12
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(325, 421)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 11
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTransaksiPackingList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(886, 512)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.gboDetail)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(902, 550)
        Me.Name = "FrmTransaksiPackingList"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Packing List"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.gboItem.ResumeLayout(False)
        Me.gboItem.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryPO As System.Windows.Forms.TextBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryUnitsDefault As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsMulti As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatchMulti As System.Windows.Forms.TextBox
    Friend WithEvents dtpETD As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryPackingList As System.Windows.Forms.TextBox
    Friend WithEvents TxtPackingListNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents cboCodeNo As System.Windows.Forms.ComboBox
    Friend WithEvents TxtRemarkItem As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents BtnInput As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtSupplier_GITNo As System.Windows.Forms.TextBox
    Friend WithEvents dtpReceivedDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents cboPurchaseOrder As System.Windows.Forms.ComboBox
    Friend WithEvents cboSupplierCodeNo As System.Windows.Forms.ComboBox
    Friend WithEvents TxtSupplierDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents LblUnits2 As System.Windows.Forms.Label
    Friend WithEvents TxtOrderBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtRatioDefault As System.Windows.Forms.TextBox
    Friend WithEvents TxtRatioReference As System.Windows.Forms.TextBox
    Friend WithEvents LlbQty As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtQtyperBatchDefault As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsDefault As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitCodeMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyPerMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMasterUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtPrimaryUnitsMulti As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtToleranceUp As System.Windows.Forms.TextBox
    Friend WithEvents ChkIgnoreTolerance As System.Windows.Forms.CheckBox
End Class
