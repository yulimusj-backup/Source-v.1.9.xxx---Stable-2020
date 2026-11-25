<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockOutAdj
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiStockOutAdj))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.lblWarehouse = New System.Windows.Forms.Label
        Me.lblLocation = New System.Windows.Forms.Label
        Me.lblSecQty = New System.Windows.Forms.Label
        Me.lblExpiryDate = New System.Windows.Forms.Label
        Me.txtPrimaryPDO = New System.Windows.Forms.TextBox
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.txtPrimaryLocation = New System.Windows.Forms.TextBox
        Me.chkLocationControlled = New System.Windows.Forms.CheckBox
        Me.chkExpiryDateControlled = New System.Windows.Forms.CheckBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.chkBatchNoControlled = New System.Windows.Forms.CheckBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONoDONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppDONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecondaryQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecondaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWareHouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPOnoDoNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecondaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPdO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMultiUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsKain = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label6 = New System.Windows.Forms.Label
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.lblQuickInput3 = New System.Windows.Forms.LinkLabel
        Me.llbQuickInput = New System.Windows.Forms.LinkLabel
        Me.TxtAdjustmentTypeDesc = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAdjType = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.CboAdjType = New System.Windows.Forms.ComboBox
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.TxtStockInNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.llbAdvance = New System.Windows.Forms.LinkLabel
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtQtyperBatch = New System.Windows.Forms.TextBox
        Me.txtQtyOnHand = New System.Windows.Forms.TextBox
        Me.LblQtyOnHand = New System.Windows.Forms.Label
        Me.lblUnits = New System.Windows.Forms.Label
        Me.lblBatchNo = New System.Windows.Forms.Label
        Me.chkPdoControlled = New System.Windows.Forms.CheckBox
        Me.chkUseSecondaryUnit = New System.Windows.Forms.CheckBox
        Me.txtPrimarySecondaryUnits = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnits = New System.Windows.Forms.TextBox
        Me.txtSecondaryQty = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.DtpExpired = New System.Windows.Forms.DateTimePicker
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.TxtRemarkItem = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtWarehouse = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.BtnFindCodeNo = New System.Windows.Forms.Button
        Me.Label10 = New System.Windows.Forms.Label
        Me.BtnFindPONo = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.BtnInput = New System.Windows.Forms.Button
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.cboBatchNo = New System.Windows.Forms.ComboBox
        Me.cboSecondaryUnits = New System.Windows.Forms.ComboBox
        Me.lblSecondaryUnits = New System.Windows.Forms.Label
        Me.cboLocation = New System.Windows.Forms.ComboBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.llbQuickClear = New System.Windows.Forms.LinkLabel
        Me.cms1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsm_AddByBarcode = New System.Windows.Forms.ToolStripMenuItem
        Me.tsm_AddByRefDoc = New System.Windows.Forms.ToolStripMenuItem
        Me.tsm_AddByBatchNo = New System.Windows.Forms.ToolStripMenuItem
        Me.tsm_ByOnHand = New System.Windows.Forms.ToolStripMenuItem
        Me.AddItemByCustom = New System.Windows.Forms.ToolStripMenuItem
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboDetail.SuspendLayout()
        Me.gboItem.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.cms1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblWarehouse
        '
        Me.lblWarehouse.AutoSize = True
        Me.lblWarehouse.Location = New System.Drawing.Point(561, 45)
        Me.lblWarehouse.Name = "lblWarehouse"
        Me.lblWarehouse.Size = New System.Drawing.Size(24, 13)
        Me.lblWarehouse.TabIndex = 100
        Me.lblWarehouse.Text = "n/a"
        Me.lblWarehouse.Visible = False
        '
        'lblLocation
        '
        Me.lblLocation.AutoSize = True
        Me.lblLocation.Location = New System.Drawing.Point(561, 19)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(24, 13)
        Me.lblLocation.TabIndex = 99
        Me.lblLocation.Text = "n/a"
        Me.lblLocation.Visible = False
        '
        'lblSecQty
        '
        Me.lblSecQty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecQty.AutoSize = True
        Me.lblSecQty.Location = New System.Drawing.Point(764, 71)
        Me.lblSecQty.Name = "lblSecQty"
        Me.lblSecQty.Size = New System.Drawing.Size(18, 13)
        Me.lblSecQty.TabIndex = 98
        Me.lblSecQty.Text = "@"
        '
        'lblExpiryDate
        '
        Me.lblExpiryDate.AutoSize = True
        Me.lblExpiryDate.Location = New System.Drawing.Point(105, 99)
        Me.lblExpiryDate.Name = "lblExpiryDate"
        Me.lblExpiryDate.Size = New System.Drawing.Size(24, 13)
        Me.lblExpiryDate.TabIndex = 97
        Me.lblExpiryDate.Text = "n/a"
        Me.lblExpiryDate.Visible = False
        '
        'txtPrimaryPDO
        '
        Me.txtPrimaryPDO.Location = New System.Drawing.Point(440, 26)
        Me.txtPrimaryPDO.Name = "txtPrimaryPDO"
        Me.txtPrimaryPDO.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryPDO.TabIndex = 96
        Me.txtPrimaryPDO.TabStop = False
        Me.txtPrimaryPDO.Visible = False
        '
        'txtPrimaryCategory
        '
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(145, 16)
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCategory.TabIndex = 95
        Me.txtPrimaryCategory.TabStop = False
        Me.txtPrimaryCategory.Visible = False
        '
        'txtPrimaryLocation
        '
        Me.txtPrimaryLocation.Location = New System.Drawing.Point(530, 17)
        Me.txtPrimaryLocation.Name = "txtPrimaryLocation"
        Me.txtPrimaryLocation.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryLocation.TabIndex = 94
        Me.txtPrimaryLocation.TabStop = False
        Me.txtPrimaryLocation.Visible = False
        '
        'chkLocationControlled
        '
        Me.chkLocationControlled.AutoSize = True
        Me.chkLocationControlled.Location = New System.Drawing.Point(369, 54)
        Me.chkLocationControlled.Name = "chkLocationControlled"
        Me.chkLocationControlled.Size = New System.Drawing.Size(67, 17)
        Me.chkLocationControlled.TabIndex = 93
        Me.chkLocationControlled.TabStop = False
        Me.chkLocationControlled.Text = "Location"
        Me.chkLocationControlled.UseVisualStyleBackColor = True
        Me.chkLocationControlled.Visible = False
        '
        'chkExpiryDateControlled
        '
        Me.chkExpiryDateControlled.AutoSize = True
        Me.chkExpiryDateControlled.Location = New System.Drawing.Point(159, 54)
        Me.chkExpiryDateControlled.Name = "chkExpiryDateControlled"
        Me.chkExpiryDateControlled.Size = New System.Drawing.Size(80, 17)
        Me.chkExpiryDateControlled.TabIndex = 92
        Me.chkExpiryDateControlled.TabStop = False
        Me.chkExpiryDateControlled.Text = "Expiry Date"
        Me.chkExpiryDateControlled.UseVisualStyleBackColor = True
        Me.chkExpiryDateControlled.Visible = False
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = CType(resources.GetObject("btnUP.BackgroundImage"), System.Drawing.Image)
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 457)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 9
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = CType(resources.GetObject("btnDown.BackgroundImage"), System.Drawing.Image)
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(40, 457)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 10
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'chkBatchNoControlled
        '
        Me.chkBatchNoControlled.AutoSize = True
        Me.chkBatchNoControlled.Location = New System.Drawing.Point(517, 43)
        Me.chkBatchNoControlled.Name = "chkBatchNoControlled"
        Me.chkBatchNoControlled.Size = New System.Drawing.Size(71, 17)
        Me.chkBatchNoControlled.TabIndex = 91
        Me.chkBatchNoControlled.TabStop = False
        Me.chkBatchNoControlled.Text = "Batch No"
        Me.chkBatchNoControlled.UseVisualStyleBackColor = True
        Me.chkBatchNoControlled.Visible = False
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPONoDONo, Me.ColSuppDONo, Me.ColCodeNo, Me.ColDescription, Me.ColQty, Me.ColUnit, Me.ColSecondaryQty, Me.ColSecondaryUnit, Me.ColBatch, Me.ColExpiredDate, Me.ColLocation, Me.ColWareHouse, Me.ColRemark, Me.ColPrimaryPOnoDoNo, Me.ColPrimaryCodeNo, Me.ColPrimaryLocation, Me.ColPrimaryUnits, Me.ColPrimarySecondaryUnits, Me.ColPrimaryPdO, Me.ColPrimaryMultiUnit, Me.ColIsKain})
        Me.dgvItem.Location = New System.Drawing.Point(6, 200)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(917, 251)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 6
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
        'ColPONoDONo
        '
        Me.ColPONoDONo.HeaderText = "PO. No"
        Me.ColPONoDONo.Name = "ColPONoDONo"
        Me.ColPONoDONo.Visible = False
        Me.ColPONoDONo.Width = 79
        '
        'ColSuppDONo
        '
        Me.ColSuppDONo.HeaderText = "Supp. DO No"
        Me.ColSuppDONo.Name = "ColSuppDONo"
        Me.ColSuppDONo.Visible = False
        Me.ColSuppDONo.Width = 117
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
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Width = 48
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 56
        '
        'ColSecondaryQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecondaryQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColSecondaryQty.HeaderText = "@"
        Me.ColSecondaryQty.Name = "ColSecondaryQty"
        Me.ColSecondaryQty.Width = 43
        '
        'ColSecondaryUnit
        '
        Me.ColSecondaryUnit.HeaderText = "Units"
        Me.ColSecondaryUnit.Name = "ColSecondaryUnit"
        Me.ColSecondaryUnit.Width = 56
        '
        'ColBatch
        '
        Me.ColBatch.HeaderText = "Batch No"
        Me.ColBatch.Name = "ColBatch"
        Me.ColBatch.Width = 77
        '
        'ColExpiredDate
        '
        Me.ColExpiredDate.HeaderText = "Expiry Date"
        Me.ColExpiredDate.Name = "ColExpiredDate"
        Me.ColExpiredDate.Width = 86
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.Width = 73
        '
        'ColWareHouse
        '
        Me.ColWareHouse.HeaderText = "WareHouse"
        Me.ColWareHouse.Name = "ColWareHouse"
        Me.ColWareHouse.Width = 89
        '
        'ColRemark
        '
        Me.ColRemark.HeaderText = "Remarks"
        Me.ColRemark.Name = "ColRemark"
        Me.ColRemark.Width = 74
        '
        'ColPrimaryPOnoDoNo
        '
        Me.ColPrimaryPOnoDoNo.HeaderText = "PrimaryPOno"
        Me.ColPrimaryPOnoDoNo.Name = "ColPrimaryPOnoDoNo"
        Me.ColPrimaryPOnoDoNo.Visible = False
        Me.ColPrimaryPOnoDoNo.Width = 117
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 132
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 139
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 113
        '
        'ColPrimarySecondaryUnits
        '
        Me.ColPrimarySecondaryUnits.HeaderText = "PrimarySecondaryUnits"
        Me.ColPrimarySecondaryUnits.Name = "ColPrimarySecondaryUnits"
        Me.ColPrimarySecondaryUnits.Visible = False
        Me.ColPrimarySecondaryUnits.Width = 181
        '
        'ColPrimaryPdO
        '
        Me.ColPrimaryPdO.HeaderText = "Primary PdO"
        Me.ColPrimaryPdO.Name = "ColPrimaryPdO"
        Me.ColPrimaryPdO.Visible = False
        Me.ColPrimaryPdO.Width = 113
        '
        'ColPrimaryMultiUnit
        '
        Me.ColPrimaryMultiUnit.HeaderText = "Primary Multi Unit"
        Me.ColPrimaryMultiUnit.Name = "ColPrimaryMultiUnit"
        Me.ColPrimaryMultiUnit.ReadOnly = True
        Me.ColPrimaryMultiUnit.Visible = False
        Me.ColPrimaryMultiUnit.Width = 143
        '
        'ColIsKain
        '
        Me.ColIsKain.HeaderText = "Is Kain"
        Me.ColIsKain.Name = "ColIsKain"
        Me.ColIsKain.ReadOnly = True
        Me.ColIsKain.Visible = False
        Me.ColIsKain.Width = 75
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(463, 45)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 65
        Me.Label6.Text = "Remarks"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.lblQuickInput3)
        Me.gboDetail.Controls.Add(Me.llbQuickInput)
        Me.gboDetail.Controls.Add(Me.TxtAdjustmentTypeDesc)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryAdjType)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.CboAdjType)
        Me.gboDetail.Controls.Add(Me.btnUP)
        Me.gboDetail.Controls.Add(Me.btnDown)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.TxtRemark)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.dtpDate)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryStockIn)
        Me.gboDetail.Controls.Add(Me.TxtStockInNo)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.gboItem)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtCustomerName)
        Me.gboDetail.Location = New System.Drawing.Point(4, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(935, 489)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'lblQuickInput3
        '
        Me.lblQuickInput3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblQuickInput3.AutoSize = True
        Me.lblQuickInput3.Location = New System.Drawing.Point(787, 464)
        Me.lblQuickInput3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblQuickInput3.Name = "lblQuickInput3"
        Me.lblQuickInput3.Size = New System.Drawing.Size(133, 13)
        Me.lblQuickInput3.TabIndex = 105
        Me.lblQuickInput3.TabStop = True
        Me.lblQuickInput3.Text = "Quick Input from Batch No"
        '
        'llbQuickInput
        '
        Me.llbQuickInput.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbQuickInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbQuickInput.AutoSize = True
        Me.llbQuickInput.Location = New System.Drawing.Point(84, 464)
        Me.llbQuickInput.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbQuickInput.Name = "llbQuickInput"
        Me.llbQuickInput.Size = New System.Drawing.Size(62, 13)
        Me.llbQuickInput.TabIndex = 24
        Me.llbQuickInput.TabStop = True
        Me.llbQuickInput.Text = "Quick Input"
        '
        'TxtAdjustmentTypeDesc
        '
        Me.TxtAdjustmentTypeDesc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAdjustmentTypeDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAdjustmentTypeDesc.Location = New System.Drawing.Point(697, 16)
        Me.TxtAdjustmentTypeDesc.Name = "TxtAdjustmentTypeDesc"
        Me.TxtAdjustmentTypeDesc.ReadOnly = True
        Me.TxtAdjustmentTypeDesc.Size = New System.Drawing.Size(229, 20)
        Me.TxtAdjustmentTypeDesc.TabIndex = 3
        '
        'TxtPrimaryAdjType
        '
        Me.TxtPrimaryAdjType.Location = New System.Drawing.Point(464, 16)
        Me.TxtPrimaryAdjType.Name = "TxtPrimaryAdjType"
        Me.TxtPrimaryAdjType.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryAdjType.TabIndex = 104
        Me.TxtPrimaryAdjType.TabStop = False
        Me.TxtPrimaryAdjType.Text = "-1"
        Me.TxtPrimaryAdjType.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(463, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 13)
        Me.Label1.TabIndex = 103
        Me.Label1.Text = "Adjustment Type"
        '
        'CboAdjType
        '
        Me.CboAdjType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboAdjType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAdjType.FormattingEnabled = True
        Me.CboAdjType.Location = New System.Drawing.Point(554, 16)
        Me.CboAdjType.Name = "CboAdjType"
        Me.CboAdjType.Size = New System.Drawing.Size(138, 21)
        Me.CboAdjType.TabIndex = 2
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(554, 42)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(372, 18)
        Me.TxtRemark.TabIndex = 4
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(114, 42)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(316, 20)
        Me.dtpDate.TabIndex = 1
        '
        'TxtPrimaryStockIn
        '
        Me.TxtPrimaryStockIn.Location = New System.Drawing.Point(114, 19)
        Me.TxtPrimaryStockIn.Name = "TxtPrimaryStockIn"
        Me.TxtPrimaryStockIn.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryStockIn.TabIndex = 47
        Me.TxtPrimaryStockIn.TabStop = False
        Me.TxtPrimaryStockIn.Visible = False
        '
        'TxtStockInNo
        '
        Me.TxtStockInNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtStockInNo.Location = New System.Drawing.Point(114, 19)
        Me.TxtStockInNo.Name = "TxtStockInNo"
        Me.TxtStockInNo.ReadOnly = True
        Me.TxtStockInNo.Size = New System.Drawing.Size(316, 20)
        Me.TxtStockInNo.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 19)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Stock Out No"
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.llbAdvance)
        Me.gboItem.Controls.Add(Me.SplitContainer1)
        Me.gboItem.Controls.Add(Me.lblUnits)
        Me.gboItem.Controls.Add(Me.lblBatchNo)
        Me.gboItem.Controls.Add(Me.lblWarehouse)
        Me.gboItem.Controls.Add(Me.lblLocation)
        Me.gboItem.Controls.Add(Me.lblSecQty)
        Me.gboItem.Controls.Add(Me.lblExpiryDate)
        Me.gboItem.Controls.Add(Me.txtPrimaryPDO)
        Me.gboItem.Controls.Add(Me.txtPrimaryCategory)
        Me.gboItem.Controls.Add(Me.txtPrimaryLocation)
        Me.gboItem.Controls.Add(Me.chkLocationControlled)
        Me.gboItem.Controls.Add(Me.chkExpiryDateControlled)
        Me.gboItem.Controls.Add(Me.chkBatchNoControlled)
        Me.gboItem.Controls.Add(Me.chkPdoControlled)
        Me.gboItem.Controls.Add(Me.chkUseSecondaryUnit)
        Me.gboItem.Controls.Add(Me.txtPrimarySecondaryUnits)
        Me.gboItem.Controls.Add(Me.txtPrimaryUnits)
        Me.gboItem.Controls.Add(Me.txtSecondaryQty)
        Me.gboItem.Controls.Add(Me.Label19)
        Me.gboItem.Controls.Add(Me.DtpExpired)
        Me.gboItem.Controls.Add(Me.Label12)
        Me.gboItem.Controls.Add(Me.Label15)
        Me.gboItem.Controls.Add(Me.txtDesc)
        Me.gboItem.Controls.Add(Me.TxtRemarkItem)
        Me.gboItem.Controls.Add(Me.Label16)
        Me.gboItem.Controls.Add(Me.TxtWarehouse)
        Me.gboItem.Controls.Add(Me.Label14)
        Me.gboItem.Controls.Add(Me.Label13)
        Me.gboItem.Controls.Add(Me.Label11)
        Me.gboItem.Controls.Add(Me.BtnFindCodeNo)
        Me.gboItem.Controls.Add(Me.Label10)
        Me.gboItem.Controls.Add(Me.BtnFindPONo)
        Me.gboItem.Controls.Add(Me.Label9)
        Me.gboItem.Controls.Add(Me.BtnInput)
        Me.gboItem.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.gboItem.Controls.Add(Me.TxtCodeNo)
        Me.gboItem.Controls.Add(Me.cboBatchNo)
        Me.gboItem.Controls.Add(Me.cboSecondaryUnits)
        Me.gboItem.Controls.Add(Me.lblSecondaryUnits)
        Me.gboItem.Controls.Add(Me.cboLocation)
        Me.gboItem.Location = New System.Drawing.Point(6, 69)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Size = New System.Drawing.Size(917, 125)
        Me.gboItem.TabIndex = 5
        Me.gboItem.TabStop = False
        '
        'llbAdvance
        '
        Me.llbAdvance.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbAdvance.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAdvance.AutoSize = True
        Me.llbAdvance.Location = New System.Drawing.Point(864, 0)
        Me.llbAdvance.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbAdvance.Name = "llbAdvance"
        Me.llbAdvance.Size = New System.Drawing.Size(50, 13)
        Me.llbAdvance.TabIndex = 103
        Me.llbAdvance.TabStop = True
        Me.llbAdvance.Text = "Advance"
        Me.llbAdvance.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(558, 64)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtQtyperBatch)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtQtyOnHand)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LblQtyOnHand)
        Me.SplitContainer1.Size = New System.Drawing.Size(274, 26)
        Me.SplitContainer1.SplitterDistance = 102
        Me.SplitContainer1.TabIndex = 5
        Me.SplitContainer1.TabStop = False
        '
        'TxtQtyperBatch
        '
        Me.TxtQtyperBatch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatch.Location = New System.Drawing.Point(3, 3)
        Me.TxtQtyperBatch.MaxLength = 18
        Me.TxtQtyperBatch.Name = "TxtQtyperBatch"
        Me.TxtQtyperBatch.Size = New System.Drawing.Size(89, 20)
        Me.TxtQtyperBatch.TabIndex = 5
        Me.TxtQtyperBatch.Text = "0"
        Me.TxtQtyperBatch.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtQtyOnHand
        '
        Me.txtQtyOnHand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQtyOnHand.Location = New System.Drawing.Point(75, 3)
        Me.txtQtyOnHand.MaxLength = 18
        Me.txtQtyOnHand.Name = "txtQtyOnHand"
        Me.txtQtyOnHand.ReadOnly = True
        Me.txtQtyOnHand.Size = New System.Drawing.Size(90, 20)
        Me.txtQtyOnHand.TabIndex = 102
        Me.txtQtyOnHand.TabStop = False
        Me.txtQtyOnHand.Text = "0"
        Me.txtQtyOnHand.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblQtyOnHand
        '
        Me.LblQtyOnHand.AutoSize = True
        Me.LblQtyOnHand.Location = New System.Drawing.Point(15, 6)
        Me.LblQtyOnHand.Name = "LblQtyOnHand"
        Me.LblQtyOnHand.Size = New System.Drawing.Size(50, 13)
        Me.LblQtyOnHand.TabIndex = 102
        Me.LblQtyOnHand.Text = "On Hand"
        '
        'lblUnits
        '
        Me.lblUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnits.AutoSize = True
        Me.lblUnits.Location = New System.Drawing.Point(838, 71)
        Me.lblUnits.Name = "lblUnits"
        Me.lblUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblUnits.TabIndex = 84
        Me.lblUnits.Text = "Units"
        '
        'lblBatchNo
        '
        Me.lblBatchNo.AutoSize = True
        Me.lblBatchNo.Location = New System.Drawing.Point(105, 73)
        Me.lblBatchNo.Name = "lblBatchNo"
        Me.lblBatchNo.Size = New System.Drawing.Size(24, 13)
        Me.lblBatchNo.TabIndex = 101
        Me.lblBatchNo.Text = "n/a"
        Me.lblBatchNo.Visible = False
        '
        'chkPdoControlled
        '
        Me.chkPdoControlled.AutoSize = True
        Me.chkPdoControlled.Location = New System.Drawing.Point(517, 62)
        Me.chkPdoControlled.Name = "chkPdoControlled"
        Me.chkPdoControlled.Size = New System.Drawing.Size(47, 17)
        Me.chkPdoControlled.TabIndex = 90
        Me.chkPdoControlled.TabStop = False
        Me.chkPdoControlled.Text = "PdO"
        Me.chkPdoControlled.UseVisualStyleBackColor = True
        Me.chkPdoControlled.Visible = False
        '
        'chkUseSecondaryUnit
        '
        Me.chkUseSecondaryUnit.AutoSize = True
        Me.chkUseSecondaryUnit.Location = New System.Drawing.Point(517, 85)
        Me.chkUseSecondaryUnit.Name = "chkUseSecondaryUnit"
        Me.chkUseSecondaryUnit.Size = New System.Drawing.Size(121, 17)
        Me.chkUseSecondaryUnit.TabIndex = 89
        Me.chkUseSecondaryUnit.TabStop = False
        Me.chkUseSecondaryUnit.Text = "Use Secondary Unit"
        Me.chkUseSecondaryUnit.UseVisualStyleBackColor = True
        Me.chkUseSecondaryUnit.Visible = False
        '
        'txtPrimarySecondaryUnits
        '
        Me.txtPrimarySecondaryUnits.Location = New System.Drawing.Point(481, 53)
        Me.txtPrimarySecondaryUnits.Name = "txtPrimarySecondaryUnits"
        Me.txtPrimarySecondaryUnits.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimarySecondaryUnits.TabIndex = 88
        Me.txtPrimarySecondaryUnits.TabStop = False
        Me.txtPrimarySecondaryUnits.Visible = False
        '
        'txtPrimaryUnits
        '
        Me.txtPrimaryUnits.Location = New System.Drawing.Point(439, 53)
        Me.txtPrimaryUnits.Name = "txtPrimaryUnits"
        Me.txtPrimaryUnits.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryUnits.TabIndex = 87
        Me.txtPrimaryUnits.TabStop = False
        Me.txtPrimaryUnits.Visible = False
        '
        'txtSecondaryQty
        '
        Me.txtSecondaryQty.Location = New System.Drawing.Point(520, 68)
        Me.txtSecondaryQty.MaxLength = 30
        Me.txtSecondaryQty.Name = "txtSecondaryQty"
        Me.txtSecondaryQty.Size = New System.Drawing.Size(35, 20)
        Me.txtSecondaryQty.TabIndex = 9
        Me.txtSecondaryQty.TabStop = False
        Me.txtSecondaryQty.Text = "0"
        Me.txtSecondaryQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSecondaryQty.Visible = False
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(16, 73)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(52, 13)
        Me.Label19.TabIndex = 79
        Me.Label19.Text = "Batch No"
        '
        'DtpExpired
        '
        Me.DtpExpired.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpExpired.Enabled = False
        Me.DtpExpired.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpExpired.Location = New System.Drawing.Point(105, 96)
        Me.DtpExpired.Name = "DtpExpired"
        Me.DtpExpired.Size = New System.Drawing.Size(327, 20)
        Me.DtpExpired.TabIndex = 3
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(16, 99)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(61, 13)
        Me.Label12.TabIndex = 77
        Me.Label12.Text = "Expiry Date"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(457, 19)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(48, 13)
        Me.Label15.TabIndex = 73
        Me.Label15.Text = "Location"
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(105, 42)
        Me.txtDesc.MaxLength = 6
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(327, 20)
        Me.txtDesc.TabIndex = 2
        Me.txtDesc.TabStop = False
        '
        'TxtRemarkItem
        '
        Me.TxtRemarkItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItem.Location = New System.Drawing.Point(561, 96)
        Me.TxtRemarkItem.MaxLength = 50
        Me.TxtRemarkItem.Name = "TxtRemarkItem"
        Me.TxtRemarkItem.Size = New System.Drawing.Size(285, 20)
        Me.TxtRemarkItem.TabIndex = 7
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(457, 101)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(49, 13)
        Me.Label16.TabIndex = 63
        Me.Label16.Text = "Remarks"
        '
        'TxtWarehouse
        '
        Me.TxtWarehouse.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtWarehouse.Location = New System.Drawing.Point(561, 42)
        Me.TxtWarehouse.MaxLength = 50
        Me.TxtWarehouse.Name = "TxtWarehouse"
        Me.TxtWarehouse.ReadOnly = True
        Me.TxtWarehouse.Size = New System.Drawing.Size(343, 20)
        Me.TxtWarehouse.TabIndex = 7
        Me.TxtWarehouse.TabStop = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(457, 42)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(62, 13)
        Me.Label14.TabIndex = 65
        Me.Label14.Text = "Warehouse"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(457, 73)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(23, 13)
        Me.Label13.TabIndex = 63
        Me.Label13.Text = "Qty"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(16, 45)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 13)
        Me.Label11.TabIndex = 62
        Me.Label11.Text = "Description"
        '
        'BtnFindCodeNo
        '
        Me.BtnFindCodeNo.Location = New System.Drawing.Point(400, 17)
        Me.BtnFindCodeNo.Name = "BtnFindCodeNo"
        Me.BtnFindCodeNo.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindCodeNo.TabIndex = 1
        Me.BtnFindCodeNo.TabStop = False
        Me.BtnFindCodeNo.Text = "...."
        Me.BtnFindCodeNo.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(16, 19)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(49, 13)
        Me.Label10.TabIndex = 52
        Me.Label10.Text = "Code No"
        '
        'BtnFindPONo
        '
        Me.BtnFindPONo.Location = New System.Drawing.Point(400, 16)
        Me.BtnFindPONo.Name = "BtnFindPONo"
        Me.BtnFindPONo.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindPONo.TabIndex = 1
        Me.BtnFindPONo.Text = "...."
        Me.BtnFindPONo.UseVisualStyleBackColor = True
        Me.BtnFindPONo.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(16, 19)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(39, 13)
        Me.Label9.TabIndex = 48
        Me.Label9.Text = "PO No"
        Me.Label9.Visible = False
        '
        'BtnInput
        '
        Me.BtnInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInput.Location = New System.Drawing.Point(852, 93)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(52, 25)
        Me.BtnInput.TabIndex = 8
        Me.BtnInput.Text = "&Input"
        Me.BtnInput.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(114, 16)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCodeNo.TabIndex = 53
        Me.TxtPrimaryCodeNo.TabStop = False
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(105, 16)
        Me.TxtCodeNo.MaxLength = 50
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(287, 20)
        Me.TxtCodeNo.TabIndex = 0
        '
        'cboBatchNo
        '
        Me.cboBatchNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBatchNo.FormattingEnabled = True
        Me.cboBatchNo.Location = New System.Drawing.Point(105, 68)
        Me.cboBatchNo.Name = "cboBatchNo"
        Me.cboBatchNo.Size = New System.Drawing.Size(327, 21)
        Me.cboBatchNo.TabIndex = 2
        '
        'cboSecondaryUnits
        '
        Me.cboSecondaryUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboSecondaryUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSecondaryUnits.FormattingEnabled = True
        Me.cboSecondaryUnits.Location = New System.Drawing.Point(788, 68)
        Me.cboSecondaryUnits.Name = "cboSecondaryUnits"
        Me.cboSecondaryUnits.Size = New System.Drawing.Size(70, 21)
        Me.cboSecondaryUnits.TabIndex = 6
        '
        'lblSecondaryUnits
        '
        Me.lblSecondaryUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecondaryUnits.AutoSize = True
        Me.lblSecondaryUnits.Location = New System.Drawing.Point(864, 71)
        Me.lblSecondaryUnits.Name = "lblSecondaryUnits"
        Me.lblSecondaryUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblSecondaryUnits.TabIndex = 86
        Me.lblSecondaryUnits.Text = "Units"
        '
        'cboLocation
        '
        Me.cboLocation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLocation.FormattingEnabled = True
        Me.cboLocation.Location = New System.Drawing.Point(561, 15)
        Me.cboLocation.Name = "cboLocation"
        Me.cboLocation.Size = New System.Drawing.Size(343, 21)
        Me.cboLocation.TabIndex = 4
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(473, 457)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(357, 457)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 7
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Location = New System.Drawing.Point(116, 95)
        Me.TxtCustomerName.MaxLength = 30
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.ReadOnly = True
        Me.TxtCustomerName.Size = New System.Drawing.Size(316, 20)
        Me.TxtCustomerName.TabIndex = 4
        Me.TxtCustomerName.TabStop = False
        Me.TxtCustomerName.Visible = False
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(232, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(637, 496)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 535)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(944, 22)
        Me.StatusStrip1.TabIndex = 22
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(232, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(232, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(232, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(465, 496)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(379, 496)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(206, 496)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(293, 496)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(551, 496)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'llbQuickClear
        '
        Me.llbQuickClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbQuickClear.AutoSize = True
        Me.llbQuickClear.Location = New System.Drawing.Point(2, 496)
        Me.llbQuickClear.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbQuickClear.Name = "llbQuickClear"
        Me.llbQuickClear.Size = New System.Drawing.Size(62, 13)
        Me.llbQuickClear.TabIndex = 23
        Me.llbQuickClear.TabStop = True
        Me.llbQuickClear.Text = "Quick Clear"
        '
        'cms1
        '
        Me.cms1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsm_AddByBarcode, Me.tsm_AddByRefDoc, Me.tsm_AddByBatchNo, Me.tsm_ByOnHand, Me.AddItemByCustom})
        Me.cms1.Name = "cms1"
        Me.cms1.Size = New System.Drawing.Size(227, 114)
        '
        'tsm_AddByBarcode
        '
        Me.tsm_AddByBarcode.Name = "tsm_AddByBarcode"
        Me.tsm_AddByBarcode.Size = New System.Drawing.Size(226, 22)
        Me.tsm_AddByBarcode.Text = "Add By Barcode"
        '
        'tsm_AddByRefDoc
        '
        Me.tsm_AddByRefDoc.Name = "tsm_AddByRefDoc"
        Me.tsm_AddByRefDoc.Size = New System.Drawing.Size(226, 22)
        Me.tsm_AddByRefDoc.Text = "Add By Reference Document"
        '
        'tsm_AddByBatchNo
        '
        Me.tsm_AddByBatchNo.Name = "tsm_AddByBatchNo"
        Me.tsm_AddByBatchNo.Size = New System.Drawing.Size(226, 22)
        Me.tsm_AddByBatchNo.Text = "Add By Batch No"
        '
        'tsm_ByOnHand
        '
        Me.tsm_ByOnHand.Name = "tsm_ByOnHand"
        Me.tsm_ByOnHand.Size = New System.Drawing.Size(226, 22)
        Me.tsm_ByOnHand.Text = "Add By On Hand Stock"
        '
        'AddItemByCustom
        '
        Me.AddItemByCustom.Name = "AddItemByCustom"
        Me.AddItemByCustom.Size = New System.Drawing.Size(226, 22)
        Me.AddItemByCustom.Text = "Add Item By Custom"
        '
        'FrmTransaksiStockOutAdj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(944, 557)
        Me.Controls.Add(Me.llbQuickClear)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.MinimumSize = New System.Drawing.Size(960, 595)
        Me.Name = "FrmTransaksiStockOutAdj"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock Out - Adjust"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.gboItem.ResumeLayout(False)
        Me.gboItem.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.cms1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblWarehouse As System.Windows.Forms.Label
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents lblSecQty As System.Windows.Forms.Label
    Friend WithEvents lblExpiryDate As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryPDO As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryLocation As System.Windows.Forms.TextBox
    Friend WithEvents chkLocationControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkExpiryDateControlled As System.Windows.Forms.CheckBox
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents chkBatchNoControlled As System.Windows.Forms.CheckBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtPrimaryStockIn As System.Windows.Forms.TextBox
    Friend WithEvents TxtStockInNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents lblBatchNo As System.Windows.Forms.Label
    Friend WithEvents chkPdoControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseSecondaryUnit As System.Windows.Forms.CheckBox
    Friend WithEvents txtPrimarySecondaryUnits As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnits As System.Windows.Forms.TextBox
    Friend WithEvents lblSecondaryUnits As System.Windows.Forms.Label
    Friend WithEvents txtSecondaryQty As System.Windows.Forms.TextBox
    Friend WithEvents lblUnits As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatch As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents DtpExpired As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarkItem As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouse As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents BtnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents BtnFindPONo As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents BtnInput As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents cboBatchNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboLocation As System.Windows.Forms.ComboBox
    Friend WithEvents cboSecondaryUnits As System.Windows.Forms.ComboBox
    Public WithEvents TxtRemark As System.Windows.Forms.TextBox
    Public WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Public WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents btnFind As System.Windows.Forms.Button
    Public WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryAdjType As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CboAdjType As System.Windows.Forms.ComboBox
    Friend WithEvents TxtAdjustmentTypeDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents llbQuickClear As System.Windows.Forms.LinkLabel
    Friend WithEvents llbQuickInput As System.Windows.Forms.LinkLabel
    Friend WithEvents lblQuickInput3 As System.Windows.Forms.LinkLabel
    Friend WithEvents LblQtyOnHand As System.Windows.Forms.Label
    Friend WithEvents txtQtyOnHand As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents llbAdvance As System.Windows.Forms.LinkLabel
    Friend WithEvents cms1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsm_AddByBarcode As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_AddByRefDoc As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_AddByBatchNo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_ByOnHand As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddItemByCustom As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONoDONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppDONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecondaryQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecondaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWareHouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPOnoDoNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecondaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPdO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMultiUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsKain As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
