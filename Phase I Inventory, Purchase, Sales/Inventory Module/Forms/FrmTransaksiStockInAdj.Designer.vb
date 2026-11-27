<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockInAdj
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiStockInAdj))
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TxtRemarkItem = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtWarehouse = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.BtnFindCodeNo = New System.Windows.Forms.Button
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.txtBatchNo = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.DtpExpired = New System.Windows.Forms.DateTimePicker
        Me.Label12 = New System.Windows.Forms.Label
        Me.btnFindLocation = New System.Windows.Forms.Button
        Me.TxtLocation = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.llbAdvance = New System.Windows.Forms.LinkLabel
        Me.lblBatchNo = New System.Windows.Forms.Label
        Me.lblWarehouse = New System.Windows.Forms.Label
        Me.lblLocation = New System.Windows.Forms.Label
        Me.lblSecQty = New System.Windows.Forms.Label
        Me.lblExpiryDate = New System.Windows.Forms.Label
        Me.txtPrimaryPDO = New System.Windows.Forms.TextBox
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.txtPrimaryLocation = New System.Windows.Forms.TextBox
        Me.chkLocationControlled = New System.Windows.Forms.CheckBox
        Me.chkExpiryDateControlled = New System.Windows.Forms.CheckBox
        Me.chkBatchNoControlled = New System.Windows.Forms.CheckBox
        Me.chkPdoControlled = New System.Windows.Forms.CheckBox
        Me.chkUseSecondaryUnit = New System.Windows.Forms.CheckBox
        Me.txtPrimarySecondaryUnits = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnits = New System.Windows.Forms.TextBox
        Me.lblSecondaryUnits = New System.Windows.Forms.Label
        Me.txtSecondaryQty = New System.Windows.Forms.TextBox
        Me.lblUnits = New System.Windows.Forms.Label
        Me.TxtQtyperBatch = New System.Windows.Forms.TextBox
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.BtnFindPONo = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.BtnInput = New System.Windows.Forms.Button
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.llbQuickInputBundling = New System.Windows.Forms.LinkLabel
        Me.lblQuickInput2 = New System.Windows.Forms.LinkLabel
        Me.llbSetAmount = New System.Windows.Forms.LinkLabel
        Me.TxtAdjustmentTypeDesc = New System.Windows.Forms.TextBox
        Me.lblQuickInput = New System.Windows.Forms.LinkLabel
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtPrimarySetting = New System.Windows.Forms.TextBox
        Me.CboAdjType = New System.Windows.Forms.ComboBox
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONo = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.ColPrimaryPOno = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecondaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPdO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.TxtStockInNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.LblNo = New System.Windows.Forms.Label
        Me.btnPrint = New System.Windows.Forms.Button
        Me.LlbUpdDecProd = New System.Windows.Forms.LinkLabel
        Me.cms1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsm_AddByBarcode = New System.Windows.Forms.ToolStripMenuItem
        Me.tsm_AddByRefDoc = New System.Windows.Forms.ToolStripMenuItem
        Me.tsm_QuickInput = New System.Windows.Forms.ToolStripMenuItem
        Me.AddItemByCustom = New System.Windows.Forms.ToolStripMenuItem
        Me.gboItem.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cms1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtRemarkItem
        '
        Me.TxtRemarkItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItem.Location = New System.Drawing.Point(561, 96)
        Me.TxtRemarkItem.MaxLength = 50
        Me.TxtRemarkItem.Name = "TxtRemarkItem"
        Me.TxtRemarkItem.Size = New System.Drawing.Size(214, 20)
        Me.TxtRemarkItem.TabIndex = 8
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
        Me.TxtWarehouse.Size = New System.Drawing.Size(272, 20)
        Me.TxtWarehouse.TabIndex = 5
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
        'TxtCodeNo
        '
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(105, 16)
        Me.TxtCodeNo.MaxLength = 50
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(287, 20)
        Me.TxtCodeNo.TabIndex = 0
        '
        'txtBatchNo
        '
        Me.txtBatchNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtBatchNo.Location = New System.Drawing.Point(105, 70)
        Me.txtBatchNo.MaxLength = 50
        Me.txtBatchNo.Name = "txtBatchNo"
        Me.txtBatchNo.Size = New System.Drawing.Size(327, 20)
        Me.txtBatchNo.TabIndex = 2
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
        'btnFindLocation
        '
        Me.btnFindLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindLocation.Location = New System.Drawing.Point(801, 16)
        Me.btnFindLocation.Name = "btnFindLocation"
        Me.btnFindLocation.Size = New System.Drawing.Size(32, 20)
        Me.btnFindLocation.TabIndex = 6
        Me.btnFindLocation.TabStop = False
        Me.btnFindLocation.Text = "...."
        Me.btnFindLocation.UseVisualStyleBackColor = True
        '
        'TxtLocation
        '
        Me.TxtLocation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLocation.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtLocation.Location = New System.Drawing.Point(561, 16)
        Me.TxtLocation.MaxLength = 50
        Me.TxtLocation.Name = "TxtLocation"
        Me.TxtLocation.Size = New System.Drawing.Size(237, 20)
        Me.TxtLocation.TabIndex = 4
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
        'gboItem
        '
        Me.gboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.llbAdvance)
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
        Me.gboItem.Controls.Add(Me.lblSecondaryUnits)
        Me.gboItem.Controls.Add(Me.txtSecondaryQty)
        Me.gboItem.Controls.Add(Me.lblUnits)
        Me.gboItem.Controls.Add(Me.TxtQtyperBatch)
        Me.gboItem.Controls.Add(Me.txtBatchNo)
        Me.gboItem.Controls.Add(Me.Label19)
        Me.gboItem.Controls.Add(Me.DtpExpired)
        Me.gboItem.Controls.Add(Me.Label12)
        Me.gboItem.Controls.Add(Me.btnFindLocation)
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
        Me.gboItem.Controls.Add(Me.TxtLocation)
        Me.gboItem.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.gboItem.Controls.Add(Me.TxtCodeNo)
        Me.gboItem.Location = New System.Drawing.Point(6, 69)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Size = New System.Drawing.Size(846, 125)
        Me.gboItem.TabIndex = 5
        Me.gboItem.TabStop = False
        '
        'llbAdvance
        '
        Me.llbAdvance.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbAdvance.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAdvance.AutoSize = True
        Me.llbAdvance.Location = New System.Drawing.Point(793, -1)
        Me.llbAdvance.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbAdvance.Name = "llbAdvance"
        Me.llbAdvance.Size = New System.Drawing.Size(50, 13)
        Me.llbAdvance.TabIndex = 104
        Me.llbAdvance.TabStop = True
        Me.llbAdvance.Text = "Advance"
        Me.llbAdvance.VisitedLinkColor = System.Drawing.Color.Blue
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
        Me.lblSecQty.Location = New System.Drawing.Point(693, 71)
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
        'lblSecondaryUnits
        '
        Me.lblSecondaryUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecondaryUnits.AutoSize = True
        Me.lblSecondaryUnits.Location = New System.Drawing.Point(793, 71)
        Me.lblSecondaryUnits.Name = "lblSecondaryUnits"
        Me.lblSecondaryUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblSecondaryUnits.TabIndex = 86
        Me.lblSecondaryUnits.Text = "Units"
        '
        'txtSecondaryQty
        '
        Me.txtSecondaryQty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSecondaryQty.Location = New System.Drawing.Point(717, 68)
        Me.txtSecondaryQty.MaxLength = 18
        Me.txtSecondaryQty.Name = "txtSecondaryQty"
        Me.txtSecondaryQty.Size = New System.Drawing.Size(70, 20)
        Me.txtSecondaryQty.TabIndex = 7
        Me.txtSecondaryQty.Text = "0"
        Me.txtSecondaryQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblUnits
        '
        Me.lblUnits.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnits.AutoSize = True
        Me.lblUnits.Location = New System.Drawing.Point(656, 71)
        Me.lblUnits.Name = "lblUnits"
        Me.lblUnits.Size = New System.Drawing.Size(31, 13)
        Me.lblUnits.TabIndex = 84
        Me.lblUnits.Text = "Units"
        '
        'TxtQtyperBatch
        '
        Me.TxtQtyperBatch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatch.Location = New System.Drawing.Point(561, 68)
        Me.TxtQtyperBatch.MaxLength = 18
        Me.TxtQtyperBatch.Name = "TxtQtyperBatch"
        Me.TxtQtyperBatch.Size = New System.Drawing.Size(89, 20)
        Me.TxtQtyperBatch.TabIndex = 6
        Me.TxtQtyperBatch.Text = "0"
        Me.TxtQtyperBatch.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(105, 42)
        Me.txtDesc.MaxLength = 6
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(327, 20)
        Me.txtDesc.TabIndex = 1
        Me.txtDesc.TabStop = False
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
        Me.BtnInput.Location = New System.Drawing.Point(781, 93)
        Me.BtnInput.Name = "BtnInput"
        Me.BtnInput.Size = New System.Drawing.Size(52, 25)
        Me.BtnInput.TabIndex = 9
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
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(221, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(221, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(221, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(447, 380)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(361, 380)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(189, 380)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(275, 380)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(221, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 439)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(902, 22)
        Me.StatusStrip1.TabIndex = 14
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(620, 380)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(439, 335)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(323, 335)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 26)
        Me.btnOK.TabIndex = 6
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
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(554, 42)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(300, 18)
        Me.TxtRemark.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(463, 42)
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
        Me.gboDetail.Controls.Add(Me.llbQuickInputBundling)
        Me.gboDetail.Controls.Add(Me.lblQuickInput2)
        Me.gboDetail.Controls.Add(Me.llbSetAmount)
        Me.gboDetail.Controls.Add(Me.TxtAdjustmentTypeDesc)
        Me.gboDetail.Controls.Add(Me.lblQuickInput)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.TxtPrimarySetting)
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
        Me.gboDetail.Controls.Add(Me.LblNo)
        Me.gboDetail.Controls.Add(Me.gboItem)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtCustomerName)
        Me.gboDetail.Location = New System.Drawing.Point(14, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(864, 367)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'llbQuickInputBundling
        '
        Me.llbQuickInputBundling.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbQuickInputBundling.AutoSize = True
        Me.llbQuickInputBundling.Location = New System.Drawing.Point(778, 341)
        Me.llbQuickInputBundling.Name = "llbQuickInputBundling"
        Me.llbQuickInputBundling.Size = New System.Drawing.Size(90, 13)
        Me.llbQuickInputBundling.TabIndex = 74
        Me.llbQuickInputBundling.TabStop = True
        Me.llbQuickInputBundling.Text = "Input By Bundling"
        '
        'lblQuickInput2
        '
        Me.lblQuickInput2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblQuickInput2.AutoSize = True
        Me.lblQuickInput2.Location = New System.Drawing.Point(140, 342)
        Me.lblQuickInput2.Name = "lblQuickInput2"
        Me.lblQuickInput2.Size = New System.Drawing.Size(71, 13)
        Me.lblQuickInput2.TabIndex = 73
        Me.lblQuickInput2.TabStop = True
        Me.lblQuickInput2.Text = "Quick Input 2"
        '
        'llbSetAmount
        '
        Me.llbSetAmount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbSetAmount.AutoSize = True
        Me.llbSetAmount.Location = New System.Drawing.Point(216, 342)
        Me.llbSetAmount.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbSetAmount.Name = "llbSetAmount"
        Me.llbSetAmount.Size = New System.Drawing.Size(62, 13)
        Me.llbSetAmount.TabIndex = 72
        Me.llbSetAmount.TabStop = True
        Me.llbSetAmount.Text = "Set Amount"
        '
        'TxtAdjustmentTypeDesc
        '
        Me.TxtAdjustmentTypeDesc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAdjustmentTypeDesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAdjustmentTypeDesc.Location = New System.Drawing.Point(702, 16)
        Me.TxtAdjustmentTypeDesc.Name = "TxtAdjustmentTypeDesc"
        Me.TxtAdjustmentTypeDesc.ReadOnly = True
        Me.TxtAdjustmentTypeDesc.Size = New System.Drawing.Size(154, 20)
        Me.TxtAdjustmentTypeDesc.TabIndex = 3
        Me.TxtAdjustmentTypeDesc.TabStop = False
        '
        'lblQuickInput
        '
        Me.lblQuickInput.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblQuickInput.AutoSize = True
        Me.lblQuickInput.Location = New System.Drawing.Point(806, 0)
        Me.lblQuickInput.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblQuickInput.Name = "lblQuickInput"
        Me.lblQuickInput.Size = New System.Drawing.Size(62, 13)
        Me.lblQuickInput.TabIndex = 15
        Me.lblQuickInput.TabStop = True
        Me.lblQuickInput.Text = "Import Data"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(463, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 13)
        Me.Label1.TabIndex = 71
        Me.Label1.Text = "Adjustment Type"
        '
        'TxtPrimarySetting
        '
        Me.TxtPrimarySetting.Location = New System.Drawing.Point(465, 16)
        Me.TxtPrimarySetting.Name = "TxtPrimarySetting"
        Me.TxtPrimarySetting.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimarySetting.TabIndex = 70
        Me.TxtPrimarySetting.TabStop = False
        Me.TxtPrimarySetting.Text = "-1"
        Me.TxtPrimarySetting.Visible = False
        '
        'CboAdjType
        '
        Me.CboAdjType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboAdjType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAdjType.FormattingEnabled = True
        Me.CboAdjType.Location = New System.Drawing.Point(554, 15)
        Me.CboAdjType.Name = "CboAdjType"
        Me.CboAdjType.Size = New System.Drawing.Size(138, 21)
        Me.CboAdjType.TabIndex = 2
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = CType(resources.GetObject("btnUP.BackgroundImage"), System.Drawing.Image)
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(6, 335)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 8
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
        Me.btnDown.Location = New System.Drawing.Point(40, 335)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 9
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
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
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColPONo, Me.ColSuppDONo, Me.ColCodeNo, Me.ColDescription, Me.ColQty, Me.ColUnit, Me.ColSecondaryQty, Me.ColSecondaryUnit, Me.ColBatch, Me.ColExpiredDate, Me.ColLocation, Me.ColWareHouse, Me.ColRemark, Me.ColPrimaryPOno, Me.ColPrimaryCodeNo, Me.ColPrimaryLocation, Me.ColPrimaryUnits, Me.ColPrimarySecondaryUnits, Me.ColPrimaryPdO, Me.ColUnitCost, Me.ColAmount})
        Me.dgvItem.Location = New System.Drawing.Point(6, 200)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(846, 129)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 4
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
        'ColPONo
        '
        Me.ColPONo.Frozen = True
        Me.ColPONo.HeaderText = "PO. No"
        Me.ColPONo.Name = "ColPONo"
        Me.ColPONo.Visible = False
        Me.ColPONo.Width = 84
        '
        'ColSuppDONo
        '
        Me.ColSuppDONo.Frozen = True
        Me.ColSuppDONo.HeaderText = "Supp. DO No"
        Me.ColSuppDONo.Name = "ColSuppDONo"
        Me.ColSuppDONo.Visible = False
        Me.ColSuppDONo.Width = 128
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
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle7
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
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecondaryQty.DefaultCellStyle = DataGridViewCellStyle8
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
        'ColPrimaryPOno
        '
        Me.ColPrimaryPOno.HeaderText = "PrimaryPOno"
        Me.ColPrimaryPOno.Name = "ColPrimaryPOno"
        Me.ColPrimaryPOno.Visible = False
        Me.ColPrimaryPOno.Width = 126
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 144
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 151
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 123
        '
        'ColPrimarySecondaryUnits
        '
        Me.ColPrimarySecondaryUnits.HeaderText = "PrimarySecondaryUnits"
        Me.ColPrimarySecondaryUnits.Name = "ColPrimarySecondaryUnits"
        Me.ColPrimarySecondaryUnits.Visible = False
        Me.ColPrimarySecondaryUnits.Width = 199
        '
        'ColPrimaryPdO
        '
        Me.ColPrimaryPdO.HeaderText = "Primary PdO"
        Me.ColPrimaryPdO.Name = "ColPrimaryPdO"
        Me.ColPrimaryPdO.Visible = False
        Me.ColPrimaryPdO.Width = 121
        '
        'ColUnitCost
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitCost.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColUnitCost.HeaderText = "Unit Cost"
        Me.ColUnitCost.Name = "ColUnitCost"
        Me.ColUnitCost.Width = 75
        '
        'ColAmount
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.Width = 68
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
        Me.TxtStockInNo.Location = New System.Drawing.Point(114, 15)
        Me.TxtStockInNo.Name = "TxtStockInNo"
        Me.TxtStockInNo.ReadOnly = True
        Me.TxtStockInNo.Size = New System.Drawing.Size(316, 20)
        Me.TxtStockInNo.TabIndex = 0
        Me.TxtStockInNo.TabStop = False
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
        'LblNo
        '
        Me.LblNo.AutoSize = True
        Me.LblNo.Location = New System.Drawing.Point(16, 19)
        Me.LblNo.Name = "LblNo"
        Me.LblNo.Size = New System.Drawing.Size(64, 13)
        Me.LblNo.TabIndex = 45
        Me.LblNo.Text = "Stock In No"
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(534, 380)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'LlbUpdDecProd
        '
        Me.LlbUpdDecProd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LlbUpdDecProd.AutoSize = True
        Me.LlbUpdDecProd.Location = New System.Drawing.Point(7, 456)
        Me.LlbUpdDecProd.Name = "LlbUpdDecProd"
        Me.LlbUpdDecProd.Size = New System.Drawing.Size(123, 13)
        Me.LlbUpdDecProd.TabIndex = 74
        Me.LlbUpdDecProd.TabStop = True
        Me.LlbUpdDecProd.Text = "Update Decimal Product"
        '
        'cms1
        '
        Me.cms1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsm_AddByBarcode, Me.tsm_AddByRefDoc, Me.tsm_QuickInput, Me.AddItemByCustom})
        Me.cms1.Name = "cms1"
        Me.cms1.Size = New System.Drawing.Size(227, 92)
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
        'tsm_QuickInput
        '
        Me.tsm_QuickInput.Name = "tsm_QuickInput"
        Me.tsm_QuickInput.Size = New System.Drawing.Size(226, 22)
        Me.tsm_QuickInput.Text = "Add By Product Level"
        '
        'AddItemByCustom
        '
        Me.AddItemByCustom.Name = "AddItemByCustom"
        Me.AddItemByCustom.Size = New System.Drawing.Size(226, 22)
        Me.AddItemByCustom.Text = "Add Item By Custom"
        '
        'FrmTransaksiStockInAdj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(902, 461)
        Me.Controls.Add(Me.LlbUpdDecProd)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnNew)
        Me.MinimumSize = New System.Drawing.Size(910, 492)
        Me.Name = "FrmTransaksiStockInAdj"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock In - Adjust"
        Me.gboItem.ResumeLayout(False)
        Me.gboItem.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cms1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtRemarkItem As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouse As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents BtnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents txtBatchNo As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents DtpExpired As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents btnFindLocation As System.Windows.Forms.Button
    Friend WithEvents TxtLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents BtnFindPONo As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents BtnInput As System.Windows.Forms.Button
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtPrimaryStockIn As System.Windows.Forms.TextBox
    Friend WithEvents TxtStockInNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LblNo As System.Windows.Forms.Label
    Friend WithEvents lblSecondaryUnits As System.Windows.Forms.Label
    Friend WithEvents txtSecondaryQty As System.Windows.Forms.TextBox
    Friend WithEvents lblUnits As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatch As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimarySecondaryUnits As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnits As System.Windows.Forms.TextBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents chkPdoControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseSecondaryUnit As System.Windows.Forms.CheckBox
    Friend WithEvents chkLocationControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkExpiryDateControlled As System.Windows.Forms.CheckBox
    Friend WithEvents chkBatchNoControlled As System.Windows.Forms.CheckBox
    Friend WithEvents txtPrimaryLocation As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryPDO As System.Windows.Forms.TextBox
    Friend WithEvents lblExpiryDate As System.Windows.Forms.Label
    Friend WithEvents lblSecQty As System.Windows.Forms.Label
    Friend WithEvents lblWarehouse As System.Windows.Forms.Label
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents lblBatchNo As System.Windows.Forms.Label
    Public WithEvents btnNew As System.Windows.Forms.Button
    Public WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents TxtRemark As System.Windows.Forms.TextBox
    Public WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimarySetting As System.Windows.Forms.TextBox
    Friend WithEvents CboAdjType As System.Windows.Forms.ComboBox
    Friend WithEvents TxtAdjustmentTypeDesc As System.Windows.Forms.TextBox
    Friend WithEvents lblQuickInput As System.Windows.Forms.LinkLabel
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents llbSetAmount As System.Windows.Forms.LinkLabel
    Friend WithEvents lblQuickInput2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbUpdDecProd As System.Windows.Forms.LinkLabel
    Friend WithEvents llbQuickInputBundling As System.Windows.Forms.LinkLabel
    Friend WithEvents llbAdvance As System.Windows.Forms.LinkLabel
    Friend WithEvents cms1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsm_AddByBarcode As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_AddByRefDoc As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsm_QuickInput As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddItemByCustom As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONo As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents ColPrimaryPOno As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecondaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPdO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
