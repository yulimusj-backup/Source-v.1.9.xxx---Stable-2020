<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockBundling_BOM
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
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.gboItemb4 = New System.Windows.Forms.GroupBox
        Me.dgvItemb4 = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecondaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.cboLocationb4 = New System.Windows.Forms.ComboBox
        Me.lblBatchNob4 = New System.Windows.Forms.Label
        Me.lblWarehouseb4 = New System.Windows.Forms.Label
        Me.lblLocationb4 = New System.Windows.Forms.Label
        Me.lblSecQtyb4 = New System.Windows.Forms.Label
        Me.lblExpiryDateb4 = New System.Windows.Forms.Label
        Me.txtPrimaryPDOb4 = New System.Windows.Forms.TextBox
        Me.txtPrimaryCategoryb4 = New System.Windows.Forms.TextBox
        Me.txtPrimaryLocationb4 = New System.Windows.Forms.TextBox
        Me.chkLocationControlledb4 = New System.Windows.Forms.CheckBox
        Me.chkExpiryDateControlledb4 = New System.Windows.Forms.CheckBox
        Me.chkBatchNoControlledb4 = New System.Windows.Forms.CheckBox
        Me.chkPdoControlledb4 = New System.Windows.Forms.CheckBox
        Me.chkUseSecondaryUnitb4 = New System.Windows.Forms.CheckBox
        Me.txtPrimarySecondaryUnitsb4 = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnitsb4 = New System.Windows.Forms.TextBox
        Me.txtSecondaryQtyb4 = New System.Windows.Forms.TextBox
        Me.lblUnitsb4 = New System.Windows.Forms.Label
        Me.TxtQtyperBatchb4 = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.DtpExpiredb4 = New System.Windows.Forms.DateTimePicker
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtDescb4 = New System.Windows.Forms.TextBox
        Me.TxtRemarkItemb4 = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtWarehouseb4 = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.BtnInputb4 = New System.Windows.Forms.Button
        Me.TxtPrimaryCodeNob4 = New System.Windows.Forms.TextBox
        Me.TxtCodeNob4 = New System.Windows.Forms.TextBox
        Me.cboBatchNob4 = New System.Windows.Forms.ComboBox
        Me.cboSecondaryUnitsb4 = New System.Windows.Forms.ComboBox
        Me.lblSecondaryUnitsb4 = New System.Windows.Forms.Label
        Me.BtnFindCodeNob4 = New System.Windows.Forms.Button
        Me.gboItemAfter = New System.Windows.Forms.GroupBox
        Me.txtFaktorPemberat = New System.Windows.Forms.TextBox
        Me.lblFaktorPemberat = New System.Windows.Forms.Label
        Me.lblBatchNoAfter = New System.Windows.Forms.Label
        Me.lblWarehouseAfter = New System.Windows.Forms.Label
        Me.lblLocationAfter = New System.Windows.Forms.Label
        Me.lblSecQtyAfter = New System.Windows.Forms.Label
        Me.lblExpiryDateAfter = New System.Windows.Forms.Label
        Me.txtPrimaryPDOAfter = New System.Windows.Forms.TextBox
        Me.txtPrimaryCategoryAfter = New System.Windows.Forms.TextBox
        Me.txtPrimaryLocationAfter = New System.Windows.Forms.TextBox
        Me.chkLocationControlledAfter = New System.Windows.Forms.CheckBox
        Me.chkExpiryDateControlledAfter = New System.Windows.Forms.CheckBox
        Me.chkBatchNoControlledAfter = New System.Windows.Forms.CheckBox
        Me.chkPdoControlledAfter = New System.Windows.Forms.CheckBox
        Me.chkUseSecondaryUnitAfter = New System.Windows.Forms.CheckBox
        Me.txtPrimarySecondaryUnitsAfter = New System.Windows.Forms.TextBox
        Me.txtPrimaryUnitsAfter = New System.Windows.Forms.TextBox
        Me.lblSecondaryUnitsAfter = New System.Windows.Forms.Label
        Me.txtSecondaryQtyAfter = New System.Windows.Forms.TextBox
        Me.lblUnitsAfter = New System.Windows.Forms.Label
        Me.TxtQtyperBatchAfter = New System.Windows.Forms.TextBox
        Me.txtBatchNoAfter = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.DtpExpiredAfter = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnFindLocationAfter = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDescAfter = New System.Windows.Forms.TextBox
        Me.TxtRemarkItemAfter = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtWarehouseAfter = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.BtnFindCodeNoAfter = New System.Windows.Forms.Button
        Me.Label20 = New System.Windows.Forms.Label
        Me.BtnFindPONo = New System.Windows.Forms.Button
        Me.Label21 = New System.Windows.Forms.Label
        Me.TxtLocationAfter = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCodeNoAfter = New System.Windows.Forms.TextBox
        Me.TxtCodeNoAfter = New System.Windows.Forms.TextBox
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.gboItemb4.SuspendLayout()
        CType(Me.dgvItemb4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboItemAfter.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(794, 445)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(713, 445)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 6
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'gboItemb4
        '
        Me.gboItemb4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItemb4.Controls.Add(Me.dgvItemb4)
        Me.gboItemb4.Controls.Add(Me.cboLocationb4)
        Me.gboItemb4.Controls.Add(Me.lblBatchNob4)
        Me.gboItemb4.Controls.Add(Me.lblWarehouseb4)
        Me.gboItemb4.Controls.Add(Me.lblLocationb4)
        Me.gboItemb4.Controls.Add(Me.lblSecQtyb4)
        Me.gboItemb4.Controls.Add(Me.lblExpiryDateb4)
        Me.gboItemb4.Controls.Add(Me.txtPrimaryPDOb4)
        Me.gboItemb4.Controls.Add(Me.txtPrimaryCategoryb4)
        Me.gboItemb4.Controls.Add(Me.txtPrimaryLocationb4)
        Me.gboItemb4.Controls.Add(Me.chkLocationControlledb4)
        Me.gboItemb4.Controls.Add(Me.chkExpiryDateControlledb4)
        Me.gboItemb4.Controls.Add(Me.chkBatchNoControlledb4)
        Me.gboItemb4.Controls.Add(Me.chkPdoControlledb4)
        Me.gboItemb4.Controls.Add(Me.chkUseSecondaryUnitb4)
        Me.gboItemb4.Controls.Add(Me.txtPrimarySecondaryUnitsb4)
        Me.gboItemb4.Controls.Add(Me.txtPrimaryUnitsb4)
        Me.gboItemb4.Controls.Add(Me.txtSecondaryQtyb4)
        Me.gboItemb4.Controls.Add(Me.lblUnitsb4)
        Me.gboItemb4.Controls.Add(Me.TxtQtyperBatchb4)
        Me.gboItemb4.Controls.Add(Me.Label19)
        Me.gboItemb4.Controls.Add(Me.DtpExpiredb4)
        Me.gboItemb4.Controls.Add(Me.Label12)
        Me.gboItemb4.Controls.Add(Me.Label15)
        Me.gboItemb4.Controls.Add(Me.txtDescb4)
        Me.gboItemb4.Controls.Add(Me.TxtRemarkItemb4)
        Me.gboItemb4.Controls.Add(Me.Label16)
        Me.gboItemb4.Controls.Add(Me.TxtWarehouseb4)
        Me.gboItemb4.Controls.Add(Me.Label14)
        Me.gboItemb4.Controls.Add(Me.Label13)
        Me.gboItemb4.Controls.Add(Me.Label11)
        Me.gboItemb4.Controls.Add(Me.Label10)
        Me.gboItemb4.Controls.Add(Me.Label9)
        Me.gboItemb4.Controls.Add(Me.BtnInputb4)
        Me.gboItemb4.Controls.Add(Me.TxtPrimaryCodeNob4)
        Me.gboItemb4.Controls.Add(Me.TxtCodeNob4)
        Me.gboItemb4.Controls.Add(Me.cboBatchNob4)
        Me.gboItemb4.Controls.Add(Me.cboSecondaryUnitsb4)
        Me.gboItemb4.Controls.Add(Me.lblSecondaryUnitsb4)
        Me.gboItemb4.Controls.Add(Me.BtnFindCodeNob4)
        Me.gboItemb4.Location = New System.Drawing.Point(12, 143)
        Me.gboItemb4.Name = "gboItemb4"
        Me.gboItemb4.Size = New System.Drawing.Size(859, 296)
        Me.gboItemb4.TabIndex = 9
        Me.gboItemb4.TabStop = False
        Me.gboItemb4.Text = "Item Before"
        '
        'dgvItemb4
        '
        Me.dgvItemb4.AllowUserToAddRows = False
        Me.dgvItemb4.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItemb4.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItemb4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemb4.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemb4.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItemb4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvItemb4.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColQty, Me.ColUnit, Me.ColSecondaryQty, Me.ColSecondaryUnit, Me.ColBatch, Me.ColExpiredDate, Me.ColLocation, Me.ColWareHouse, Me.ColRemark, Me.ColPrimaryCodeNo, Me.ColPrimaryLocation, Me.ColPrimaryUnits, Me.ColPrimarySecondaryUnits})
        Me.dgvItemb4.Location = New System.Drawing.Point(6, 124)
        Me.dgvItemb4.MultiSelect = False
        Me.dgvItemb4.Name = "dgvItemb4"
        Me.dgvItemb4.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItemb4.Size = New System.Drawing.Size(847, 166)
        Me.dgvItemb4.StandardTab = True
        Me.dgvItemb4.TabIndex = 102
        Me.dgvItemb4.TabStop = False
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
        'cboLocationb4
        '
        Me.cboLocationb4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboLocationb4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLocationb4.FormattingEnabled = True
        Me.cboLocationb4.Location = New System.Drawing.Point(561, 15)
        Me.cboLocationb4.Name = "cboLocationb4"
        Me.cboLocationb4.Size = New System.Drawing.Size(285, 21)
        Me.cboLocationb4.TabIndex = 10
        '
        'lblBatchNob4
        '
        Me.lblBatchNob4.AutoSize = True
        Me.lblBatchNob4.Location = New System.Drawing.Point(105, 73)
        Me.lblBatchNob4.Name = "lblBatchNob4"
        Me.lblBatchNob4.Size = New System.Drawing.Size(24, 13)
        Me.lblBatchNob4.TabIndex = 101
        Me.lblBatchNob4.Text = "n/a"
        Me.lblBatchNob4.Visible = False
        '
        'lblWarehouseb4
        '
        Me.lblWarehouseb4.AutoSize = True
        Me.lblWarehouseb4.Location = New System.Drawing.Point(561, 45)
        Me.lblWarehouseb4.Name = "lblWarehouseb4"
        Me.lblWarehouseb4.Size = New System.Drawing.Size(24, 13)
        Me.lblWarehouseb4.TabIndex = 100
        Me.lblWarehouseb4.Text = "n/a"
        Me.lblWarehouseb4.Visible = False
        '
        'lblLocationb4
        '
        Me.lblLocationb4.AutoSize = True
        Me.lblLocationb4.Location = New System.Drawing.Point(561, 19)
        Me.lblLocationb4.Name = "lblLocationb4"
        Me.lblLocationb4.Size = New System.Drawing.Size(24, 13)
        Me.lblLocationb4.TabIndex = 99
        Me.lblLocationb4.Text = "n/a"
        Me.lblLocationb4.Visible = False
        '
        'lblSecQtyb4
        '
        Me.lblSecQtyb4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecQtyb4.AutoSize = True
        Me.lblSecQtyb4.Location = New System.Drawing.Point(706, 71)
        Me.lblSecQtyb4.Name = "lblSecQtyb4"
        Me.lblSecQtyb4.Size = New System.Drawing.Size(18, 13)
        Me.lblSecQtyb4.TabIndex = 98
        Me.lblSecQtyb4.Text = "@"
        '
        'lblExpiryDateb4
        '
        Me.lblExpiryDateb4.AutoSize = True
        Me.lblExpiryDateb4.Location = New System.Drawing.Point(105, 99)
        Me.lblExpiryDateb4.Name = "lblExpiryDateb4"
        Me.lblExpiryDateb4.Size = New System.Drawing.Size(24, 13)
        Me.lblExpiryDateb4.TabIndex = 97
        Me.lblExpiryDateb4.Text = "n/a"
        Me.lblExpiryDateb4.Visible = False
        '
        'txtPrimaryPDOb4
        '
        Me.txtPrimaryPDOb4.Location = New System.Drawing.Point(440, 26)
        Me.txtPrimaryPDOb4.Name = "txtPrimaryPDOb4"
        Me.txtPrimaryPDOb4.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryPDOb4.TabIndex = 96
        Me.txtPrimaryPDOb4.TabStop = False
        Me.txtPrimaryPDOb4.Visible = False
        '
        'txtPrimaryCategoryb4
        '
        Me.txtPrimaryCategoryb4.Location = New System.Drawing.Point(145, 16)
        Me.txtPrimaryCategoryb4.Name = "txtPrimaryCategoryb4"
        Me.txtPrimaryCategoryb4.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCategoryb4.TabIndex = 1
        Me.txtPrimaryCategoryb4.TabStop = False
        Me.txtPrimaryCategoryb4.Visible = False
        '
        'txtPrimaryLocationb4
        '
        Me.txtPrimaryLocationb4.Location = New System.Drawing.Point(530, 17)
        Me.txtPrimaryLocationb4.Name = "txtPrimaryLocationb4"
        Me.txtPrimaryLocationb4.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryLocationb4.TabIndex = 94
        Me.txtPrimaryLocationb4.TabStop = False
        Me.txtPrimaryLocationb4.Visible = False
        '
        'chkLocationControlledb4
        '
        Me.chkLocationControlledb4.AutoSize = True
        Me.chkLocationControlledb4.Location = New System.Drawing.Point(369, 54)
        Me.chkLocationControlledb4.Name = "chkLocationControlledb4"
        Me.chkLocationControlledb4.Size = New System.Drawing.Size(67, 17)
        Me.chkLocationControlledb4.TabIndex = 8
        Me.chkLocationControlledb4.TabStop = False
        Me.chkLocationControlledb4.Text = "Location"
        Me.chkLocationControlledb4.UseVisualStyleBackColor = True
        Me.chkLocationControlledb4.Visible = False
        '
        'chkExpiryDateControlledb4
        '
        Me.chkExpiryDateControlledb4.AutoSize = True
        Me.chkExpiryDateControlledb4.Location = New System.Drawing.Point(159, 54)
        Me.chkExpiryDateControlledb4.Name = "chkExpiryDateControlledb4"
        Me.chkExpiryDateControlledb4.Size = New System.Drawing.Size(80, 17)
        Me.chkExpiryDateControlledb4.TabIndex = 6
        Me.chkExpiryDateControlledb4.TabStop = False
        Me.chkExpiryDateControlledb4.Text = "Expiry Date"
        Me.chkExpiryDateControlledb4.UseVisualStyleBackColor = True
        Me.chkExpiryDateControlledb4.Visible = False
        '
        'chkBatchNoControlledb4
        '
        Me.chkBatchNoControlledb4.AutoSize = True
        Me.chkBatchNoControlledb4.Location = New System.Drawing.Point(517, 43)
        Me.chkBatchNoControlledb4.Name = "chkBatchNoControlledb4"
        Me.chkBatchNoControlledb4.Size = New System.Drawing.Size(71, 17)
        Me.chkBatchNoControlledb4.TabIndex = 91
        Me.chkBatchNoControlledb4.TabStop = False
        Me.chkBatchNoControlledb4.Text = "Batch No"
        Me.chkBatchNoControlledb4.UseVisualStyleBackColor = True
        Me.chkBatchNoControlledb4.Visible = False
        '
        'chkPdoControlledb4
        '
        Me.chkPdoControlledb4.AutoSize = True
        Me.chkPdoControlledb4.Location = New System.Drawing.Point(517, 62)
        Me.chkPdoControlledb4.Name = "chkPdoControlledb4"
        Me.chkPdoControlledb4.Size = New System.Drawing.Size(47, 17)
        Me.chkPdoControlledb4.TabIndex = 90
        Me.chkPdoControlledb4.TabStop = False
        Me.chkPdoControlledb4.Text = "PdO"
        Me.chkPdoControlledb4.UseVisualStyleBackColor = True
        Me.chkPdoControlledb4.Visible = False
        '
        'chkUseSecondaryUnitb4
        '
        Me.chkUseSecondaryUnitb4.AutoSize = True
        Me.chkUseSecondaryUnitb4.Location = New System.Drawing.Point(517, 85)
        Me.chkUseSecondaryUnitb4.Name = "chkUseSecondaryUnitb4"
        Me.chkUseSecondaryUnitb4.Size = New System.Drawing.Size(121, 17)
        Me.chkUseSecondaryUnitb4.TabIndex = 89
        Me.chkUseSecondaryUnitb4.TabStop = False
        Me.chkUseSecondaryUnitb4.Text = "Use Secondary Unit"
        Me.chkUseSecondaryUnitb4.UseVisualStyleBackColor = True
        Me.chkUseSecondaryUnitb4.Visible = False
        '
        'txtPrimarySecondaryUnitsb4
        '
        Me.txtPrimarySecondaryUnitsb4.Location = New System.Drawing.Point(481, 53)
        Me.txtPrimarySecondaryUnitsb4.Name = "txtPrimarySecondaryUnitsb4"
        Me.txtPrimarySecondaryUnitsb4.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimarySecondaryUnitsb4.TabIndex = 88
        Me.txtPrimarySecondaryUnitsb4.TabStop = False
        Me.txtPrimarySecondaryUnitsb4.Visible = False
        '
        'txtPrimaryUnitsb4
        '
        Me.txtPrimaryUnitsb4.Location = New System.Drawing.Point(439, 53)
        Me.txtPrimaryUnitsb4.Name = "txtPrimaryUnitsb4"
        Me.txtPrimaryUnitsb4.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryUnitsb4.TabIndex = 87
        Me.txtPrimaryUnitsb4.TabStop = False
        Me.txtPrimaryUnitsb4.Visible = False
        '
        'txtSecondaryQtyb4
        '
        Me.txtSecondaryQtyb4.Location = New System.Drawing.Point(520, 68)
        Me.txtSecondaryQtyb4.MaxLength = 30
        Me.txtSecondaryQtyb4.Name = "txtSecondaryQtyb4"
        Me.txtSecondaryQtyb4.Size = New System.Drawing.Size(35, 20)
        Me.txtSecondaryQtyb4.TabIndex = 9
        Me.txtSecondaryQtyb4.TabStop = False
        Me.txtSecondaryQtyb4.Text = "0"
        Me.txtSecondaryQtyb4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSecondaryQtyb4.Visible = False
        '
        'lblUnitsb4
        '
        Me.lblUnitsb4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsb4.AutoSize = True
        Me.lblUnitsb4.Location = New System.Drawing.Point(669, 71)
        Me.lblUnitsb4.Name = "lblUnitsb4"
        Me.lblUnitsb4.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsb4.TabIndex = 13
        Me.lblUnitsb4.Text = "Units"
        '
        'TxtQtyperBatchb4
        '
        Me.TxtQtyperBatchb4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchb4.Location = New System.Drawing.Point(561, 68)
        Me.TxtQtyperBatchb4.MaxLength = 21
        Me.TxtQtyperBatchb4.Name = "TxtQtyperBatchb4"
        Me.TxtQtyperBatchb4.Size = New System.Drawing.Size(102, 20)
        Me.TxtQtyperBatchb4.TabIndex = 12
        Me.TxtQtyperBatchb4.Text = "0"
        Me.TxtQtyperBatchb4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
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
        'DtpExpiredb4
        '
        Me.DtpExpiredb4.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpExpiredb4.Enabled = False
        Me.DtpExpiredb4.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpExpiredb4.Location = New System.Drawing.Point(105, 96)
        Me.DtpExpiredb4.Name = "DtpExpiredb4"
        Me.DtpExpiredb4.Size = New System.Drawing.Size(327, 20)
        Me.DtpExpiredb4.TabIndex = 9
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
        'txtDescb4
        '
        Me.txtDescb4.Location = New System.Drawing.Point(105, 42)
        Me.txtDescb4.MaxLength = 6
        Me.txtDescb4.Name = "txtDescb4"
        Me.txtDescb4.ReadOnly = True
        Me.txtDescb4.Size = New System.Drawing.Size(327, 20)
        Me.txtDescb4.TabIndex = 4
        Me.txtDescb4.TabStop = False
        '
        'TxtRemarkItemb4
        '
        Me.TxtRemarkItemb4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItemb4.Location = New System.Drawing.Point(561, 96)
        Me.TxtRemarkItemb4.MaxLength = 50
        Me.TxtRemarkItemb4.Name = "TxtRemarkItemb4"
        Me.TxtRemarkItemb4.Size = New System.Drawing.Size(227, 20)
        Me.TxtRemarkItemb4.TabIndex = 15
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
        'TxtWarehouseb4
        '
        Me.TxtWarehouseb4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtWarehouseb4.Location = New System.Drawing.Point(561, 42)
        Me.TxtWarehouseb4.MaxLength = 50
        Me.TxtWarehouseb4.Name = "TxtWarehouseb4"
        Me.TxtWarehouseb4.ReadOnly = True
        Me.TxtWarehouseb4.Size = New System.Drawing.Size(285, 20)
        Me.TxtWarehouseb4.TabIndex = 11
        Me.TxtWarehouseb4.TabStop = False
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
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(16, 19)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(49, 13)
        Me.Label10.TabIndex = 52
        Me.Label10.Text = "Code No"
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
        'BtnInputb4
        '
        Me.BtnInputb4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnInputb4.Location = New System.Drawing.Point(794, 93)
        Me.BtnInputb4.Name = "BtnInputb4"
        Me.BtnInputb4.Size = New System.Drawing.Size(52, 25)
        Me.BtnInputb4.TabIndex = 16
        Me.BtnInputb4.Text = "&Input"
        Me.BtnInputb4.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCodeNob4
        '
        Me.TxtPrimaryCodeNob4.Location = New System.Drawing.Point(114, 16)
        Me.TxtPrimaryCodeNob4.Name = "TxtPrimaryCodeNob4"
        Me.TxtPrimaryCodeNob4.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCodeNob4.TabIndex = 2
        Me.TxtPrimaryCodeNob4.TabStop = False
        Me.TxtPrimaryCodeNob4.Visible = False
        '
        'TxtCodeNob4
        '
        Me.TxtCodeNob4.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNob4.Location = New System.Drawing.Point(105, 16)
        Me.TxtCodeNob4.MaxLength = 50
        Me.TxtCodeNob4.Name = "TxtCodeNob4"
        Me.TxtCodeNob4.Size = New System.Drawing.Size(287, 20)
        Me.TxtCodeNob4.TabIndex = 0
        '
        'cboBatchNob4
        '
        Me.cboBatchNob4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBatchNob4.FormattingEnabled = True
        Me.cboBatchNob4.Location = New System.Drawing.Point(105, 68)
        Me.cboBatchNob4.Name = "cboBatchNob4"
        Me.cboBatchNob4.Size = New System.Drawing.Size(327, 21)
        Me.cboBatchNob4.TabIndex = 7
        '
        'cboSecondaryUnitsb4
        '
        Me.cboSecondaryUnitsb4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboSecondaryUnitsb4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSecondaryUnitsb4.FormattingEnabled = True
        Me.cboSecondaryUnitsb4.Location = New System.Drawing.Point(730, 68)
        Me.cboSecondaryUnitsb4.Name = "cboSecondaryUnitsb4"
        Me.cboSecondaryUnitsb4.Size = New System.Drawing.Size(70, 21)
        Me.cboSecondaryUnitsb4.TabIndex = 14
        '
        'lblSecondaryUnitsb4
        '
        Me.lblSecondaryUnitsb4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecondaryUnitsb4.AutoSize = True
        Me.lblSecondaryUnitsb4.Location = New System.Drawing.Point(806, 71)
        Me.lblSecondaryUnitsb4.Name = "lblSecondaryUnitsb4"
        Me.lblSecondaryUnitsb4.Size = New System.Drawing.Size(31, 13)
        Me.lblSecondaryUnitsb4.TabIndex = 86
        Me.lblSecondaryUnitsb4.Text = "Units"
        '
        'BtnFindCodeNob4
        '
        Me.BtnFindCodeNob4.Location = New System.Drawing.Point(400, 17)
        Me.BtnFindCodeNob4.Name = "BtnFindCodeNob4"
        Me.BtnFindCodeNob4.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindCodeNob4.TabIndex = 3
        Me.BtnFindCodeNob4.TabStop = False
        Me.BtnFindCodeNob4.Text = "...."
        Me.BtnFindCodeNob4.UseVisualStyleBackColor = True
        '
        'gboItemAfter
        '
        Me.gboItemAfter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItemAfter.Controls.Add(Me.txtFaktorPemberat)
        Me.gboItemAfter.Controls.Add(Me.lblFaktorPemberat)
        Me.gboItemAfter.Controls.Add(Me.lblBatchNoAfter)
        Me.gboItemAfter.Controls.Add(Me.lblWarehouseAfter)
        Me.gboItemAfter.Controls.Add(Me.lblLocationAfter)
        Me.gboItemAfter.Controls.Add(Me.lblSecQtyAfter)
        Me.gboItemAfter.Controls.Add(Me.lblExpiryDateAfter)
        Me.gboItemAfter.Controls.Add(Me.txtPrimaryPDOAfter)
        Me.gboItemAfter.Controls.Add(Me.txtPrimaryCategoryAfter)
        Me.gboItemAfter.Controls.Add(Me.txtPrimaryLocationAfter)
        Me.gboItemAfter.Controls.Add(Me.chkLocationControlledAfter)
        Me.gboItemAfter.Controls.Add(Me.chkExpiryDateControlledAfter)
        Me.gboItemAfter.Controls.Add(Me.chkBatchNoControlledAfter)
        Me.gboItemAfter.Controls.Add(Me.chkPdoControlledAfter)
        Me.gboItemAfter.Controls.Add(Me.chkUseSecondaryUnitAfter)
        Me.gboItemAfter.Controls.Add(Me.txtPrimarySecondaryUnitsAfter)
        Me.gboItemAfter.Controls.Add(Me.txtPrimaryUnitsAfter)
        Me.gboItemAfter.Controls.Add(Me.lblSecondaryUnitsAfter)
        Me.gboItemAfter.Controls.Add(Me.txtSecondaryQtyAfter)
        Me.gboItemAfter.Controls.Add(Me.lblUnitsAfter)
        Me.gboItemAfter.Controls.Add(Me.TxtQtyperBatchAfter)
        Me.gboItemAfter.Controls.Add(Me.txtBatchNoAfter)
        Me.gboItemAfter.Controls.Add(Me.Label1)
        Me.gboItemAfter.Controls.Add(Me.DtpExpiredAfter)
        Me.gboItemAfter.Controls.Add(Me.Label2)
        Me.gboItemAfter.Controls.Add(Me.btnFindLocationAfter)
        Me.gboItemAfter.Controls.Add(Me.Label5)
        Me.gboItemAfter.Controls.Add(Me.txtDescAfter)
        Me.gboItemAfter.Controls.Add(Me.TxtRemarkItemAfter)
        Me.gboItemAfter.Controls.Add(Me.Label7)
        Me.gboItemAfter.Controls.Add(Me.TxtWarehouseAfter)
        Me.gboItemAfter.Controls.Add(Me.Label8)
        Me.gboItemAfter.Controls.Add(Me.Label17)
        Me.gboItemAfter.Controls.Add(Me.Label18)
        Me.gboItemAfter.Controls.Add(Me.BtnFindCodeNoAfter)
        Me.gboItemAfter.Controls.Add(Me.Label20)
        Me.gboItemAfter.Controls.Add(Me.BtnFindPONo)
        Me.gboItemAfter.Controls.Add(Me.Label21)
        Me.gboItemAfter.Controls.Add(Me.TxtLocationAfter)
        Me.gboItemAfter.Controls.Add(Me.TxtPrimaryCodeNoAfter)
        Me.gboItemAfter.Controls.Add(Me.TxtCodeNoAfter)
        Me.gboItemAfter.Location = New System.Drawing.Point(11, 12)
        Me.gboItemAfter.Name = "gboItemAfter"
        Me.gboItemAfter.Size = New System.Drawing.Size(859, 125)
        Me.gboItemAfter.TabIndex = 11
        Me.gboItemAfter.TabStop = False
        Me.gboItemAfter.Text = "Item After"
        '
        'txtFaktorPemberat
        '
        Me.txtFaktorPemberat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFaktorPemberat.Location = New System.Drawing.Point(561, 122)
        Me.txtFaktorPemberat.MaxLength = 50
        Me.txtFaktorPemberat.Name = "txtFaktorPemberat"
        Me.txtFaktorPemberat.Size = New System.Drawing.Size(227, 20)
        Me.txtFaktorPemberat.TabIndex = 11
        Me.txtFaktorPemberat.Text = "0"
        Me.txtFaktorPemberat.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtFaktorPemberat.Visible = False
        '
        'lblFaktorPemberat
        '
        Me.lblFaktorPemberat.AutoSize = True
        Me.lblFaktorPemberat.Location = New System.Drawing.Point(457, 127)
        Me.lblFaktorPemberat.Name = "lblFaktorPemberat"
        Me.lblFaktorPemberat.Size = New System.Drawing.Size(49, 13)
        Me.lblFaktorPemberat.TabIndex = 103
        Me.lblFaktorPemberat.Text = "Remarks"
        Me.lblFaktorPemberat.Visible = False
        '
        'lblBatchNoAfter
        '
        Me.lblBatchNoAfter.AutoSize = True
        Me.lblBatchNoAfter.Location = New System.Drawing.Point(105, 73)
        Me.lblBatchNoAfter.Name = "lblBatchNoAfter"
        Me.lblBatchNoAfter.Size = New System.Drawing.Size(24, 13)
        Me.lblBatchNoAfter.TabIndex = 101
        Me.lblBatchNoAfter.Text = "n/a"
        Me.lblBatchNoAfter.Visible = False
        '
        'lblWarehouseAfter
        '
        Me.lblWarehouseAfter.AutoSize = True
        Me.lblWarehouseAfter.Location = New System.Drawing.Point(561, 45)
        Me.lblWarehouseAfter.Name = "lblWarehouseAfter"
        Me.lblWarehouseAfter.Size = New System.Drawing.Size(24, 13)
        Me.lblWarehouseAfter.TabIndex = 100
        Me.lblWarehouseAfter.Text = "n/a"
        Me.lblWarehouseAfter.Visible = False
        '
        'lblLocationAfter
        '
        Me.lblLocationAfter.AutoSize = True
        Me.lblLocationAfter.Location = New System.Drawing.Point(561, 19)
        Me.lblLocationAfter.Name = "lblLocationAfter"
        Me.lblLocationAfter.Size = New System.Drawing.Size(24, 13)
        Me.lblLocationAfter.TabIndex = 99
        Me.lblLocationAfter.Text = "n/a"
        Me.lblLocationAfter.Visible = False
        '
        'lblSecQtyAfter
        '
        Me.lblSecQtyAfter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecQtyAfter.AutoSize = True
        Me.lblSecQtyAfter.Location = New System.Drawing.Point(706, 71)
        Me.lblSecQtyAfter.Name = "lblSecQtyAfter"
        Me.lblSecQtyAfter.Size = New System.Drawing.Size(18, 13)
        Me.lblSecQtyAfter.TabIndex = 98
        Me.lblSecQtyAfter.Text = "@"
        '
        'lblExpiryDateAfter
        '
        Me.lblExpiryDateAfter.AutoSize = True
        Me.lblExpiryDateAfter.Location = New System.Drawing.Point(105, 99)
        Me.lblExpiryDateAfter.Name = "lblExpiryDateAfter"
        Me.lblExpiryDateAfter.Size = New System.Drawing.Size(24, 13)
        Me.lblExpiryDateAfter.TabIndex = 97
        Me.lblExpiryDateAfter.Text = "n/a"
        Me.lblExpiryDateAfter.Visible = False
        '
        'txtPrimaryPDOAfter
        '
        Me.txtPrimaryPDOAfter.Location = New System.Drawing.Point(440, 26)
        Me.txtPrimaryPDOAfter.Name = "txtPrimaryPDOAfter"
        Me.txtPrimaryPDOAfter.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryPDOAfter.TabIndex = 96
        Me.txtPrimaryPDOAfter.TabStop = False
        Me.txtPrimaryPDOAfter.Visible = False
        '
        'txtPrimaryCategoryAfter
        '
        Me.txtPrimaryCategoryAfter.Location = New System.Drawing.Point(145, 16)
        Me.txtPrimaryCategoryAfter.Name = "txtPrimaryCategoryAfter"
        Me.txtPrimaryCategoryAfter.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCategoryAfter.TabIndex = 95
        Me.txtPrimaryCategoryAfter.TabStop = False
        Me.txtPrimaryCategoryAfter.Visible = False
        '
        'txtPrimaryLocationAfter
        '
        Me.txtPrimaryLocationAfter.Location = New System.Drawing.Point(530, 17)
        Me.txtPrimaryLocationAfter.Name = "txtPrimaryLocationAfter"
        Me.txtPrimaryLocationAfter.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryLocationAfter.TabIndex = 94
        Me.txtPrimaryLocationAfter.TabStop = False
        Me.txtPrimaryLocationAfter.Visible = False
        '
        'chkLocationControlledAfter
        '
        Me.chkLocationControlledAfter.AutoSize = True
        Me.chkLocationControlledAfter.Location = New System.Drawing.Point(369, 54)
        Me.chkLocationControlledAfter.Name = "chkLocationControlledAfter"
        Me.chkLocationControlledAfter.Size = New System.Drawing.Size(67, 17)
        Me.chkLocationControlledAfter.TabIndex = 93
        Me.chkLocationControlledAfter.TabStop = False
        Me.chkLocationControlledAfter.Text = "Location"
        Me.chkLocationControlledAfter.UseVisualStyleBackColor = True
        Me.chkLocationControlledAfter.Visible = False
        '
        'chkExpiryDateControlledAfter
        '
        Me.chkExpiryDateControlledAfter.AutoSize = True
        Me.chkExpiryDateControlledAfter.Location = New System.Drawing.Point(159, 54)
        Me.chkExpiryDateControlledAfter.Name = "chkExpiryDateControlledAfter"
        Me.chkExpiryDateControlledAfter.Size = New System.Drawing.Size(80, 17)
        Me.chkExpiryDateControlledAfter.TabIndex = 92
        Me.chkExpiryDateControlledAfter.TabStop = False
        Me.chkExpiryDateControlledAfter.Text = "Expiry Date"
        Me.chkExpiryDateControlledAfter.UseVisualStyleBackColor = True
        Me.chkExpiryDateControlledAfter.Visible = False
        '
        'chkBatchNoControlledAfter
        '
        Me.chkBatchNoControlledAfter.AutoSize = True
        Me.chkBatchNoControlledAfter.Location = New System.Drawing.Point(517, 43)
        Me.chkBatchNoControlledAfter.Name = "chkBatchNoControlledAfter"
        Me.chkBatchNoControlledAfter.Size = New System.Drawing.Size(71, 17)
        Me.chkBatchNoControlledAfter.TabIndex = 91
        Me.chkBatchNoControlledAfter.TabStop = False
        Me.chkBatchNoControlledAfter.Text = "Batch No"
        Me.chkBatchNoControlledAfter.UseVisualStyleBackColor = True
        Me.chkBatchNoControlledAfter.Visible = False
        '
        'chkPdoControlledAfter
        '
        Me.chkPdoControlledAfter.AutoSize = True
        Me.chkPdoControlledAfter.Location = New System.Drawing.Point(517, 62)
        Me.chkPdoControlledAfter.Name = "chkPdoControlledAfter"
        Me.chkPdoControlledAfter.Size = New System.Drawing.Size(47, 17)
        Me.chkPdoControlledAfter.TabIndex = 90
        Me.chkPdoControlledAfter.TabStop = False
        Me.chkPdoControlledAfter.Text = "PdO"
        Me.chkPdoControlledAfter.UseVisualStyleBackColor = True
        Me.chkPdoControlledAfter.Visible = False
        '
        'chkUseSecondaryUnitAfter
        '
        Me.chkUseSecondaryUnitAfter.AutoSize = True
        Me.chkUseSecondaryUnitAfter.Location = New System.Drawing.Point(517, 85)
        Me.chkUseSecondaryUnitAfter.Name = "chkUseSecondaryUnitAfter"
        Me.chkUseSecondaryUnitAfter.Size = New System.Drawing.Size(121, 17)
        Me.chkUseSecondaryUnitAfter.TabIndex = 89
        Me.chkUseSecondaryUnitAfter.TabStop = False
        Me.chkUseSecondaryUnitAfter.Text = "Use Secondary Unit"
        Me.chkUseSecondaryUnitAfter.UseVisualStyleBackColor = True
        Me.chkUseSecondaryUnitAfter.Visible = False
        '
        'txtPrimarySecondaryUnitsAfter
        '
        Me.txtPrimarySecondaryUnitsAfter.Location = New System.Drawing.Point(481, 53)
        Me.txtPrimarySecondaryUnitsAfter.Name = "txtPrimarySecondaryUnitsAfter"
        Me.txtPrimarySecondaryUnitsAfter.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimarySecondaryUnitsAfter.TabIndex = 88
        Me.txtPrimarySecondaryUnitsAfter.TabStop = False
        Me.txtPrimarySecondaryUnitsAfter.Visible = False
        '
        'txtPrimaryUnitsAfter
        '
        Me.txtPrimaryUnitsAfter.Location = New System.Drawing.Point(439, 53)
        Me.txtPrimaryUnitsAfter.Name = "txtPrimaryUnitsAfter"
        Me.txtPrimaryUnitsAfter.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimaryUnitsAfter.TabIndex = 87
        Me.txtPrimaryUnitsAfter.TabStop = False
        Me.txtPrimaryUnitsAfter.Visible = False
        '
        'lblSecondaryUnitsAfter
        '
        Me.lblSecondaryUnitsAfter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSecondaryUnitsAfter.AutoSize = True
        Me.lblSecondaryUnitsAfter.Location = New System.Drawing.Point(806, 71)
        Me.lblSecondaryUnitsAfter.Name = "lblSecondaryUnitsAfter"
        Me.lblSecondaryUnitsAfter.Size = New System.Drawing.Size(31, 13)
        Me.lblSecondaryUnitsAfter.TabIndex = 86
        Me.lblSecondaryUnitsAfter.Text = "Units"
        '
        'txtSecondaryQtyAfter
        '
        Me.txtSecondaryQtyAfter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSecondaryQtyAfter.Location = New System.Drawing.Point(730, 68)
        Me.txtSecondaryQtyAfter.MaxLength = 30
        Me.txtSecondaryQtyAfter.Name = "txtSecondaryQtyAfter"
        Me.txtSecondaryQtyAfter.Size = New System.Drawing.Size(70, 20)
        Me.txtSecondaryQtyAfter.TabIndex = 9
        Me.txtSecondaryQtyAfter.Text = "0"
        Me.txtSecondaryQtyAfter.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblUnitsAfter
        '
        Me.lblUnitsAfter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnitsAfter.AutoSize = True
        Me.lblUnitsAfter.Location = New System.Drawing.Point(669, 71)
        Me.lblUnitsAfter.Name = "lblUnitsAfter"
        Me.lblUnitsAfter.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitsAfter.TabIndex = 84
        Me.lblUnitsAfter.Text = "Units"
        '
        'TxtQtyperBatchAfter
        '
        Me.TxtQtyperBatchAfter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyperBatchAfter.Location = New System.Drawing.Point(561, 68)
        Me.TxtQtyperBatchAfter.MaxLength = 50
        Me.TxtQtyperBatchAfter.Name = "TxtQtyperBatchAfter"
        Me.TxtQtyperBatchAfter.Size = New System.Drawing.Size(102, 20)
        Me.TxtQtyperBatchAfter.TabIndex = 8
        Me.TxtQtyperBatchAfter.Text = "0"
        Me.TxtQtyperBatchAfter.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtBatchNoAfter
        '
        Me.txtBatchNoAfter.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtBatchNoAfter.Location = New System.Drawing.Point(105, 70)
        Me.txtBatchNoAfter.MaxLength = 50
        Me.txtBatchNoAfter.Name = "txtBatchNoAfter"
        Me.txtBatchNoAfter.Size = New System.Drawing.Size(327, 20)
        Me.txtBatchNoAfter.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 73)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 13)
        Me.Label1.TabIndex = 79
        Me.Label1.Text = "Batch No"
        '
        'DtpExpiredAfter
        '
        Me.DtpExpiredAfter.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpExpiredAfter.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpExpiredAfter.Location = New System.Drawing.Point(105, 96)
        Me.DtpExpiredAfter.Name = "DtpExpiredAfter"
        Me.DtpExpiredAfter.Size = New System.Drawing.Size(327, 20)
        Me.DtpExpiredAfter.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 99)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(61, 13)
        Me.Label2.TabIndex = 77
        Me.Label2.Text = "Expiry Date"
        '
        'btnFindLocationAfter
        '
        Me.btnFindLocationAfter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindLocationAfter.Location = New System.Drawing.Point(814, 16)
        Me.btnFindLocationAfter.Name = "btnFindLocationAfter"
        Me.btnFindLocationAfter.Size = New System.Drawing.Size(32, 20)
        Me.btnFindLocationAfter.TabIndex = 6
        Me.btnFindLocationAfter.TabStop = False
        Me.btnFindLocationAfter.Text = "...."
        Me.btnFindLocationAfter.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(457, 19)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 13)
        Me.Label5.TabIndex = 73
        Me.Label5.Text = "Location"
        '
        'txtDescAfter
        '
        Me.txtDescAfter.Location = New System.Drawing.Point(105, 42)
        Me.txtDescAfter.MaxLength = 6
        Me.txtDescAfter.Name = "txtDescAfter"
        Me.txtDescAfter.ReadOnly = True
        Me.txtDescAfter.Size = New System.Drawing.Size(327, 20)
        Me.txtDescAfter.TabIndex = 2
        Me.txtDescAfter.TabStop = False
        '
        'TxtRemarkItemAfter
        '
        Me.TxtRemarkItemAfter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarkItemAfter.Location = New System.Drawing.Point(561, 96)
        Me.TxtRemarkItemAfter.MaxLength = 50
        Me.TxtRemarkItemAfter.Name = "TxtRemarkItemAfter"
        Me.TxtRemarkItemAfter.Size = New System.Drawing.Size(285, 20)
        Me.TxtRemarkItemAfter.TabIndex = 10
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(457, 101)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(49, 13)
        Me.Label7.TabIndex = 63
        Me.Label7.Text = "Remarks"
        '
        'TxtWarehouseAfter
        '
        Me.TxtWarehouseAfter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtWarehouseAfter.Location = New System.Drawing.Point(561, 42)
        Me.TxtWarehouseAfter.MaxLength = 50
        Me.TxtWarehouseAfter.Name = "TxtWarehouseAfter"
        Me.TxtWarehouseAfter.ReadOnly = True
        Me.TxtWarehouseAfter.Size = New System.Drawing.Size(285, 20)
        Me.TxtWarehouseAfter.TabIndex = 7
        Me.TxtWarehouseAfter.TabStop = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(457, 42)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(62, 13)
        Me.Label8.TabIndex = 65
        Me.Label8.Text = "Warehouse"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(457, 73)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(23, 13)
        Me.Label17.TabIndex = 63
        Me.Label17.Text = "Qty"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(16, 45)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(60, 13)
        Me.Label18.TabIndex = 62
        Me.Label18.Text = "Description"
        '
        'BtnFindCodeNoAfter
        '
        Me.BtnFindCodeNoAfter.Location = New System.Drawing.Point(400, 17)
        Me.BtnFindCodeNoAfter.Name = "BtnFindCodeNoAfter"
        Me.BtnFindCodeNoAfter.Size = New System.Drawing.Size(32, 20)
        Me.BtnFindCodeNoAfter.TabIndex = 1
        Me.BtnFindCodeNoAfter.TabStop = False
        Me.BtnFindCodeNoAfter.Text = "...."
        Me.BtnFindCodeNoAfter.UseVisualStyleBackColor = True
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(16, 19)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(49, 13)
        Me.Label20.TabIndex = 52
        Me.Label20.Text = "Code No"
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
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(16, 19)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(39, 13)
        Me.Label21.TabIndex = 48
        Me.Label21.Text = "PO No"
        Me.Label21.Visible = False
        '
        'TxtLocationAfter
        '
        Me.TxtLocationAfter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLocationAfter.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtLocationAfter.Location = New System.Drawing.Point(561, 16)
        Me.TxtLocationAfter.MaxLength = 50
        Me.TxtLocationAfter.Name = "TxtLocationAfter"
        Me.TxtLocationAfter.Size = New System.Drawing.Size(250, 20)
        Me.TxtLocationAfter.TabIndex = 5
        '
        'TxtPrimaryCodeNoAfter
        '
        Me.TxtPrimaryCodeNoAfter.Location = New System.Drawing.Point(114, 16)
        Me.TxtPrimaryCodeNoAfter.Name = "TxtPrimaryCodeNoAfter"
        Me.TxtPrimaryCodeNoAfter.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCodeNoAfter.TabIndex = 53
        Me.TxtPrimaryCodeNoAfter.TabStop = False
        Me.TxtPrimaryCodeNoAfter.Visible = False
        '
        'TxtCodeNoAfter
        '
        Me.TxtCodeNoAfter.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNoAfter.Location = New System.Drawing.Point(105, 16)
        Me.TxtCodeNoAfter.MaxLength = 50
        Me.TxtCodeNoAfter.Name = "TxtCodeNoAfter"
        Me.TxtCodeNoAfter.Size = New System.Drawing.Size(287, 20)
        Me.TxtCodeNoAfter.TabIndex = 0
        '
        'TxtPrimaryStockIn
        '
        Me.TxtPrimaryStockIn.Location = New System.Drawing.Point(115, 2)
        Me.TxtPrimaryStockIn.Name = "TxtPrimaryStockIn"
        Me.TxtPrimaryStockIn.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryStockIn.TabIndex = 104
        Me.TxtPrimaryStockIn.TabStop = False
        Me.TxtPrimaryStockIn.Text = "-1"
        Me.TxtPrimaryStockIn.Visible = False
        '
        'FrmTransaksiStockBundling_BOM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(881, 480)
        Me.Controls.Add(Me.TxtPrimaryStockIn)
        Me.Controls.Add(Me.gboItemAfter)
        Me.Controls.Add(Me.gboItemb4)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(507, 404)
        Me.Name = "FrmTransaksiStockBundling_BOM"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Input By Bill Of Material"
        Me.gboItemb4.ResumeLayout(False)
        Me.gboItemb4.PerformLayout()
        CType(Me.dgvItemb4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboItemAfter.ResumeLayout(False)
        Me.gboItemAfter.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents gboItemb4 As System.Windows.Forms.GroupBox
    Friend WithEvents cboLocationb4 As System.Windows.Forms.ComboBox
    Friend WithEvents lblBatchNob4 As System.Windows.Forms.Label
    Friend WithEvents lblWarehouseb4 As System.Windows.Forms.Label
    Friend WithEvents lblLocationb4 As System.Windows.Forms.Label
    Friend WithEvents lblSecQtyb4 As System.Windows.Forms.Label
    Friend WithEvents lblExpiryDateb4 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryPDOb4 As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCategoryb4 As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryLocationb4 As System.Windows.Forms.TextBox
    Friend WithEvents chkLocationControlledb4 As System.Windows.Forms.CheckBox
    Friend WithEvents chkExpiryDateControlledb4 As System.Windows.Forms.CheckBox
    Friend WithEvents chkBatchNoControlledb4 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPdoControlledb4 As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseSecondaryUnitb4 As System.Windows.Forms.CheckBox
    Friend WithEvents txtPrimarySecondaryUnitsb4 As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnitsb4 As System.Windows.Forms.TextBox
    Friend WithEvents txtSecondaryQtyb4 As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsb4 As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatchb4 As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents DtpExpiredb4 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtDescb4 As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarkItemb4 As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouseb4 As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents BtnInputb4 As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCodeNob4 As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeNob4 As System.Windows.Forms.TextBox
    Friend WithEvents cboBatchNob4 As System.Windows.Forms.ComboBox
    Friend WithEvents cboSecondaryUnitsb4 As System.Windows.Forms.ComboBox
    Friend WithEvents lblSecondaryUnitsb4 As System.Windows.Forms.Label
    Friend WithEvents BtnFindCodeNob4 As System.Windows.Forms.Button
    Public WithEvents dgvItemb4 As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecondaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents gboItemAfter As System.Windows.Forms.GroupBox
    Friend WithEvents txtFaktorPemberat As System.Windows.Forms.TextBox
    Friend WithEvents lblFaktorPemberat As System.Windows.Forms.Label
    Friend WithEvents lblBatchNoAfter As System.Windows.Forms.Label
    Friend WithEvents lblWarehouseAfter As System.Windows.Forms.Label
    Friend WithEvents lblLocationAfter As System.Windows.Forms.Label
    Friend WithEvents lblSecQtyAfter As System.Windows.Forms.Label
    Friend WithEvents lblExpiryDateAfter As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryPDOAfter As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCategoryAfter As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryLocationAfter As System.Windows.Forms.TextBox
    Friend WithEvents chkLocationControlledAfter As System.Windows.Forms.CheckBox
    Friend WithEvents chkExpiryDateControlledAfter As System.Windows.Forms.CheckBox
    Friend WithEvents chkBatchNoControlledAfter As System.Windows.Forms.CheckBox
    Friend WithEvents chkPdoControlledAfter As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseSecondaryUnitAfter As System.Windows.Forms.CheckBox
    Friend WithEvents txtPrimarySecondaryUnitsAfter As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryUnitsAfter As System.Windows.Forms.TextBox
    Friend WithEvents lblSecondaryUnitsAfter As System.Windows.Forms.Label
    Friend WithEvents txtSecondaryQtyAfter As System.Windows.Forms.TextBox
    Friend WithEvents lblUnitsAfter As System.Windows.Forms.Label
    Friend WithEvents TxtQtyperBatchAfter As System.Windows.Forms.TextBox
    Friend WithEvents txtBatchNoAfter As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DtpExpiredAfter As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFindLocationAfter As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDescAfter As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemarkItemAfter As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtWarehouseAfter As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents BtnFindCodeNoAfter As System.Windows.Forms.Button
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents BtnFindPONo As System.Windows.Forms.Button
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TxtLocationAfter As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCodeNoAfter As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeNoAfter As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryStockIn As System.Windows.Forms.TextBox
End Class
