<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiModifyControl
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
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnFindCodeNo = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.TxtQty = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.LblUnit = New System.Windows.Forms.Label
        Me.gboInput = New System.Windows.Forms.GroupBox
        Me.LblUnit2 = New System.Windows.Forms.Label
        Me.btnInput = New System.Windows.Forms.Button
        Me.TxtQtyPerBatch = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.btnFindLocation = New System.Windows.Forms.Button
        Me.TxtLocationCode = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtPrimaryLocation = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.dtpED = New System.Windows.Forms.DateTimePicker
        Me.TxtBatch = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.chkBatch = New System.Windows.Forms.CheckBox
        Me.chkExpiryDate = New System.Windows.Forms.CheckBox
        Me.chkLocation = New System.Windows.Forms.CheckBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColBatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkED_Ori = New System.Windows.Forms.CheckBox
        Me.chk_Location_Ori = New System.Windows.Forms.CheckBox
        Me.chkBatch_Ori = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.gboInput.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtDesc
        '
        Me.txtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDesc.Location = New System.Drawing.Point(105, 38)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(408, 20)
        Me.txtDesc.TabIndex = 2
        Me.txtDesc.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 41)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "Description"
        '
        'btnFindCodeNo
        '
        Me.btnFindCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCodeNo.Location = New System.Drawing.Point(470, 12)
        Me.btnFindCodeNo.Name = "btnFindCodeNo"
        Me.btnFindCodeNo.Size = New System.Drawing.Size(43, 20)
        Me.btnFindCodeNo.TabIndex = 1
        Me.btnFindCodeNo.TabStop = False
        Me.btnFindCodeNo.Text = "...."
        Me.btnFindCodeNo.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(357, 510)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(438, 510)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "&Close"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(105, 12)
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(359, 20)
        Me.TxtCodeNo.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Code No"
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(67, 12)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(39, 20)
        Me.TxtPrimaryCodeNo.TabIndex = 0
        Me.TxtPrimaryCodeNo.TabStop = False
        Me.TxtPrimaryCodeNo.Text = "-1"
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'TxtQty
        '
        Me.TxtQty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQty.Location = New System.Drawing.Point(105, 64)
        Me.TxtQty.Name = "TxtQty"
        Me.TxtQty.ReadOnly = True
        Me.TxtQty.Size = New System.Drawing.Size(350, 20)
        Me.TxtQty.TabIndex = 3
        Me.TxtQty.TabStop = False
        Me.TxtQty.Text = "0"
        Me.TxtQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 67)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(23, 13)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Qty"
        '
        'LblUnit
        '
        Me.LblUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnit.AutoSize = True
        Me.LblUnit.Location = New System.Drawing.Point(461, 67)
        Me.LblUnit.Name = "LblUnit"
        Me.LblUnit.Size = New System.Drawing.Size(26, 13)
        Me.LblUnit.TabIndex = 4
        Me.LblUnit.Text = "Unit"
        '
        'gboInput
        '
        Me.gboInput.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboInput.Controls.Add(Me.LblUnit2)
        Me.gboInput.Controls.Add(Me.btnInput)
        Me.gboInput.Controls.Add(Me.TxtQtyPerBatch)
        Me.gboInput.Controls.Add(Me.Label8)
        Me.gboInput.Controls.Add(Me.btnFindLocation)
        Me.gboInput.Controls.Add(Me.TxtLocationCode)
        Me.gboInput.Controls.Add(Me.Label7)
        Me.gboInput.Controls.Add(Me.TxtPrimaryLocation)
        Me.gboInput.Controls.Add(Me.Label6)
        Me.gboInput.Controls.Add(Me.dtpED)
        Me.gboInput.Controls.Add(Me.TxtBatch)
        Me.gboInput.Controls.Add(Me.Label4)
        Me.gboInput.Location = New System.Drawing.Point(15, 186)
        Me.gboInput.Name = "gboInput"
        Me.gboInput.Size = New System.Drawing.Size(498, 148)
        Me.gboInput.TabIndex = 7
        Me.gboInput.TabStop = False
        '
        'LblUnit2
        '
        Me.LblUnit2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblUnit2.AutoSize = True
        Me.LblUnit2.Location = New System.Drawing.Point(446, 100)
        Me.LblUnit2.Name = "LblUnit2"
        Me.LblUnit2.Size = New System.Drawing.Size(26, 13)
        Me.LblUnit2.TabIndex = 6
        Me.LblUnit2.Text = "Unit"
        '
        'btnInput
        '
        Me.btnInput.Location = New System.Drawing.Point(90, 123)
        Me.btnInput.Name = "btnInput"
        Me.btnInput.Size = New System.Drawing.Size(55, 20)
        Me.btnInput.TabIndex = 7
        Me.btnInput.Text = "Input"
        Me.btnInput.UseVisualStyleBackColor = True
        '
        'TxtQtyPerBatch
        '
        Me.TxtQtyPerBatch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtQtyPerBatch.Location = New System.Drawing.Point(90, 97)
        Me.TxtQtyPerBatch.Name = "TxtQtyPerBatch"
        Me.TxtQtyPerBatch.Size = New System.Drawing.Size(350, 20)
        Me.TxtQtyPerBatch.TabIndex = 5
        Me.TxtQtyPerBatch.Text = "0"
        Me.TxtQtyPerBatch.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(17, 100)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(23, 13)
        Me.Label8.TabIndex = 32
        Me.Label8.Text = "Qty"
        '
        'btnFindLocation
        '
        Me.btnFindLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindLocation.Enabled = False
        Me.btnFindLocation.Location = New System.Drawing.Point(446, 71)
        Me.btnFindLocation.Name = "btnFindLocation"
        Me.btnFindLocation.Size = New System.Drawing.Size(43, 20)
        Me.btnFindLocation.TabIndex = 4
        Me.btnFindLocation.TabStop = False
        Me.btnFindLocation.Text = "...."
        Me.btnFindLocation.UseVisualStyleBackColor = True
        '
        'TxtLocationCode
        '
        Me.TxtLocationCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLocationCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtLocationCode.Enabled = False
        Me.TxtLocationCode.Location = New System.Drawing.Point(90, 71)
        Me.TxtLocationCode.Name = "TxtLocationCode"
        Me.TxtLocationCode.Size = New System.Drawing.Size(350, 20)
        Me.TxtLocationCode.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(11, 74)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(48, 13)
        Me.Label7.TabIndex = 28
        Me.Label7.Text = "Location"
        '
        'TxtPrimaryLocation
        '
        Me.TxtPrimaryLocation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryLocation.Location = New System.Drawing.Point(52, 71)
        Me.TxtPrimaryLocation.Name = "TxtPrimaryLocation"
        Me.TxtPrimaryLocation.Size = New System.Drawing.Size(39, 20)
        Me.TxtPrimaryLocation.TabIndex = 3
        Me.TxtPrimaryLocation.TabStop = False
        Me.TxtPrimaryLocation.Text = "-1"
        Me.TxtPrimaryLocation.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(11, 49)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(61, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Expiry Date"
        '
        'dtpED
        '
        Me.dtpED.CustomFormat = "dd MMM yyyy"
        Me.dtpED.Enabled = False
        Me.dtpED.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpED.Location = New System.Drawing.Point(90, 45)
        Me.dtpED.Name = "dtpED"
        Me.dtpED.Size = New System.Drawing.Size(399, 20)
        Me.dtpED.TabIndex = 1
        '
        'TxtBatch
        '
        Me.TxtBatch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBatch.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBatch.Enabled = False
        Me.TxtBatch.Location = New System.Drawing.Point(90, 19)
        Me.TxtBatch.Name = "TxtBatch"
        Me.TxtBatch.Size = New System.Drawing.Size(399, 20)
        Me.TxtBatch.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(75, 13)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Batch Number"
        '
        'chkBatch
        '
        Me.chkBatch.AutoSize = True
        Me.chkBatch.Location = New System.Drawing.Point(10, 13)
        Me.chkBatch.Name = "chkBatch"
        Me.chkBatch.Size = New System.Drawing.Size(130, 17)
        Me.chkBatch.TabIndex = 0
        Me.chkBatch.Text = "Batch Number Control"
        Me.chkBatch.UseVisualStyleBackColor = True
        '
        'chkExpiryDate
        '
        Me.chkExpiryDate.AutoSize = True
        Me.chkExpiryDate.Enabled = False
        Me.chkExpiryDate.Location = New System.Drawing.Point(148, 13)
        Me.chkExpiryDate.Name = "chkExpiryDate"
        Me.chkExpiryDate.Size = New System.Drawing.Size(130, 17)
        Me.chkExpiryDate.TabIndex = 1
        Me.chkExpiryDate.Text = "Expiry Date Controlled"
        Me.chkExpiryDate.UseVisualStyleBackColor = True
        '
        'chkLocation
        '
        Me.chkLocation.AutoSize = True
        Me.chkLocation.Location = New System.Drawing.Point(286, 13)
        Me.chkLocation.Name = "chkLocation"
        Me.chkLocation.Size = New System.Drawing.Size(103, 17)
        Me.chkLocation.TabIndex = 2
        Me.chkLocation.Text = "Location Control"
        Me.chkLocation.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColBatchNo, Me.ColExpiryDate, Me.ColLocation, Me.ColQty, Me.ColUnit, Me.ColPrimaryLocation})
        Me.dgvItem.Location = New System.Drawing.Point(15, 340)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.Size = New System.Drawing.Size(498, 164)
        Me.dgvItem.TabIndex = 8
        '
        'ColBatchNo
        '
        Me.ColBatchNo.HeaderText = "Batch Number"
        Me.ColBatchNo.Name = "ColBatchNo"
        Me.ColBatchNo.Visible = False
        '
        'ColExpiryDate
        '
        Me.ColExpiryDate.HeaderText = "Expiry Date"
        Me.ColExpiryDate.Name = "ColExpiryDate"
        Me.ColExpiryDate.Visible = False
        Me.ColExpiryDate.Width = 86
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.Visible = False
        Me.ColLocation.Width = 73
        '
        'ColQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle1
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
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 110
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.chkED_Ori)
        Me.GroupBox1.Controls.Add(Me.chk_Location_Ori)
        Me.GroupBox1.Controls.Add(Me.chkBatch_Ori)
        Me.GroupBox1.Enabled = False
        Me.GroupBox1.Location = New System.Drawing.Point(105, 90)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(399, 42)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Original State"
        '
        'chkED_Ori
        '
        Me.chkED_Ori.AutoSize = True
        Me.chkED_Ori.Enabled = False
        Me.chkED_Ori.Location = New System.Drawing.Point(148, 13)
        Me.chkED_Ori.Name = "chkED_Ori"
        Me.chkED_Ori.Size = New System.Drawing.Size(130, 17)
        Me.chkED_Ori.TabIndex = 1
        Me.chkED_Ori.Text = "Expiry Date Controlled"
        Me.chkED_Ori.UseVisualStyleBackColor = True
        '
        'chk_Location_Ori
        '
        Me.chk_Location_Ori.AutoSize = True
        Me.chk_Location_Ori.Location = New System.Drawing.Point(286, 13)
        Me.chk_Location_Ori.Name = "chk_Location_Ori"
        Me.chk_Location_Ori.Size = New System.Drawing.Size(103, 17)
        Me.chk_Location_Ori.TabIndex = 2
        Me.chk_Location_Ori.Text = "Location Control"
        Me.chk_Location_Ori.UseVisualStyleBackColor = True
        '
        'chkBatch_Ori
        '
        Me.chkBatch_Ori.AutoSize = True
        Me.chkBatch_Ori.Location = New System.Drawing.Point(10, 13)
        Me.chkBatch_Ori.Name = "chkBatch_Ori"
        Me.chkBatch_Ori.Size = New System.Drawing.Size(130, 17)
        Me.chkBatch_Ori.TabIndex = 0
        Me.chkBatch_Ori.Text = "Batch Number Control"
        Me.chkBatch_Ori.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.chkBatch)
        Me.GroupBox2.Controls.Add(Me.chkLocation)
        Me.GroupBox2.Controls.Add(Me.chkExpiryDate)
        Me.GroupBox2.Location = New System.Drawing.Point(105, 138)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(399, 42)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Modified State"
        '
        'FrmTransaksiModifyControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(525, 545)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.gboInput)
        Me.Controls.Add(Me.LblUnit)
        Me.Controls.Add(Me.TxtQty)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtDesc)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.btnFindCodeNo)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.TxtCodeNo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.Name = "FrmTransaksiModifyControl"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Modify Product Control"
        Me.gboInput.ResumeLayout(False)
        Me.gboInput.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtQty As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblUnit As System.Windows.Forms.Label
    Friend WithEvents gboInput As System.Windows.Forms.GroupBox
    Friend WithEvents chkBatch As System.Windows.Forms.CheckBox
    Friend WithEvents chkExpiryDate As System.Windows.Forms.CheckBox
    Friend WithEvents chkLocation As System.Windows.Forms.CheckBox
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnFindLocation As System.Windows.Forms.Button
    Friend WithEvents TxtLocationCode As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpED As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtBatch As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnInput As System.Windows.Forms.Button
    Friend WithEvents LblUnit2 As System.Windows.Forms.Label
    Friend WithEvents TxtQtyPerBatch As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkED_Ori As System.Windows.Forms.CheckBox
    Friend WithEvents chk_Location_Ori As System.Windows.Forms.CheckBox
    Friend WithEvents chkBatch_Ori As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ColBatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
