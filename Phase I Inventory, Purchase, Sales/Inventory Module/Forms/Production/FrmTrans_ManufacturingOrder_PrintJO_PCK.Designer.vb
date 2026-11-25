<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_PrintJO_PCK
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpPrintDate = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtCustomerCode = New System.Windows.Forms.TextBox
        Me.btnFindCustomer = New System.Windows.Forms.Button
        Me.txtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.txtCustomerDesc = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboRevision = New System.Windows.Forms.ComboBox
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.txtFU = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtOuterPack = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtShrinkPack = New System.Windows.Forms.TextBox
        Me.txtPoSo = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.dtpPCKTo = New System.Windows.Forms.DateTimePicker
        Me.Label14 = New System.Windows.Forms.Label
        Me.dtpPCKFrom = New System.Windows.Forms.DateTimePicker
        Me.chkPCK = New System.Windows.Forms.CheckBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.txtRemarks1 = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Print Date:"
        '
        'dtpPrintDate
        '
        Me.dtpPrintDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpPrintDate.CustomFormat = "dd MMM yyyy"
        Me.dtpPrintDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPrintDate.Location = New System.Drawing.Point(66, 7)
        Me.dtpPrintDate.Name = "dtpPrintDate"
        Me.dtpPrintDate.Size = New System.Drawing.Size(183, 20)
        Me.dtpPrintDate.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Customer:"
        '
        'txtCustomerCode
        '
        Me.txtCustomerCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerCode.Location = New System.Drawing.Point(66, 33)
        Me.txtCustomerCode.MaxLength = 150
        Me.txtCustomerCode.Name = "txtCustomerCode"
        Me.txtCustomerCode.Size = New System.Drawing.Size(140, 20)
        Me.txtCustomerCode.TabIndex = 1
        '
        'btnFindCustomer
        '
        Me.btnFindCustomer.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCustomer.Location = New System.Drawing.Point(212, 33)
        Me.btnFindCustomer.Name = "btnFindCustomer"
        Me.btnFindCustomer.Size = New System.Drawing.Size(37, 21)
        Me.btnFindCustomer.TabIndex = 2
        Me.btnFindCustomer.TabStop = False
        Me.btnFindCustomer.Text = "..."
        Me.btnFindCustomer.UseVisualStyleBackColor = True
        '
        'txtPrimaryCustomer
        '
        Me.txtPrimaryCustomer.Location = New System.Drawing.Point(66, 33)
        Me.txtPrimaryCustomer.Name = "txtPrimaryCustomer"
        Me.txtPrimaryCustomer.Size = New System.Drawing.Size(31, 20)
        Me.txtPrimaryCustomer.TabIndex = 5
        Me.txtPrimaryCustomer.Text = "-1"
        Me.txtPrimaryCustomer.Visible = False
        '
        'txtCustomerDesc
        '
        Me.txtCustomerDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustomerDesc.Location = New System.Drawing.Point(66, 59)
        Me.txtCustomerDesc.Name = "txtCustomerDesc"
        Me.txtCustomerDesc.ReadOnly = True
        Me.txtCustomerDesc.Size = New System.Drawing.Size(183, 20)
        Me.txtCustomerDesc.TabIndex = 6
        Me.txtCustomerDesc.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 141)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Revision:"
        '
        'cboRevision
        '
        Me.cboRevision.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboRevision.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRevision.FormattingEnabled = True
        Me.cboRevision.Items.AddRange(New Object() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"})
        Me.cboRevision.Location = New System.Drawing.Point(66, 137)
        Me.cboRevision.Name = "cboRevision"
        Me.cboRevision.Size = New System.Drawing.Size(183, 21)
        Me.cboRevision.TabIndex = 6
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(66, 7)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(31, 20)
        Me.txtPrimary.TabIndex = 9
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'txtFU
        '
        Me.txtFU.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFU.Location = New System.Drawing.Point(59, 7)
        Me.txtFU.MaxLength = 150
        Me.txtFU.Name = "txtFU"
        Me.txtFU.Size = New System.Drawing.Size(193, 20)
        Me.txtFU.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 11)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(24, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "FU:"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(1, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtOuterPack)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtShrinkPack)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCustomerDesc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCustomerCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboRevision)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpPrintDate)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPoSo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label17)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dtpPCKTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label14)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dtpPCKFrom)
        Me.SplitContainer1.Panel2.Controls.Add(Me.chkPCK)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtFU)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Size = New System.Drawing.Size(524, 164)
        Me.SplitContainer1.SplitterDistance = 258
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(217, 115)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(32, 13)
        Me.Label13.TabIndex = 20
        Me.Label13.Text = "Pack"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(217, 88)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(32, 13)
        Me.Label12.TabIndex = 19
        Me.Label12.Text = "Pack"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 115)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(54, 13)
        Me.Label11.TabIndex = 18
        Me.Label11.Text = "1 Outer = "
        '
        'txtOuterPack
        '
        Me.txtOuterPack.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOuterPack.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtOuterPack.Location = New System.Drawing.Point(66, 111)
        Me.txtOuterPack.MaxLength = 150
        Me.txtOuterPack.Name = "txtOuterPack"
        Me.txtOuterPack.Size = New System.Drawing.Size(145, 20)
        Me.txtOuterPack.TabIndex = 5
        Me.txtOuterPack.Text = "1.0000"
        Me.txtOuterPack.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 89)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(58, 13)
        Me.Label10.TabIndex = 16
        Me.Label10.Text = "1 Shrink = "
        '
        'txtShrinkPack
        '
        Me.txtShrinkPack.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtShrinkPack.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtShrinkPack.Location = New System.Drawing.Point(66, 85)
        Me.txtShrinkPack.MaxLength = 150
        Me.txtShrinkPack.Name = "txtShrinkPack"
        Me.txtShrinkPack.Size = New System.Drawing.Size(145, 20)
        Me.txtShrinkPack.TabIndex = 4
        Me.txtShrinkPack.Text = "1.0000"
        Me.txtShrinkPack.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtPoSo
        '
        Me.txtPoSo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPoSo.Location = New System.Drawing.Point(59, 33)
        Me.txtPoSo.MaxLength = 150
        Me.txtPoSo.Name = "txtPoSo"
        Me.txtPoSo.Size = New System.Drawing.Size(193, 20)
        Me.txtPoSo.TabIndex = 2
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(3, 37)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(45, 13)
        Me.Label17.TabIndex = 31
        Me.Label17.Text = "PO/SO:"
        '
        'dtpPCKTo
        '
        Me.dtpPCKTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpPCKTo.CustomFormat = "dd MMM yyyy"
        Me.dtpPCKTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPCKTo.Location = New System.Drawing.Point(165, 59)
        Me.dtpPCKTo.Name = "dtpPCKTo"
        Me.dtpPCKTo.Size = New System.Drawing.Size(87, 20)
        Me.dtpPCKTo.TabIndex = 5
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(152, 63)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(10, 13)
        Me.Label14.TabIndex = 21
        Me.Label14.Text = "-"
        '
        'dtpPCKFrom
        '
        Me.dtpPCKFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpPCKFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpPCKFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPCKFrom.Location = New System.Drawing.Point(59, 59)
        Me.dtpPCKFrom.Name = "dtpPCKFrom"
        Me.dtpPCKFrom.Size = New System.Drawing.Size(87, 20)
        Me.dtpPCKFrom.TabIndex = 4
        '
        'chkPCK
        '
        Me.chkPCK.AutoSize = True
        Me.chkPCK.Checked = True
        Me.chkPCK.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPCK.Location = New System.Drawing.Point(6, 61)
        Me.chkPCK.Name = "chkPCK"
        Me.chkPCK.Size = New System.Drawing.Size(47, 17)
        Me.chkPCK.TabIndex = 3
        Me.chkPCK.Text = "PCK"
        Me.chkPCK.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColJONo, Me.ColCodeNo, Me.ColDesc, Me.ColMO, Me.ColIdentifierJO})
        Me.dgvItem.Location = New System.Drawing.Point(4, 172)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.Size = New System.Drawing.Size(518, 135)
        Me.dgvItem.TabIndex = 1
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColJONo
        '
        Me.ColJONo.Frozen = True
        Me.ColJONo.HeaderText = "Job Order Number"
        Me.ColJONo.Name = "ColJONo"
        Me.ColJONo.Width = 118
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
        'ColMO
        '
        Me.ColMO.HeaderText = "MO No"
        Me.ColMO.Name = "ColMO"
        Me.ColMO.ReadOnly = True
        Me.ColMO.Width = 66
        '
        'ColIdentifierJO
        '
        Me.ColIdentifierJO.HeaderText = "Identifier JO"
        Me.ColIdentifierJO.Name = "ColIdentifierJO"
        Me.ColIdentifierJO.ReadOnly = True
        Me.ColIdentifierJO.Visible = False
        Me.ColIdentifierJO.Width = 88
        '
        'txtRemarks1
        '
        Me.txtRemarks1.Location = New System.Drawing.Point(4, 326)
        Me.txtRemarks1.MaxLength = 150
        Me.txtRemarks1.Multiline = True
        Me.txtRemarks1.Name = "txtRemarks1"
        Me.txtRemarks1.Size = New System.Drawing.Size(518, 66)
        Me.txtRemarks1.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(4, 310)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(49, 13)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "Remarks"
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(390, 398)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(63, 23)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(459, 398)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(63, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(6, 401)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(52, 20)
        Me.TxtReadOnly.TabIndex = 25
        Me.TxtReadOnly.Visible = False
        '
        'BackgroundWorker1
        '
        '
        'FrmTrans_ManufacturingOrder_PrintJO_PCK
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(527, 428)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtRemarks1)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(543, 466)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(543, 466)
        Me.Name = "FrmTrans_ManufacturingOrder_PrintJO_PCK"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print JO Packaging"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpPrintDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCustomer As System.Windows.Forms.Button
    Friend WithEvents txtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboRevision As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents txtFU As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents txtRemarks1 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtOuterPack As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtShrinkPack As System.Windows.Forms.TextBox
    Friend WithEvents dtpPCKTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents dtpPCKFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkPCK As System.Windows.Forms.CheckBox
    Friend WithEvents txtPoSo As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
