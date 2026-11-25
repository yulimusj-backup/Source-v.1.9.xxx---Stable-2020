<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiExpiryDateUpdate
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
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.dtpOld = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboBatchNo = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpNew = New System.Windows.Forms.DateTimePicker
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnFindCodeNo = New System.Windows.Forms.Button
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.TxtBatchNew = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.chkExpiry = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Code No"
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(118, 12)
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(234, 20)
        Me.TxtCodeNo.TabIndex = 0
        '
        'dtpOld
        '
        Me.dtpOld.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpOld.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpOld.Enabled = False
        Me.dtpOld.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOld.Location = New System.Drawing.Point(118, 91)
        Me.dtpOld.Name = "dtpOld"
        Me.dtpOld.Size = New System.Drawing.Size(290, 20)
        Me.dtpOld.TabIndex = 5
        Me.dtpOld.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Batch Number"
        '
        'cboBatchNo
        '
        Me.cboBatchNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboBatchNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBatchNo.FormattingEnabled = True
        Me.cboBatchNo.Location = New System.Drawing.Point(118, 38)
        Me.cboBatchNo.Name = "cboBatchNo"
        Me.cboBatchNo.Size = New System.Drawing.Size(290, 21)
        Me.cboBatchNo.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 97)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Old Expiry Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 123)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(86, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "New Expiry Date"
        '
        'dtpNew
        '
        Me.dtpNew.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpNew.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpNew.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpNew.Location = New System.Drawing.Point(118, 117)
        Me.dtpNew.Name = "dtpNew"
        Me.dtpNew.Size = New System.Drawing.Size(290, 20)
        Me.dtpNew.TabIndex = 6
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(335, 143)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "&Close"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(254, 143)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 7
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnFindCodeNo
        '
        Me.btnFindCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCodeNo.Location = New System.Drawing.Point(356, 12)
        Me.btnFindCodeNo.Name = "btnFindCodeNo"
        Me.btnFindCodeNo.Size = New System.Drawing.Size(52, 20)
        Me.btnFindCodeNo.TabIndex = 1
        Me.btnFindCodeNo.TabStop = False
        Me.btnFindCodeNo.Text = "...."
        Me.btnFindCodeNo.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(87, 27)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(37, 20)
        Me.TxtPrimaryCodeNo.TabIndex = 13
        Me.TxtPrimaryCodeNo.Text = "-1"
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'TxtBatchNew
        '
        Me.TxtBatchNew.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBatchNew.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBatchNew.Location = New System.Drawing.Point(118, 65)
        Me.TxtBatchNew.Name = "TxtBatchNew"
        Me.TxtBatchNew.Size = New System.Drawing.Size(290, 20)
        Me.TxtBatchNew.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 68)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "New Batch Number"
        '
        'chkExpiry
        '
        Me.chkExpiry.AutoSize = True
        Me.chkExpiry.Location = New System.Drawing.Point(31, 11)
        Me.chkExpiry.Name = "chkExpiry"
        Me.chkExpiry.Size = New System.Drawing.Size(81, 17)
        Me.chkExpiry.TabIndex = 15
        Me.chkExpiry.TabStop = False
        Me.chkExpiry.Text = "CheckBox1"
        Me.chkExpiry.UseVisualStyleBackColor = True
        Me.chkExpiry.Visible = False
        '
        'FrmTransaksiExpiryDateUpdate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(420, 178)
        Me.Controls.Add(Me.chkExpiry)
        Me.Controls.Add(Me.TxtBatchNew)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnFindCodeNo)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.dtpNew)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboBatchNo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpOld)
        Me.Controls.Add(Me.TxtCodeNo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(436, 191)
        Me.Name = "FrmTransaksiExpiryDateUpdate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Batch Number & Expiry Date Change"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents dtpOld As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboBatchNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpNew As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtBatchNew As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents chkExpiry As System.Windows.Forms.CheckBox
End Class
