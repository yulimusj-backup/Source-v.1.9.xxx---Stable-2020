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
        Me.TxtProductCode = New System.Windows.Forms.TextBox
        Me.dtpOldExpiry = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboBatchNo = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.DtpNewExpiry = New System.Windows.Forms.DateTimePicker
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnFindCodeNo = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.txtPrimaryCode = New System.Windows.Forms.TextBox
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
        'TxtProductCode
        '
        Me.TxtProductCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtProductCode.Location = New System.Drawing.Point(105, 12)
        Me.TxtProductCode.Name = "TxtProductCode"
        Me.TxtProductCode.Size = New System.Drawing.Size(247, 20)
        Me.TxtProductCode.TabIndex = 0
        '
        'dtpOldExpiry
        '
        Me.dtpOldExpiry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpOldExpiry.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpOldExpiry.Enabled = False
        Me.dtpOldExpiry.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOldExpiry.Location = New System.Drawing.Point(105, 92)
        Me.dtpOldExpiry.Name = "dtpOldExpiry"
        Me.dtpOldExpiry.Size = New System.Drawing.Size(305, 20)
        Me.dtpOldExpiry.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 68)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Batch No"
        '
        'cboBatchNo
        '
        Me.cboBatchNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboBatchNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBatchNo.FormattingEnabled = True
        Me.cboBatchNo.Location = New System.Drawing.Point(105, 65)
        Me.cboBatchNo.Name = "cboBatchNo"
        Me.cboBatchNo.Size = New System.Drawing.Size(305, 21)
        Me.cboBatchNo.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 98)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Old Expiry Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 124)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(86, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "New Expiry Date"
        '
        'DtpNewExpiry
        '
        Me.DtpNewExpiry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpNewExpiry.CustomFormat = "ddd, dd MMM yyyy"
        Me.DtpNewExpiry.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpNewExpiry.Location = New System.Drawing.Point(105, 118)
        Me.DtpNewExpiry.Name = "DtpNewExpiry"
        Me.DtpNewExpiry.Size = New System.Drawing.Size(305, 20)
        Me.DtpNewExpiry.TabIndex = 5
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(335, 145)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "&Close"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(254, 145)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 6
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
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 41)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Description"
        '
        'txtDesc
        '
        Me.txtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDesc.Location = New System.Drawing.Point(105, 38)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(303, 20)
        Me.txtDesc.TabIndex = 2
        Me.txtDesc.TabStop = False
        '
        'txtPrimaryCode
        '
        Me.txtPrimaryCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryCode.Location = New System.Drawing.Point(67, 12)
        Me.txtPrimaryCode.Name = "txtPrimaryCode"
        Me.txtPrimaryCode.Size = New System.Drawing.Size(37, 20)
        Me.txtPrimaryCode.TabIndex = 13
        Me.txtPrimaryCode.Text = "-1"
        Me.txtPrimaryCode.Visible = False
        '
        'FrmTransaksiExpiryDateUpdate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(420, 180)
        Me.Controls.Add(Me.txtDesc)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.btnFindCodeNo)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.DtpNewExpiry)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboBatchNo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpOldExpiry)
        Me.Controls.Add(Me.TxtProductCode)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtPrimaryCode)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(436, 191)
        Me.Name = "FrmTransaksiExpiryDateUpdate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Expiry Date Change"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtProductCode As System.Windows.Forms.TextBox
    Friend WithEvents dtpOldExpiry As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboBatchNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DtpNewExpiry As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCode As System.Windows.Forms.TextBox
End Class
