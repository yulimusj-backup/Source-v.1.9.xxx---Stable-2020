<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiLicensing2
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
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpExpiry = New System.Windows.Forms.DateTimePicker
        Me.TxtNewLicenseKey = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.TxtLicenseKey = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtProductKey = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtpNewExpiry = New System.Windows.Forms.DateTimePicker
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(22, 84)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 17)
        Me.Label4.TabIndex = 31
        Me.Label4.Text = "Expiry"
        '
        'dtpExpiry
        '
        Me.dtpExpiry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpExpiry.Enabled = False
        Me.dtpExpiry.Location = New System.Drawing.Point(150, 77)
        Me.dtpExpiry.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpExpiry.Name = "dtpExpiry"
        Me.dtpExpiry.Size = New System.Drawing.Size(319, 22)
        Me.dtpExpiry.TabIndex = 6
        Me.dtpExpiry.TabStop = False
        '
        'TxtNewLicenseKey
        '
        Me.TxtNewLicenseKey.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNewLicenseKey.Location = New System.Drawing.Point(150, 109)
        Me.TxtNewLicenseKey.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtNewLicenseKey.Name = "TxtNewLicenseKey"
        Me.TxtNewLicenseKey.Size = New System.Drawing.Size(319, 22)
        Me.TxtNewLicenseKey.TabIndex = 1
        Me.TxtNewLicenseKey.TabStop = False
        Me.TxtNewLicenseKey.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(22, 113)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(116, 17)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "New License Key"
        Me.Label3.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(258, 141)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(100, 28)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Close"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'TxtLicenseKey
        '
        Me.TxtLicenseKey.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtLicenseKey.Location = New System.Drawing.Point(150, 45)
        Me.TxtLicenseKey.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtLicenseKey.Name = "TxtLicenseKey"
        Me.TxtLicenseKey.ReadOnly = True
        Me.TxtLicenseKey.Size = New System.Drawing.Size(319, 22)
        Me.TxtLicenseKey.TabIndex = 5
        Me.TxtLicenseKey.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(22, 49)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 17)
        Me.Label2.TabIndex = 28
        Me.Label2.Text = "License Key"
        '
        'txtProductKey
        '
        Me.txtProductKey.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProductKey.Location = New System.Drawing.Point(150, 13)
        Me.txtProductKey.Margin = New System.Windows.Forms.Padding(4)
        Me.txtProductKey.Name = "txtProductKey"
        Me.txtProductKey.ReadOnly = True
        Me.txtProductKey.Size = New System.Drawing.Size(319, 22)
        Me.txtProductKey.TabIndex = 4
        Me.txtProductKey.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(22, 17)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 17)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "Product Key"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(150, 141)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(100, 28)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(22, 114)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 17)
        Me.Label5.TabIndex = 33
        Me.Label5.Text = "New Expiry"
        '
        'dtpNewExpiry
        '
        Me.dtpNewExpiry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpNewExpiry.Location = New System.Drawing.Point(150, 107)
        Me.dtpNewExpiry.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpNewExpiry.Name = "dtpNewExpiry"
        Me.dtpNewExpiry.Size = New System.Drawing.Size(319, 22)
        Me.dtpNewExpiry.TabIndex = 0
        '
        'FrmTransaksiLicensing2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(491, 183)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.dtpExpiry)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.TxtLicenseKey)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtProductKey)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dtpNewExpiry)
        Me.Controls.Add(Me.TxtNewLicenseKey)
        Me.MinimumSize = New System.Drawing.Size(499, 220)
        Me.Name = "FrmTransaksiLicensing2"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set License"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpExpiry As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtNewLicenseKey As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents TxtLicenseKey As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtProductKey As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtpNewExpiry As System.Windows.Forms.DateTimePicker
End Class
