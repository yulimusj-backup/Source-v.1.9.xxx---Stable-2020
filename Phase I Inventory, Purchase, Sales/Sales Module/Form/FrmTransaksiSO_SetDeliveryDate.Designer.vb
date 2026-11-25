<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSO_SetDeliveryDate
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
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.BtnOK = New System.Windows.Forms.Button
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.RboSetDeliveryDate = New System.Windows.Forms.RadioButton
        Me.RboNoDeliveryDate = New System.Windows.Forms.RadioButton
        Me.SuspendLayout()
        '
        'DtpDate
        '
        Me.DtpDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(168, 13)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.Size = New System.Drawing.Size(127, 20)
        Me.DtpDate.TabIndex = 1
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(79, 71)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(75, 23)
        Me.BtnOK.TabIndex = 2
        Me.BtnOK.Text = "&OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(160, 71)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(75, 23)
        Me.BtnCancel.TabIndex = 3
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'RboSetDeliveryDate
        '
        Me.RboSetDeliveryDate.AutoSize = True
        Me.RboSetDeliveryDate.Checked = True
        Me.RboSetDeliveryDate.Location = New System.Drawing.Point(12, 14)
        Me.RboSetDeliveryDate.Name = "RboSetDeliveryDate"
        Me.RboSetDeliveryDate.Size = New System.Drawing.Size(126, 17)
        Me.RboSetDeliveryDate.TabIndex = 4
        Me.RboSetDeliveryDate.TabStop = True
        Me.RboSetDeliveryDate.Text = "Set Delivery Date  to:"
        Me.RboSetDeliveryDate.UseVisualStyleBackColor = True
        '
        'RboNoDeliveryDate
        '
        Me.RboNoDeliveryDate.AutoSize = True
        Me.RboNoDeliveryDate.Location = New System.Drawing.Point(12, 37)
        Me.RboNoDeliveryDate.Name = "RboNoDeliveryDate"
        Me.RboNoDeliveryDate.Size = New System.Drawing.Size(142, 17)
        Me.RboNoDeliveryDate.TabIndex = 5
        Me.RboNoDeliveryDate.Text = "No Delivery Date Control"
        Me.RboNoDeliveryDate.UseVisualStyleBackColor = True
        '
        'FrmTransaksiSO_SetDeliveryDate
        '
        Me.AcceptButton = Me.BtnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(307, 106)
        Me.Controls.Add(Me.RboNoDeliveryDate)
        Me.Controls.Add(Me.RboSetDeliveryDate)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.DtpDate)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(323, 144)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(16, 133)
        Me.Name = "FrmTransaksiSO_SetDeliveryDate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set Delivery Date"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents RboSetDeliveryDate As System.Windows.Forms.RadioButton
    Friend WithEvents RboNoDeliveryDate As System.Windows.Forms.RadioButton
End Class
