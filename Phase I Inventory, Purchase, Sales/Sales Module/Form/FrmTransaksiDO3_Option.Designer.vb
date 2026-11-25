<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiDO3_Option
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
        Me.btn_simpan = New System.Windows.Forms.Button
        Me.btn_cancel = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.cbo_DeliveryDate = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'btn_simpan
        '
        Me.btn_simpan.Location = New System.Drawing.Point(107, 64)
        Me.btn_simpan.Name = "btn_simpan"
        Me.btn_simpan.Size = New System.Drawing.Size(75, 23)
        Me.btn_simpan.TabIndex = 6
        Me.btn_simpan.Text = "OK"
        Me.btn_simpan.UseVisualStyleBackColor = True
        '
        'btn_cancel
        '
        Me.btn_cancel.Location = New System.Drawing.Point(188, 64)
        Me.btn_cancel.Name = "btn_cancel"
        Me.btn_cancel.Size = New System.Drawing.Size(75, 23)
        Me.btn_cancel.TabIndex = 7
        Me.btn_cancel.Text = "Cancel"
        Me.btn_cancel.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 28)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(71, 13)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "Delivery Date"
        '
        'cbo_DeliveryDate
        '
        Me.cbo_DeliveryDate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbo_DeliveryDate.FormattingEnabled = True
        Me.cbo_DeliveryDate.Location = New System.Drawing.Point(107, 25)
        Me.cbo_DeliveryDate.Name = "cbo_DeliveryDate"
        Me.cbo_DeliveryDate.Size = New System.Drawing.Size(156, 21)
        Me.cbo_DeliveryDate.TabIndex = 5
        '
        'FrmTransaksiDO3_Option
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(270, 95)
        Me.Controls.Add(Me.cbo_DeliveryDate)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.btn_cancel)
        Me.Controls.Add(Me.btn_simpan)
        Me.MaximumSize = New System.Drawing.Size(286, 133)
        Me.MinimumSize = New System.Drawing.Size(286, 133)
        Me.Name = "FrmTransaksiDO3_Option"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Option Delivery Date"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_simpan As System.Windows.Forms.Button
    Friend WithEvents btn_cancel As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cbo_DeliveryDate As System.Windows.Forms.ComboBox
End Class
