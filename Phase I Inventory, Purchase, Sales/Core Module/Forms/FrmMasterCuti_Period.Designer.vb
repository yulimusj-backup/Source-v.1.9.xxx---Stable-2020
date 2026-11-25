<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterCuti_Period
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
        Me.DtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.BtnOK = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Period"
        '
        'DtpPeriod
        '
        Me.DtpPeriod.CustomFormat = "MMMM"
        Me.DtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpPeriod.Location = New System.Drawing.Point(58, 11)
        Me.DtpPeriod.Name = "DtpPeriod"
        Me.DtpPeriod.ShowUpDown = True
        Me.DtpPeriod.Size = New System.Drawing.Size(85, 20)
        Me.DtpPeriod.TabIndex = 1
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(149, 10)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(85, 23)
        Me.BtnOK.TabIndex = 2
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'FrmMasterCuti_Period
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(247, 49)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.DtpPeriod)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FrmMasterCuti_Period"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Period"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents BtnOK As System.Windows.Forms.Button
End Class
