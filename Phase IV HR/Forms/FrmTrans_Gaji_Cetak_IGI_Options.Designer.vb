<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_Cetak_IGI_Options
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
        Me.btnExit = New System.Windows.Forms.Button
        Me.BtnPreview = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.CboFormat = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(276, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(86, 20)
        Me.btnExit.TabIndex = 280
        Me.btnExit.TabStop = False
        Me.btnExit.Text = "&Close"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'BtnPreview
        '
        Me.BtnPreview.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnPreview.Location = New System.Drawing.Point(184, 11)
        Me.BtnPreview.Name = "BtnPreview"
        Me.BtnPreview.Size = New System.Drawing.Size(86, 20)
        Me.BtnPreview.TabIndex = 281
        Me.BtnPreview.TabStop = False
        Me.BtnPreview.Text = "&Preview"
        Me.BtnPreview.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 282
        Me.Label1.Text = "Format"
        '
        'CboFormat
        '
        Me.CboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormat.FormattingEnabled = True
        Me.CboFormat.Items.AddRange(New Object() {"A5", "Continuous"})
        Me.CboFormat.Location = New System.Drawing.Point(57, 11)
        Me.CboFormat.Name = "CboFormat"
        Me.CboFormat.Size = New System.Drawing.Size(121, 21)
        Me.CboFormat.TabIndex = 283
        '
        'FrmTrans_Gaji_Cetak_IGI_Options
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(374, 47)
        Me.ControlBox = False
        Me.Controls.Add(Me.CboFormat)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BtnPreview)
        Me.Controls.Add(Me.btnExit)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(390, 85)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(390, 85)
        Me.Name = "FrmTrans_Gaji_Cetak_IGI_Options"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Printed Options"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents BtnPreview As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CboFormat As System.Windows.Forms.ComboBox
End Class
