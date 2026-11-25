<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPL_Budget_Opt
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
        Me.ChkORedColor = New System.Windows.Forms.CheckBox
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ChkORedColor
        '
        Me.ChkORedColor.AutoSize = True
        Me.ChkORedColor.Location = New System.Drawing.Point(12, 12)
        Me.ChkORedColor.Name = "ChkORedColor"
        Me.ChkORedColor.Size = New System.Drawing.Size(160, 17)
        Me.ChkORedColor.TabIndex = 0
        Me.ChkORedColor.Text = "Red color text if over budget"
        Me.ChkORedColor.UseVisualStyleBackColor = True
        '
        'BtnPrint
        '
        Me.BtnPrint.Location = New System.Drawing.Point(178, 8)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(44, 22)
        Me.BtnPrint.TabIndex = 1
        Me.BtnPrint.Text = "Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'FrmTrans_SPL_Budget_Opt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(235, 42)
        Me.Controls.Add(Me.BtnPrint)
        Me.Controls.Add(Me.ChkORedColor)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(251, 80)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(251, 80)
        Me.Name = "FrmTrans_SPL_Budget_Opt"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Option"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ChkORedColor As System.Windows.Forms.CheckBox
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
End Class
