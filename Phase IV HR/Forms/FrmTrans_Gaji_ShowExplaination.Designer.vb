<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_ShowExplaination
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
        Me.TxtFormula = New System.Windows.Forms.TextBox
        Me.BtnClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TxtFormula
        '
        Me.TxtFormula.BackColor = System.Drawing.Color.White
        Me.TxtFormula.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtFormula.Location = New System.Drawing.Point(0, 0)
        Me.TxtFormula.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtFormula.MaxLength = 250
        Me.TxtFormula.Multiline = True
        Me.TxtFormula.Name = "TxtFormula"
        Me.TxtFormula.ReadOnly = True
        Me.TxtFormula.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TxtFormula.Size = New System.Drawing.Size(1109, 545)
        Me.TxtFormula.TabIndex = 1
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(949, 48)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(100, 28)
        Me.BtnClose.TabIndex = 7
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'FrmTrans_Gaji_ShowExplaination
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(1109, 545)
        Me.Controls.Add(Me.TxtFormula)
        Me.Controls.Add(Me.BtnClose)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1117, 582)
        Me.Name = "FrmTrans_Gaji_ShowExplaination"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Show Explanation"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtFormula As System.Windows.Forms.TextBox
    Friend WithEvents BtnClose As System.Windows.Forms.Button
End Class
