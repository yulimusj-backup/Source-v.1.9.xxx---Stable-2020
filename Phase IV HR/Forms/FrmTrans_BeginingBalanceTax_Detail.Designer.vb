<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_BeginingBalanceTax_Detail
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
        Me.ChkShowNIK = New System.Windows.Forms.CheckBox
        Me.ChkShowDepartment = New System.Windows.Forms.CheckBox
        Me.ChkShowPosition = New System.Windows.Forms.CheckBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ChkShowNIK
        '
        Me.ChkShowNIK.AutoSize = True
        Me.ChkShowNIK.Location = New System.Drawing.Point(12, 12)
        Me.ChkShowNIK.Name = "ChkShowNIK"
        Me.ChkShowNIK.Size = New System.Drawing.Size(74, 17)
        Me.ChkShowNIK.TabIndex = 0
        Me.ChkShowNIK.Text = "Show NIK"
        Me.ChkShowNIK.UseVisualStyleBackColor = True
        '
        'ChkShowDepartment
        '
        Me.ChkShowDepartment.AutoSize = True
        Me.ChkShowDepartment.Location = New System.Drawing.Point(12, 35)
        Me.ChkShowDepartment.Name = "ChkShowDepartment"
        Me.ChkShowDepartment.Size = New System.Drawing.Size(111, 17)
        Me.ChkShowDepartment.TabIndex = 1
        Me.ChkShowDepartment.Text = "Show Department"
        Me.ChkShowDepartment.UseVisualStyleBackColor = True
        '
        'ChkShowPosition
        '
        Me.ChkShowPosition.AutoSize = True
        Me.ChkShowPosition.Location = New System.Drawing.Point(12, 58)
        Me.ChkShowPosition.Name = "ChkShowPosition"
        Me.ChkShowPosition.Size = New System.Drawing.Size(93, 17)
        Me.ChkShowPosition.TabIndex = 2
        Me.ChkShowPosition.Text = "Show Position"
        Me.ChkShowPosition.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(92, 29)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 95
        Me.btnExit.Text = "&Close"
        Me.btnExit.UseVisualStyleBackColor = True
        Me.btnExit.Visible = False
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(92, 3)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 94
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        Me.btnOK.Visible = False
        '
        'FrmTrans_BeginingBalanceTax_Detail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(178, 80)
        Me.Controls.Add(Me.ChkShowDepartment)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.ChkShowPosition)
        Me.Controls.Add(Me.ChkShowNIK)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(194, 118)
        Me.Name = "FrmTrans_BeginingBalanceTax_Detail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Setting"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ChkShowNIK As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowDepartment As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowPosition As System.Windows.Forms.CheckBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
