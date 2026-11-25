<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSO_SetVisible
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
        Me.chkTolerance = New System.Windows.Forms.CheckBox
        Me.chkUnitsDefault = New System.Windows.Forms.CheckBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'chkTolerance
        '
        Me.chkTolerance.AutoSize = True
        Me.chkTolerance.Location = New System.Drawing.Point(11, 30)
        Me.chkTolerance.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTolerance.Name = "chkTolerance"
        Me.chkTolerance.Size = New System.Drawing.Size(74, 17)
        Me.chkTolerance.TabIndex = 8
        Me.chkTolerance.Text = "Tolerance"
        Me.chkTolerance.UseVisualStyleBackColor = True
        '
        'chkUnitsDefault
        '
        Me.chkUnitsDefault.AutoSize = True
        Me.chkUnitsDefault.Location = New System.Drawing.Point(11, 9)
        Me.chkUnitsDefault.Margin = New System.Windows.Forms.Padding(2)
        Me.chkUnitsDefault.Name = "chkUnitsDefault"
        Me.chkUnitsDefault.Size = New System.Drawing.Size(110, 17)
        Me.chkUnitsDefault.TabIndex = 7
        Me.chkUnitsDefault.Text = "Qty && Unit Default"
        Me.chkUnitsDefault.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(33, 51)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 21)
        Me.btnOK.TabIndex = 6
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(93, 51)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 21)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'FrmTransaksiSO_SetVisible
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(183, 81)
        Me.Controls.Add(Me.chkTolerance)
        Me.Controls.Add(Me.chkUnitsDefault)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FrmTransaksiSO_SetVisible"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Columns Setting"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents chkUnitsDefault As System.Windows.Forms.CheckBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
