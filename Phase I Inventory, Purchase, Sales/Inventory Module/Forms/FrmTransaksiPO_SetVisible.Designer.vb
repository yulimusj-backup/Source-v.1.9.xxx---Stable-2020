<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPO_SetVisible
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
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.chkQtyDefault = New System.Windows.Forms.CheckBox
        Me.chkUnitCodeDefault = New System.Windows.Forms.CheckBox
        Me.chkToleranceDown = New System.Windows.Forms.CheckBox
        Me.chkToleranceUp = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(174, 60)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 21)
        Me.btnCancel.TabIndex = 0
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(114, 60)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 21)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'chkQtyDefault
        '
        Me.chkQtyDefault.AutoSize = True
        Me.chkQtyDefault.Location = New System.Drawing.Point(7, 11)
        Me.chkQtyDefault.Margin = New System.Windows.Forms.Padding(2)
        Me.chkQtyDefault.Name = "chkQtyDefault"
        Me.chkQtyDefault.Size = New System.Drawing.Size(79, 17)
        Me.chkQtyDefault.TabIndex = 2
        Me.chkQtyDefault.Text = "Qty Default"
        Me.chkQtyDefault.UseVisualStyleBackColor = True
        '
        'chkUnitCodeDefault
        '
        Me.chkUnitCodeDefault.AutoSize = True
        Me.chkUnitCodeDefault.Location = New System.Drawing.Point(7, 33)
        Me.chkUnitCodeDefault.Margin = New System.Windows.Forms.Padding(2)
        Me.chkUnitCodeDefault.Name = "chkUnitCodeDefault"
        Me.chkUnitCodeDefault.Size = New System.Drawing.Size(82, 17)
        Me.chkUnitCodeDefault.TabIndex = 3
        Me.chkUnitCodeDefault.Text = "Unit Default"
        Me.chkUnitCodeDefault.UseVisualStyleBackColor = True
        '
        'chkToleranceDown
        '
        Me.chkToleranceDown.AutoSize = True
        Me.chkToleranceDown.Location = New System.Drawing.Point(125, 11)
        Me.chkToleranceDown.Margin = New System.Windows.Forms.Padding(2)
        Me.chkToleranceDown.Name = "chkToleranceDown"
        Me.chkToleranceDown.Size = New System.Drawing.Size(105, 17)
        Me.chkToleranceDown.TabIndex = 4
        Me.chkToleranceDown.Text = "Tolerance Down"
        Me.chkToleranceDown.UseVisualStyleBackColor = True
        '
        'chkToleranceUp
        '
        Me.chkToleranceUp.AutoSize = True
        Me.chkToleranceUp.Location = New System.Drawing.Point(125, 33)
        Me.chkToleranceUp.Margin = New System.Windows.Forms.Padding(2)
        Me.chkToleranceUp.Name = "chkToleranceUp"
        Me.chkToleranceUp.Size = New System.Drawing.Size(91, 17)
        Me.chkToleranceUp.TabIndex = 5
        Me.chkToleranceUp.Text = "Tolerance Up"
        Me.chkToleranceUp.UseVisualStyleBackColor = True
        '
        'FrmTransaksiPO_SetVisible
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(237, 87)
        Me.Controls.Add(Me.chkToleranceUp)
        Me.Controls.Add(Me.chkToleranceDown)
        Me.Controls.Add(Me.chkUnitCodeDefault)
        Me.Controls.Add(Me.chkQtyDefault)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(253, 126)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(253, 126)
        Me.Name = "FrmTransaksiPO_SetVisible"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Column Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents chkQtyDefault As System.Windows.Forms.CheckBox
    Friend WithEvents chkUnitCodeDefault As System.Windows.Forms.CheckBox
    Friend WithEvents chkToleranceDown As System.Windows.Forms.CheckBox
    Friend WithEvents chkToleranceUp As System.Windows.Forms.CheckBox
End Class
