<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_DetailWO_ColSetting
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
        Me.ChkCodeNo = New System.Windows.Forms.CheckBox
        Me.ChkDesc = New System.Windows.Forms.CheckBox
        Me.ChkProcess = New System.Windows.Forms.CheckBox
        Me.ChkProcessDesc = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ChkCodeNo
        '
        Me.ChkCodeNo.AutoSize = True
        Me.ChkCodeNo.Location = New System.Drawing.Point(12, 12)
        Me.ChkCodeNo.Name = "ChkCodeNo"
        Me.ChkCodeNo.Size = New System.Drawing.Size(68, 17)
        Me.ChkCodeNo.TabIndex = 0
        Me.ChkCodeNo.Text = "Code No"
        Me.ChkCodeNo.UseVisualStyleBackColor = True
        '
        'ChkDesc
        '
        Me.ChkDesc.AutoSize = True
        Me.ChkDesc.Location = New System.Drawing.Point(12, 35)
        Me.ChkDesc.Name = "ChkDesc"
        Me.ChkDesc.Size = New System.Drawing.Size(79, 17)
        Me.ChkDesc.TabIndex = 1
        Me.ChkDesc.Text = "Description"
        Me.ChkDesc.UseVisualStyleBackColor = True
        '
        'ChkProcess
        '
        Me.ChkProcess.AutoSize = True
        Me.ChkProcess.Location = New System.Drawing.Point(111, 12)
        Me.ChkProcess.Name = "ChkProcess"
        Me.ChkProcess.Size = New System.Drawing.Size(64, 17)
        Me.ChkProcess.TabIndex = 2
        Me.ChkProcess.Text = "Process"
        Me.ChkProcess.UseVisualStyleBackColor = True
        '
        'ChkProcessDesc
        '
        Me.ChkProcessDesc.AutoSize = True
        Me.ChkProcessDesc.Location = New System.Drawing.Point(111, 35)
        Me.ChkProcessDesc.Name = "ChkProcessDesc"
        Me.ChkProcessDesc.Size = New System.Drawing.Size(120, 17)
        Me.ChkProcessDesc.TabIndex = 3
        Me.ChkProcessDesc.Text = "Process Description"
        Me.ChkProcessDesc.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(90, 58)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(63, 21)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "&OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(159, 58)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(63, 21)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "&Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'FrmTrans_ProductionSchedule_DetailWO_ColSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(234, 85)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ChkProcessDesc)
        Me.Controls.Add(Me.ChkProcess)
        Me.Controls.Add(Me.ChkDesc)
        Me.Controls.Add(Me.ChkCodeNo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTrans_ProductionSchedule_DetailWO_ColSetting"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Column Setting"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ChkCodeNo As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDesc As System.Windows.Forms.CheckBox
    Friend WithEvents ChkProcess As System.Windows.Forms.CheckBox
    Friend WithEvents ChkProcessDesc As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
