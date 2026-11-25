<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_WIP_ColumnSettings
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
        Me.ChkWO = New System.Windows.Forms.CheckBox
        Me.ChkProcessCode = New System.Windows.Forms.CheckBox
        Me.ChkCodeNo = New System.Windows.Forms.CheckBox
        Me.ChkDesc = New System.Windows.Forms.CheckBox
        Me.BtnOK = New System.Windows.Forms.Button
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ChkWO
        '
        Me.ChkWO.AutoSize = True
        Me.ChkWO.Location = New System.Drawing.Point(16, 12)
        Me.ChkWO.Name = "ChkWO"
        Me.ChkWO.Size = New System.Drawing.Size(62, 17)
        Me.ChkWO.TabIndex = 0
        Me.ChkWO.Text = "WO No"
        Me.ChkWO.UseVisualStyleBackColor = True
        '
        'ChkProcessCode
        '
        Me.ChkProcessCode.AutoSize = True
        Me.ChkProcessCode.Location = New System.Drawing.Point(16, 35)
        Me.ChkProcessCode.Name = "ChkProcessCode"
        Me.ChkProcessCode.Size = New System.Drawing.Size(92, 17)
        Me.ChkProcessCode.TabIndex = 1
        Me.ChkProcessCode.Text = "Process Code"
        Me.ChkProcessCode.UseVisualStyleBackColor = True
        '
        'ChkCodeNo
        '
        Me.ChkCodeNo.AutoSize = True
        Me.ChkCodeNo.Location = New System.Drawing.Point(122, 12)
        Me.ChkCodeNo.Name = "ChkCodeNo"
        Me.ChkCodeNo.Size = New System.Drawing.Size(68, 17)
        Me.ChkCodeNo.TabIndex = 2
        Me.ChkCodeNo.Text = "Code No"
        Me.ChkCodeNo.UseVisualStyleBackColor = True
        '
        'ChkDesc
        '
        Me.ChkDesc.AutoSize = True
        Me.ChkDesc.Location = New System.Drawing.Point(122, 35)
        Me.ChkDesc.Name = "ChkDesc"
        Me.ChkDesc.Size = New System.Drawing.Size(79, 17)
        Me.ChkDesc.TabIndex = 3
        Me.ChkDesc.Text = "Description"
        Me.ChkDesc.UseVisualStyleBackColor = True
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(40, 60)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(65, 21)
        Me.BtnOK.TabIndex = 4
        Me.BtnOK.Text = "&OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(111, 60)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(65, 21)
        Me.BtnCancel.TabIndex = 5
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'FrmTrans_ManufacturingOrder_WIP_ColumnSettings
        '
        Me.AcceptButton = Me.BtnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(226, 91)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.ChkDesc)
        Me.Controls.Add(Me.ChkCodeNo)
        Me.Controls.Add(Me.ChkProcessCode)
        Me.Controls.Add(Me.ChkWO)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTrans_ManufacturingOrder_WIP_ColumnSettings"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Columns Setting"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ChkWO As System.Windows.Forms.CheckBox
    Friend WithEvents ChkProcessCode As System.Windows.Forms.CheckBox
    Friend WithEvents ChkCodeNo As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDesc As System.Windows.Forms.CheckBox
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
End Class
