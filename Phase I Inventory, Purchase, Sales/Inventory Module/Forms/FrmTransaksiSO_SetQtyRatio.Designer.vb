<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSO_SetQtyRatio
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
        Me.TxtRatioMulti = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.LlbUnitMulti = New System.Windows.Forms.Label
        Me.LlbUnitDefault = New System.Windows.Forms.Label
        Me.TxtRatioDefault = New System.Windows.Forms.TextBox
        Me.BtnOK = New System.Windows.Forms.Button
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'TxtRatioMulti
        '
        Me.TxtRatioMulti.Location = New System.Drawing.Point(8, 8)
        Me.TxtRatioMulti.Name = "TxtRatioMulti"
        Me.TxtRatioMulti.Size = New System.Drawing.Size(74, 20)
        Me.TxtRatioMulti.TabIndex = 0
        Me.TxtRatioMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(136, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(13, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "="
        '
        'LlbUnitMulti
        '
        Me.LlbUnitMulti.AutoSize = True
        Me.LlbUnitMulti.Location = New System.Drawing.Point(85, 12)
        Me.LlbUnitMulti.Name = "LlbUnitMulti"
        Me.LlbUnitMulti.Size = New System.Drawing.Size(51, 13)
        Me.LlbUnitMulti.TabIndex = 3
        Me.LlbUnitMulti.Text = "Unit Multi"
        '
        'LlbUnitDefault
        '
        Me.LlbUnitDefault.AutoSize = True
        Me.LlbUnitDefault.Location = New System.Drawing.Point(228, 12)
        Me.LlbUnitDefault.Name = "LlbUnitDefault"
        Me.LlbUnitDefault.Size = New System.Drawing.Size(63, 13)
        Me.LlbUnitDefault.TabIndex = 5
        Me.LlbUnitDefault.Text = "Unit Default"
        '
        'TxtRatioDefault
        '
        Me.TxtRatioDefault.Location = New System.Drawing.Point(151, 8)
        Me.TxtRatioDefault.Name = "TxtRatioDefault"
        Me.TxtRatioDefault.Size = New System.Drawing.Size(74, 20)
        Me.TxtRatioDefault.TabIndex = 1
        Me.TxtRatioDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(211, 39)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(75, 23)
        Me.BtnOK.TabIndex = 3
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(130, 39)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(75, 23)
        Me.BtnCancel.TabIndex = 2
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(8, 42)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryCodeNo.TabIndex = 109
        Me.TxtPrimaryCodeNo.TabStop = False
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'FrmTransaksiSO_SetQtyRatio
        '
        Me.AcceptButton = Me.BtnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(294, 71)
        Me.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.LlbUnitDefault)
        Me.Controls.Add(Me.TxtRatioDefault)
        Me.Controls.Add(Me.LlbUnitMulti)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtRatioMulti)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiSO_SetQtyRatio"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set Ratio"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtRatioMulti As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LlbUnitMulti As System.Windows.Forms.Label
    Friend WithEvents LlbUnitDefault As System.Windows.Forms.Label
    Friend WithEvents TxtRatioDefault As System.Windows.Forms.TextBox
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
End Class
