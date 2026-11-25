<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockInPurchase2_SetQtyRatio
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
        Me.TxtRatioMulti.Location = New System.Drawing.Point(11, 10)
        Me.TxtRatioMulti.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtRatioMulti.Name = "TxtRatioMulti"
        Me.TxtRatioMulti.Size = New System.Drawing.Size(97, 22)
        Me.TxtRatioMulti.TabIndex = 0
        Me.TxtRatioMulti.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(181, 15)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "="
        '
        'LlbUnitMulti
        '
        Me.LlbUnitMulti.AutoSize = True
        Me.LlbUnitMulti.Location = New System.Drawing.Point(113, 15)
        Me.LlbUnitMulti.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LlbUnitMulti.Name = "LlbUnitMulti"
        Me.LlbUnitMulti.Size = New System.Drawing.Size(66, 17)
        Me.LlbUnitMulti.TabIndex = 3
        Me.LlbUnitMulti.Text = "Unit Multi"
        '
        'LlbUnitDefault
        '
        Me.LlbUnitDefault.AutoSize = True
        Me.LlbUnitDefault.Location = New System.Drawing.Point(304, 15)
        Me.LlbUnitDefault.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LlbUnitDefault.Name = "LlbUnitDefault"
        Me.LlbUnitDefault.Size = New System.Drawing.Size(82, 17)
        Me.LlbUnitDefault.TabIndex = 5
        Me.LlbUnitDefault.Text = "Unit Default"
        '
        'TxtRatioDefault
        '
        Me.TxtRatioDefault.Location = New System.Drawing.Point(201, 10)
        Me.TxtRatioDefault.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtRatioDefault.Name = "TxtRatioDefault"
        Me.TxtRatioDefault.Size = New System.Drawing.Size(97, 22)
        Me.TxtRatioDefault.TabIndex = 1
        Me.TxtRatioDefault.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(173, 48)
        Me.BtnOK.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(100, 28)
        Me.BtnOK.TabIndex = 3
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.Location = New System.Drawing.Point(281, 48)
        Me.BtnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(100, 28)
        Me.BtnCancel.TabIndex = 2
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(11, 52)
        Me.TxtPrimaryCodeNo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(36, 22)
        Me.TxtPrimaryCodeNo.TabIndex = 110
        Me.TxtPrimaryCodeNo.TabStop = False
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'FrmTransaksiStockInPurchase2_SetQtyRatio
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(392, 87)
        Me.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.LlbUnitDefault)
        Me.Controls.Add(Me.TxtRatioDefault)
        Me.Controls.Add(Me.LlbUnitMulti)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtRatioMulti)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiStockInPurchase2_SetQtyRatio"
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
