<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ManufacturingOrder_ColumnSetting
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkProcessDesc = New System.Windows.Forms.CheckBox
        Me.ChkProcessCode = New System.Windows.Forms.CheckBox
        Me.ChkDescription = New System.Windows.Forms.CheckBox
        Me.ChkCodeNo = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.ChkProcessDesc)
        Me.GroupBox1.Controls.Add(Me.ChkProcessCode)
        Me.GroupBox1.Controls.Add(Me.ChkDescription)
        Me.GroupBox1.Controls.Add(Me.ChkCodeNo)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(268, 68)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Please specify visible columns"
        '
        'ChkProcessDesc
        '
        Me.ChkProcessDesc.AutoSize = True
        Me.ChkProcessDesc.Location = New System.Drawing.Point(138, 42)
        Me.ChkProcessDesc.Name = "ChkProcessDesc"
        Me.ChkProcessDesc.Size = New System.Drawing.Size(120, 17)
        Me.ChkProcessDesc.TabIndex = 5
        Me.ChkProcessDesc.Text = "Process Description"
        Me.ChkProcessDesc.UseVisualStyleBackColor = True
        '
        'ChkProcessCode
        '
        Me.ChkProcessCode.AutoSize = True
        Me.ChkProcessCode.Location = New System.Drawing.Point(138, 19)
        Me.ChkProcessCode.Name = "ChkProcessCode"
        Me.ChkProcessCode.Size = New System.Drawing.Size(92, 17)
        Me.ChkProcessCode.TabIndex = 4
        Me.ChkProcessCode.Text = "Process Code"
        Me.ChkProcessCode.UseVisualStyleBackColor = True
        '
        'ChkDescription
        '
        Me.ChkDescription.AutoSize = True
        Me.ChkDescription.Location = New System.Drawing.Point(6, 42)
        Me.ChkDescription.Name = "ChkDescription"
        Me.ChkDescription.Size = New System.Drawing.Size(79, 17)
        Me.ChkDescription.TabIndex = 1
        Me.ChkDescription.Text = "Description"
        Me.ChkDescription.UseVisualStyleBackColor = True
        '
        'ChkCodeNo
        '
        Me.ChkCodeNo.AutoSize = True
        Me.ChkCodeNo.Location = New System.Drawing.Point(6, 19)
        Me.ChkCodeNo.Name = "ChkCodeNo"
        Me.ChkCodeNo.Size = New System.Drawing.Size(68, 17)
        Me.ChkCodeNo.TabIndex = 0
        Me.ChkCodeNo.Text = "Code No"
        Me.ChkCodeNo.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(211, 78)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(66, 22)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "&OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmTrans_ManufacturingOrder_ColumnSetting
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(283, 104)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTrans_ManufacturingOrder_ColumnSetting"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Column Settings"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkCodeNo As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ChkDescription As System.Windows.Forms.CheckBox
    Friend WithEvents ChkProcessCode As System.Windows.Forms.CheckBox
    Friend WithEvents ChkProcessDesc As System.Windows.Forms.CheckBox
End Class
