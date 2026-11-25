<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_MP_SetVisible
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
        Me.chkProcurementStrategy = New System.Windows.Forms.CheckBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.chkOrderStrategy = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'chkProcurementStrategy
        '
        Me.chkProcurementStrategy.AutoSize = True
        Me.chkProcurementStrategy.Location = New System.Drawing.Point(13, 8)
        Me.chkProcurementStrategy.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkProcurementStrategy.Name = "chkProcurementStrategy"
        Me.chkProcurementStrategy.Size = New System.Drawing.Size(168, 21)
        Me.chkProcurementStrategy.TabIndex = 7
        Me.chkProcurementStrategy.Text = "Procurement Strategy"
        Me.chkProcurementStrategy.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(257, 11)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 6
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(257, 41)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'chkOrderStrategy
        '
        Me.chkOrderStrategy.AutoSize = True
        Me.chkOrderStrategy.Location = New System.Drawing.Point(12, 33)
        Me.chkOrderStrategy.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkOrderStrategy.Name = "chkOrderStrategy"
        Me.chkOrderStrategy.Size = New System.Drawing.Size(147, 21)
        Me.chkOrderStrategy.TabIndex = 8
        Me.chkOrderStrategy.Text = "Procurement Type"
        Me.chkOrderStrategy.UseVisualStyleBackColor = True
        '
        'FrmTrans_MaterialPlanningDashboard_MP_SetVisible
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(345, 91)
        Me.Controls.Add(Me.chkOrderStrategy)
        Me.Controls.Add(Me.chkProcurementStrategy)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(353, 128)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(353, 128)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_MP_SetVisible"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Columns Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkProcurementStrategy As System.Windows.Forms.CheckBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents chkOrderStrategy As System.Windows.Forms.CheckBox
End Class
