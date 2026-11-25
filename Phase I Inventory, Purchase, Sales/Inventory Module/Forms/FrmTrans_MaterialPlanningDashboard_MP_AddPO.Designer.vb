<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_MP_AddPO
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.rboNoDeliveryDate = New System.Windows.Forms.RadioButton
        Me.btnCancel = New System.Windows.Forms.Button
        Me.rboSetDeliveryDate = New System.Windows.Forms.RadioButton
        Me.rboSetDeliveryDate2 = New System.Windows.Forms.RadioButton
        Me.btnOK = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(348, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Please specify how to specify PO delivery date control"
        '
        'rboNoDeliveryDate
        '
        Me.rboNoDeliveryDate.AutoSize = True
        Me.rboNoDeliveryDate.Checked = True
        Me.rboNoDeliveryDate.Location = New System.Drawing.Point(12, 29)
        Me.rboNoDeliveryDate.Name = "rboNoDeliveryDate"
        Me.rboNoDeliveryDate.Size = New System.Drawing.Size(195, 21)
        Me.rboNoDeliveryDate.TabIndex = 1
        Me.rboNoDeliveryDate.TabStop = True
        Me.rboNoDeliveryDate.Text = "Don't specify delivery date"
        Me.rboNoDeliveryDate.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(372, 43)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 25)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'rboSetDeliveryDate
        '
        Me.rboSetDeliveryDate.AutoSize = True
        Me.rboSetDeliveryDate.Location = New System.Drawing.Point(12, 56)
        Me.rboSetDeliveryDate.Name = "rboSetDeliveryDate"
        Me.rboSetDeliveryDate.Size = New System.Drawing.Size(228, 21)
        Me.rboSetDeliveryDate.TabIndex = 3
        Me.rboSetDeliveryDate.Text = "Specify delivery date if required"
        Me.rboSetDeliveryDate.UseVisualStyleBackColor = True
        '
        'rboSetDeliveryDate2
        '
        Me.rboSetDeliveryDate2.AutoSize = True
        Me.rboSetDeliveryDate2.Location = New System.Drawing.Point(12, 83)
        Me.rboSetDeliveryDate2.Name = "rboSetDeliveryDate2"
        Me.rboSetDeliveryDate2.Size = New System.Drawing.Size(348, 21)
        Me.rboSetDeliveryDate2.TabIndex = 4
        Me.rboSetDeliveryDate2.Text = "Differentiate purchase order for each delivery date"
        Me.rboSetDeliveryDate2.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(372, 12)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 25)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "Ok"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTrans_MaterialPlanningDashboard_MP_AddPO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(459, 118)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.rboSetDeliveryDate2)
        Me.Controls.Add(Me.rboSetDeliveryDate)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.rboNoDeliveryDate)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(465, 153)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_MP_AddPO"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add PO Mode"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents rboNoDeliveryDate As System.Windows.Forms.RadioButton
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents rboSetDeliveryDate As System.Windows.Forms.RadioButton
    Friend WithEvents rboSetDeliveryDate2 As System.Windows.Forms.RadioButton
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
