<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_SF_SetVisible
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
        Me.chkSONumber = New System.Windows.Forms.CheckBox
        Me.chkCustomerPONumber = New System.Windows.Forms.CheckBox
        Me.chkProductDesc = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(192, 46)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 19)
        Me.btnCancel.TabIndex = 0
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(192, 22)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 19)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'chkSONumber
        '
        Me.chkSONumber.AutoSize = True
        Me.chkSONumber.Location = New System.Drawing.Point(9, 10)
        Me.chkSONumber.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSONumber.Name = "chkSONumber"
        Me.chkSONumber.Size = New System.Drawing.Size(81, 17)
        Me.chkSONumber.TabIndex = 2
        Me.chkSONumber.Text = "SO Number"
        Me.chkSONumber.UseVisualStyleBackColor = True
        '
        'chkCustomerPONumber
        '
        Me.chkCustomerPONumber.AutoSize = True
        Me.chkCustomerPONumber.Location = New System.Drawing.Point(9, 32)
        Me.chkCustomerPONumber.Margin = New System.Windows.Forms.Padding(2)
        Me.chkCustomerPONumber.Name = "chkCustomerPONumber"
        Me.chkCustomerPONumber.Size = New System.Drawing.Size(128, 17)
        Me.chkCustomerPONumber.TabIndex = 3
        Me.chkCustomerPONumber.Text = "Customer PO Number"
        Me.chkCustomerPONumber.UseVisualStyleBackColor = True
        '
        'chkProductDesc
        '
        Me.chkProductDesc.AutoSize = True
        Me.chkProductDesc.Location = New System.Drawing.Point(9, 54)
        Me.chkProductDesc.Margin = New System.Windows.Forms.Padding(2)
        Me.chkProductDesc.Name = "chkProductDesc"
        Me.chkProductDesc.Size = New System.Drawing.Size(119, 17)
        Me.chkProductDesc.TabIndex = 4
        Me.chkProductDesc.Text = "Product Description"
        Me.chkProductDesc.UseVisualStyleBackColor = True
        '
        'FrmTrans_MaterialPlanningDashboard_SF_SetVisible
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(251, 73)
        Me.Controls.Add(Me.chkProductDesc)
        Me.Controls.Add(Me.chkCustomerPONumber)
        Me.Controls.Add(Me.chkSONumber)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(267, 111)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(267, 111)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_SF_SetVisible"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Column Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents chkSONumber As System.Windows.Forms.CheckBox
    Friend WithEvents chkCustomerPONumber As System.Windows.Forms.CheckBox
    Friend WithEvents chkProductDesc As System.Windows.Forms.CheckBox
End Class
