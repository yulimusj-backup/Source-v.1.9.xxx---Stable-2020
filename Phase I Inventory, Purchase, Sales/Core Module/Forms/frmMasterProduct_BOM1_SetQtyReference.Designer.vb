<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMasterProduct_BOM1_SetQtyReference
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
        Me.lblQtyReference = New System.Windows.Forms.Label
        Me.txtQtyReference = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lblUnit = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblQtyReference
        '
        Me.lblQtyReference.AutoSize = True
        Me.lblQtyReference.Location = New System.Drawing.Point(3, 14)
        Me.lblQtyReference.Name = "lblQtyReference"
        Me.lblQtyReference.Size = New System.Drawing.Size(83, 13)
        Me.lblQtyReference.TabIndex = 0
        Me.lblQtyReference.Text = "lblQtyReference"
        '
        'txtQtyReference
        '
        Me.txtQtyReference.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQtyReference.Location = New System.Drawing.Point(102, 10)
        Me.txtQtyReference.Name = "txtQtyReference"
        Me.txtQtyReference.Size = New System.Drawing.Size(86, 20)
        Me.txtQtyReference.TabIndex = 0
        Me.txtQtyReference.Text = "0.00"
        Me.txtQtyReference.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(173, 38)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(65, 21)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(102, 38)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(65, 21)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lblUnit
        '
        Me.lblUnit.AutoSize = True
        Me.lblUnit.Location = New System.Drawing.Point(196, 14)
        Me.lblUnit.Name = "lblUnit"
        Me.lblUnit.Size = New System.Drawing.Size(36, 13)
        Me.lblUnit.TabIndex = 4
        Me.lblUnit.Text = "lblUnit"
        '
        'frmMasterProduct_BOM1_SetQtyReference
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(258, 67)
        Me.Controls.Add(Me.lblUnit)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.txtQtyReference)
        Me.Controls.Add(Me.lblQtyReference)
        Me.Controls.Add(Me.btnCancel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(264, 96)
        Me.Name = "frmMasterProduct_BOM1_SetQtyReference"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmMasterProduct_BOM1_SetQtyReference"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents lblQtyReference As System.Windows.Forms.Label
    Public WithEvents txtQtyReference As System.Windows.Forms.TextBox
    Public WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents lblUnit As System.Windows.Forms.Label
End Class
