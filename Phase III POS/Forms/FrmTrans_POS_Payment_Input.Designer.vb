<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_POS_Payment_Input
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
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtApprovalCode = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtAccountNumber = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtTotalPayment = New System.Windows.Forms.TextBox
        Me.TxtPrimaryPayment = New System.Windows.Forms.TextBox
        Me.cboPaymentTipe = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 120)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(77, 13)
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "Approval Code"
        '
        'TxtApprovalCode
        '
        Me.TxtApprovalCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtApprovalCode.Location = New System.Drawing.Point(108, 117)
        Me.TxtApprovalCode.Name = "TxtApprovalCode"
        Me.TxtApprovalCode.Size = New System.Drawing.Size(237, 20)
        Me.TxtApprovalCode.TabIndex = 4
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 94)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(87, 13)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Account Number"
        '
        'TxtAccountNumber
        '
        Me.TxtAccountNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountNumber.Location = New System.Drawing.Point(108, 91)
        Me.TxtAccountNumber.Name = "TxtAccountNumber"
        Me.TxtAccountNumber.Size = New System.Drawing.Size(237, 20)
        Me.TxtAccountNumber.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 42)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(87, 13)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Payment Amount"
        '
        'TxtTotalPayment
        '
        Me.TxtTotalPayment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotalPayment.Location = New System.Drawing.Point(108, 39)
        Me.TxtTotalPayment.MaxLength = 12
        Me.TxtTotalPayment.Name = "TxtTotalPayment"
        Me.TxtTotalPayment.Size = New System.Drawing.Size(237, 20)
        Me.TxtTotalPayment.TabIndex = 1
        Me.TxtTotalPayment.Text = "0"
        Me.TxtTotalPayment.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtPrimaryPayment
        '
        Me.TxtPrimaryPayment.Location = New System.Drawing.Point(71, 12)
        Me.TxtPrimaryPayment.Name = "TxtPrimaryPayment"
        Me.TxtPrimaryPayment.Size = New System.Drawing.Size(31, 20)
        Me.TxtPrimaryPayment.TabIndex = 17
        Me.TxtPrimaryPayment.TabStop = False
        Me.TxtPrimaryPayment.Text = "-1"
        Me.TxtPrimaryPayment.Visible = False
        '
        'cboPaymentTipe
        '
        Me.cboPaymentTipe.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboPaymentTipe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPaymentTipe.FormattingEnabled = True
        Me.cboPaymentTipe.Location = New System.Drawing.Point(108, 12)
        Me.cboPaymentTipe.Name = "cboPaymentTipe"
        Me.cboPaymentTipe.Size = New System.Drawing.Size(237, 21)
        Me.cboPaymentTipe.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Payment Type"
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(189, 142)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(270, 142)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 68)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 29
        Me.Label1.Text = "Remarks"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(108, 65)
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(237, 20)
        Me.TxtRemarks.TabIndex = 2
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(108, 142)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(27, 20)
        Me.TxtReadOnly.TabIndex = 67
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'FrmTrans_POS_Payment_Input
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(357, 177)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtRemarks)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.TxtApprovalCode)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TxtAccountNumber)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TxtTotalPayment)
        Me.Controls.Add(Me.TxtPrimaryPayment)
        Me.Controls.Add(Me.cboPaymentTipe)
        Me.Controls.Add(Me.Label2)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(373, 215)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(373, 215)
        Me.Name = "FrmTrans_POS_Payment_Input"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Payment"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtApprovalCode As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtAccountNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtTotalPayment As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryPayment As System.Windows.Forms.TextBox
    Friend WithEvents cboPaymentTipe As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
End Class
