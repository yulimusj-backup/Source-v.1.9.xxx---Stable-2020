<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmAddReplace
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
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.TxtOverrideAmount = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtRemarksItem = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtPrimaryLoanPayment = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.rboByBS = New System.Windows.Forms.RadioButton
        Me.TxtBS = New System.Windows.Forms.TextBox
        Me.rboFixedValue = New System.Windows.Forms.RadioButton
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.OK_Button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OK_Button.Location = New System.Drawing.Point(49, 186)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(89, 28)
        Me.OK_Button.TabIndex = 2
        Me.OK_Button.Text = "Add"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Yes
        Me.Cancel_Button.Location = New System.Drawing.Point(147, 186)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(89, 28)
        Me.Cancel_Button.TabIndex = 3
        Me.Cancel_Button.Text = "Replace"
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(244, 186)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(89, 28)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Cancel"
        '
        'TxtOverrideAmount
        '
        Me.TxtOverrideAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtOverrideAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtOverrideAmount.Location = New System.Drawing.Point(141, 22)
        Me.TxtOverrideAmount.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtOverrideAmount.MaxLength = 100
        Me.TxtOverrideAmount.Name = "TxtOverrideAmount"
        Me.TxtOverrideAmount.Size = New System.Drawing.Size(227, 22)
        Me.TxtOverrideAmount.TabIndex = 0
        Me.TxtOverrideAmount.Text = "0"
        Me.TxtOverrideAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 99)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(116, 17)
        Me.Label1.TabIndex = 40
        Me.Label1.Text = "Override Reason"
        '
        'TxtRemarksItem
        '
        Me.TxtRemarksItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarksItem.Location = New System.Drawing.Point(153, 99)
        Me.TxtRemarksItem.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtRemarksItem.MaxLength = 250
        Me.TxtRemarksItem.Multiline = True
        Me.TxtRemarksItem.Name = "TxtRemarksItem"
        Me.TxtRemarksItem.Size = New System.Drawing.Size(233, 79)
        Me.TxtRemarksItem.TabIndex = 1
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(25, 126)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtReadOnly.MaxLength = 30
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(37, 22)
        Me.TxtReadOnly.TabIndex = 48
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtPrimaryLoanPayment
        '
        Me.TxtPrimaryLoanPayment.Location = New System.Drawing.Point(72, 99)
        Me.TxtPrimaryLoanPayment.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryLoanPayment.Name = "TxtPrimaryLoanPayment"
        Me.TxtPrimaryLoanPayment.Size = New System.Drawing.Size(32, 22)
        Me.TxtPrimaryLoanPayment.TabIndex = 49
        Me.TxtPrimaryLoanPayment.TabStop = False
        Me.TxtPrimaryLoanPayment.Text = "-1"
        Me.TxtPrimaryLoanPayment.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.rboByBS)
        Me.GroupBox1.Controls.Add(Me.TxtBS)
        Me.GroupBox1.Controls.Add(Me.rboFixedValue)
        Me.GroupBox1.Controls.Add(Me.TxtOverrideAmount)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(375, 81)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Override Amount"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(272, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 17)
        Me.Label2.TabIndex = 42
        Me.Label2.Text = "x Basic Salary"
        '
        'rboByBS
        '
        Me.rboByBS.AutoSize = True
        Me.rboByBS.Location = New System.Drawing.Point(7, 53)
        Me.rboByBS.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rboByBS.Name = "rboByBS"
        Me.rboByBS.Size = New System.Drawing.Size(127, 21)
        Me.rboByBS.TabIndex = 41
        Me.rboByBS.Text = "By Basic Salary"
        Me.rboByBS.UseVisualStyleBackColor = True
        '
        'TxtBS
        '
        Me.TxtBS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBS.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBS.Enabled = False
        Me.TxtBS.Location = New System.Drawing.Point(141, 52)
        Me.TxtBS.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtBS.MaxLength = 100
        Me.TxtBS.Name = "TxtBS"
        Me.TxtBS.Size = New System.Drawing.Size(124, 22)
        Me.TxtBS.TabIndex = 1
        Me.TxtBS.Text = "1"
        Me.TxtBS.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'rboFixedValue
        '
        Me.rboFixedValue.AutoSize = True
        Me.rboFixedValue.Checked = True
        Me.rboFixedValue.Location = New System.Drawing.Point(7, 23)
        Me.rboFixedValue.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rboFixedValue.Name = "rboFixedValue"
        Me.rboFixedValue.Size = New System.Drawing.Size(102, 21)
        Me.rboFixedValue.TabIndex = 39
        Me.rboFixedValue.TabStop = True
        Me.rboFixedValue.Text = "Fixed Value"
        Me.rboFixedValue.UseVisualStyleBackColor = True
        '
        'FrmAddReplace
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(396, 226)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.TxtPrimaryLoanPayment)
        Me.Controls.Add(Me.TxtRemarksItem)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(402, 261)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(402, 261)
        Me.Name = "FrmAddReplace"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add / Replace"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TxtOverrideAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarksItem As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryLoanPayment As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rboByBS As System.Windows.Forms.RadioButton
    Friend WithEvents TxtBS As System.Windows.Forms.TextBox
    Friend WithEvents rboFixedValue As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label

End Class
