<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPO_DetailShipTo
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
        Me.btnExit = New System.Windows.Forms.Button
        Me.cboCustAddCode = New System.Windows.Forms.ComboBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.txtAdd = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.TxtIdentifier = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(314, 155)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 20)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "&Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'cboCustAddCode
        '
        Me.cboCustAddCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustAddCode.FormattingEnabled = True
        Me.cboCustAddCode.Location = New System.Drawing.Point(13, 13)
        Me.cboCustAddCode.Name = "cboCustAddCode"
        Me.cboCustAddCode.Size = New System.Drawing.Size(139, 21)
        Me.cboCustAddCode.TabIndex = 5
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(314, 155)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 20)
        Me.btnOK.TabIndex = 6
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'txtAdd
        '
        Me.txtAdd.Location = New System.Drawing.Point(158, 13)
        Me.txtAdd.Multiline = True
        Me.txtAdd.Name = "txtAdd"
        Me.txtAdd.ReadOnly = True
        Me.txtAdd.Size = New System.Drawing.Size(231, 136)
        Me.txtAdd.TabIndex = 8
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(279, 76)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(65, 37)
        Me.Button1.TabIndex = 9
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TxtIdentifier
        '
        Me.TxtIdentifier.Location = New System.Drawing.Point(13, 40)
        Me.TxtIdentifier.Name = "TxtIdentifier"
        Me.TxtIdentifier.Size = New System.Drawing.Size(60, 20)
        Me.TxtIdentifier.TabIndex = 10
        Me.TxtIdentifier.TabStop = False
        Me.TxtIdentifier.Visible = False
        '
        'FrmTransaksiPO_DetailShipTo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(403, 184)
        Me.Controls.Add(Me.TxtIdentifier)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.cboCustAddCode)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.txtAdd)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(409, 208)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(409, 208)
        Me.Name = "FrmTransaksiPO_DetailShipTo"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrmDetailAddressSupplier"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents btnExit As System.Windows.Forms.Button
    Public WithEvents cboCustAddCode As System.Windows.Forms.ComboBox
    Public WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents txtAdd As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TxtIdentifier As System.Windows.Forms.TextBox
End Class
