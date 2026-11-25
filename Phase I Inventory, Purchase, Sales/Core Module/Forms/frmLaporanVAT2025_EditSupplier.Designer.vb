<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLaporanVAT2025_EditSupplier
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
        Me.gboTax = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboTypeNPWP = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.MTxtNPWP = New System.Windows.Forms.MaskedTextBox
        Me.txtPrimary = New System.Windows.Forms.TextBox
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.TxtEmail = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtTKUID = New System.Windows.Forms.TextBox
        Me.TxtNPWP_Name = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtAddressPajak1 = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.TxtAddressPajak2 = New System.Windows.Forms.TextBox
        Me.TxtTaxCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.gboTax.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboTax
        '
        Me.gboTax.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboTax.Controls.Add(Me.Label4)
        Me.gboTax.Controls.Add(Me.cboTypeNPWP)
        Me.gboTax.Controls.Add(Me.Label2)
        Me.gboTax.Controls.Add(Me.MTxtNPWP)
        Me.gboTax.Controls.Add(Me.txtPrimary)
        Me.gboTax.Controls.Add(Me.BtnCancel)
        Me.gboTax.Controls.Add(Me.btnSave)
        Me.gboTax.Controls.Add(Me.TxtEmail)
        Me.gboTax.Controls.Add(Me.Label15)
        Me.gboTax.Controls.Add(Me.Label16)
        Me.gboTax.Controls.Add(Me.TxtTKUID)
        Me.gboTax.Controls.Add(Me.TxtNPWP_Name)
        Me.gboTax.Controls.Add(Me.Label14)
        Me.gboTax.Controls.Add(Me.TxtAddressPajak1)
        Me.gboTax.Controls.Add(Me.Label22)
        Me.gboTax.Controls.Add(Me.Label20)
        Me.gboTax.Controls.Add(Me.Label21)
        Me.gboTax.Controls.Add(Me.TxtAddressPajak2)
        Me.gboTax.Controls.Add(Me.TxtTaxCode)
        Me.gboTax.Location = New System.Drawing.Point(1, 36)
        Me.gboTax.Name = "gboTax"
        Me.gboTax.Size = New System.Drawing.Size(476, 190)
        Me.gboTax.TabIndex = 14
        Me.gboTax.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 46)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(57, 13)
        Me.Label4.TabIndex = 96
        Me.Label4.Text = "NPWP No"
        '
        'cboTypeNPWP
        '
        Me.cboTypeNPWP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboTypeNPWP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTypeNPWP.FormattingEnabled = True
        Me.cboTypeNPWP.Items.AddRange(New Object() {"Company", "Individual"})
        Me.cboTypeNPWP.Location = New System.Drawing.Point(110, 15)
        Me.cboTypeNPWP.Name = "cboTypeNPWP"
        Me.cboTypeNPWP.Size = New System.Drawing.Size(361, 21)
        Me.cboTypeNPWP.TabIndex = 94
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 13)
        Me.Label2.TabIndex = 95
        Me.Label2.Text = "NPWP Type"
        '
        'MTxtNPWP
        '
        Me.MTxtNPWP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MTxtNPWP.Location = New System.Drawing.Point(110, 43)
        Me.MTxtNPWP.Mask = "0000 0000 0000 0000"
        Me.MTxtNPWP.Name = "MTxtNPWP"
        Me.MTxtNPWP.Size = New System.Drawing.Size(360, 20)
        Me.MTxtNPWP.TabIndex = 93
        '
        'txtPrimary
        '
        Me.txtPrimary.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtPrimary.Location = New System.Drawing.Point(88, 15)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(16, 20)
        Me.txtPrimary.TabIndex = 90
        Me.txtPrimary.TabStop = False
        Me.txtPrimary.Visible = False
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(397, 158)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(75, 23)
        Me.BtnCancel.TabIndex = 87
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnSave.Location = New System.Drawing.Point(316, 157)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 86
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'TxtEmail
        '
        Me.TxtEmail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmail.Location = New System.Drawing.Point(224, 94)
        Me.TxtEmail.MaxLength = 100
        Me.TxtEmail.Name = "TxtEmail"
        Me.TxtEmail.Size = New System.Drawing.Size(246, 20)
        Me.TxtEmail.TabIndex = 4
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(186, 96)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(32, 13)
        Me.Label15.TabIndex = 84
        Me.Label15.Text = "Email"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(8, 96)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(43, 13)
        Me.Label16.TabIndex = 83
        Me.Label16.Text = "TKU ID"
        '
        'TxtTKUID
        '
        Me.TxtTKUID.Location = New System.Drawing.Point(110, 94)
        Me.TxtTKUID.MaxLength = 6
        Me.TxtTKUID.Name = "TxtTKUID"
        Me.TxtTKUID.Size = New System.Drawing.Size(70, 20)
        Me.TxtTKUID.TabIndex = 3
        '
        'TxtNPWP_Name
        '
        Me.TxtNPWP_Name.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNPWP_Name.Location = New System.Drawing.Point(110, 67)
        Me.TxtNPWP_Name.MaxLength = 250
        Me.TxtNPWP_Name.Name = "TxtNPWP_Name"
        Me.TxtNPWP_Name.Size = New System.Drawing.Size(360, 20)
        Me.TxtNPWP_Name.TabIndex = 2
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(8, 71)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(71, 13)
        Me.Label14.TabIndex = 72
        Me.Label14.Text = "NPWP Name"
        '
        'TxtAddressPajak1
        '
        Me.TxtAddressPajak1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddressPajak1.Location = New System.Drawing.Point(110, 120)
        Me.TxtAddressPajak1.MaxLength = 100
        Me.TxtAddressPajak1.Name = "TxtAddressPajak1"
        Me.TxtAddressPajak1.Size = New System.Drawing.Size(360, 20)
        Me.TxtAddressPajak1.TabIndex = 6
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(8, 124)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(61, 13)
        Me.Label22.TabIndex = 66
        Me.Label22.Text = "Tax Addres"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(8, 246)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(90, 13)
        Me.Label20.TabIndex = 64
        Me.Label20.Text = "Default Tax Code"
        Me.Label20.Visible = False
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(8, 271)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(95, 13)
        Me.Label21.TabIndex = 68
        Me.Label21.Text = "Tax Address Line2"
        Me.Label21.Visible = False
        '
        'TxtAddressPajak2
        '
        Me.TxtAddressPajak2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddressPajak2.Location = New System.Drawing.Point(110, 268)
        Me.TxtAddressPajak2.MaxLength = 100
        Me.TxtAddressPajak2.Name = "TxtAddressPajak2"
        Me.TxtAddressPajak2.Size = New System.Drawing.Size(360, 20)
        Me.TxtAddressPajak2.TabIndex = 7
        Me.TxtAddressPajak2.Visible = False
        '
        'TxtTaxCode
        '
        Me.TxtTaxCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTaxCode.Location = New System.Drawing.Point(110, 242)
        Me.TxtTaxCode.MaxLength = 7
        Me.TxtTaxCode.Name = "TxtTaxCode"
        Me.TxtTaxCode.Size = New System.Drawing.Size(360, 20)
        Me.TxtTaxCode.TabIndex = 5
        Me.TxtTaxCode.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 9)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(331, 30)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Based On Government Regulation No.: PMK 112/PMK.03/2022 " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "NPWP No. Contains of 16" & _
            " Digit."
        Me.Label3.UseCompatibleTextRendering = True
        '
        'frmLaporanVAT2025_EditSupplier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(480, 231)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.gboTax)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(496, 270)
        Me.Name = "frmLaporanVAT2025_EditSupplier"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit Supplier"
        Me.gboTax.ResumeLayout(False)
        Me.gboTax.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboTax As System.Windows.Forms.GroupBox
    Friend WithEvents TxtEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtTKUID As System.Windows.Forms.TextBox
    Friend WithEvents TxtNPWP_Name As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtAddressPajak1 As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TxtAddressPajak2 As System.Windows.Forms.TextBox
    Friend WithEvents TxtTaxCode As System.Windows.Forms.TextBox
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboTypeNPWP As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents MTxtNPWP As System.Windows.Forms.MaskedTextBox
End Class
