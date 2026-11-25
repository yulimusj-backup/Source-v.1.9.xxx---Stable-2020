<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterUser
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
        Me.GBDetail = New System.Windows.Forms.GroupBox
        Me.txtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtPrimaryUserGroup = New System.Windows.Forms.TextBox
        Me.btnFindUserGroup = New System.Windows.Forms.Button
        Me.txtUserGroupCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimaryUnit = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.tbKeterangan = New System.Windows.Forms.TextBox
        Me.tbKonfimasiPassword = New System.Windows.Forms.TextBox
        Me.tbPassword = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmdBatal = New System.Windows.Forms.Button
        Me.cmdSimpan = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.tbNamaUser = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbKodeUser = New System.Windows.Forms.TextBox
        Me.btnHapus = New System.Windows.Forms.Button
        Me.cmdCari = New System.Windows.Forms.Button
        Me.cmdKeluar = New System.Windows.Forms.Button
        Me.cmdEdit = New System.Windows.Forms.Button
        Me.cmdNew = New System.Windows.Forms.Button
        Me.GBDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'GBDetail
        '
        Me.GBDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GBDetail.BackColor = System.Drawing.SystemColors.Control
        Me.GBDetail.Controls.Add(Me.txtReadOnly)
        Me.GBDetail.Controls.Add(Me.TxtPrimaryUserGroup)
        Me.GBDetail.Controls.Add(Me.btnFindUserGroup)
        Me.GBDetail.Controls.Add(Me.txtUserGroupCode)
        Me.GBDetail.Controls.Add(Me.Label3)
        Me.GBDetail.Controls.Add(Me.TxtPrimaryUnit)
        Me.GBDetail.Controls.Add(Me.Label12)
        Me.GBDetail.Controls.Add(Me.tbKeterangan)
        Me.GBDetail.Controls.Add(Me.tbKonfimasiPassword)
        Me.GBDetail.Controls.Add(Me.tbPassword)
        Me.GBDetail.Controls.Add(Me.Label5)
        Me.GBDetail.Controls.Add(Me.Label6)
        Me.GBDetail.Controls.Add(Me.cmdBatal)
        Me.GBDetail.Controls.Add(Me.cmdSimpan)
        Me.GBDetail.Controls.Add(Me.Label2)
        Me.GBDetail.Controls.Add(Me.tbNamaUser)
        Me.GBDetail.Controls.Add(Me.Label1)
        Me.GBDetail.Controls.Add(Me.tbKodeUser)
        Me.GBDetail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.GBDetail.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GBDetail.Location = New System.Drawing.Point(12, -1)
        Me.GBDetail.Name = "GBDetail"
        Me.GBDetail.Size = New System.Drawing.Size(759, 180)
        Me.GBDetail.TabIndex = 5
        Me.GBDetail.TabStop = False
        '
        'txtReadOnly
        '
        Me.txtReadOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.txtReadOnly.Location = New System.Drawing.Point(134, 90)
        Me.txtReadOnly.Name = "txtReadOnly"
        Me.txtReadOnly.ReadOnly = True
        Me.txtReadOnly.Size = New System.Drawing.Size(23, 20)
        Me.txtReadOnly.TabIndex = 249
        Me.txtReadOnly.TabStop = False
        Me.txtReadOnly.Visible = False
        '
        'TxtPrimaryUserGroup
        '
        Me.TxtPrimaryUserGroup.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.TxtPrimaryUserGroup.Location = New System.Drawing.Point(114, 152)
        Me.TxtPrimaryUserGroup.Name = "TxtPrimaryUserGroup"
        Me.TxtPrimaryUserGroup.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimaryUserGroup.TabIndex = 47
        Me.TxtPrimaryUserGroup.TabStop = False
        Me.TxtPrimaryUserGroup.Visible = False
        '
        'btnFindUserGroup
        '
        Me.btnFindUserGroup.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindUserGroup.Enabled = False
        Me.btnFindUserGroup.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnFindUserGroup.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnFindUserGroup.ForeColor = System.Drawing.Color.Black
        Me.btnFindUserGroup.Location = New System.Drawing.Point(339, 149)
        Me.btnFindUserGroup.Name = "btnFindUserGroup"
        Me.btnFindUserGroup.Size = New System.Drawing.Size(30, 21)
        Me.btnFindUserGroup.TabIndex = 5
        Me.btnFindUserGroup.TabStop = False
        Me.btnFindUserGroup.Text = "..."
        Me.btnFindUserGroup.UseVisualStyleBackColor = True
        '
        'txtUserGroupCode
        '
        Me.txtUserGroupCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUserGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtUserGroupCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.txtUserGroupCode.Location = New System.Drawing.Point(114, 149)
        Me.txtUserGroupCode.MaxLength = 20
        Me.txtUserGroupCode.Name = "txtUserGroupCode"
        Me.txtUserGroupCode.ReadOnly = True
        Me.txtUserGroupCode.Size = New System.Drawing.Size(219, 20)
        Me.txtUserGroupCode.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Label3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label3.Location = New System.Drawing.Point(16, 152)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 13)
        Me.Label3.TabIndex = 45
        Me.Label3.Text = "User Group"
        '
        'TxtPrimaryUnit
        '
        Me.TxtPrimaryUnit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.TxtPrimaryUnit.Location = New System.Drawing.Point(112, 30)
        Me.TxtPrimaryUnit.Name = "TxtPrimaryUnit"
        Me.TxtPrimaryUnit.Size = New System.Drawing.Size(43, 20)
        Me.TxtPrimaryUnit.TabIndex = 43
        Me.TxtPrimaryUnit.TabStop = False
        Me.TxtPrimaryUnit.Visible = False
        '
        'Label12
        '
        Me.Label12.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Label12.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label12.Location = New System.Drawing.Point(383, 31)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(49, 13)
        Me.Label12.TabIndex = 41
        Me.Label12.Text = "Remarks"
        '
        'tbKeterangan
        '
        Me.tbKeterangan.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbKeterangan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tbKeterangan.Location = New System.Drawing.Point(463, 28)
        Me.tbKeterangan.Multiline = True
        Me.tbKeterangan.Name = "tbKeterangan"
        Me.tbKeterangan.ReadOnly = True
        Me.tbKeterangan.Size = New System.Drawing.Size(267, 79)
        Me.tbKeterangan.TabIndex = 6
        '
        'tbKonfimasiPassword
        '
        Me.tbKonfimasiPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbKonfimasiPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tbKonfimasiPassword.Location = New System.Drawing.Point(114, 122)
        Me.tbKonfimasiPassword.MaxLength = 20
        Me.tbKonfimasiPassword.Name = "tbKonfimasiPassword"
        Me.tbKonfimasiPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.tbKonfimasiPassword.ReadOnly = True
        Me.tbKonfimasiPassword.Size = New System.Drawing.Size(255, 20)
        Me.tbKonfimasiPassword.TabIndex = 3
        Me.tbKonfimasiPassword.UseSystemPasswordChar = True
        '
        'tbPassword
        '
        Me.tbPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tbPassword.Location = New System.Drawing.Point(114, 90)
        Me.tbPassword.MaxLength = 20
        Me.tbPassword.Name = "tbPassword"
        Me.tbPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.tbPassword.ReadOnly = True
        Me.tbPassword.Size = New System.Drawing.Size(255, 20)
        Me.tbPassword.TabIndex = 2
        Me.tbPassword.UseSystemPasswordChar = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Label5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label5.Location = New System.Drawing.Point(16, 125)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(91, 13)
        Me.Label5.TabIndex = 36
        Me.Label5.Text = "Confirm Password"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Label6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label6.Location = New System.Drawing.Point(16, 93)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 13)
        Me.Label6.TabIndex = 34
        Me.Label6.Text = "Password"
        '
        'cmdBatal
        '
        Me.cmdBatal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBatal.Enabled = False
        Me.cmdBatal.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdBatal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.cmdBatal.ForeColor = System.Drawing.Color.Black
        Me.cmdBatal.Location = New System.Drawing.Point(623, 125)
        Me.cmdBatal.Name = "cmdBatal"
        Me.cmdBatal.Size = New System.Drawing.Size(107, 30)
        Me.cmdBatal.TabIndex = 8
        Me.cmdBatal.Text = "Cancel"
        Me.cmdBatal.UseVisualStyleBackColor = True
        '
        'cmdSimpan
        '
        Me.cmdSimpan.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSimpan.Enabled = False
        Me.cmdSimpan.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdSimpan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.cmdSimpan.ForeColor = System.Drawing.Color.Black
        Me.cmdSimpan.Location = New System.Drawing.Point(463, 125)
        Me.cmdSimpan.Name = "cmdSimpan"
        Me.cmdSimpan.Size = New System.Drawing.Size(105, 30)
        Me.cmdSimpan.TabIndex = 7
        Me.cmdSimpan.Text = "OK"
        Me.cmdSimpan.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Label2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.Location = New System.Drawing.Point(15, 63)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Name"
        '
        'tbNamaUser
        '
        Me.tbNamaUser.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbNamaUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tbNamaUser.Location = New System.Drawing.Point(114, 60)
        Me.tbNamaUser.MaxLength = 200
        Me.tbNamaUser.Name = "tbNamaUser"
        Me.tbNamaUser.ReadOnly = True
        Me.tbNamaUser.Size = New System.Drawing.Size(255, 20)
        Me.tbNamaUser.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(15, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 13)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "User Code"
        '
        'tbKodeUser
        '
        Me.tbKodeUser.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbKodeUser.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.tbKodeUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tbKodeUser.Location = New System.Drawing.Point(114, 28)
        Me.tbKodeUser.MaxLength = 20
        Me.tbKodeUser.Name = "tbKodeUser"
        Me.tbKodeUser.ReadOnly = True
        Me.tbKodeUser.Size = New System.Drawing.Size(255, 20)
        Me.tbKodeUser.TabIndex = 0
        '
        'btnHapus
        '
        Me.btnHapus.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnHapus.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnHapus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnHapus.Location = New System.Drawing.Point(475, 185)
        Me.btnHapus.Name = "btnHapus"
        Me.btnHapus.Size = New System.Drawing.Size(139, 26)
        Me.btnHapus.TabIndex = 3
        Me.btnHapus.Text = "Delete"
        Me.btnHapus.UseVisualStyleBackColor = True
        '
        'cmdCari
        '
        Me.cmdCari.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cmdCari.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdCari.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.cmdCari.Location = New System.Drawing.Point(16, 185)
        Me.cmdCari.Name = "cmdCari"
        Me.cmdCari.Size = New System.Drawing.Size(139, 26)
        Me.cmdCari.TabIndex = 0
        Me.cmdCari.Text = "Find"
        Me.cmdCari.UseVisualStyleBackColor = True
        '
        'cmdKeluar
        '
        Me.cmdKeluar.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cmdKeluar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdKeluar.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdKeluar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.cmdKeluar.Location = New System.Drawing.Point(628, 185)
        Me.cmdKeluar.Name = "cmdKeluar"
        Me.cmdKeluar.Size = New System.Drawing.Size(139, 26)
        Me.cmdKeluar.TabIndex = 4
        Me.cmdKeluar.Text = "Exit"
        Me.cmdKeluar.UseVisualStyleBackColor = True
        '
        'cmdEdit
        '
        Me.cmdEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cmdEdit.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.cmdEdit.Location = New System.Drawing.Point(322, 185)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(139, 26)
        Me.cmdEdit.TabIndex = 2
        Me.cmdEdit.Text = "Edit"
        Me.cmdEdit.UseVisualStyleBackColor = True
        '
        'cmdNew
        '
        Me.cmdNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cmdNew.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdNew.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.cmdNew.Location = New System.Drawing.Point(169, 185)
        Me.cmdNew.Name = "cmdNew"
        Me.cmdNew.Size = New System.Drawing.Size(139, 26)
        Me.cmdNew.TabIndex = 1
        Me.cmdNew.Text = "New"
        Me.cmdNew.UseVisualStyleBackColor = True
        '
        'FrmMasterUser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdKeluar
        Me.ClientSize = New System.Drawing.Size(783, 223)
        Me.Controls.Add(Me.btnHapus)
        Me.Controls.Add(Me.cmdCari)
        Me.Controls.Add(Me.cmdKeluar)
        Me.Controls.Add(Me.cmdEdit)
        Me.Controls.Add(Me.cmdNew)
        Me.Controls.Add(Me.GBDetail)
        Me.MinimumSize = New System.Drawing.Size(789, 249)
        Me.Name = "FrmMasterUser"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "User Data"
        Me.GBDetail.ResumeLayout(False)
        Me.GBDetail.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GBDetail As System.Windows.Forms.GroupBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents tbKeterangan As System.Windows.Forms.TextBox
    Friend WithEvents tbKonfimasiPassword As System.Windows.Forms.TextBox
    Friend WithEvents tbPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdBatal As System.Windows.Forms.Button
    Friend WithEvents cmdSimpan As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbNamaUser As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbKodeUser As System.Windows.Forms.TextBox
    Friend WithEvents btnHapus As System.Windows.Forms.Button
    Friend WithEvents cmdCari As System.Windows.Forms.Button
    Friend WithEvents cmdKeluar As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents cmdNew As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryUnit As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryUserGroup As System.Windows.Forms.TextBox
    Friend WithEvents btnFindUserGroup As System.Windows.Forms.Button
    Friend WithEvents txtUserGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtReadOnly As System.Windows.Forms.TextBox
End Class
