<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmGantiPassword
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnKeluar = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.tbKonfirmasiPassword = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbPasswordBaru = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbKodeUser = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.tbPasswordLama = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkShowPassword = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'btnKeluar
        '
        Me.btnKeluar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnKeluar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnKeluar.Location = New System.Drawing.Point(185, 120)
        Me.btnKeluar.Margin = New System.Windows.Forms.Padding(2)
        Me.btnKeluar.Name = "btnKeluar"
        Me.btnKeluar.Size = New System.Drawing.Size(87, 27)
        Me.btnKeluar.TabIndex = 6
        Me.btnKeluar.TabStop = False
        Me.btnKeluar.Text = "&Keluar"
        Me.btnKeluar.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.Location = New System.Drawing.Point(72, 120)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(87, 27)
        Me.btnOK.TabIndex = 5
        Me.btnOK.TabStop = False
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'tbKonfirmasiPassword
        '
        Me.tbKonfirmasiPassword.Location = New System.Drawing.Point(155, 91)
        Me.tbKonfirmasiPassword.Margin = New System.Windows.Forms.Padding(2)
        Me.tbKonfirmasiPassword.Name = "tbKonfirmasiPassword"
        Me.tbKonfirmasiPassword.Size = New System.Drawing.Size(178, 20)
        Me.tbKonfirmasiPassword.TabIndex = 4
        Me.tbKonfirmasiPassword.UseSystemPasswordChar = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(8, 93)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(129, 13)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Konfirmasi Password Baru"
        '
        'tbPasswordBaru
        '
        Me.tbPasswordBaru.Location = New System.Drawing.Point(155, 69)
        Me.tbPasswordBaru.Margin = New System.Windows.Forms.Padding(2)
        Me.tbPasswordBaru.Name = "tbPasswordBaru"
        Me.tbPasswordBaru.Size = New System.Drawing.Size(178, 20)
        Me.tbPasswordBaru.TabIndex = 3
        Me.tbPasswordBaru.UseSystemPasswordChar = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(8, 71)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Password Baru"
        '
        'tbKodeUser
        '
        Me.tbKodeUser.Location = New System.Drawing.Point(155, 4)
        Me.tbKodeUser.Margin = New System.Windows.Forms.Padding(2)
        Me.tbKodeUser.Name = "tbKodeUser"
        Me.tbKodeUser.ReadOnly = True
        Me.tbKodeUser.Size = New System.Drawing.Size(178, 20)
        Me.tbKodeUser.TabIndex = 0
        Me.tbKodeUser.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 6)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Kode User"
        '
        'tbPasswordLama
        '
        Me.tbPasswordLama.Location = New System.Drawing.Point(155, 26)
        Me.tbPasswordLama.Margin = New System.Windows.Forms.Padding(2)
        Me.tbPasswordLama.Name = "tbPasswordLama"
        Me.tbPasswordLama.Size = New System.Drawing.Size(178, 20)
        Me.tbPasswordLama.TabIndex = 1
        Me.tbPasswordLama.UseSystemPasswordChar = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 28)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Password Lama"
        '
        'chkShowPassword
        '
        Me.chkShowPassword.AutoSize = True
        Me.chkShowPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowPassword.Location = New System.Drawing.Point(8, 48)
        Me.chkShowPassword.Margin = New System.Windows.Forms.Padding(2)
        Me.chkShowPassword.Name = "chkShowPassword"
        Me.chkShowPassword.Size = New System.Drawing.Size(125, 17)
        Me.chkShowPassword.TabIndex = 2
        Me.chkShowPassword.Text = "Tunjukkan password"
        Me.chkShowPassword.UseVisualStyleBackColor = True
        '
        'FrmGantiPassword
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnKeluar
        Me.ClientSize = New System.Drawing.Size(344, 154)
        Me.Controls.Add(Me.chkShowPassword)
        Me.Controls.Add(Me.btnKeluar)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.tbKonfirmasiPassword)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbPasswordBaru)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbKodeUser)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbPasswordLama)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmGantiPassword"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ganti Password"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnKeluar As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents tbKonfirmasiPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbPasswordBaru As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbKodeUser As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbPasswordLama As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkShowPassword As System.Windows.Forms.CheckBox
End Class
