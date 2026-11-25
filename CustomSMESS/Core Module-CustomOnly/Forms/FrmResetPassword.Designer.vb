<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmResetPassword
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
        Me.btnKeluar = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboUser = New System.Windows.Forms.ComboBox
        Me.txtNamaUser = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnBatal = New System.Windows.Forms.Button
        Me.chkShowPassword = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtKonfirmasiPassword = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtPasswordBaru = New System.Windows.Forms.TextBox
        Me.btnGanti = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnKeluar
        '
        Me.btnKeluar.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnKeluar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnKeluar.Location = New System.Drawing.Point(198, 179)
        Me.btnKeluar.Margin = New System.Windows.Forms.Padding(2)
        Me.btnKeluar.Name = "btnKeluar"
        Me.btnKeluar.Size = New System.Drawing.Size(87, 27)
        Me.btnKeluar.TabIndex = 1
        Me.btnKeluar.TabStop = False
        Me.btnKeluar.Text = "&Exit"
        Me.btnKeluar.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Enabled = False
        Me.btnOK.Location = New System.Drawing.Point(92, 86)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(87, 27)
        Me.btnOK.TabIndex = 2
        Me.btnOK.TabStop = False
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 32)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "User Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 10)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "User Code"
        '
        'cboUser
        '
        Me.cboUser.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUser.FormattingEnabled = True
        Me.cboUser.Location = New System.Drawing.Point(131, 8)
        Me.cboUser.Margin = New System.Windows.Forms.Padding(2)
        Me.cboUser.Name = "cboUser"
        Me.cboUser.Size = New System.Drawing.Size(229, 21)
        Me.cboUser.TabIndex = 2
        '
        'txtNamaUser
        '
        Me.txtNamaUser.Location = New System.Drawing.Point(131, 30)
        Me.txtNamaUser.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNamaUser.Name = "txtNamaUser"
        Me.txtNamaUser.ReadOnly = True
        Me.txtNamaUser.Size = New System.Drawing.Size(229, 20)
        Me.txtNamaUser.TabIndex = 3
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnBatal)
        Me.GroupBox1.Controls.Add(Me.chkShowPassword)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtKonfirmasiPassword)
        Me.GroupBox1.Controls.Add(Me.btnOK)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtPasswordBaru)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 52)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(357, 120)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Password Detail"
        '
        'btnBatal
        '
        Me.btnBatal.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnBatal.Enabled = False
        Me.btnBatal.Location = New System.Drawing.Point(193, 86)
        Me.btnBatal.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBatal.Name = "btnBatal"
        Me.btnBatal.Size = New System.Drawing.Size(87, 27)
        Me.btnBatal.TabIndex = 3
        Me.btnBatal.TabStop = False
        Me.btnBatal.Text = "&Cancel"
        Me.btnBatal.UseVisualStyleBackColor = True
        '
        'chkShowPassword
        '
        Me.chkShowPassword.AutoSize = True
        Me.chkShowPassword.Enabled = False
        Me.chkShowPassword.Location = New System.Drawing.Point(10, 17)
        Me.chkShowPassword.Margin = New System.Windows.Forms.Padding(2)
        Me.chkShowPassword.Name = "chkShowPassword"
        Me.chkShowPassword.Size = New System.Drawing.Size(102, 17)
        Me.chkShowPassword.TabIndex = 4
        Me.chkShowPassword.Text = "Show Password"
        Me.chkShowPassword.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 44)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 34
        Me.Label1.Text = "New Password"
        '
        'txtKonfirmasiPassword
        '
        Me.txtKonfirmasiPassword.Enabled = False
        Me.txtKonfirmasiPassword.Location = New System.Drawing.Point(140, 64)
        Me.txtKonfirmasiPassword.Margin = New System.Windows.Forms.Padding(2)
        Me.txtKonfirmasiPassword.Name = "txtKonfirmasiPassword"
        Me.txtKonfirmasiPassword.Size = New System.Drawing.Size(213, 20)
        Me.txtKonfirmasiPassword.TabIndex = 1
        Me.txtKonfirmasiPassword.UseSystemPasswordChar = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 66)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(116, 13)
        Me.Label4.TabIndex = 33
        Me.Label4.Text = "Confirm New Password"
        '
        'txtPasswordBaru
        '
        Me.txtPasswordBaru.Enabled = False
        Me.txtPasswordBaru.Location = New System.Drawing.Point(140, 42)
        Me.txtPasswordBaru.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPasswordBaru.Name = "txtPasswordBaru"
        Me.txtPasswordBaru.Size = New System.Drawing.Size(213, 20)
        Me.txtPasswordBaru.TabIndex = 0
        Me.txtPasswordBaru.UseSystemPasswordChar = True
        '
        'btnGanti
        '
        Me.btnGanti.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnGanti.Location = New System.Drawing.Point(97, 179)
        Me.btnGanti.Margin = New System.Windows.Forms.Padding(2)
        Me.btnGanti.Name = "btnGanti"
        Me.btnGanti.Size = New System.Drawing.Size(87, 27)
        Me.btnGanti.TabIndex = 0
        Me.btnGanti.TabStop = False
        Me.btnGanti.Text = "&Reset"
        Me.btnGanti.UseVisualStyleBackColor = True
        '
        'FrmResetPassword
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnKeluar
        Me.ClientSize = New System.Drawing.Size(371, 210)
        Me.Controls.Add(Me.btnGanti)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cboUser)
        Me.Controls.Add(Me.btnKeluar)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtNamaUser)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmResetPassword"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Password Reset"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnKeluar As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboUser As System.Windows.Forms.ComboBox
    Friend WithEvents txtNamaUser As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnBatal As System.Windows.Forms.Button
    Friend WithEvents chkShowPassword As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtKonfirmasiPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtPasswordBaru As System.Windows.Forms.TextBox
    Friend WithEvents btnGanti As System.Windows.Forms.Button
End Class
