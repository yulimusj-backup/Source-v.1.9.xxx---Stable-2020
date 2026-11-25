<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmManajemenHakKhusus
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
        Me.BtnEdit = New System.Windows.Forms.Button
        Me.TxtNamaUser = New System.Windows.Forms.TextBox
        Me.CBOkodeuser = New System.Windows.Forms.ComboBox
        Me.BtnKeluar = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnBatal = New System.Windows.Forms.Button
        Me.BtnSimpan = New System.Windows.Forms.Button
        Me.btnKiriSemua = New System.Windows.Forms.Button
        Me.btnKananSemua = New System.Windows.Forms.Button
        Me.btnKiri = New System.Windows.Forms.Button
        Me.btnKanan = New System.Windows.Forms.Button
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.ListFormAktif = New System.Windows.Forms.ListBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.ListSemuaForm = New System.Windows.Forms.ListBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnEdit
        '
        Me.BtnEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.BtnEdit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.BtnEdit.Location = New System.Drawing.Point(430, 12)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(69, 21)
        Me.BtnEdit.TabIndex = 17
        Me.BtnEdit.Text = "&Edit"
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'TxtNamaUser
        '
        Me.TxtNamaUser.Enabled = False
        Me.TxtNamaUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.TxtNamaUser.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.TxtNamaUser.Location = New System.Drawing.Point(109, 43)
        Me.TxtNamaUser.Name = "TxtNamaUser"
        Me.TxtNamaUser.Size = New System.Drawing.Size(315, 20)
        Me.TxtNamaUser.TabIndex = 15
        '
        'CBOkodeuser
        '
        Me.CBOkodeuser.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBOkodeuser.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.CBOkodeuser.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CBOkodeuser.FormattingEnabled = True
        Me.CBOkodeuser.Location = New System.Drawing.Point(109, 12)
        Me.CBOkodeuser.Name = "CBOkodeuser"
        Me.CBOkodeuser.Size = New System.Drawing.Size(315, 21)
        Me.CBOkodeuser.TabIndex = 14
        '
        'BtnKeluar
        '
        Me.BtnKeluar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnKeluar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.BtnKeluar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.BtnKeluar.Location = New System.Drawing.Point(430, 44)
        Me.BtnKeluar.Name = "BtnKeluar"
        Me.BtnKeluar.Size = New System.Drawing.Size(69, 19)
        Me.BtnKeluar.TabIndex = 16
        Me.BtnKeluar.Text = "&Keluar"
        Me.BtnKeluar.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Label2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.Location = New System.Drawing.Point(11, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Nama User"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(11, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Kode User"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.BtnBatal)
        Me.GroupBox1.Controls.Add(Me.BtnSimpan)
        Me.GroupBox1.Controls.Add(Me.btnKiriSemua)
        Me.GroupBox1.Controls.Add(Me.btnKananSemua)
        Me.GroupBox1.Controls.Add(Me.btnKiri)
        Me.GroupBox1.Controls.Add(Me.btnKanan)
        Me.GroupBox1.Controls.Add(Me.GroupBox6)
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GroupBox1.Location = New System.Drawing.Point(2, 80)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(924, 358)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Management Hak Khusus"
        '
        'BtnBatal
        '
        Me.BtnBatal.Enabled = False
        Me.BtnBatal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.BtnBatal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.BtnBatal.Location = New System.Drawing.Point(571, 326)
        Me.BtnBatal.Name = "BtnBatal"
        Me.BtnBatal.Size = New System.Drawing.Size(167, 26)
        Me.BtnBatal.TabIndex = 58
        Me.BtnBatal.Text = "&Batal"
        Me.BtnBatal.UseVisualStyleBackColor = True
        '
        'BtnSimpan
        '
        Me.BtnSimpan.Enabled = False
        Me.BtnSimpan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.BtnSimpan.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.BtnSimpan.Location = New System.Drawing.Point(186, 326)
        Me.BtnSimpan.Name = "BtnSimpan"
        Me.BtnSimpan.Size = New System.Drawing.Size(167, 26)
        Me.BtnSimpan.TabIndex = 57
        Me.BtnSimpan.Text = "&Simpan"
        Me.BtnSimpan.UseVisualStyleBackColor = True
        '
        'btnKiriSemua
        '
        Me.btnKiriSemua.Enabled = False
        Me.btnKiriSemua.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnKiriSemua.ForeColor = System.Drawing.Color.Black
        Me.btnKiriSemua.Location = New System.Drawing.Point(428, 240)
        Me.btnKiriSemua.Name = "btnKiriSemua"
        Me.btnKiriSemua.Size = New System.Drawing.Size(69, 43)
        Me.btnKiriSemua.TabIndex = 56
        Me.btnKiriSemua.Text = "<<"
        Me.btnKiriSemua.UseVisualStyleBackColor = True
        '
        'btnKananSemua
        '
        Me.btnKananSemua.Enabled = False
        Me.btnKananSemua.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnKananSemua.ForeColor = System.Drawing.Color.Black
        Me.btnKananSemua.Location = New System.Drawing.Point(428, 183)
        Me.btnKananSemua.Name = "btnKananSemua"
        Me.btnKananSemua.Size = New System.Drawing.Size(69, 43)
        Me.btnKananSemua.TabIndex = 55
        Me.btnKananSemua.Text = ">>"
        Me.btnKananSemua.UseVisualStyleBackColor = True
        '
        'btnKiri
        '
        Me.btnKiri.Enabled = False
        Me.btnKiri.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnKiri.ForeColor = System.Drawing.Color.Black
        Me.btnKiri.Location = New System.Drawing.Point(428, 125)
        Me.btnKiri.Name = "btnKiri"
        Me.btnKiri.Size = New System.Drawing.Size(69, 43)
        Me.btnKiri.TabIndex = 54
        Me.btnKiri.Text = "<"
        Me.btnKiri.UseVisualStyleBackColor = True
        '
        'btnKanan
        '
        Me.btnKanan.Enabled = False
        Me.btnKanan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnKanan.ForeColor = System.Drawing.Color.Black
        Me.btnKanan.Location = New System.Drawing.Point(428, 69)
        Me.btnKanan.Name = "btnKanan"
        Me.btnKanan.Size = New System.Drawing.Size(69, 43)
        Me.btnKanan.TabIndex = 53
        Me.btnKanan.Text = ">"
        Me.btnKanan.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox6.Controls.Add(Me.ListFormAktif)
        Me.GroupBox6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.GroupBox6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GroupBox6.Location = New System.Drawing.Point(503, 26)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(406, 294)
        Me.GroupBox6.TabIndex = 50
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Daftar hak khusus yang diaktifkan"
        '
        'ListFormAktif
        '
        Me.ListFormAktif.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.ListFormAktif.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ListFormAktif.FormattingEnabled = True
        Me.ListFormAktif.Location = New System.Drawing.Point(6, 25)
        Me.ListFormAktif.Name = "ListFormAktif"
        Me.ListFormAktif.Size = New System.Drawing.Size(394, 264)
        Me.ListFormAktif.TabIndex = 2
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox5.Controls.Add(Me.ListSemuaForm)
        Me.GroupBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.GroupBox5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GroupBox5.Location = New System.Drawing.Point(16, 26)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(406, 294)
        Me.GroupBox5.TabIndex = 49
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Daftar semua hak khusus"
        '
        'ListSemuaForm
        '
        Me.ListSemuaForm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.ListSemuaForm.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ListSemuaForm.FormattingEnabled = True
        Me.ListSemuaForm.Location = New System.Drawing.Point(6, 25)
        Me.ListSemuaForm.Name = "ListSemuaForm"
        Me.ListSemuaForm.Size = New System.Drawing.Size(394, 264)
        Me.ListSemuaForm.TabIndex = 2
        '
        'FrmManajemenHakKhusus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(938, 450)
        Me.Controls.Add(Me.BtnEdit)
        Me.Controls.Add(Me.TxtNamaUser)
        Me.Controls.Add(Me.CBOkodeuser)
        Me.Controls.Add(Me.BtnKeluar)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(954, 486)
        Me.Name = "FrmManajemenHakKhusus"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Management Hak Khusus"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents TxtNamaUser As System.Windows.Forms.TextBox
    Friend WithEvents CBOkodeuser As System.Windows.Forms.ComboBox
    Friend WithEvents BtnKeluar As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnBatal As System.Windows.Forms.Button
    Friend WithEvents BtnSimpan As System.Windows.Forms.Button
    Friend WithEvents btnKiriSemua As System.Windows.Forms.Button
    Friend WithEvents btnKananSemua As System.Windows.Forms.Button
    Friend WithEvents btnKiri As System.Windows.Forms.Button
    Friend WithEvents btnKanan As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents ListFormAktif As System.Windows.Forms.ListBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents ListSemuaForm As System.Windows.Forms.ListBox
End Class
