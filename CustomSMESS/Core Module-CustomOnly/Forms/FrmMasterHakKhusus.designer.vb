<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterHakKhusus
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
        Me.cmdKeluar = New System.Windows.Forms.Button
        Me.btnHapus = New System.Windows.Forms.Button
        Me.cmdCari = New System.Windows.Forms.Button
        Me.cmdEdt = New System.Windows.Forms.Button
        Me.cmdNew = New System.Windows.Forms.Button
        Me.GBMain = New System.Windows.Forms.GroupBox
        Me.TxtKodeHakKhusus = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.txtDeskripsi = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtNamaHakKhusus = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GBMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdKeluar
        '
        Me.cmdKeluar.Location = New System.Drawing.Point(205, 222)
        Me.cmdKeluar.Name = "cmdKeluar"
        Me.cmdKeluar.Size = New System.Drawing.Size(192, 30)
        Me.cmdKeluar.TabIndex = 4
        Me.cmdKeluar.Text = "&Keluar"
        Me.cmdKeluar.UseVisualStyleBackColor = True
        '
        'btnHapus
        '
        Me.btnHapus.Location = New System.Drawing.Point(271, 186)
        Me.btnHapus.Name = "btnHapus"
        Me.btnHapus.Size = New System.Drawing.Size(127, 30)
        Me.btnHapus.TabIndex = 2
        Me.btnHapus.Text = "&Hapus"
        Me.btnHapus.UseVisualStyleBackColor = True
        '
        'cmdCari
        '
        Me.cmdCari.Location = New System.Drawing.Point(6, 222)
        Me.cmdCari.Name = "cmdCari"
        Me.cmdCari.Size = New System.Drawing.Size(192, 30)
        Me.cmdCari.TabIndex = 3
        Me.cmdCari.Text = "&Cari"
        Me.cmdCari.UseVisualStyleBackColor = True
        '
        'cmdEdt
        '
        Me.cmdEdt.Location = New System.Drawing.Point(138, 186)
        Me.cmdEdt.Name = "cmdEdt"
        Me.cmdEdt.Size = New System.Drawing.Size(127, 30)
        Me.cmdEdt.TabIndex = 1
        Me.cmdEdt.Text = "&Edit Hak Khusus"
        Me.cmdEdt.UseVisualStyleBackColor = True
        '
        'cmdNew
        '
        Me.cmdNew.Location = New System.Drawing.Point(5, 186)
        Me.cmdNew.Name = "cmdNew"
        Me.cmdNew.Size = New System.Drawing.Size(127, 30)
        Me.cmdNew.TabIndex = 0
        Me.cmdNew.Text = "Hak Khusus &Baru"
        Me.cmdNew.UseVisualStyleBackColor = True
        '
        'GBMain
        '
        Me.GBMain.Controls.Add(Me.TxtKodeHakKhusus)
        Me.GBMain.Controls.Add(Me.Label2)
        Me.GBMain.Controls.Add(Me.cmdCancel)
        Me.GBMain.Controls.Add(Me.cmdOK)
        Me.GBMain.Controls.Add(Me.txtDeskripsi)
        Me.GBMain.Controls.Add(Me.Label3)
        Me.GBMain.Controls.Add(Me.TxtNamaHakKhusus)
        Me.GBMain.Controls.Add(Me.Label1)
        Me.GBMain.Location = New System.Drawing.Point(7, 7)
        Me.GBMain.Margin = New System.Windows.Forms.Padding(4)
        Me.GBMain.Name = "GBMain"
        Me.GBMain.Padding = New System.Windows.Forms.Padding(4)
        Me.GBMain.Size = New System.Drawing.Size(393, 172)
        Me.GBMain.TabIndex = 5
        Me.GBMain.TabStop = False
        Me.GBMain.Text = "Data Hak Khusus"
        '
        'TxtKodeHakKhusus
        '
        Me.TxtKodeHakKhusus.BackColor = System.Drawing.SystemColors.Window
        Me.TxtKodeHakKhusus.Enabled = False
        Me.TxtKodeHakKhusus.Location = New System.Drawing.Point(170, 27)
        Me.TxtKodeHakKhusus.MaxLength = 30
        Me.TxtKodeHakKhusus.Name = "TxtKodeHakKhusus"
        Me.TxtKodeHakKhusus.Size = New System.Drawing.Size(200, 20)
        Me.TxtKodeHakKhusus.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 30)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 13)
        Me.Label2.TabIndex = 78
        Me.Label2.Text = "Kode Hak Khusus"
        '
        'cmdCancel
        '
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.Location = New System.Drawing.Point(198, 135)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(135, 30)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "Bata&l"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Enabled = False
        Me.cmdOK.Location = New System.Drawing.Point(50, 135)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(135, 30)
        Me.cmdOK.TabIndex = 3
        Me.cmdOK.Text = "&Simpan"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'txtDeskripsi
        '
        Me.txtDeskripsi.BackColor = System.Drawing.SystemColors.Window
        Me.txtDeskripsi.Enabled = False
        Me.txtDeskripsi.Location = New System.Drawing.Point(170, 93)
        Me.txtDeskripsi.MaxLength = 50
        Me.txtDeskripsi.Name = "txtDeskripsi"
        Me.txtDeskripsi.Size = New System.Drawing.Size(200, 20)
        Me.txtDeskripsi.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(20, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Description"
        '
        'TxtNamaHakKhusus
        '
        Me.TxtNamaHakKhusus.BackColor = System.Drawing.SystemColors.Window
        Me.TxtNamaHakKhusus.Enabled = False
        Me.TxtNamaHakKhusus.Location = New System.Drawing.Point(170, 61)
        Me.TxtNamaHakKhusus.MaxLength = 30
        Me.TxtNamaHakKhusus.Name = "TxtNamaHakKhusus"
        Me.TxtNamaHakKhusus.Size = New System.Drawing.Size(200, 20)
        Me.TxtNamaHakKhusus.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Nama Hak Khusus"
        '
        'FrmMasterHakKhusus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(407, 260)
        Me.Controls.Add(Me.cmdKeluar)
        Me.Controls.Add(Me.btnHapus)
        Me.Controls.Add(Me.cmdCari)
        Me.Controls.Add(Me.cmdEdt)
        Me.Controls.Add(Me.cmdNew)
        Me.Controls.Add(Me.GBMain)
        Me.Name = "FrmMasterHakKhusus"
        Me.Text = "FrmMasterHakKhusus"
        Me.GBMain.ResumeLayout(False)
        Me.GBMain.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdKeluar As System.Windows.Forms.Button
    Friend WithEvents btnHapus As System.Windows.Forms.Button
    Friend WithEvents cmdCari As System.Windows.Forms.Button
    Friend WithEvents cmdEdt As System.Windows.Forms.Button
    Friend WithEvents cmdNew As System.Windows.Forms.Button
    Friend WithEvents GBMain As System.Windows.Forms.GroupBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents txtDeskripsi As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtNamaHakKhusus As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtKodeHakKhusus As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
