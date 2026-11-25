<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterHakAkses
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
        Me.GBMain = New System.Windows.Forms.GroupBox
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.txtDeskripsi = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtNamaHakAkses = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtKodeHakAkses = New System.Windows.Forms.TextBox
        Me.cmdKeluar = New System.Windows.Forms.Button
        Me.btnHapus = New System.Windows.Forms.Button
        Me.cmdCari = New System.Windows.Forms.Button
        Me.cmdEdt = New System.Windows.Forms.Button
        Me.cmdNew = New System.Windows.Forms.Button
        Me.GBMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'GBMain
        '
        Me.GBMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GBMain.Controls.Add(Me.cmdCancel)
        Me.GBMain.Controls.Add(Me.cmdOK)
        Me.GBMain.Controls.Add(Me.txtDeskripsi)
        Me.GBMain.Controls.Add(Me.Label3)
        Me.GBMain.Controls.Add(Me.TxtNamaHakAkses)
        Me.GBMain.Controls.Add(Me.Label1)
        Me.GBMain.Controls.Add(Me.TxtKodeHakAkses)
        Me.GBMain.Location = New System.Drawing.Point(9, 6)
        Me.GBMain.Margin = New System.Windows.Forms.Padding(4)
        Me.GBMain.Name = "GBMain"
        Me.GBMain.Padding = New System.Windows.Forms.Padding(4)
        Me.GBMain.Size = New System.Drawing.Size(393, 168)
        Me.GBMain.TabIndex = 5
        Me.GBMain.TabStop = False
        Me.GBMain.Text = "Data Hak Akses"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.Location = New System.Drawing.Point(200, 131)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(135, 30)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Bata&l"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cmdOK.Enabled = False
        Me.cmdOK.Location = New System.Drawing.Point(52, 131)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(135, 30)
        Me.cmdOK.TabIndex = 2
        Me.cmdOK.Text = "&Simpan"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'txtDeskripsi
        '
        Me.txtDeskripsi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDeskripsi.BackColor = System.Drawing.SystemColors.Window
        Me.txtDeskripsi.Enabled = False
        Me.txtDeskripsi.Location = New System.Drawing.Point(120, 60)
        Me.txtDeskripsi.MaxLength = 150
        Me.txtDeskripsi.Multiline = True
        Me.txtDeskripsi.Name = "txtDeskripsi"
        Me.txtDeskripsi.Size = New System.Drawing.Size(254, 65)
        Me.txtDeskripsi.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(24, 61)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Description"
        '
        'TxtNamaHakAkses
        '
        Me.TxtNamaHakAkses.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNamaHakAkses.BackColor = System.Drawing.SystemColors.Window
        Me.TxtNamaHakAkses.Enabled = False
        Me.TxtNamaHakAkses.Location = New System.Drawing.Point(120, 27)
        Me.TxtNamaHakAkses.MaxLength = 50
        Me.TxtNamaHakAkses.Name = "TxtNamaHakAkses"
        Me.TxtNamaHakAkses.Size = New System.Drawing.Size(254, 20)
        Me.TxtNamaHakAkses.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Nama Hak Akses"
        '
        'TxtKodeHakAkses
        '
        Me.TxtKodeHakAkses.BackColor = System.Drawing.SystemColors.Window
        Me.TxtKodeHakAkses.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtKodeHakAkses.Enabled = False
        Me.TxtKodeHakAkses.Location = New System.Drawing.Point(120, 27)
        Me.TxtKodeHakAkses.MaxLength = 0
        Me.TxtKodeHakAkses.Name = "TxtKodeHakAkses"
        Me.TxtKodeHakAkses.Size = New System.Drawing.Size(254, 20)
        Me.TxtKodeHakAkses.TabIndex = 81
        Me.TxtKodeHakAkses.Text = "01"
        Me.TxtKodeHakAkses.Visible = False
        '
        'cmdKeluar
        '
        Me.cmdKeluar.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cmdKeluar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdKeluar.Location = New System.Drawing.Point(209, 217)
        Me.cmdKeluar.Name = "cmdKeluar"
        Me.cmdKeluar.Size = New System.Drawing.Size(192, 30)
        Me.cmdKeluar.TabIndex = 4
        Me.cmdKeluar.Text = "&Keluar"
        Me.cmdKeluar.UseVisualStyleBackColor = True
        '
        'btnHapus
        '
        Me.btnHapus.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnHapus.Location = New System.Drawing.Point(275, 181)
        Me.btnHapus.Name = "btnHapus"
        Me.btnHapus.Size = New System.Drawing.Size(127, 30)
        Me.btnHapus.TabIndex = 2
        Me.btnHapus.Text = "&Hapus"
        Me.btnHapus.UseVisualStyleBackColor = True
        '
        'cmdCari
        '
        Me.cmdCari.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cmdCari.Location = New System.Drawing.Point(10, 217)
        Me.cmdCari.Name = "cmdCari"
        Me.cmdCari.Size = New System.Drawing.Size(192, 30)
        Me.cmdCari.TabIndex = 3
        Me.cmdCari.Text = "&Cari"
        Me.cmdCari.UseVisualStyleBackColor = True
        '
        'cmdEdt
        '
        Me.cmdEdt.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cmdEdt.Location = New System.Drawing.Point(142, 181)
        Me.cmdEdt.Name = "cmdEdt"
        Me.cmdEdt.Size = New System.Drawing.Size(127, 30)
        Me.cmdEdt.TabIndex = 1
        Me.cmdEdt.Text = "&Edit Hak Khusus"
        Me.cmdEdt.UseVisualStyleBackColor = True
        '
        'cmdNew
        '
        Me.cmdNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cmdNew.Location = New System.Drawing.Point(9, 181)
        Me.cmdNew.Name = "cmdNew"
        Me.cmdNew.Size = New System.Drawing.Size(127, 30)
        Me.cmdNew.TabIndex = 0
        Me.cmdNew.Text = "Hak Akses &Baru"
        Me.cmdNew.UseVisualStyleBackColor = True
        '
        'FrmMasterHakAkses
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdKeluar
        Me.ClientSize = New System.Drawing.Size(408, 262)
        Me.Controls.Add(Me.cmdKeluar)
        Me.Controls.Add(Me.btnHapus)
        Me.Controls.Add(Me.cmdCari)
        Me.Controls.Add(Me.cmdEdt)
        Me.Controls.Add(Me.cmdNew)
        Me.Controls.Add(Me.GBMain)
        Me.Name = "FrmMasterHakAkses"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Master Hak Akses"
        Me.GBMain.ResumeLayout(False)
        Me.GBMain.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GBMain As System.Windows.Forms.GroupBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents txtDeskripsi As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtNamaHakAkses As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdKeluar As System.Windows.Forms.Button
    Friend WithEvents btnHapus As System.Windows.Forms.Button
    Friend WithEvents cmdCari As System.Windows.Forms.Button
    Friend WithEvents cmdEdt As System.Windows.Forms.Button
    Friend WithEvents cmdNew As System.Windows.Forms.Button
    Friend WithEvents TxtKodeHakAkses As System.Windows.Forms.TextBox
End Class
