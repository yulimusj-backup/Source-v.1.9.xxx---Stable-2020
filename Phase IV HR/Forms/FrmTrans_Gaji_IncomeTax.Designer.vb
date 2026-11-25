<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_IncomeTax
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.btnExit = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtAlamat2 = New System.Windows.Forms.TextBox
        Me.TxtAlamat1 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtNama = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtNIK = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtNPWP = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.ChkKaryawanAsing = New System.Windows.Forms.CheckBox
        Me.TxtKodeKaryawanAsing = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtNamaJabatan = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.CboStatusPTKP = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtJenisKelamin = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtNomor = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.TxtNamaPemotong = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.TxtNPWPPemotong = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.DtpTanggalTransaksi = New System.Windows.Forms.DateTimePicker
        Me.Label14 = New System.Windows.Forms.Label
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUraian = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 30
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColUraian, Me.ColAmount})
        Me.dgvItem.Location = New System.Drawing.Point(7, 15)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(975, 310)
        Me.dgvItem.TabIndex = 2
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(452, 440)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.SplitContainer1)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 49)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(974, 150)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "A. Identitas Penerima Penghasilan Yang Dipotong"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtAlamat2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtAlamat1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtNama)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtNIK)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtNPWP)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.ChkKaryawanAsing)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtKodeKaryawanAsing)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtNamaJabatan)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.CboStatusPTKP)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtJenisKelamin)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Size = New System.Drawing.Size(968, 131)
        Me.SplitContainer1.SplitterDistance = 488
        Me.SplitContainer1.TabIndex = 0
        '
        'TxtAlamat2
        '
        Me.TxtAlamat2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAlamat2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAlamat2.Location = New System.Drawing.Point(83, 107)
        Me.TxtAlamat2.MaxLength = 50
        Me.TxtAlamat2.Name = "TxtAlamat2"
        Me.TxtAlamat2.ReadOnly = True
        Me.TxtAlamat2.Size = New System.Drawing.Size(406, 20)
        Me.TxtAlamat2.TabIndex = 39
        '
        'TxtAlamat1
        '
        Me.TxtAlamat1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAlamat1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAlamat1.Location = New System.Drawing.Point(83, 81)
        Me.TxtAlamat1.MaxLength = 50
        Me.TxtAlamat1.Name = "TxtAlamat1"
        Me.TxtAlamat1.ReadOnly = True
        Me.TxtAlamat1.Size = New System.Drawing.Size(406, 20)
        Me.TxtAlamat1.TabIndex = 37
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(4, 81)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 13)
        Me.Label5.TabIndex = 38
        Me.Label5.Tag = ""
        Me.Label5.Text = "Alamat"
        '
        'TxtNama
        '
        Me.TxtNama.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNama.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNama.Location = New System.Drawing.Point(83, 55)
        Me.TxtNama.MaxLength = 50
        Me.TxtNama.Name = "TxtNama"
        Me.TxtNama.ReadOnly = True
        Me.TxtNama.Size = New System.Drawing.Size(406, 20)
        Me.TxtNama.TabIndex = 35
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 55)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(35, 13)
        Me.Label4.TabIndex = 36
        Me.Label4.Tag = ""
        Me.Label4.Text = "Nama"
        '
        'TxtNIK
        '
        Me.TxtNIK.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNIK.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNIK.Location = New System.Drawing.Point(83, 29)
        Me.TxtNIK.MaxLength = 50
        Me.TxtNIK.Name = "TxtNIK"
        Me.TxtNIK.ReadOnly = True
        Me.TxtNIK.Size = New System.Drawing.Size(406, 20)
        Me.TxtNIK.TabIndex = 33
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(25, 13)
        Me.Label3.TabIndex = 34
        Me.Label3.Tag = ""
        Me.Label3.Text = "NIK"
        '
        'TxtNPWP
        '
        Me.TxtNPWP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNPWP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNPWP.Location = New System.Drawing.Point(82, 3)
        Me.TxtNPWP.MaxLength = 50
        Me.TxtNPWP.Name = "TxtNPWP"
        Me.TxtNPWP.ReadOnly = True
        Me.TxtNPWP.Size = New System.Drawing.Size(406, 20)
        Me.TxtNPWP.TabIndex = 31
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 3)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 13)
        Me.Label2.TabIndex = 32
        Me.Label2.Tag = ""
        Me.Label2.Text = "NPWP"
        '
        'ChkKaryawanAsing
        '
        Me.ChkKaryawanAsing.AutoSize = True
        Me.ChkKaryawanAsing.Enabled = False
        Me.ChkKaryawanAsing.Location = New System.Drawing.Point(108, 81)
        Me.ChkKaryawanAsing.Name = "ChkKaryawanAsing"
        Me.ChkKaryawanAsing.Size = New System.Drawing.Size(39, 17)
        Me.ChkKaryawanAsing.TabIndex = 51
        Me.ChkKaryawanAsing.Text = "Ya"
        Me.ChkKaryawanAsing.UseVisualStyleBackColor = True
        '
        'TxtKodeKaryawanAsing
        '
        Me.TxtKodeKaryawanAsing.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtKodeKaryawanAsing.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtKodeKaryawanAsing.Location = New System.Drawing.Point(108, 107)
        Me.TxtKodeKaryawanAsing.MaxLength = 50
        Me.TxtKodeKaryawanAsing.Name = "TxtKodeKaryawanAsing"
        Me.TxtKodeKaryawanAsing.ReadOnly = True
        Me.TxtKodeKaryawanAsing.Size = New System.Drawing.Size(366, 20)
        Me.TxtKodeKaryawanAsing.TabIndex = 49
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 107)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(99, 13)
        Me.Label11.TabIndex = 50
        Me.Label11.Tag = ""
        Me.Label11.Text = "Kode Negara Asing"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(4, 82)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(83, 13)
        Me.Label10.TabIndex = 47
        Me.Label10.Tag = ""
        Me.Label10.Text = "Karyawan Asing"
        '
        'TxtNamaJabatan
        '
        Me.TxtNamaJabatan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNamaJabatan.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNamaJabatan.Location = New System.Drawing.Point(108, 52)
        Me.TxtNamaJabatan.MaxLength = 50
        Me.TxtNamaJabatan.Name = "TxtNamaJabatan"
        Me.TxtNamaJabatan.ReadOnly = True
        Me.TxtNamaJabatan.Size = New System.Drawing.Size(366, 20)
        Me.TxtNamaJabatan.TabIndex = 45
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 55)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(76, 13)
        Me.Label9.TabIndex = 46
        Me.Label9.Tag = ""
        Me.Label9.Text = "Nama Jabatan"
        '
        'CboStatusPTKP
        '
        Me.CboStatusPTKP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboStatusPTKP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboStatusPTKP.Enabled = False
        Me.CboStatusPTKP.FormattingEnabled = True
        Me.CboStatusPTKP.Items.AddRange(New Object() {"TK0", "TK1", "TK2", "TK3", "K0", "K1", "K2", "K3"})
        Me.CboStatusPTKP.Location = New System.Drawing.Point(108, 25)
        Me.CboStatusPTKP.Name = "CboStatusPTKP"
        Me.CboStatusPTKP.Size = New System.Drawing.Size(366, 21)
        Me.CboStatusPTKP.TabIndex = 44
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 26)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(68, 13)
        Me.Label8.TabIndex = 43
        Me.Label8.Tag = ""
        Me.Label8.Text = "Status PTKP"
        '
        'TxtJenisKelamin
        '
        Me.TxtJenisKelamin.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtJenisKelamin.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtJenisKelamin.Location = New System.Drawing.Point(108, 3)
        Me.TxtJenisKelamin.MaxLength = 50
        Me.TxtJenisKelamin.Name = "TxtJenisKelamin"
        Me.TxtJenisKelamin.ReadOnly = True
        Me.TxtJenisKelamin.Size = New System.Drawing.Size(366, 20)
        Me.TxtJenisKelamin.TabIndex = 41
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(71, 13)
        Me.Label7.TabIndex = 42
        Me.Label7.Tag = ""
        Me.Label7.Text = "Jenis Kelamin"
        '
        'TxtNomor
        '
        Me.TxtNomor.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNomor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNomor.Location = New System.Drawing.Point(85, 16)
        Me.TxtNomor.MaxLength = 50
        Me.TxtNomor.Name = "TxtNomor"
        Me.TxtNomor.ReadOnly = True
        Me.TxtNomor.Size = New System.Drawing.Size(881, 20)
        Me.TxtNomor.TabIndex = 29
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 30
        Me.Label1.Tag = ""
        Me.Label1.Text = "Nomor"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Controls.Add(Me.dgvItem)
        Me.GroupBox2.Location = New System.Drawing.Point(0, 202)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(981, 418)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "B. Rincian Penghasilan Dan Perhitungan PPh Pasal 21"
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.SplitContainer2)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 314)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(975, 98)
        Me.GroupBox3.TabIndex = 10
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "C. Identitas Pemotong"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtNamaPemotong)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtNPWPPemotong)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label13)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.DtpTanggalTransaksi)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label14)
        Me.SplitContainer2.Size = New System.Drawing.Size(969, 79)
        Me.SplitContainer2.SplitterDistance = 490
        Me.SplitContainer2.TabIndex = 0
        '
        'TxtNamaPemotong
        '
        Me.TxtNamaPemotong.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNamaPemotong.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNamaPemotong.Location = New System.Drawing.Point(82, 42)
        Me.TxtNamaPemotong.MaxLength = 50
        Me.TxtNamaPemotong.Name = "TxtNamaPemotong"
        Me.TxtNamaPemotong.ReadOnly = True
        Me.TxtNamaPemotong.Size = New System.Drawing.Size(403, 20)
        Me.TxtNamaPemotong.TabIndex = 35
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(6, 42)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(35, 13)
        Me.Label12.TabIndex = 36
        Me.Label12.Tag = ""
        Me.Label12.Text = "Nama"
        '
        'TxtNPWPPemotong
        '
        Me.TxtNPWPPemotong.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNPWPPemotong.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNPWPPemotong.Location = New System.Drawing.Point(82, 16)
        Me.TxtNPWPPemotong.MaxLength = 50
        Me.TxtNPWPPemotong.Name = "TxtNPWPPemotong"
        Me.TxtNPWPPemotong.ReadOnly = True
        Me.TxtNPWPPemotong.Size = New System.Drawing.Size(403, 20)
        Me.TxtNPWPPemotong.TabIndex = 33
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 16)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(40, 13)
        Me.Label13.TabIndex = 34
        Me.Label13.Tag = ""
        Me.Label13.Text = "NPWP"
        '
        'DtpTanggalTransaksi
        '
        Me.DtpTanggalTransaksi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTanggalTransaksi.CustomFormat = "dd-MMM-yyyy"
        Me.DtpTanggalTransaksi.Enabled = False
        Me.DtpTanggalTransaksi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTanggalTransaksi.Location = New System.Drawing.Point(106, 16)
        Me.DtpTanggalTransaksi.Name = "DtpTanggalTransaksi"
        Me.DtpTanggalTransaksi.Size = New System.Drawing.Size(366, 20)
        Me.DtpTanggalTransaksi.TabIndex = 36
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(3, 19)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(95, 13)
        Me.Label14.TabIndex = 35
        Me.Label14.Tag = ""
        Me.Label14.Text = "Tanggal Transaksi"
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Controls.Add(Me.TxtNomor)
        Me.GroupBox4.Location = New System.Drawing.Point(6, 7)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(974, 39)
        Me.GroupBox4.TabIndex = 10
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Formulir 1721 - A1"
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColUraian
        '
        Me.ColUraian.HeaderText = "Uraian"
        Me.ColUraian.Name = "ColUraian"
        Me.ColUraian.ReadOnly = True
        Me.ColUraian.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUraian.Width = 44
        '
        'ColAmount
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.ReadOnly = True
        Me.ColAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAmount.Width = 49
        '
        'FrmTrans_Gaji_IncomeTax
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(984, 632)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnExit)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmTrans_Gaji_IncomeTax"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Income Tax"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtNomor As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtAlamat2 As System.Windows.Forms.TextBox
    Friend WithEvents TxtAlamat1 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtNama As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtNIK As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtNPWP As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtJenisKelamin As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtNamaJabatan As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CboStatusPTKP As System.Windows.Forms.ComboBox
    Friend WithEvents TxtKodeKaryawanAsing As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtNamaPemotong As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtNPWPPemotong As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents DtpTanggalTransaksi As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkKaryawanAsing As System.Windows.Forms.CheckBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUraian As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
