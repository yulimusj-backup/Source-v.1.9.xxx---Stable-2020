<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_BeginingBalanceTax
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle17 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle18 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle19 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle20 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label6 = New System.Windows.Forms.Label
        Me.LblDetailSetting = New System.Windows.Forms.LinkLabel
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.dtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtDocumentNo = New System.Windows.Forms.TextBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColEmp_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKodeKaryawan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNamaKaryawan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColNoNPWP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDepartement = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPosition = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGaji = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTunjanganPPH = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTunjanganLainnya = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHonor = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPremi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNatura = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTantiem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPenghasilBruto = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBiayaJabatan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIuranPensiun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJumlahPengurangan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJumlahPenghasilanNetto = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJumlahPenghasilNettoMasaPajakSebelum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJumlahPenghasilanNettoUntukPPH = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPTKP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPenghasilanKenaPajak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPPhPs21AtasPKP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPPHps21TelahDipotong = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CollPPHTerutang = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPPhTelahDipotong = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryKaryawan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(2, 6)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 145
        Me.Label6.Text = "Remarks"
        '
        'LblDetailSetting
        '
        Me.LblDetailSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblDetailSetting.AutoSize = True
        Me.LblDetailSetting.Location = New System.Drawing.Point(4, 433)
        Me.LblDetailSetting.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblDetailSetting.Name = "LblDetailSetting"
        Me.LblDetailSetting.Size = New System.Drawing.Size(73, 13)
        Me.LblDetailSetting.TabIndex = 37
        Me.LblDetailSetting.TabStop = True
        Me.LblDetailSetting.Text = "Detail settings"
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(638, 427)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(77, 20)
        Me.TxtReadOnly.TabIndex = 35
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'dtpPeriod
        '
        Me.dtpPeriod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpPeriod.CustomFormat = "MMM yyyy"
        Me.dtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriod.Location = New System.Drawing.Point(80, 28)
        Me.dtpPeriod.Name = "dtpPeriod"
        Me.dtpPeriod.ShowUpDown = True
        Me.dtpPeriod.Size = New System.Drawing.Size(268, 20)
        Me.dtpPeriod.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 13)
        Me.Label2.TabIndex = 50
        Me.Label2.Text = "Period"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 11)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpPeriod)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDocumentNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Size = New System.Drawing.Size(712, 51)
        Me.SplitContainer1.SplitterDistance = 350
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'DtpDate
        '
        Me.DtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(229, 3)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.Size = New System.Drawing.Size(117, 20)
        Me.DtpDate.TabIndex = 1
        '
        'TxtDocumentNo
        '
        Me.TxtDocumentNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDocumentNo.Location = New System.Drawing.Point(80, 3)
        Me.TxtDocumentNo.MaxLength = 50
        Me.TxtDocumentNo.Name = "TxtDocumentNo"
        Me.TxtDocumentNo.ReadOnly = True
        Me.TxtDocumentNo.Size = New System.Drawing.Size(144, 20)
        Me.TxtDocumentNo.TabIndex = 0
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(6, 3)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimary.TabIndex = 0
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Document No"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(56, 3)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(303, 43)
        Me.TxtRemarks.TabIndex = 0
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 493)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(748, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 95
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(183, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(183, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(183, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(183, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(213, 465)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(472, 465)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 26)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 30
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColEmp_No, Me.ColKodeKaryawan, Me.ColNamaKaryawan, Me.ColStatus, Me.ColNoNPWP, Me.ColNIK, Me.ColDepartement, Me.ColPosition, Me.ColGaji, Me.ColTunjanganPPH, Me.ColTunjanganLainnya, Me.ColHonor, Me.ColPremi, Me.ColNatura, Me.ColTantiem, Me.ColPenghasilBruto, Me.ColBiayaJabatan, Me.ColIuranPensiun, Me.ColJumlahPengurangan, Me.ColJumlahPenghasilanNetto, Me.ColJumlahPenghasilNettoMasaPajakSebelum, Me.ColJumlahPenghasilanNettoUntukPPH, Me.ColPTKP, Me.ColPenghasilanKenaPajak, Me.ColPPhPs21AtasPKP, Me.ColPPHps21TelahDipotong, Me.CollPPHTerutang, Me.ColPPhTelahDipotong, Me.ColRemarks, Me.ColPrimaryKaryawan})
        Me.dgvItem.Location = New System.Drawing.Point(6, 67)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(712, 355)
        Me.dgvItem.TabIndex = 1
        '
        'ColEmp_No
        '
        Me.ColEmp_No.HeaderText = "No"
        Me.ColEmp_No.Name = "ColEmp_No"
        Me.ColEmp_No.Width = 46
        '
        'ColKodeKaryawan
        '
        Me.ColKodeKaryawan.HeaderText = "Employee Code"
        Me.ColKodeKaryawan.Name = "ColKodeKaryawan"
        Me.ColKodeKaryawan.Width = 106
        '
        'ColNamaKaryawan
        '
        Me.ColNamaKaryawan.HeaderText = "Name"
        Me.ColNamaKaryawan.Name = "ColNamaKaryawan"
        Me.ColNamaKaryawan.Width = 60
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Category"
        Me.ColStatus.Items.AddRange(New Object() {"TK0", "TK1", "TK2", "TK3", "K0", "K1", "K2", "K3"})
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.Width = 55
        '
        'ColNoNPWP
        '
        Me.ColNoNPWP.HeaderText = "No NPWP"
        Me.ColNoNPWP.Name = "ColNoNPWP"
        Me.ColNoNPWP.Width = 82
        '
        'ColNIK
        '
        Me.ColNIK.HeaderText = "NIK"
        Me.ColNIK.Name = "ColNIK"
        Me.ColNIK.Width = 50
        '
        'ColDepartement
        '
        Me.ColDepartement.HeaderText = "Departement"
        Me.ColDepartement.Name = "ColDepartement"
        Me.ColDepartement.Width = 93
        '
        'ColPosition
        '
        Me.ColPosition.HeaderText = "Position"
        Me.ColPosition.Name = "ColPosition"
        Me.ColPosition.Width = 69
        '
        'ColGaji
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColGaji.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColGaji.HeaderText = "Gaji/Pensiunan Atau THT/JHT"
        Me.ColGaji.Name = "ColGaji"
        Me.ColGaji.Width = 180
        '
        'ColTunjanganPPH
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTunjanganPPH.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColTunjanganPPH.HeaderText = "Tunjangan PPh"
        Me.ColTunjanganPPH.Name = "ColTunjanganPPH"
        Me.ColTunjanganPPH.Width = 106
        '
        'ColTunjanganLainnya
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTunjanganLainnya.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColTunjanganLainnya.HeaderText = "Tunjangan Lainnya. Uang lembur dan sebagainya"
        Me.ColTunjanganLainnya.Name = "ColTunjanganLainnya"
        Me.ColTunjanganLainnya.Width = 267
        '
        'ColHonor
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColHonor.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColHonor.HeaderText = "Honorarium dan Imbalan Sejenisnya"
        Me.ColHonor.Name = "ColHonor"
        Me.ColHonor.Width = 201
        '
        'ColPremi
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPremi.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColPremi.HeaderText = "Premi Asuransi yang Dibayarkan Pemberi Kerja"
        Me.ColPremi.Name = "ColPremi"
        Me.ColPremi.Width = 252
        '
        'ColNatura
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNatura.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColNatura.HeaderText = "Penerimaan Dalam Bentuk Natura dan Kenikmatan Lainnya yang Dikenakan Pemotongan P" & _
            "Ph Pasal 21"
        Me.ColNatura.Name = "ColNatura"
        Me.ColNatura.Width = 524
        '
        'ColTantiem
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTantiem.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColTantiem.HeaderText = "Tantiem, Bonus, Gratifikasi, Jasa Produksi dan THR"
        Me.ColTantiem.Name = "ColTantiem"
        Me.ColTantiem.Width = 277
        '
        'ColPenghasilBruto
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPenghasilBruto.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColPenghasilBruto.HeaderText = "Jumlah Penghasilan Bruto"
        Me.ColPenghasilBruto.Name = "ColPenghasilBruto"
        Me.ColPenghasilBruto.Width = 154
        '
        'ColBiayaJabatan
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBiayaJabatan.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColBiayaJabatan.HeaderText = "Biaya Jabatan"
        Me.ColBiayaJabatan.Name = "ColBiayaJabatan"
        Me.ColBiayaJabatan.Width = 99
        '
        'ColIuranPensiun
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColIuranPensiun.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColIuranPensiun.HeaderText = "Iuran Pensiun atau Iuran JHT/THT"
        Me.ColIuranPensiun.Name = "ColIuranPensiun"
        Me.ColIuranPensiun.Width = 198
        '
        'ColJumlahPengurangan
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColJumlahPengurangan.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColJumlahPengurangan.HeaderText = "Jumlah Pengurangan"
        Me.ColJumlahPengurangan.Name = "ColJumlahPengurangan"
        Me.ColJumlahPengurangan.Width = 132
        '
        'ColJumlahPenghasilanNetto
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColJumlahPenghasilanNetto.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColJumlahPenghasilanNetto.HeaderText = "Jumlah Penghasilan Netto"
        Me.ColJumlahPenghasilanNetto.Name = "ColJumlahPenghasilanNetto"
        Me.ColJumlahPenghasilanNetto.Width = 155
        '
        'ColJumlahPenghasilNettoMasaPajakSebelum
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColJumlahPenghasilNettoMasaPajakSebelum.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColJumlahPenghasilNettoMasaPajakSebelum.HeaderText = "Jumlah Penghasil Netto Masa Pajak Sebelumnya"
        Me.ColJumlahPenghasilNettoMasaPajakSebelum.Name = "ColJumlahPenghasilNettoMasaPajakSebelum"
        Me.ColJumlahPenghasilNettoMasaPajakSebelum.Width = 263
        '
        'ColJumlahPenghasilanNettoUntukPPH
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColJumlahPenghasilanNettoUntukPPH.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColJumlahPenghasilanNettoUntukPPH.HeaderText = "Jumlah Penghasilan Netto untuk Perhitungan PPh Pasal 21 (Setahun/Disetahunkan)"
        Me.ColJumlahPenghasilanNettoUntukPPH.Name = "ColJumlahPenghasilanNettoUntukPPH"
        Me.ColJumlahPenghasilanNettoUntukPPH.Width = 432
        '
        'ColPTKP
        '
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPTKP.DefaultCellStyle = DataGridViewCellStyle15
        Me.ColPTKP.HeaderText = "Penghasilan Tidak Kena Pajak (PTKP)"
        Me.ColPTKP.Name = "ColPTKP"
        Me.ColPTKP.Width = 215
        '
        'ColPenghasilanKenaPajak
        '
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPenghasilanKenaPajak.DefaultCellStyle = DataGridViewCellStyle16
        Me.ColPenghasilanKenaPajak.HeaderText = "Penghasilan Kena Pajak Setahun atau Disetahunkan"
        Me.ColPenghasilanKenaPajak.Name = "ColPenghasilanKenaPajak"
        Me.ColPenghasilanKenaPajak.Width = 284
        '
        'ColPPhPs21AtasPKP
        '
        DataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPPhPs21AtasPKP.DefaultCellStyle = DataGridViewCellStyle17
        Me.ColPPhPs21AtasPKP.HeaderText = "PPh Pasal 21 Atas Penghasilan Kena Pajak"
        Me.ColPPhPs21AtasPKP.Name = "ColPPhPs21AtasPKP"
        Me.ColPPhPs21AtasPKP.Width = 239
        '
        'ColPPHps21TelahDipotong
        '
        DataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPPHps21TelahDipotong.DefaultCellStyle = DataGridViewCellStyle18
        Me.ColPPHps21TelahDipotong.HeaderText = "PPh Pasal 21 yang Telah Dipotong Masa Pajak Sebelumnya"
        Me.ColPPHps21TelahDipotong.Name = "ColPPHps21TelahDipotong"
        Me.ColPPHps21TelahDipotong.Width = 318
        '
        'CollPPHTerutang
        '
        DataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CollPPHTerutang.DefaultCellStyle = DataGridViewCellStyle19
        Me.CollPPHTerutang.HeaderText = "PPh Pasal 21 Terutang"
        Me.CollPPHTerutang.Name = "CollPPHTerutang"
        Me.CollPPHTerutang.Width = 142
        '
        'ColPPhTelahDipotong
        '
        DataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPPhTelahDipotong.DefaultCellStyle = DataGridViewCellStyle20
        Me.ColPPhTelahDipotong.HeaderText = "PPh Pasal 21 dan PPh Pasal 26 yang Telah Dipotong dan Dilunasi"
        Me.ColPPhTelahDipotong.Name = "ColPPhTelahDipotong"
        Me.ColPPhTelahDipotong.Width = 347
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.MaxInputLength = 250
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'ColPrimaryKaryawan
        '
        Me.ColPrimaryKaryawan.HeaderText = "Primary Karyawan"
        Me.ColPrimaryKaryawan.Name = "ColPrimaryKaryawan"
        Me.ColPrimaryKaryawan.Visible = False
        Me.ColPrimaryKaryawan.Width = 116
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(299, 465)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(127, 465)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(558, 465)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.LblDetailSetting)
        Me.gboDetail.Controls.Add(Me.TxtReadOnly)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(9, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(724, 458)
        Me.gboDetail.TabIndex = 6
        Me.gboDetail.TabStop = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(372, 427)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(272, 427)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(386, 465)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'FrmTrans_BeginingBalanceTax
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(748, 515)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnDelete)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(758, 552)
        Me.Name = "FrmTrans_BeginingBalanceTax"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tax Begining Balance"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LblDetailSetting As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents dtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtDocumentNo As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColEmp_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKodeKaryawan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNamaKaryawan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColNoNPWP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDepartement As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPosition As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGaji As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTunjanganPPH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTunjanganLainnya As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHonor As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPremi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNatura As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTantiem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPenghasilBruto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBiayaJabatan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIuranPensiun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJumlahPengurangan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJumlahPenghasilanNetto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJumlahPenghasilNettoMasaPajakSebelum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJumlahPenghasilanNettoUntukPPH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPTKP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPenghasilanKenaPajak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPPhPs21AtasPKP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPPHps21TelahDipotong As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CollPPHTerutang As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPPhTelahDipotong As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryKaryawan As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
