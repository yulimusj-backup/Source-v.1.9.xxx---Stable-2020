<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiFakturPajak
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DGVFaktur = New System.Windows.Forms.DataGridView
        Me.ColBaris = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTanggalFaktur = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJenisfaktur = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKodeTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKetTambahan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDokPendukung = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPeriodDokPendukung = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReferensi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCapFasilitas = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIDTKUPenjual = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNPWPorNIKPembeli = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJenisIDPembeli = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNegaraPembeli = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNomorDokPembeli = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNamaPembeli = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlamatPembeli = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmailPembeli = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIDTKUPembeli = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DGVDetailFaktur = New System.Windows.Forms.DataGridView
        Me.ColDetail_Baris = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_BarangJasa = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_KodeBarangJasa = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_NamaBarangJasa = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_NamaSatuanUkuran = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_HargaSatuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_JumlahBarangJasa = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_TotalDiskon = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_DPP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_DPPNilaiLain = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_TarifPPN = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_PPN = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_TarifPPnBM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_PPnBM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_PrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtNPWPPenjual = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DGVFaktur, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DGVDetailFaktur, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(9, 31)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(733, 446)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DGVFaktur)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Size = New System.Drawing.Size(725, 420)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Faktur"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DGVFaktur
        '
        Me.DGVFaktur.AllowUserToAddRows = False
        Me.DGVFaktur.AllowUserToDeleteRows = False
        Me.DGVFaktur.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGVFaktur.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVFaktur.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGVFaktur.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.DGVFaktur.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DGVFaktur.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColBaris, Me.ColTanggalFaktur, Me.ColJenisfaktur, Me.ColKodeTransaksi, Me.ColKetTambahan, Me.ColDokPendukung, Me.ColPeriodDokPendukung, Me.ColReferensi, Me.ColCapFasilitas, Me.ColIDTKUPenjual, Me.ColNPWPorNIKPembeli, Me.ColJenisIDPembeli, Me.ColNegaraPembeli, Me.ColNomorDokPembeli, Me.ColNamaPembeli, Me.ColAlamatPembeli, Me.ColEmailPembeli, Me.ColIDTKUPembeli, Me.ColPrimaryTransaksi})
        Me.DGVFaktur.Location = New System.Drawing.Point(4, 4)
        Me.DGVFaktur.Margin = New System.Windows.Forms.Padding(2)
        Me.DGVFaktur.Name = "DGVFaktur"
        Me.DGVFaktur.RowTemplate.Height = 24
        Me.DGVFaktur.Size = New System.Drawing.Size(718, 413)
        Me.DGVFaktur.TabIndex = 0
        '
        'ColBaris
        '
        Me.ColBaris.HeaderText = "Baris"
        Me.ColBaris.Name = "ColBaris"
        Me.ColBaris.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBaris.Width = 36
        '
        'ColTanggalFaktur
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.Format = "d"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.ColTanggalFaktur.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColTanggalFaktur.HeaderText = "Tanggal Faktur"
        Me.ColTanggalFaktur.Name = "ColTanggalFaktur"
        Me.ColTanggalFaktur.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTanggalFaktur.Width = 85
        '
        'ColJenisfaktur
        '
        Me.ColJenisfaktur.HeaderText = "Jenis Faktur"
        Me.ColJenisfaktur.Name = "ColJenisfaktur"
        Me.ColJenisfaktur.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJenisfaktur.Width = 70
        '
        'ColKodeTransaksi
        '
        Me.ColKodeTransaksi.HeaderText = "Kode Transaksi"
        Me.ColKodeTransaksi.Name = "ColKodeTransaksi"
        Me.ColKodeTransaksi.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColKodeTransaksi.Width = 87
        '
        'ColKetTambahan
        '
        Me.ColKetTambahan.HeaderText = "Keterangan Tambahan"
        Me.ColKetTambahan.Name = "ColKetTambahan"
        Me.ColKetTambahan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColKetTambahan.Width = 122
        '
        'ColDokPendukung
        '
        Me.ColDokPendukung.HeaderText = "Dokumen Pendukung"
        Me.ColDokPendukung.Name = "ColDokPendukung"
        Me.ColDokPendukung.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDokPendukung.Width = 117
        '
        'ColPeriodDokPendukung
        '
        Me.ColPeriodDokPendukung.HeaderText = "Period Dokumen Pendukung"
        Me.ColPeriodDokPendukung.Name = "ColPeriodDokPendukung"
        Me.ColPeriodDokPendukung.Width = 169
        '
        'ColReferensi
        '
        Me.ColReferensi.HeaderText = "Referensi"
        Me.ColReferensi.Name = "ColReferensi"
        Me.ColReferensi.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColReferensi.Width = 58
        '
        'ColCapFasilitas
        '
        Me.ColCapFasilitas.HeaderText = "Cap Fasilitas"
        Me.ColCapFasilitas.Name = "ColCapFasilitas"
        Me.ColCapFasilitas.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCapFasilitas.Width = 72
        '
        'ColIDTKUPenjual
        '
        Me.ColIDTKUPenjual.HeaderText = "ID TKU Penjual"
        Me.ColIDTKUPenjual.Name = "ColIDTKUPenjual"
        Me.ColIDTKUPenjual.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColIDTKUPenjual.Width = 87
        '
        'ColNPWPorNIKPembeli
        '
        Me.ColNPWPorNIKPembeli.HeaderText = "NPWP/NIK Pembeli"
        Me.ColNPWPorNIKPembeli.Name = "ColNPWPorNIKPembeli"
        Me.ColNPWPorNIKPembeli.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNPWPorNIKPembeli.Width = 109
        '
        'ColJenisIDPembeli
        '
        Me.ColJenisIDPembeli.HeaderText = "Jenis ID Pembeli"
        Me.ColJenisIDPembeli.Name = "ColJenisIDPembeli"
        Me.ColJenisIDPembeli.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColJenisIDPembeli.Width = 91
        '
        'ColNegaraPembeli
        '
        Me.ColNegaraPembeli.HeaderText = "Negara Pembeli"
        Me.ColNegaraPembeli.Name = "ColNegaraPembeli"
        Me.ColNegaraPembeli.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNegaraPembeli.Width = 88
        '
        'ColNomorDokPembeli
        '
        Me.ColNomorDokPembeli.HeaderText = "Nomor Dokumen Pembeli"
        Me.ColNomorDokPembeli.Name = "ColNomorDokPembeli"
        Me.ColNomorDokPembeli.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNomorDokPembeli.Width = 133
        '
        'ColNamaPembeli
        '
        Me.ColNamaPembeli.HeaderText = "Nama Pembeli"
        Me.ColNamaPembeli.Name = "ColNamaPembeli"
        Me.ColNamaPembeli.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNamaPembeli.Width = 81
        '
        'ColAlamatPembeli
        '
        Me.ColAlamatPembeli.HeaderText = "AlamatPembeli"
        Me.ColAlamatPembeli.Name = "ColAlamatPembeli"
        Me.ColAlamatPembeli.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAlamatPembeli.Width = 82
        '
        'ColEmailPembeli
        '
        Me.ColEmailPembeli.HeaderText = "EmailPembeli"
        Me.ColEmailPembeli.Name = "ColEmailPembeli"
        Me.ColEmailPembeli.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEmailPembeli.Width = 75
        '
        'ColIDTKUPembeli
        '
        Me.ColIDTKUPembeli.HeaderText = "ID TKU Pembeli"
        Me.ColIDTKUPembeli.Name = "ColIDTKUPembeli"
        Me.ColIDTKUPembeli.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColIDTKUPembeli.Width = 89
        '
        'ColPrimaryTransaksi
        '
        Me.ColPrimaryTransaksi.HeaderText = "Primary Transaksi"
        Me.ColPrimaryTransaksi.Name = "ColPrimaryTransaksi"
        Me.ColPrimaryTransaksi.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryTransaksi.Visible = False
        Me.ColPrimaryTransaksi.Width = 96
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DGVDetailFaktur)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Size = New System.Drawing.Size(725, 420)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Detail Faktur"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DGVDetailFaktur
        '
        Me.DGVDetailFaktur.AllowUserToAddRows = False
        Me.DGVDetailFaktur.AllowUserToDeleteRows = False
        Me.DGVDetailFaktur.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVDetailFaktur.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGVDetailFaktur.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.DGVDetailFaktur.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DGVDetailFaktur.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDetail_Baris, Me.ColDetail_BarangJasa, Me.ColDetail_KodeBarangJasa, Me.ColDetail_NamaBarangJasa, Me.ColDetail_NamaSatuanUkuran, Me.ColDetail_HargaSatuan, Me.ColDetail_JumlahBarangJasa, Me.ColDetail_TotalDiskon, Me.ColDetail_DPP, Me.ColDetail_DPPNilaiLain, Me.ColDetail_TarifPPN, Me.ColDetail_PPN, Me.ColDetail_TarifPPnBM, Me.ColDetail_PPnBM, Me.ColDetail_PrimaryProduct})
        Me.DGVDetailFaktur.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVDetailFaktur.Location = New System.Drawing.Point(2, 2)
        Me.DGVDetailFaktur.Margin = New System.Windows.Forms.Padding(2)
        Me.DGVDetailFaktur.Name = "DGVDetailFaktur"
        Me.DGVDetailFaktur.RowTemplate.Height = 24
        Me.DGVDetailFaktur.Size = New System.Drawing.Size(721, 416)
        Me.DGVDetailFaktur.TabIndex = 0
        '
        'ColDetail_Baris
        '
        Me.ColDetail_Baris.HeaderText = "Baris"
        Me.ColDetail_Baris.Name = "ColDetail_Baris"
        Me.ColDetail_Baris.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_Baris.Width = 36
        '
        'ColDetail_BarangJasa
        '
        Me.ColDetail_BarangJasa.HeaderText = "Barang/ Jasa"
        Me.ColDetail_BarangJasa.Name = "ColDetail_BarangJasa"
        Me.ColDetail_BarangJasa.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_BarangJasa.Width = 77
        '
        'ColDetail_KodeBarangJasa
        '
        Me.ColDetail_KodeBarangJasa.HeaderText = "Kode Barang/ Jasa"
        Me.ColDetail_KodeBarangJasa.Name = "ColDetail_KodeBarangJasa"
        Me.ColDetail_KodeBarangJasa.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_KodeBarangJasa.Width = 105
        '
        'ColDetail_NamaBarangJasa
        '
        Me.ColDetail_NamaBarangJasa.HeaderText = "Nama Barang/ Jasa"
        Me.ColDetail_NamaBarangJasa.Name = "ColDetail_NamaBarangJasa"
        Me.ColDetail_NamaBarangJasa.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_NamaBarangJasa.Width = 108
        '
        'ColDetail_NamaSatuanUkuran
        '
        Me.ColDetail_NamaSatuanUkuran.HeaderText = "Nama Satuan Ukuruan"
        Me.ColDetail_NamaSatuanUkuran.Name = "ColDetail_NamaSatuanUkuran"
        Me.ColDetail_NamaSatuanUkuran.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_NamaSatuanUkuran.Width = 122
        '
        'ColDetail_HargaSatuan
        '
        Me.ColDetail_HargaSatuan.HeaderText = "Harga Satuan"
        Me.ColDetail_HargaSatuan.Name = "ColDetail_HargaSatuan"
        Me.ColDetail_HargaSatuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_HargaSatuan.Width = 79
        '
        'ColDetail_JumlahBarangJasa
        '
        Me.ColDetail_JumlahBarangJasa.HeaderText = "Jumlah Barang/ Jasa"
        Me.ColDetail_JumlahBarangJasa.Name = "ColDetail_JumlahBarangJasa"
        Me.ColDetail_JumlahBarangJasa.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_JumlahBarangJasa.Width = 113
        '
        'ColDetail_TotalDiskon
        '
        Me.ColDetail_TotalDiskon.HeaderText = "Total Diskon"
        Me.ColDetail_TotalDiskon.Name = "ColDetail_TotalDiskon"
        Me.ColDetail_TotalDiskon.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_TotalDiskon.Width = 73
        '
        'ColDetail_DPP
        '
        Me.ColDetail_DPP.HeaderText = "DPP"
        Me.ColDetail_DPP.Name = "ColDetail_DPP"
        Me.ColDetail_DPP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_DPP.Width = 35
        '
        'ColDetail_DPPNilaiLain
        '
        Me.ColDetail_DPPNilaiLain.HeaderText = "DPP Nilai Lain"
        Me.ColDetail_DPPNilaiLain.Name = "ColDetail_DPPNilaiLain"
        Me.ColDetail_DPPNilaiLain.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_DPPNilaiLain.Width = 81
        '
        'ColDetail_TarifPPN
        '
        Me.ColDetail_TarifPPN.HeaderText = "Tarif PPN"
        Me.ColDetail_TarifPPN.Name = "ColDetail_TarifPPN"
        Me.ColDetail_TarifPPN.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_TarifPPN.Width = 59
        '
        'ColDetail_PPN
        '
        Me.ColDetail_PPN.HeaderText = "PPN"
        Me.ColDetail_PPN.Name = "ColDetail_PPN"
        Me.ColDetail_PPN.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_PPN.Width = 35
        '
        'ColDetail_TarifPPnBM
        '
        Me.ColDetail_TarifPPnBM.HeaderText = "Tarfi PPnBM"
        Me.ColDetail_TarifPPnBM.Name = "ColDetail_TarifPPnBM"
        Me.ColDetail_TarifPPnBM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_TarifPPnBM.Width = 73
        '
        'ColDetail_PPnBM
        '
        Me.ColDetail_PPnBM.HeaderText = "PPnBM"
        Me.ColDetail_PPnBM.Name = "ColDetail_PPnBM"
        Me.ColDetail_PPnBM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_PPnBM.Width = 49
        '
        'ColDetail_PrimaryProduct
        '
        Me.ColDetail_PrimaryProduct.HeaderText = "Primary Product"
        Me.ColDetail_PrimaryProduct.Name = "ColDetail_PrimaryProduct"
        Me.ColDetail_PrimaryProduct.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_PrimaryProduct.Visible = False
        Me.ColDetail_PrimaryProduct.Width = 87
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 9)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "NPWP Penjual"
        '
        'TxtNPWPPenjual
        '
        Me.TxtNPWPPenjual.Location = New System.Drawing.Point(94, 7)
        Me.TxtNPWPPenjual.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtNPWPPenjual.Name = "TxtNPWPPenjual"
        Me.TxtNPWPPenjual.Size = New System.Drawing.Size(192, 20)
        Me.TxtNPWPPenjual.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(299, 6)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 20)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Filter"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'FrmTransaksiFakturPajak
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(751, 486)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TxtNPWPPenjual)
        Me.Controls.Add(Me.TabControl1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(759, 513)
        Me.Name = "FrmTransaksiFakturPajak"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "XML Convertion for CoreTax"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DGVFaktur, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DGVDetailFaktur, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents DGVFaktur As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtNPWPPenjual As System.Windows.Forms.TextBox
    Friend WithEvents DGVDetailFaktur As System.Windows.Forms.DataGridView
    Friend WithEvents ColDetail_Baris As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_BarangJasa As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_KodeBarangJasa As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_NamaBarangJasa As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_NamaSatuanUkuran As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_HargaSatuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_JumlahBarangJasa As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_TotalDiskon As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_DPP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_DPPNilaiLain As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_TarifPPN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_PPN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_TarifPPnBM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_PPnBM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_PrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBaris As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTanggalFaktur As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJenisfaktur As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKodeTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKetTambahan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDokPendukung As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPeriodDokPendukung As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReferensi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCapFasilitas As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIDTKUPenjual As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNPWPorNIKPembeli As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJenisIDPembeli As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNegaraPembeli As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNomorDokPembeli As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNamaPembeli As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlamatPembeli As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmailPembeli As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIDTKUPembeli As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
