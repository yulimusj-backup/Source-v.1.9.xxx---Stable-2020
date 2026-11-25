''inf-001
Option Explicit On
Option Strict Off

Imports System.Data
Imports Microsoft.VisualBasic.FileIO
Public Class FrmTransaksiFakturPajak
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim UseTax As Boolean = True
    Dim DefPrecBase As Integer = ObjCur.GetDecimal_DefaultCurrency
    Dim DTAOut As New DseLaporanVAT2025TableAdapters.DataTable2TableAdapter
    Dim TBLOut As New DseLaporanVAT2025.DataTable2DataTable
    Dim DTAOut_Sales As New DseLaporanVAT2025TableAdapters.InvoiceTableAdapter
    Dim TBLOut_Sales As New DseLaporanVAT2025.InvoiceDataTable
    Dim rowOut_Sales As DseLaporanVAT2025.InvoiceRow
    Dim DTAReturn As New DseLaporanVAT2025TableAdapters.Item_Invoice_ReturnTableAdapter
    Dim TBLReturn As New DseLaporanVAT2025.Item_Invoice_ReturnDataTable
    Dim rowReturn As DseLaporanVAT2025.Item_Invoice_ReturnRow
    Dim DTA As New DseLaporanVAT2025TableAdapters.Item_InvoiceTableAdapter
    Dim TBL As New DseLaporanVAT2025.Item_InvoiceDataTable
    Dim row As DseLaporanVAT2025.Item_InvoiceRow
    Dim DTAItemDP As New DseLaporanVAT2025TableAdapters.Item_Invoice_DPTableAdapter
    Dim TBLItemDP As New DseLaporanVAT2025.Item_Invoice_DPDataTable
    Dim rowItemDP As DseLaporanVAT2025.Item_Invoice_DPRow

    'Dim DTAOut_PR As New DseLaporanVAT2025TableAdapters.Purchase_ReturnTableAdapter
    'Dim TBLOut_PR As New DseLaporanVAT2025.Purchase_ReturnDataTable
    'Dim rowOut_PR As DseLaporanVAT2025.Purchase_ReturnRow
    'Dim DTAOut_PR_Ref As New DseLaporanVAT2025TableAdapters.Purchase_Return_ReferenceTableAdapter
    'Dim TBLOut_PR_Ref As New DseLaporanVAT2025.Purchase_Return_ReferenceDataTable
    'Dim rowOut_PR_Ref As DseLaporanVAT2025.Purchase_Return_ReferenceRow
    Dim ObjInv As New Sales_Module.ClsTransaksiSalesInvoice

    Dim DTAItem4DP As New DseLaporanVAT2025_1TableAdapters.TblItemSO4DPTableAdapter
    Dim TblItem4DP As New DseLaporanVAT2025_1.TblItemSO4DPDataTable
    Dim rowItem4DP As DseLaporanVAT2025_1.TblItemSO4DPRow

    '~ab~
    'Remarks : tambah hak akses bisa edit tax number
    '03/02/2020
    Dim BisaEditTaxNo As Boolean = False
    Dim BufferTax As String = ""
    '~

    '~ab~
    'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
    '03/02/2020 
    Public tblDummy As New DseLaporanVAT2025.Tbl_VATDataTable
    Public dtpFrom As Date
    Public dtpTo As Date
    Public Berhasil As Boolean = True
    Public showDiscount As Boolean = True


    Public Function FormatNPWP(ByVal str As String) As String
        Dim _NPWP As String
        Try
            If String.IsNullOrEmpty(str) = False Then
                _NPWP = str.Replace(".", "")
                _NPWP = _NPWP.Replace("-", "").Replace(" ", "")
                Return _NPWP
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub FrmTransaksiFakturPajak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim objCompany As New Core_Module.ClsMasterCompanyData

        Try
            Me.DGVFaktur.Rows.Clear()
            Me.DGVDetailFaktur.Rows.Clear()
            Me.DGVDetailFaktur.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVFaktur.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.TxtNPWPPenjual.ReadOnly = True
            Me.TxtNPWPPenjual.Text = objCompany.GetCompany_NPWP
            Me.DGVFaktur.Columns(Me.ColAlamatPembeli.Index).HeaderText = "Alamat Pembeli"
            Me.DGVFaktur.Columns(Me.ColEmailPembeli.Index).HeaderText = "Email Pembeli"
            Me.DGVDetailFaktur.Columns(Me.ColDetail_TarifPPnBM.Index).HeaderText = "Tarif PPnBM"
            Me.DGVDetailFaktur.Columns(Me.ColDetail_NamaSatuanUkuran.Index).HeaderText = "Nama Satuan Ukuran"
            Me.DGVDetailFaktur.Columns(Me.ColDetail_HargaSatuan.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.DGVDetailFaktur.Columns(Me.ColDetail_JumlahBarangJasa.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.DGVDetailFaktur.Columns(Me.ColDetail_TotalDiskon.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.DGVDetailFaktur.Columns(Me.ColDetail_TotalDiskon.Index).DefaultCellStyle.Format = "F2"
            Me.DGVDetailFaktur.Columns(Me.ColDetail_DPP.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.DGVDetailFaktur.Columns(Me.ColDetail_DPPNilaiLain.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.DGVDetailFaktur.Columns(Me.ColDetail_TarifPPN.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.DGVDetailFaktur.Columns(Me.ColDetail_PPN.Index).DefaultCellStyle.Format = "F2"
            Me.DGVDetailFaktur.Columns(Me.ColDetail_PPN.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.DGVDetailFaktur.Columns(Me.ColDetail_TarifPPnBM.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.DGVDetailFaktur.Columns(Me.ColDetail_PPnBM.Index).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.loadfaktur()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub loadfaktur()
        Dim rowDummy As DseLaporanVAT2025.Tbl_VATRow
        Dim DTA5 As New DseLaporanVAT2025TableAdapters.DataTable5TableAdapter
        Dim Tbl5 As New DseLaporanVAT2025.DataTable5DataTable
        Dim DTA6 As New DseLaporanVAT2025TableAdapters.DataTable6TableAdapter
        Dim Tbl6 As New DseLaporanVAT2025.DataTable6DataTable
        Dim Row6 As DseLaporanVAT2025.DataTable6Row
        Dim rowDGVFaktur As DataGridViewRow
        Dim rowDGVDetail As DataGridViewRow

        Dim NoFakturDummy As String = ""
        Dim NoNPWP As String = ""
        Dim DPP_DP As Decimal = 0
        Dim PPN_DP As Decimal = 0
        Dim DefPrecIDR As Integer = 0
        Dim FG_UANG_MUKA As Integer = 0
        Dim DPRef As String = ""
        Dim DummyDPP As Decimal = 0
        Dim DummyPPN As Decimal = 0
        Dim ItemCount As Integer = 0
        Dim Dummy1 As Decimal = 0
        Dim dummy As String = ""
        Dim valtxt As String = ""
        Dim OngkosKirim As Decimal = 0
        Dim isShipmentWithVAT As Boolean = False
        Dim TKU_ID_Penjual As String = "000000"
        Dim ketemu As Boolean = False
        Dim errorCount As Integer = 1
        Dim msgError As String = "Please check tax measurement reference for the following transaction :  "
        Dim DefaultPrecCoreTax As Integer = 2
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DGVFaktur.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVFaktur.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DGVDetailFaktur.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVDetailFaktur.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            'Me.DGVDetailFaktur.Columns(Me.ColTanggalFaktur.Index).DefaultCellStyle.Format = "mm-dd-yyyy"

            'load data item
            Me.DTA.Fill(Me.TBL, Me.dtpTo.Date, Me.dtpFrom.Date)
            Me.DTAReturn.Fill(Me.TBLReturn, Me.dtpTo, Me.dtpFrom)
            Me.KurangkanItemReturn()
            '2026-10-16, IM setelah kurangkan item return, sum yang primary transaksi, primary product, primary unit, dan nett price sama
            Me.SatukanProdukSama()

            Me.DTAItemDP.Fill(Me.TBLItemDP, Me.dtpTo, Me.dtpFrom)

            'Load data item SO for DP
            Me.DTAItem4DP.Fill(Me.TblItem4DP, Me.dtpTo, Me.dtpFrom)

            DTA5.Connection.ConnectionString = Core_Module.GetConstring
            DTA6.Connection.ConnectionString = Core_Module.GetConstring

            DTA5.Fill(Tbl5, Me.dtpFrom, Me.dtpTo)
            DTA6.Fill(Tbl6)

            Me.DGVFaktur.Rows.Clear()
            Me.DGVDetailFaktur.Rows.Clear()

            'TKU_ID_Penjual = me.ObjCompany.GetTKU_ID()
            Dim idxBaris As Integer = 1
            For Each rowDummy In tblDummy.Rows
                Me.DGVFaktur.Rows.Add()
                rowDGVFaktur = Me.DGVFaktur.Rows(Me.DGVFaktur.RowCount - 1)
                With rowDummy
                    NoFakturDummy = .Tax_Invoice_No
                    NoFakturDummy = NoFakturDummy.Trim
                    NoFakturDummy = NoFakturDummy.Replace("-", "")
                    NoFakturDummy = NoFakturDummy.Replace(".", "")
                    NoNPWP = .NPWP
                    NoNPWP = NoNPWP.Trim
                    NoNPWP = NoNPWP.Replace("-", "")
                    NoNPWP = NoNPWP.Replace(".", "").Replace(" ", "")
                    dummy = .Alamat_Pajak_1 + " " + .Alamat_Pajak_2
                    If dummy.Length > 255 Then
                        dummy = dummy.Substring(0, 255)
                    End If
                    .Cust_Supp_Name = .Cust_Supp_Name.Replace(";", "")
                    .Cust_Supp_Name = .Cust_Supp_Name.Replace("""", "")
                    .Invoice_No = .Invoice_No.Replace(";", "")
                    .Invoice_No = .Invoice_No.Replace("""", "")
                    .NPWP = .NPWP.Replace(";", "")
                    .NPWP = .NPWP.Replace("""", "")
                    .Retur = .Retur.Replace(";", "")
                    .Retur = .Retur.Replace("""", "")
                    .Tax_Address = .Tax_Address.Replace(";", "")
                    .Tax_Address = .Tax_Address.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .Tax_Invoice_No = .Tax_Invoice_No.Replace(";", "")
                    .Tax_Invoice_No = .Tax_Invoice_No.Replace("""", "")
                    .Type = .Type.Replace(";", "")
                    .Type = .Type.Replace("""", "")
                    .Nama_Jalan = .Nama_Jalan.Replace(";", "")
                    .Nama_Jalan = .Nama_Jalan.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .Blok = .Blok.Replace(";", "")
                    .Blok = .Blok.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .Nomor = .Nomor.Replace(";", "")
                    .Nomor = .Nomor.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .RT = .RT.Replace(";", "")
                    .RT = .RT.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .RW = .RW.Replace(";", "")
                    .RW = .RW.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .Kelurahan = .Kelurahan.Replace(";", "")
                    .Kelurahan = .Kelurahan.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .Kecamatan = .Kecamatan.Replace(";", "")
                    .Kecamatan = .Kecamatan.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .Kabupaten = .Kabupaten.Replace(";", "")
                    .Kabupaten = .Kabupaten.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .Propinsi = .Propinsi.Replace(";", "")
                    .Propinsi = .Propinsi.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .Kode_Pos = .Kode_Pos.Replace(";", "")
                    .Kode_Pos = .Kode_Pos.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    .No_Telepon = .No_Telepon.Replace(";", "")
                    .No_Telepon = .No_Telepon.Replace("""", "").Replace(vbCr, "").Replace(vbLf, " ")
                    '//TR : Tambahan baru Nik/Pas
                    .NikPass = .NikPass.Replace(";", "").Replace("""", "")
                    .SPPB_NO = .SPPB_NO.Replace(";", "").Replace("""", "")
                    TxtNPWPPenjual.Text = TxtNPWPPenjual.Text.Trim
                    TxtNPWPPenjual.Text = TxtNPWPPenjual.Text.Replace("-", "").Replace(".", "").Replace(" ", "")
                    '.NikPass = .NikPass.Replace("-", "")
                    '.NikPass = .NikPass.Replace(".", "").Replace(" ", "")

                    rowDGVFaktur.Cells(Me.ColBaris.Index).Value = idxBaris
                    rowDGVFaktur.Cells(Me.ColPrimaryTransaksi.Index).Value = .primary_sales_invoice
                    rowDGVFaktur.Cells(Me.ColTanggalFaktur.Index).ValueType = GetType(Date)
                    rowDGVFaktur.Cells(Me.ColTanggalFaktur.Index).Value = CDate(._Date) 'DateTime.FromOADate(._Date).ToString("YYYY-MM-DD")
                    rowDGVFaktur.Cells(Me.ColJenisfaktur.Index).Value = "Normal" 'Untuk sementara diisi normal terlebih dahulu
                    rowDGVFaktur.Cells(Me.ColKodeTransaksi.Index).Value = "04" 'Untuk sementara diisi 04 karena perhitungan dgn dpp lain
                    rowDGVFaktur.Cells(Me.ColKetTambahan.Index).Value = "" 'berhubungan dengan jenis faktur
                    rowDGVFaktur.Cells(Me.ColDokPendukung.Index).Value = "" 'berhubungan dengan keterangan tambahan
                    rowDGVFaktur.Cells(Me.ColPeriodDokPendukung.Index).Value = "" '2025-03-06, Tambahan Period keterangan tambahan (BulanTahun, ex : 012025)
                    rowDGVFaktur.Cells(Me.ColNegaraPembeli.Index).Value = "IDN" 'berhubungan dengan Country Pembeli, Sementara diisi IDN
                    valtxt = ""
                    Me.MasukanRefLainnya(valtxt, Tbl5, Tbl6, .Invoice_No, ._Date, .Cust_Supp_Name)
                    'rowDGVFaktur.Cells(Me.ColReferensi.Index).Value = .Invoice_No & " " & valtxt '"Invoice No: " & .Invoice_No & " " & valtxt 'isian referensi
                    rowDGVFaktur.Cells(Me.ColReferensi.Index).Value = .Invoice_No 'isian referensi hanya no invoice saja
                    rowDGVFaktur.Cells(Me.ColCapFasilitas.Index).Value = "" 'berhubungan dengan keterangan tambahan, disamakan dengan keterangan tambahan
                    rowDGVFaktur.Cells(Me.ColIDTKUPenjual.Index).Value = Me.TxtNPWPPenjual.Text & TKU_ID_Penjual 'ketika sudah ada fungsi ambil id TKU company
                    rowDGVFaktur.Cells(Me.ColNPWPorNIKPembeli.Index).Value = NoNPWP

                    '2025 Core Tax Tahap 2.1
                    If (NoNPWP.Length = 16 And NoNPWP <> "0000000000000000") Or .NPWPType = Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP.Company Then
                        rowDGVFaktur.Cells(Me.ColJenisIDPembeli.Index).Value = "TIN"
                        rowDGVFaktur.Cells(Me.ColNomorDokPembeli.Index).Value = "-"
                        rowDGVFaktur.Cells(Me.ColIDTKUPembeli.Index).Value = NoNPWP & .ID_TKU
                    Else
                        If .isNIK = 0 Then
                            rowDGVFaktur.Cells(Me.ColJenisIDPembeli.Index).Value = "Passport"
                            rowDGVFaktur.Cells(Me.ColNomorDokPembeli.Index).Value = .NikPass
                            rowDGVFaktur.Cells(Me.ColIDTKUPembeli.Index).Value = .ID_TKU
                            rowDGVFaktur.Cells(Me.ColNPWPorNIKPembeli.Index).Value = "0000000000000000" '@2025-01-21
                        ElseIf .isNIK = 1 Then
                            rowDGVFaktur.Cells(Me.ColJenisIDPembeli.Index).Value = "National ID"
                            rowDGVFaktur.Cells(Me.ColNomorDokPembeli.Index).Value = .NikPass
                            rowDGVFaktur.Cells(Me.ColIDTKUPembeli.Index).Value = .ID_TKU
                            rowDGVFaktur.Cells(Me.ColNPWPorNIKPembeli.Index).Value = "0000000000000000"
                        ElseIf .isNIK = 2 Then '2025 Core Tax Tahap 2.1
                            rowDGVFaktur.Cells(Me.ColJenisIDPembeli.Index).Value = "Other ID"
                            rowDGVFaktur.Cells(Me.ColNomorDokPembeli.Index).Value = .NikPass
                            rowDGVFaktur.Cells(Me.ColIDTKUPembeli.Index).Value = .ID_TKU
                            rowDGVFaktur.Cells(Me.ColNPWPorNIKPembeli.Index).Value = "0000000000000000" '@2025-01-21
                        Else
                            rowDGVFaktur.Cells(Me.ColJenisIDPembeli.Index).Value = "TIN"
                            rowDGVFaktur.Cells(Me.ColNomorDokPembeli.Index).Value = "-" '@2025-01-21
                            rowDGVFaktur.Cells(Me.ColIDTKUPembeli.Index).Value = "0000000000000000" & .ID_TKU
                            rowDGVFaktur.Cells(Me.ColNPWPorNIKPembeli.Index).Value = "0000000000000000" '@2025-01-21
                        End If
                    End If
                    rowDGVFaktur.Cells(Me.ColNamaPembeli.Index).Value = .Cust_Supp_Name
                    rowDGVFaktur.Cells(Me.ColAlamatPembeli.Index).Value = dummy.Replace(vbCr, "").Replace(vbLf, "")
                    rowDGVFaktur.Cells(Me.ColEmailPembeli.Index).Value = .Email

                    '2025 Core Tax Tahap 2.2
                    If String.IsNullOrEmpty(.SPPB_NO) And .Tax_Type <> 6 And .Tax_Type <> 7 Then
                        rowDGVFaktur.Cells(Me.ColDokPendukung.Index).Value = ""
                    Else
                        rowDGVFaktur.Cells(Me.ColDokPendukung.Index).Value = .SPPB_NO
                    End If
                    If .Tax_Type <> 6 And .Tax_Type <> 7 Then
                        rowDGVFaktur.Cells(Me.ColPeriodDokPendukung.Index).Value = ""
                        rowDGVFaktur.Cells(Me.ColKetTambahan.Index).Value = ""
                        rowDGVFaktur.Cells(Me.ColCapFasilitas.Index).Value = ""
                    Else
                        rowDGVFaktur.Cells(Me.ColPeriodDokPendukung.Index).Value = CDate(.SPPB_Date).Month.ToString("X2") & CDate(.SPPB_Date).Year
                        rowDGVFaktur.Cells(Me.ColKetTambahan.Index).Value = .Additional_Info
                        rowDGVFaktur.Cells(Me.ColCapFasilitas.Index).Value = .Facility_Stamp
                    End If
                    rowDGVFaktur.Cells(Me.ColKodeTransaksi.Index).Value = getFakturString(.Tax_Type)

                    Me.TBL.DefaultView.RowFilter = "[Primary Invoice] = " + .primary_sales_invoice.ToString
                    '2206-06-19, sorting exclude dulu, baru include
                    Me.TBL.DefaultView.Sort = "[Price Exclude] DESC, [Primary Item] ASC"
                    If Me.TBL.DefaultView.Count > 0 Then

                        'Cek Item dengan HS code Not Found, maka gagal tampil
                        For Each dv As DataRowView In Me.TBL.DefaultView
                            Me.row = dv.Row
                            If Me.row.HS_Code.ToString.Trim = "Not Found" Then
                                ketemu = True
                                If msgError.Contains(rowDummy.Invoice_No) = False Then
                                    msgError = msgError & vbNewLine & errorCount & ". " & rowDummy.Invoice_No
                                    errorCount += 1
                                End If
                            End If
                        Next

                        ''header  
                        DummyDPP = Math.Round(CDbl(.Purchase_Sales), DefPrecIDR, MidpointRounding.AwayFromZero)
                        DummyPPN = Math.Round(CDbl(.VAT), DefPrecIDR, MidpointRounding.AwayFromZero)

                        '#by ahdar | 19/06/20 |
                        OngkosKirim = 0
                        isShipmentWithVAT = False
                        '#
                        ItemCount = 0
                        For Each dv As DataRowView In Me.TBL.DefaultView
                            ItemCount += 1
                            Me.row = dv.Row

                            If Me.ObjInv.IsShipmentWithTax(Me.row.Primary_Invoice) = True Then
                                If ItemCount = Me.TBL.DefaultView.Count Then
                                    '#by ahdar | 19/06/20 | @terkait dpp yang di export dari laporan vat mengambil sub total + pengiriman
                                    If Me.ObjInv.IsShipmentWithTax(Me.row.Primary_Invoice) = True Then
                                        OngkosKirim = Me.ObjInv.GetOngkosKirim(Me.row.Primary_Invoice)
                                        DummyDPP -= OngkosKirim
                                        DummyPPN = Math.Round(DummyDPP * CDec(Me.row.Tax) / 100, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                        isShipmentWithVAT = True
                                    Else
                                        isShipmentWithVAT = False
                                    End If
                                    '#
                                    Me.row.DPP_Value = DummyDPP
                                    Me.row.PPN_Value = DummyPPN
                                    '2026-06-19
                                    If Me.row.Price_Exclude = 0 Then
                                        Me.row.DPP_Value = Math.Round(Me.row.Amount_Include * Me.row.Tax_Exchange_Rate * 100 / (100 + CDec(Me.row.Tax)), DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                        Me.row.PPN_Value = Math.Round(Me.row.DPP_Value * CDec(Me.row.Tax) / 100, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    End If
                                Else

                                    Me.row.DPP_Value = Math.Round(Me.row.Sales_Amount * Me.row.Tax_Exchange_Rate, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    Me.row.PPN_Value = Math.Round(Me.row.Sales_Amount * Me.row.Tax_Exchange_Rate _
                                        * CDec(Me.row.Tax) / 100, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)

                                    If Me.row.Price_Exclude = 0 Then
                                        Me.row.DPP_Value = Math.Round(Me.row.Amount_Include * Me.row.Tax_Exchange_Rate * 100 / (100 + CDec(Me.row.Tax)), DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                        Me.row.PPN_Value = Math.Round(Me.row.DPP_Value * CDec(Me.row.Tax) / 100, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    End If
                                    DummyDPP -= Me.row.DPP_Value
                                    DummyPPN -= Me.row.PPN_Value
                                End If
                            Else
                                'Me.row.DPP_Value = Me.row.Sales_Amount 'Math.Round(Me.row.Nett_Price, DefPrecIDR, MidpointRounding.AwayFromZero) * Me.row.Total_Qty_Multi
                                'Me.row.PPN_Value = Math.Round(Me.row.Sales_Amount * CDec(Me.row.Tax) / 100, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)

                                'If ._Date > #7/24/2025# Then
                                'Me.row.DPP_Value = Math.Round(Me.row.Amount_Include * Me.row.Tax_Exchange_Rate, DefaultPrecCoreTax, MidpointRounding.AwayFromZero) 'Math.Round(Me.row.Nett_Price, DefPrecIDR, MidpointRounding.AwayFromZero) * Me.row.Total_Qty_Multi
                                'Me.row.PPN_Value = Math.Round(Me.row.Amount_Include * CDec(Me.row.Tax) / 100, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                'Me.row.PPN_Value = Math.Round(Me.row.PPN_Value * Me.row.Tax_Exchange_Rate, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)

                                Me.row.DPP_Value = Math.Round(Me.row.Sales_Amount * Me.row.Tax_Exchange_Rate, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                Me.row.PPN_Value = Math.Round(Me.row.Sales_Amount * Me.row.Tax_Exchange_Rate _
                                * CDec(Me.row.Tax) / 100, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)

                                If Me.row.Price_Exclude = 0 Then
                                    Me.row.DPP_Value = Math.Round(Me.row.Amount_Include * Me.row.Tax_Exchange_Rate * 100 / (100 + CDec(Me.row.Tax)), _
                                    DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    Me.row.PPN_Value = Math.Round(Me.row.DPP_Value * CDec(Me.row.Tax) / 100, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                End If
                                'End If

                                DummyDPP -= Me.row.DPP_Value
                                DummyPPN -= Me.row.PPN_Value
                            End If
                        Next


                        'isikan ke dgv detail
                        For Each dv As DataRowView In Me.TBL.DefaultView
                            Me.row = dv.Row
                            Me.DGVDetailFaktur.Rows.Add()
                            rowDGVDetail = Me.DGVDetailFaktur.Rows(Me.DGVDetailFaktur.RowCount - 1)
                            With rowDGVDetail
                                If Me.row.Primary_Units <> Me.row.Primary_Units_Multi Then '@2025-01-21
                                    Me.row.HS_Code = Me.row.HS_Code_Unit_Multi
                                End If

                                .Cells(Me.ColDetail_Baris.Index).Value = idxBaris 'Menentukan baris
                                .Cells(Me.ColDetail_PrimaryProduct.Index).Value = Me.row.Primary_Product 'Primary Product
                                If Me.row.Type_Product = 1 Then
                                    .Cells(Me.ColDetail_BarangJasa.Index).Value = "B" 'Jasa ketika 1
                                Else
                                    .Cells(Me.ColDetail_BarangJasa.Index).Value = "A" 'Barang ketika 0
                                End If
                                .Cells(Me.ColDetail_KodeBarangJasa.Index).Value = Me.row.Product_Reference_Code 'Referensi, ketika kosong diisi 000000
                                .Cells(Me.ColDetail_NamaBarangJasa.Index).Value = Me.row.Description
                                .Cells(Me.ColDetail_NamaSatuanUkuran.Index).Value = Me.row.HS_Code 'HS Code, ketika kosong diisi UM.0018
                                .Cells(Me.ColDetail_JumlahBarangJasa.Index).Value = Math.Round(Me.row.Total_Qty_Multi, Me.row._Decimal, MidpointRounding.AwayFromZero)
                                .Cells(Me.ColDetail_HargaSatuan.Index).Value = Math.Round(Me.row.Unit_Price, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                .Cells(Me.ColDetail_TotalDiskon.Index).Value = Math.Round((Me.row.Unit_Price - Me.row.Nett_Price) * _
                                Me.row.Total_Qty_Multi, DefaultPrecCoreTax, MidpointRounding.AwayFromZero) 'diskon diambil dari perbedaan unit price dan nett price
                                '2026-06-19
                                If Me.row.Price_Exclude = 0 Then
                                    .Cells(Me.ColDetail_TotalDiskon.Index).Value = Math.Round((Math.Round(Me.row.Unit_Price_Include * 100 / (100 + CDec(Me.row.Tax)), 2) - _
                                    Math.Round(Me.row.DPP_Value / .Cells(Me.ColDetail_JumlahBarangJasa.Index).Value, 4, MidpointRounding.AwayFromZero)) * Me.row.Total_Qty_Multi, _
                                    DefaultPrecCoreTax, MidpointRounding.AwayFromZero) 'diskon diambil dari price include yang dibagi 1.(tax) dan dikurangi unit price asli
                                End If

                                If Me.row.Tax_Exchange_Rate <> 1 Then 'ketika usd maka * tax exchange rate untuk unit price
                                    .Cells(Me.ColDetail_HargaSatuan.Index).Value = Math.Round(CDec(Me.row.Unit_Price) * _
                                    CDec(Me.row.Tax_Exchange_Rate), DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    .Cells(Me.ColDetail_TotalDiskon.Index).Value = Math.Round(CDec(.Cells(Me.ColDetail_TotalDiskon.Index).Value) * _
                                    CDec(Me.row.Tax_Exchange_Rate), DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                End If

                                '2025 Core Tax Tahap 2.1, tanpa diskon
                                If Me.showDiscount = False Then
                                    .Cells(Me.ColDetail_HargaSatuan.Index).Value = Math.Round(Me.row.Nett_Price * Me.row.Tax_Exchange_Rate, _
                                    DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    .Cells(Me.ColDetail_TotalDiskon.Index).Value = Math.Round(0, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                End If

                                'Cek dulu, jika sales amount <> unit * qty sales - diskon, unit price = (sales amount + diskon) / qty
                                If CDec(Me.row.DPP_Value) <> CDec((Me.row.Unit_Price * Me.row.Total_Qty_Multi) - .Cells(Me.ColDetail_TotalDiskon.Index).Value) Then
                                    .Cells(Me.ColDetail_HargaSatuan.Index).Value = Math.Round((CDec(Me.row.DPP_Value) + _
                                    CDec(.Cells(Me.ColDetail_TotalDiskon.Index).Value)) / _
                                    .Cells(Me.ColDetail_JumlahBarangJasa.Index).Value, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                End If


                                .Cells(Me.ColDetail_DPP.Index).Value = Math.Round(((.Cells(Me.ColDetail_HargaSatuan.Index).Value * _
                                .Cells(Me.ColDetail_JumlahBarangJasa.Index).Value) - .Cells(Me.ColDetail_TotalDiskon.Index).Value), _
                                DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                .Cells(Me.ColDetail_TarifPPN.Index).Value = "12"
                                If rowDGVFaktur.Cells(Me.ColKodeTransaksi.Index).Value = "04" Then
                                    .Cells(Me.ColDetail_DPPNilaiLain.Index).Value = Math.Round(CDec(.Cells(Me.ColDetail_DPP.Index).Value) _
                                    * 11 / 12, DefaultPrecCoreTax, MidpointRounding.AwayFromZero) 'peraturan PMK 131/2024
                                    .Cells(Me.ColDetail_PPN.Index).Value = Math.Round(CDec(Me.row.PPN_Value), DefaultPrecCoreTax, _
                                    MidpointRounding.AwayFromZero) 'perhitungan PPN dari DPPNilaiLain
                                ElseIf rowDGVFaktur.Cells(Me.ColKodeTransaksi.Index).Value = "08" Or rowDGVFaktur.Cells(Me.ColKodeTransaksi.Index).Value = "07" Then
                                    .Cells(Me.ColDetail_DPPNilaiLain.Index).Value = Math.Round(CDec(.Cells(Me.ColDetail_DPP.Index).Value) _
                                    * 11 / 12, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    .Cells(Me.ColDetail_PPN.Index).Value = Math.Round(Me.row.PPN_Value, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                Else
                                    .Cells(Me.ColDetail_TarifPPN.Index).Value = CDec(Me.row.Tax)
                                    .Cells(Me.ColDetail_DPPNilaiLain.Index).Value = Math.Round(CDec(.Cells(Me.ColDetail_DPP.Index).Value), DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    .Cells(Me.ColDetail_PPN.Index).Value = Math.Round(Me.row.PPN_Value, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                End If

                                .Cells(Me.ColDetail_TarifPPnBM.Index).Value = Math.Round(0, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                .Cells(Me.ColDetail_PPnBM.Index).Value = Math.Round(0, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                            End With
                        Next

                    Else
                        Me.TblItem4DP.DefaultView.RowFilter = "[Primary Invoice] = " + .primary_sales_invoice.ToString
                        If Me.TblItem4DP.DefaultView.Count = 0 Then
                            Continue For
                        End If

                        'Cek Item dengan HS code Not Found, maka gagal tampil
                        For Each dv As DataRowView In Me.TblItem4DP.DefaultView
                            Me.rowItem4DP = dv.Row
                            If Me.rowItem4DP.HS_Code.ToString.Trim = "Not Found" Then
                                ketemu = True
                                If msgError.Contains(rowDummy.Invoice_No) = False Then
                                    msgError = msgError & vbNewLine & errorCount & ". " & rowDummy.Invoice_No
                                    errorCount += 1
                                End If
                            End If
                        Next

                        ''header  
                        DummyDPP = Math.Round(CDbl(.Purchase_Sales), 0, MidpointRounding.AwayFromZero)
                        DummyPPN = Math.Round(CDbl(.VAT), DefPrecIDR, MidpointRounding.AwayFromZero)

                        '#by ahdar | 19/06/20 |
                        OngkosKirim = 0
                        isShipmentWithVAT = False
                        '#
                        ItemCount = 0
                        For Each dv As DataRowView In Me.TblItem4DP.DefaultView
                            ItemCount += 1
                            Me.rowItem4DP = dv.Row

                            If Me.ObjInv.IsShipmentWithTax(Me.rowItem4DP.Primary_Invoice) = True Then
                                If ItemCount = Me.TBL.DefaultView.Count Then
                                    '#by ahdar | 19/06/20 | @terkait dpp yang di export dari laporan vat mengambil sub total + pengiriman
                                    If Me.ObjInv.IsShipmentWithTax(Me.rowItem4DP.Primary_Invoice) = True Then
                                        OngkosKirim = Me.ObjInv.GetOngkosKirim(Me.rowItem4DP.Primary_Invoice)
                                        DummyDPP -= OngkosKirim
                                        DummyPPN = Math.Round(DummyDPP * CDec(Me.rowItem4DP.Tax) / 100, DefPrecIDR, MidpointRounding.AwayFromZero)
                                        isShipmentWithVAT = True
                                    Else
                                        isShipmentWithVAT = False
                                    End If
                                    '#
                                    Me.rowItem4DP.DPP_Value = DummyDPP
                                    Me.rowItem4DP.PPN_Value = DummyPPN
                                Else
                                    If rowItem4DP.Sales_Amount = 0 Then
                                        If rowItem4DP.Price_Exclude = 0 Then
                                            Me.rowItem4DP.DPP_Value = Math.Round(Me.rowItem4DP.Amount * Me.rowItem4DP.Tax_Exchange_Rate * 100 / (100 + CDec(Me.rowItem4DP.Tax)), _
                                            DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                            Me.rowItem4DP.PPN_Value = Math.Round(Me.rowItem4DP.DPP_Value * CDec(Me.rowItem4DP.Tax) / 100, _
                                            DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                        Else
                                            Me.rowItem4DP.DPP_Value = Math.Round(Me.rowItem4DP.Amount * Me.rowItem4DP.Tax_Exchange_Rate, _
                                            DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                            Me.rowItem4DP.PPN_Value = Math.Round(Me.rowItem4DP.Amount * Me.rowItem4DP.Tax_Exchange_Rate * CDec(Me.rowItem4DP.Tax) / 100, _
                                            DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                        End If
                                    Else
                                        Me.rowItem4DP.DPP_Value = Math.Round(Me.rowItem4DP.Sales_Amount * Me.rowItem4DP.Tax_Exchange_Rate, _
                                        DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                        Me.rowItem4DP.PPN_Value = Math.Round(Me.rowItem4DP.Sales_Amount * Me.rowItem4DP.Tax_Exchange_Rate * _
                                        CDec(Me.rowItem4DP.Tax) / 100, DefaultPrecCoreTax, _
                                        MidpointRounding.AwayFromZero)
                                    End If

                                    DummyDPP -= Me.rowItem4DP.DPP_Value
                                    DummyPPN -= Me.rowItem4DP.PPN_Value
                                End If
                            Else
                                If rowItem4DP.Sales_Amount = 0 Then
                                    If rowItem4DP.Price_Exclude = 0 Then
                                        Me.rowItem4DP.DPP_Value = Math.Round(Me.rowItem4DP.Amount * 100 / (100 + CDec(Me.rowItem4DP.Tax)), _
                                        DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                        Me.rowItem4DP.PPN_Value = Math.Round(Me.rowItem4DP.DPP_Value * CDec(Me.rowItem4DP.Tax) / 100, _
                                        DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    Else
                                        Me.rowItem4DP.DPP_Value = Me.rowItem4DP.Amount 'Math.Round(Me.row.Nett_Price, DefPrecIDR, MidpointRounding.AwayFromZero) * Me.row.Total_Qty_Multi
                                        Me.rowItem4DP.PPN_Value = Math.Round(Me.row.Amount * CDec(Me.rowItem4DP.Tax) / 100, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    End If
                                Else
                                    Me.rowItem4DP.DPP_Value = Me.rowItem4DP.Sales_Amount 'Math.Round(Me.row.Nett_Price, DefPrecIDR, MidpointRounding.AwayFromZero) * Me.row.Total_Qty_Multi
                                    Me.rowItem4DP.PPN_Value = Math.Round(Me.row.Sales_Amount * CDec(Me.rowItem4DP.Tax) / 100, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                End If
                                DummyDPP -= Me.row.DPP_Value
                                DummyPPN -= Me.row.PPN_Value
                            End If
                        Next


                        'isikan ke dgv detail
                        For Each dv As DataRowView In Me.TblItem4DP.DefaultView
                            Me.rowItem4DP = dv.Row
                            Me.DGVDetailFaktur.Rows.Add()
                            rowDGVDetail = Me.DGVDetailFaktur.Rows(Me.DGVDetailFaktur.RowCount - 1)
                            With rowDGVDetail
                                If Me.rowItem4DP.Primary_Units <> Me.rowItem4DP.Primary_Units_Multi Then '@2025-01-21
                                    Me.rowItem4DP.HS_Code = Me.rowItem4DP.HS_Code_Unit_Multi
                                End If

                                .Cells(Me.ColDetail_Baris.Index).Value = idxBaris 'Menentukan baris
                                .Cells(Me.ColDetail_PrimaryProduct.Index).Value = Me.rowItem4DP.Primary_Product 'Primary Product
                                If Me.rowItem4DP.Type_Product = 1 Then
                                    .Cells(Me.ColDetail_BarangJasa.Index).Value = "B" 'Jasa ketika 1
                                Else
                                    .Cells(Me.ColDetail_BarangJasa.Index).Value = "A" 'Barang ketika 0
                                End If
                                .Cells(Me.ColDetail_KodeBarangJasa.Index).Value = Me.rowItem4DP.Product_Reference_Code 'Referensi, ketika kosong diisi 000000
                                .Cells(Me.ColDetail_NamaBarangJasa.Index).Value = Me.rowItem4DP.Description
                                .Cells(Me.ColDetail_NamaSatuanUkuran.Index).Value = Me.rowItem4DP.HS_Code 'HS Code, ketika kosong diisi UM.0018
                                .Cells(Me.ColDetail_JumlahBarangJasa.Index).Value = Math.Round(Me.rowItem4DP.Total_Qty_Multi, Me.rowItem4DP._Decimal, MidpointRounding.AwayFromZero)
                                .Cells(Me.ColDetail_HargaSatuan.Index).Value = Math.Round(Me.rowItem4DP.Unit_Price, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                .Cells(Me.ColDetail_TotalDiskon.Index).Value = Math.Round((Me.rowItem4DP.Unit_Price - Me.rowItem4DP.Nett_Price) * _
                                Me.rowItem4DP.Total_Qty_Multi, DefaultPrecCoreTax, MidpointRounding.AwayFromZero) 'diskon diambil dari perbedaan unit price dan nett price
                                If Me.rowItem4DP.Tax_Exchange_Rate <> 1 Then 'ketika usd maka * tax exchange rate untuk unit price
                                    .Cells(Me.ColDetail_HargaSatuan.Index).Value = Math.Round(CDec(Me.rowItem4DP.Unit_Price) * _
                                    CDec(Me.rowItem4DP.Tax_Exchange_Rate), DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    .Cells(Me.ColDetail_TotalDiskon.Index).Value = Math.Round(CDec(.Cells(Me.ColDetail_TotalDiskon.Index).Value) * _
                                    CDec(Me.rowItem4DP.Tax_Exchange_Rate), DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                End If

                                '2025 Core Tax Tahap 2.1, tanpa diskon
                                If Me.showDiscount = False Then
                                    .Cells(Me.ColDetail_HargaSatuan.Index).Value = Math.Round(Me.rowItem4DP.Nett_Price, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    .Cells(Me.ColDetail_TotalDiskon.Index).Value = Math.Round(0, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                End If

                                'Cek dulu, jika sales amount <> unit * qty sales - diskon, unit price = (sales amount + diskon) / qty
                                If CDec(Me.rowItem4DP.DPP_Value) <> CDec((Me.rowItem4DP.Unit_Price * Me.rowItem4DP.Total_Qty_Multi) - .Cells(Me.ColDetail_TotalDiskon.Index).Value) Then
                                    .Cells(Me.ColDetail_HargaSatuan.Index).Value = Math.Round((CDec(Me.rowItem4DP.DPP_Value) + _
                                    CDec(.Cells(Me.ColDetail_TotalDiskon.Index).Value)) / _
                                    .Cells(Me.ColDetail_JumlahBarangJasa.Index).Value, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                End If

                                .Cells(Me.ColDetail_DPP.Index).Value = Math.Round(((.Cells(Me.ColDetail_HargaSatuan.Index).Value * _
                                .Cells(Me.ColDetail_JumlahBarangJasa.Index).Value) - .Cells(Me.ColDetail_TotalDiskon.Index).Value), _
                                DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                .Cells(Me.ColDetail_TarifPPN.Index).Value = "12"
                                If rowDGVFaktur.Cells(Me.ColKodeTransaksi.Index).Value = "04" Then
                                    .Cells(Me.ColDetail_DPPNilaiLain.Index).Value = Math.Round(CDec(.Cells(Me.ColDetail_DPP.Index).Value) _
                                    * 11 / 12, DefaultPrecCoreTax, MidpointRounding.AwayFromZero) 'peraturan PMK 131/2024
                                    '.Cells(Me.ColDetail_PPN.Index).Value = Math.Round(CDec(.Cells(Me.ColDetail_DPPNilaiLain.Index).Value) _
                                    '* CDec(.Cells(Me.ColDetail_TarifPPN.Index).Value) / 100, DefPrecIDR, MidpointRounding.AwayFromZero) 'perhitungan PPN dari DPPNilaiLain
                                    .Cells(Me.ColDetail_PPN.Index).Value = Math.Round(CDec(Me.rowItem4DP.PPN_Value), DefaultPrecCoreTax, _
                                    MidpointRounding.AwayFromZero) 'perhitungan PPN dari DPPNilaiLain
                                ElseIf rowDGVFaktur.Cells(Me.ColKodeTransaksi.Index).Value = "08" Then
                                    .Cells(Me.ColDetail_DPPNilaiLain.Index).Value = Math.Round(CDec(.Cells(Me.ColDetail_DPP.Index).Value) _
                                    * 11 / 12, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    .Cells(Me.ColDetail_PPN.Index).Value = Math.Round(Me.rowItem4DP.PPN_Value, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                Else
                                    .Cells(Me.ColDetail_TarifPPN.Index).Value = CDec(Me.rowItem4DP.Tax)
                                    .Cells(Me.ColDetail_DPPNilaiLain.Index).Value = Math.Round(CDec(Me.rowItem4DP.DPP_Value), DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                    .Cells(Me.ColDetail_PPN.Index).Value = Math.Round(Me.rowItem4DP.PPN_Value, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                End If

                                .Cells(Me.ColDetail_TarifPPnBM.Index).Value = Math.Round(0, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                                .Cells(Me.ColDetail_PPnBM.Index).Value = Math.Round(0, DefaultPrecCoreTax, MidpointRounding.AwayFromZero)
                            End With
                        Next
                    End If
                End With
                idxBaris += 1
            Next
            If ketemu = True Then
                MsgBox(msgError, MsgBoxStyle.Information, "Information")
                Me.Berhasil = False
                Me.Close()
            End If

            'Tulisan End
            Me.DGVFaktur.Rows.Add()
            rowDGVFaktur = Me.DGVFaktur.Rows(Me.DGVFaktur.RowCount - 1)
            rowDGVFaktur.Cells(Me.ColBaris.Index).Value = "END"
            Me.DGVDetailFaktur.Rows.Add()
            rowDGVDetail = Me.DGVDetailFaktur.Rows(Me.DGVDetailFaktur.RowCount - 1)
            rowDGVDetail.Cells(Me.ColDetail_Baris.Index).Value = "END"

            Me.DGVFaktur.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVFaktur.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DGVDetailFaktur.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVDetailFaktur.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub DGVFaktur_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVFaktur.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVFaktur_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVFaktur.CellEndEdit
        Try
            Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVDetailFaktur_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVDetailFaktur.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVDetailFaktur_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVDetailFaktur.CellEndEdit
        Try
            Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function MasukanRefLainnya(ByRef ValTxt As String, ByVal Tbl5 As DseLaporanVAT2025.DataTable5DataTable, ByVal Tbl6 As DseLaporanVAT2025.DataTable6DataTable, _
    ByVal InvNo As String, ByVal InvoiceDate As Date, ByVal CustomerName As String) As Boolean
        Dim Row5 As DseLaporanVAT2025.DataTable5Row
        Dim DONo As String = ""
        Dim SONo As String = "" '2024-12-17 IM, tambahan SO
        Dim Kurs As String = ""

        Dim Row6 As DseLaporanVAT2025.DataTable6Row
        Dim NoKMK As String = ""
        Dim TglKMK As String = ""
        Dim Txt As String = ""
        Dim DPP As String = ""
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier
        Dim DefPrecIDR As Decimal = Me.ObjCur.GetDecimal_DefaultCurrency()
        Dim PrimaryCust As Double = -1
        Try
            '~ab~
            'Remarks : penambahan hasil keluaran export ditambah no do, dpp, kurs & kmk no.
            '04/02/2020
            Tbl5.DefaultView.RowFilter = "[NoInvoice]='" & InvNo & "'"
            If Tbl5.DefaultView.Count > 0 Then
                For i As Integer = 0 To Tbl5.DefaultView.Count - 1
                    Row5 = Tbl5.DefaultView(i).Row
                    If DONo = "" Then
                        DONo = Row5.DO_No
                    Else
                        If DONo.Contains(Row5.DO_No) = False Then
                            DONo &= ", " & Row5.DO_No
                        End If
                    End If
                    '2024-12-17 IM, tambahan SO
                    If SONo = "" Then
                        SONo = Row5.SO_NO
                    Else
                        If SONo.Contains(Row5.SO_NO) = False Then
                            SONo &= ", " & Row5.SO_NO
                        End If
                    End If
                    If DPP = "" Then
                        DPP = Me.ObjCD.DigitGrouping(Row5.Sub_Total.ToString, 2) 'permintaan bu kusmi decimal di paksa 2 saja 'Row5.Default_Precision)
                    End If
                    If Kurs = "" And Row5.Primary_Currency <> 1 Then
                        Kurs = Me.ObjCD.DigitGrouping(Row5.Tax_Exchange_Rate.ToString, 0) 'permintaan bu kusmi decimal di paksa 0 saja 'DefPrecIDR)
                    End If
                    If PrimaryCust = -1 Then
                        PrimaryCust = Row5.Primary_Customer
                    End If
                Next
            End If

            Tbl6.DefaultView.RowFilter = "[start date] <= '" & Format(InvoiceDate, "dd MMM yyy") & "' and [end date] >= '" & Format(InvoiceDate, "dd MMM yyyy") & "'"
            If Tbl6.DefaultView.Count > 0 Then
                For i As Integer = 0 To Tbl6.DefaultView.Count - 1
                    Row6 = Tbl6.DefaultView(i).Row
                    If NoKMK = "" Then
                        NoKMK = Row6.kmk_number.ToString
                        TglKMK = Me.ObjCD.GetDateString_Indo_Panjang(Row6.kmk_date) 'permintaan bu kusmi untuk bulan nya ingin di tampilkan full 'Format(Row6.kmk_date, "dd MMM yyyy")
                        Exit For
                    End If
                Next
            End If

            If ObjIdentifier.GetIdentifier() = "smess_iii" Then
                If PrimaryCust = 416 Or PrimaryCust = 922 Then 'customer = PT MURAMOTO ELEKTRONIKA INDONESIA, dan PT. INDONESIA EPSON INDUSTRI
                    If DONo <> "" Then
                        Txt &= " DO No : " & DONo
                        'Txt &= vbNewLine & "DPP : " & DPP
                    End If
                    '2024-12-17 IM, tambahan SO dan DPP dipisah panggilnya
                    If SONo <> "" Then
                        Txt &= " SO No : " & SONo
                    End If
                    If DPP <> "" Then
                        Txt &= " DPP : " & DPP
                    End If
                    'end 2024-12-17

                    If Kurs <> "" Then
                        Txt &= " Kurs : Rp." & Kurs & ",-/USD 1"
                    End If
                    If NoKMK <> "" Then
                        Txt &= " Berdasarkan KMK No. " & NoKMK & " " & TglKMK
                    End If
                End If

            ElseIf ObjIdentifier.GetIdentifier() = "smess_cls" Or ObjIdentifier.GetIdentifier() = "smess_asp2" Then
                If DONo <> "" Then
                    Txt &= " DO No : " & DONo
                    Txt &= " DPP : " & DPP
                End If
                If Kurs <> "" Then
                    Txt &= " Kurs : Rp." & Kurs & ",-/USD 1"
                End If
                If NoKMK <> "" Then
                    Txt &= " Berdasarkan KMK No. " & NoKMK & " " & TglKMK
                End If
            End If

            Txt = Txt
            ValTxt = ValTxt & Txt
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub KurangkanItemReturn()
        Dim Filter As String = ""
        Dim row1 As DseLaporanVAT2025.Item_InvoiceRow
        Dim i As Integer = 0

        Try
            For Each rowReturn In Me.TBLReturn.Rows
                Filter = "[Primary Invoice] =" + rowReturn.Primary_Sales_Invoice.ToString + " AND [Primary Product] =" + _
                rowReturn.Primary_Product.ToString + " AND [Primary SO] =" + rowReturn.Primary_SO.ToString + _
                " AND [Primary DO] =" + rowReturn.Primary_DO.ToString
                Me.TBL.DefaultView.RowFilter = Filter

                i = 0
                While Me.TBL.DefaultView.Count > i And rowReturn.Total_Qty > 0
                    row1 = Me.TBL.DefaultView(i).Row
                    If row1.Total_Qty >= rowReturn.Total_Qty Then
                        row1.Total_Qty -= rowReturn.Total_Qty
                        rowReturn.Total_Qty = 0
                    Else
                        rowReturn.Total_Qty -= row1.Total_Qty
                        row1.Total_Qty = 0
                    End If

                    If row1.Total_Qty = 0 Then
                        Me.TBL.Rows.Remove(row1)
                    End If
                End While
            Next

            Me.TBL.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getFakturString(ByVal vTaxType As Integer) As String
        Dim retVal As String = ""
        Try
            If vTaxType = 0 Then
                retVal = "01"
            ElseIf vTaxType = 1 Then
                retVal = "02"
            ElseIf vTaxType = 2 Then
                retVal = "03"
            ElseIf vTaxType = 3 Then
                retVal = "04"
            ElseIf vTaxType = 4 Then
                retVal = "05"
            ElseIf vTaxType = 5 Then
                retVal = "06"
            ElseIf vTaxType = 6 Then
                retVal = "07"
            ElseIf vTaxType = 7 Then
                retVal = "08"
            ElseIf vTaxType = 8 Then
                retVal = "09"
            ElseIf vTaxType = 9 Then
                retVal = "10"
            End If
            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SatukanProdukSama()
        Dim dummyTBL As New DseLaporanVAT2025.Item_InvoiceDataTable
        ' Proses SUM berdasarkan kombinasi kolom
        Dim i As Integer
        Dim j As Integer
        Try
            ' Copy seluruh data ke tabel sementara
            For Each Me.row In Me.TBL.Rows
                dummyTBL.ImportRow(Me.row)
            Next

            For i = dummyTBL.Rows.Count - 1 To 0 Step -1
                For j = i - 1 To 0 Step -1
                    If dummyTBL.Rows(i)("Primary Invoice").ToString() = dummyTBL.Rows(j)("Primary Invoice").ToString() AndAlso _
                    dummyTBL.Rows(i)("Primary Product").ToString() = dummyTBL.Rows(j)("Primary Product").ToString() AndAlso _
                    dummyTBL.Rows(i)("Primary Units Multi").ToString() = dummyTBL.Rows(j)("Primary Units Multi").ToString() AndAlso _
                    GetSafeDouble(dummyTBL.Rows(i)("Nett Price")) = GetSafeDouble(dummyTBL.Rows(j)("Nett Price")) Then
                        dummyTBL.Rows(j)("Total Qty Multi") = GetSafeDouble(dummyTBL.Rows(j)("Total Qty Multi")) + _
                            GetSafeDouble(dummyTBL.Rows(i)("Total Qty Multi"))
                        dummyTBL.Rows(j)("Total Qty") = GetSafeDouble(dummyTBL.Rows(j)("Total Qty")) + _
                            GetSafeDouble(dummyTBL.Rows(i)("Total Qty"))
                        dummyTBL.Rows(j)("Amount Include") = GetSafeDouble(dummyTBL.Rows(j)("Amount Include")) + _
                            GetSafeDouble(dummyTBL.Rows(i)("Amount Include"))
                        dummyTBL.Rows(j)("VAT2") = GetSafeDouble(dummyTBL.Rows(j)("VAT2")) + _
                        GetSafeDouble(dummyTBL.Rows(i)("VAT2"))
                        dummyTBL.Rows(j)("Amount") = GetSafeDouble(dummyTBL.Rows(j)("Amount")) + _
                        GetSafeDouble(dummyTBL.Rows(i)("Amount"))
                        dummyTBL.Rows(j)("Sales Amount") = GetSafeDouble(dummyTBL.Rows(j)("Sales Amount")) + _
                        GetSafeDouble(dummyTBL.Rows(i)("Sales Amount"))
                        dummyTBL.Rows.RemoveAt(i)
                        Exit For
                    End If
                Next
            Next
            ' Kembalikan hasil ke tabel asli
            Me.TBL.Clear()
            For Each Me.row In dummyTBL.Rows
                Me.TBL.ImportRow(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dummyTBL.Dispose()
        End Try
    End Sub

    Private Function GetSafeDouble(ByVal Value As Object) As Double
        Try
            If Value Is Nothing OrElse IsDBNull(Value) Then
                Return 0
            End If

            Return CDbl(Value)
        Catch ex As Exception
            Return 0
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

End Class