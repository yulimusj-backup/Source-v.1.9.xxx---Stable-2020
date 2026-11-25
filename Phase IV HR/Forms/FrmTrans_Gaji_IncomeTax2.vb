Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTrans_Gaji_IncomeTax2
    Public vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable
    Public vTblSalary As DseTransaksiSalary.tbltrans_salary3DataTable

    Public vEmployeeCode As String = ""
    Public vTransactionDate As Date = #1/1/1950#
    Public vEmployeeName As String = ""
    Public vNIK_Kary As String = ""
    Public vNPWP As String = ""
    Public vStatusNPWP As String = ""
    Public vKodeNegaraDomisili As String = ""
    Public vJenisKelamin As String = ""
    Public vAlamatPajak1 As String = ""
    Public vAlamatPajak2 As String = ""
    Public vPrimaryEmployee As Double = -1
    Public vDepartement As String = ""
    Public vPosition As String = ""
    Public vFromMonthlyReport As Boolean = False
    Public vEndOfYear As Boolean = False
    Public vStartPeriod As Date = #1/1/1950#
    Public vEndPeriod As Date = #1/1/1950#
    Public vIsResign As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData

    Dim DS As New DseTransaksiSalary
    Dim TBLCetak As DseTransaksiSalary.TblCetakPajakDataTable = DS.TblCetakPajak
    Dim rowItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxRow
    Dim rowItemCetak As DseTransaksiSalary.TblCetakPajakRow
    Dim rowHeaderSalary As DseTransaksiSalary.tbltrans_salary3Row
    Dim Report As New CrpTaxIncomeCalculation
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Negatif As Integer = 0
    Dim TBLItemComponent As New DseTransaksiSalary.tblmaster_salary_formula_salarycomponentDataTable
    Dim DTAItemComponent As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
    Dim rowItemComponent As DseTransaksiSalary.tblmaster_salary_formula_salarycomponentRow
    Dim vPrimaryGroupLama As Double = -1
    Dim vPrimaryTransaksiLama As Double = -1

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Public Enum EnumDescPPHTerhutang As Integer
        Pajak_Terhutang_Disetahunkan
        Pajak_Terhutang_Disetahunkan1
        Pajak_Terhutang_Penghasilan_TidakRutin
        Pajak_Terhutang_Penghasilan_Rutin
        Pajak_Terhutang_Bulanan
        Kurang_Bayar
        Pajak_Terhutang_Disetahunkan2
        Pajak_Terhutang_Penghasilan_Rutin1
        Pajak_Terhutang_Penghasilan_TidakRutin1
        Pajak_Penghasilan_Telah_Dilunasi
        Pajak_Penghasilan_Lebih_Dipotong
        Pajak_Penghasilan_Kurang_Dipotong
        Pajak_Penghasilan_Yang_Telah_Dibayarkan_masa_Sblm
    End Enum

    Private Function GetComponentDesc_PPHTerhutang(ByVal vTipePajakTerhutang As EnumDescPPHTerhutang) As String
        Try
            If vTipePajakTerhutang = EnumDescPPHTerhutang.Kurang_Bayar Then
                Return "Kurang Bayar"
            ElseIf vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Bulanan Then
                Return "Pajak Terhutang " + Format(vTransactionDate, "MMM yyyy")
            ElseIf vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Disetahunkan Then
                Return "Pajak Terhutang Disetahunkan"
            ElseIf vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Disetahunkan2 Then
                Return "Pajak Terhutang " + Format(Me.vEndPeriod, "MMM yyyy")
                'If Format(Me.vStartPeriod, "MMM yyyy") = Format(Me.vEndPeriod, "MMM yyyy") Then
                '    Return "Pajak Terhutang " + Format(Me.vEndPeriod, "MMM yyyy")
                'Else
                '    Return "Pajak Terhutang " + Format(Me.vStartPeriod, "MMM") + " s.d " + Format(Me.vEndPeriod, "MMM yyyy")
                'End If
            ElseIf vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Disetahunkan1 Then
                Return ""
            ElseIf vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Penghasilan_Rutin Or _
            vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Penghasilan_Rutin1 Then
                Return "Pajak Terhutang Pengasilan Rutin"
            ElseIf vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Penghasilan_TidakRutin Or _
            vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Penghasilan_TidakRutin1 Then
                Return "Pajak Terhutang Pengasilan Tidak Rutin"
            ElseIf vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Penghasilan_Telah_Dilunasi Then
                If Format(Me.vStartPeriod, "MMM yyyy") = Format(Me.ObjCD.GetFirstDateofTheMonth(Me.vEndPeriod), "MMM yyyy") Then
                    Return "Pajak Penghasilan Yang Telah Dilunasi " + Format(Me.vStartPeriod, "MMM yyyy")
                Else
                    Return "Pajak Penghasilan Yang Telah Dilunasi " + Format(Me.vStartPeriod, "MMM") + " s.d " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.vEndPeriod) - TimeSpan.FromDays(1), "MMM yyyy")
                End If
            ElseIf vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Penghasilan_Kurang_Dipotong Then
                Return "Pajak Penghasilan Yang Kurang Dipotong"
            ElseIf vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Penghasilan_Lebih_Dipotong Then
                Return "Pajak Penghasilan Yang Lebih Dipotong"
            ElseIf vTipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Penghasilan_Yang_Telah_Dibayarkan_masa_Sblm Then
                Return "Pajak Penghasilan Yang Sdh Dibayarkan Masa Sebelum"
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetTaxElementString(ByVal vTaxElement As ClsTransaksiSalary_Item.EnumTaxElement) As String
        Try
            If vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT Then
                Return "GAJI/PENSIUN ATAU THT/JHT"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH Then
                Return "TUNJANGAN PPh"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb Then
                Return "TUNJANGAN LAINNYA, UANG LEMBUR DAN SEBAGAINYA"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya Then
                Return "HONORARIUM DAN IMBALAN LAIN SEJENISNYA"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja Then
                Return "PREMI ASURANSI YANG DIBAYAR PEMBERI KERJA"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH Then
                Return "PENERIMAAN DALAM BENTUK NATURA DAN KENIKMATAN LAINNYA YANG DIKENAKAN PEMOTONGAN PPh PASAL 21"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi Then
                Return "TANTIEM, BONUS, GRATIFIKASI, JASA PRODUKSI DAN THR"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No8_Jumlah_Penghasilan_Bruto Then
                Return "JUMLAH PENGHASILAN BRUTO"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No9_BiayaJabatan_BiayaPensiun Then
                Return "BIAYA JABATAN/BIAYA PENSIUN"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No10_IuranPensiun_THT_JHT Then
                Return "IURAN PENSIUN ATAU IURAN THT/JHT"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No11_JumlahPengurang Then
                Return "JUMLAH PENGURANGAN"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No12_JumlahPenghasilanUntukPerhitunganPPh Then
                Return "JUMLAH PENGHASILAN NETO"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No13_PenghasilanNettoMasaSebelum Then
                Return "PENGHASILAN NETO MASA SEBELUMNYA"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No14_PenghasilanNettoKenaPajak Then
                Return "JUMLAH PENGHASILAN NETO UNTUK PENGHITUNGAN PPh PASAL 21 (SETAHUN/DISETAHUNKAN)"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No15_PTKP Then
                Return "PENGHASILAN TIDAK KENA PAJAK (PTKP)"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No16_PKP Then
                Return "PENGHASILAN KENA PAJAK SETAHUN/DISETAHUNKAN"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No17_PPHSetahun Then
                Return "PPh PASAL 21 ATAS PENGHASILAN KENA PAJAK SETAHUN/DISETAHUNKAN"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No18_PPH21DipotongMasaPajakSebelum Then
                Return "PPh PASAL 21 YANG TELAH DIPOTONG MASA SEBELUMNYA"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang Or _
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon Or _
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No21_PPHFinalPesangon Or _
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No21_TunjPPHPesangon Then
                Return "PPh PASAL 21 TERUTANG"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No20_PPHLunas_Terpotong Then
                Return "PPh PASAL 21 DAN PPh PASAL 26 YANG TELAH DIPOTONG DAN DILUNASI"
            ElseIf vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang Then
                Return "PPh PASAL 21 FINAL PESANGON"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInlayFollowingSection(ByVal vTaxElement As ClsTransaksiSalary_Item.EnumTaxElement) As Boolean
        Try
            If vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT Or _
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb Or _
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya Or _
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja Or _
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH Or _
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi Or _
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No10_IuranPensiun_THT_JHT Or _
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang Then
                'Or vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No13_PenghasilanNettoMasaSebelum Then '#sinkron2018_pphcls
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmTrans_Gaji_IncomeTax2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim TipePajakTerhutang As EnumDescPPHTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Disetahunkan
        Dim TipeElemenPajak_PPHTerhutang As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang
        Dim TipeElemenPajak_BasicSalary As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT
        Dim TipeElemenPajak_Pesangon As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon
        Dim TipeElemenPajak_PPHFinalPesangon As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No21_PPHFinalPesangon
        Dim TipeElemenPajak_TunjPPHFinalPesangon As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No21_TunjPPHPesangon
        Dim TipeElemenPajak_PenghasilaBruto As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No8_Jumlah_Penghasilan_Bruto
        Dim TipeElemenPajak_JumlahPengurang As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No11_JumlahPengurang
        Dim TipeElemenPajak_JumlahPenghasilanUntukPerhitunganPPh As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No12_JumlahPenghasilanUntukPerhitunganPPh
        Dim TipeElemenPajak_PenghasilanNettoMasaSebelum As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No13_PenghasilanNettoMasaSebelum
        Dim TipeElemenPajak_PenghasilanNettoKenaPajak As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No14_PenghasilanNettoKenaPajak
        Dim TipeElemenPajak_PTKP As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No15_PTKP
        Dim TipeElemenPajak_PKP As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No16_PKP
        Dim TipeElemenPajak_PPHSetahun As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No17_PPHSetahun
        Dim TipeElemenPajak_PPH21DipotongMasaPajakSebelum As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No18_PPH21DipotongMasaPajakSebelum
        Dim TipeElemenPajak_PPHLunas_Terpotong As ClsTransaksiSalary_Item.EnumTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No20_PPHLunas_Terpotong
        Dim PrimaryGroup1 As Double = -1
        Dim counter As Double = 0
        Dim PrimaryGroup2 As Double = -1
        Dim IsResign As Integer = 0
        Dim EndPeriodMonth As Integer = 0

        Try
            Me.ShowIcon = False
            If Me.vEndOfYear = True Then
                TipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Disetahunkan2
            End If
            Me.DTAItemComponent.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAItemComponent.Fill(Me.TBLItemComponent)
            Me.TBLCetak.Rows.Clear()

            'urutankan berdasarkan tax element dan primary
            Me.vTBLItemTax.DefaultView.RowFilter = ""
            Me.vTBLItemTax.DefaultView.Sort = "[Primary Employee], [Tanggal Transaksi], [No Transaksi], [primary transaksi], [tipe tax], primary"

            'untuk setiap baris pada detail item tax pindahkan ke tabel cetak
            For Each DV As DataRowView In Me.vTBLItemTax.DefaultView
                Me.rowItemTax = DV.Row
                Me.rowItemCetak = Me.TBLCetak.NewTblCetakPajakRow
                counter += 1

                With Me.rowItemCetak
                    If Me.vFromMonthlyReport = True Then
                        If Me.rowItemTax.primary_transaksi <> Me.vPrimaryTransaksiLama Then
                            'ganti data header
                            Me.vPrimaryTransaksiLama = Me.rowItemTax.primary_transaksi
                            .Primary_Group = Me.GetPrimaryGroupMax + 1
                            vPrimaryGroupLama = .Primary_Group

                            If PrimaryGroup1 = -1 Then
                                PrimaryGroup1 = .Primary_Group
                            Else
                                PrimaryGroup2 = PrimaryGroup1
                                PrimaryGroup1 = .Primary_Group
                            End If
                            If PrimaryGroup1 <> PrimaryGroup2 And PrimaryGroup2 <> -1 Then
                                Me.TambahkanElemenPajakYgTidakAda(PrimaryGroup2)
                            End If

                            'If Me.rowHeaderSalary.End_Date.Month = 12 Or rowHeaderSalary.Is_Resign = 1 Then
                            '    TipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Disetahunkan2
                            'End If

                            Me.vTblSalary.DefaultView.RowFilter = "primary = " + Me.rowItemTax.primary_transaksi.ToString
                            If Me.vTblSalary.DefaultView.Count = 0 Then
                                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            Else
                                Me.rowHeaderSalary = Me.vTblSalary.DefaultView(0).Row
                                vEmployeeCode = Me.rowHeaderSalary.Employee_Code
                                vEmployeeName = Me.rowHeaderSalary.Employee_Name
                                vNIK_Kary = Me.rowHeaderSalary.A_NIK
                                vNPWP = Me.rowHeaderSalary.A_NPWP
                                vTransactionDate = Me.rowHeaderSalary.Tanggal_Transaksi
                                Me.vEndPeriod = Me.rowHeaderSalary.End_Date
                                vStatusNPWP = Me.rowHeaderSalary.A_Status_NPWP
                                vKodeNegaraDomisili = Me.rowHeaderSalary.A_Kode_Negara_Domisili
                                vJenisKelamin = Me.rowHeaderSalary.A_Jenis_Kelamin
                                vAlamatPajak1 = Me.rowHeaderSalary.A_Alamat1
                                vAlamatPajak2 = Me.rowHeaderSalary.A_Alamat2
                                vDepartement = Me.rowHeaderSalary.Dept_Desc
                                vPosition = Me.rowHeaderSalary.A_Nama_Jabatan
                                vPrimaryEmployee = Me.rowHeaderSalary.Primary

                                If rowHeaderSalary.Is_Resign = 1 Or rowHeaderSalary.End_Date.Month = 12 Then
                                    TipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Disetahunkan2
                                Else
                                    TipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Terhutang_Disetahunkan
                                End If

                                If vJenisKelamin = ClsTransaksiSalary.EnumJenisKelamin.Pria Then
                                    vJenisKelamin = "LAKI-LAKI"
                                ElseIf vJenisKelamin = ClsTransaksiSalary.EnumJenisKelamin.Wanita Then
                                    vJenisKelamin = "PEREMPUAN"
                                End If
                                IsResign = rowHeaderSalary.Is_Resign
                                EndPeriodMonth = rowHeaderSalary.End_Date.Month
                                If vStatusNPWP = ClsMasterKaryawan.GolonganPajak.TK0 Then
                                    vStatusNPWP = "TK0"
                                ElseIf vStatusNPWP = ClsMasterKaryawan.GolonganPajak.TK1 Then
                                    vStatusNPWP = "TK1"
                                ElseIf vStatusNPWP = ClsMasterKaryawan.GolonganPajak.TK2 Then
                                    vStatusNPWP = "TK2"
                                ElseIf vStatusNPWP = ClsMasterKaryawan.GolonganPajak.TK3 Then
                                    vStatusNPWP = "TK3"
                                ElseIf vStatusNPWP = ClsMasterKaryawan.GolonganPajak.K0 Then
                                    vStatusNPWP = "K0"
                                ElseIf vStatusNPWP = ClsMasterKaryawan.GolonganPajak.K1 Then
                                    vStatusNPWP = "K1"
                                ElseIf vStatusNPWP = ClsMasterKaryawan.GolonganPajak.K2 Then
                                    vStatusNPWP = "K2"
                                ElseIf vStatusNPWP = ClsMasterKaryawan.GolonganPajak.K3 Then
                                    vStatusNPWP = "K3"
                                End If
                            End If
                        Else
                            .Primary_Group = vPrimaryGroupLama
                        End If
                    Else
                        If vPrimaryGroupLama = -1 Then
                            .Primary_Group = Me.GetPrimaryGroupMax
                            vPrimaryGroupLama = .Primary_Group
                            PrimaryGroup1 = .Primary_Group
                        Else
                            .Primary_Group = vPrimaryGroupLama
                        End If
                        If vEndPeriod.Month = 12 Then
                            EndPeriodMonth = vEndPeriod.Month
                        End If
                        IsResign = vIsResign
                    End If

                    'isikan data item cetak
                    .Kode_Karyawan = vEmployeeCode
                    .Nama_Karyawan = vEmployeeName
                    .NIK = vNIK_Kary
                    .NPWP = vNPWP
                    .Period = vTransactionDate
                    .Tax_Category = vStatusNPWP
                    .Nationality = vKodeNegaraDomisili
                    .Sex = vJenisKelamin
                    .Inlay_Following_Section = Me.GetInlayFollowingSection(Me.rowItemTax.tipe_tax)
                    .Address = vAlamatPajak1
                    If vAlamatPajak2 <> "" Then
                        .Address += vbCrLf + vAlamatPajak2
                    End If
                    .Departement = vDepartement
                    .Job_Position = vPosition
                    .Primary_Karyawan = vPrimaryEmployee


                    .Amount = Me.rowItemTax.amount_month
                    'terkait dengan lebih bayar
                    If TipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Penghasilan_Lebih_Dipotong Then
                        .Amount = Math.Abs(.Amount)
                    End If


                    If Me.rowItemTax.identifier_salary_component = "" Then
                        If rowItemTax.tipe_tax = CInt(TipeElemenPajak_PPHTerhutang) Then
                            If rowItemTax.remarks_rutin <> "" And (IsResign = 1 Or EndPeriodMonth = 12) Then
                                TipePajakTerhutang = rowItemTax.remarks_rutin
                                rowItemTax.remarks_rutin = ""

                                .Tax_Component_Desc = Me.GetComponentDesc_PPHTerhutang(TipePajakTerhutang)
                                If TipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Penghasilan_Lebih_Dipotong Then
                                    If rowItemTax.amount_month > 0 Then
                                        TipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Penghasilan_Kurang_Dipotong
                                        .Tax_Component_Desc = Me.GetComponentDesc_PPHTerhutang(TipePajakTerhutang)
                                    End If
                                End If
                            Else
                                .Tax_Component_Desc = Me.GetComponentDesc_PPHTerhutang(TipePajakTerhutang)
                                If TipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Penghasilan_Lebih_Dipotong Then
                                    If rowItemTax.amount_month > 0 Then
                                        TipePajakTerhutang = EnumDescPPHTerhutang.Pajak_Penghasilan_Kurang_Dipotong
                                        .Tax_Component_Desc = Me.GetComponentDesc_PPHTerhutang(TipePajakTerhutang)
                                    End If
                                End If
                                TipePajakTerhutang += 1
                            End If
                        ElseIf rowItemTax.tipe_tax = CInt(TipeElemenPajak_BasicSalary) Then
                            .Tax_Component_Desc = "Basic Salary"
                        ElseIf rowItemTax.tipe_tax = TipeElemenPajak_Pesangon Then
                            .Tax_Component_Desc = "Pesangon"
                        ElseIf rowItemTax.tipe_tax = TipeElemenPajak_PPHFinalPesangon Then
                            .Tax_Component_Desc = "PPh Final Pesangon"
                        ElseIf rowItemTax.tipe_tax = TipeElemenPajak_TunjPPHFinalPesangon Then
                            .Tax_Component_Desc = "Tunj. PPh Final Pesangon"
                        ElseIf rowItemTax.tipe_tax = TipeElemenPajak_PenghasilanNettoMasaSebelum Then '#sinkron2018_pphcls
                            '.Tax_Component_Desc = rowItemTax.remarks_gabungan
                            'rowItemTax.remarks_gabungan = ""
                            'rowItemTax.remarks_rutin = "" 
                        Else
                            .Tax_Component_Desc = ""
                        End If
                    Else
                        Me.TBLItemComponent.DefaultView.RowFilter = "Identifier = '" + Me.rowItemTax.identifier_salary_component + "'"
                        If Me.TBLItemComponent.DefaultView.Count = 0 Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Else
                            Me.rowItemComponent = Me.TBLItemComponent.DefaultView(0).Row
                            .Tax_Component_Desc = Me.rowItemComponent.Description
                        End If
                    End If


                    .Amount_Gabungan = Me.rowItemTax.amount_gabungan_disetahunkan
                    .Amount_Routine = Me.rowItemTax.amount_rutin_disetahunkan
                    If rowItemTax.tipe_tax = TipeElemenPajak_TunjPPHFinalPesangon Or _
                    Me.rowItemTax.tipe_tax = TipeElemenPajak_Pesangon Or _
                    Me.rowItemTax.tipe_tax = TipeElemenPajak_PPHFinalPesangon Then
                        .Item_Group = CInt(TipeElemenPajak_PPHTerhutang)
                    Else
                        .Item_Group = Me.rowItemTax.tipe_tax
                    End If
                    .Item_Group_Desc = Me.GetTaxElementString(Me.rowItemTax.tipe_tax)
                    .Remarks_Gabungan = Me.rowItemTax.remarks_gabungan
                    If Me.rowItemTax.remarks_rutin <> "" Then
                        If .Remarks_Gabungan = "" Then
                            .Remarks_Gabungan += Me.rowItemTax.remarks_rutin
                        Else
                            .Remarks_Gabungan += vbCrLf + Me.rowItemTax.remarks_rutin
                        End If
                    End If
                    If rowItemTax.tipe_tax = CInt(TipeElemenPajak_PPHTerhutang) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_JumlahPenghasilanUntukPerhitunganPPh) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_JumlahPengurang) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_PenghasilaBruto) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_PenghasilanNettoKenaPajak) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_PKP) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_PPH21DipotongMasaPajakSebelum) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_PPHLunas_Terpotong) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_PenghasilanNettoMasaSebelum) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_PPHSetahun) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_PPHTerhutang) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_PTKP) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_Pesangon) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_PPHFinalPesangon) Or _
                    rowItemTax.tipe_tax = CInt(TipeElemenPajak_TunjPPHFinalPesangon) Then
                        .Routine_NotRoutine = ""
                    Else
                        If Me.rowItemTax.is_rutin = 1 Then
                            .Routine_NotRoutine = "Routine"
                        Else
                            .Routine_NotRoutine = "Not Routine"
                        End If
                    End If
                End With
                Me.TBLCetak.Rows.Add(Me.rowItemCetak)

                
                If counter = vTBLItemTax.DefaultView.Count Then
                    Me.TambahkanElemenPajakYgTidakAda(PrimaryGroup1)
                End If
            Next

            'tampilkan cetakan
            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_DefPrec.ParameterFieldName, 0)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Report.SetParameterValue(Report.Parameter_UseBracketsForNegative.ParameterFieldName, 1)
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.Report.SetParameterValue(Report.Parameter_UseBracketsForNegative.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Report
            Me.crv1.Zoom(75)
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanElemenPajakYgTidakAda(ByVal vPrimaryGroup As Double)
        Dim vTaxElement As ClsTransaksiSalary_Item.EnumTaxElement

        Try
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH
            If Me.ApakahAda(vTaxElement, vPrimaryGroup) = False Then
                Me.TambahTaxElemenTidakAda1(vTaxElement, vPrimaryGroup)
            End If
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb
            If Me.ApakahAda(vTaxElement, vPrimaryGroup) = False Then
                Me.TambahTaxElemenTidakAda1(vTaxElement, vPrimaryGroup)
            End If
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya
            If Me.ApakahAda(vTaxElement, vPrimaryGroup) = False Then
                Me.TambahTaxElemenTidakAda1(vTaxElement, vPrimaryGroup)
            End If
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja
            If Me.ApakahAda(vTaxElement, vPrimaryGroup) = False Then
                Me.TambahTaxElemenTidakAda1(vTaxElement, vPrimaryGroup)
            End If
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH
            If Me.ApakahAda(vTaxElement, vPrimaryGroup) = False Then
                Me.TambahTaxElemenTidakAda1(vTaxElement, vPrimaryGroup)
            End If
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi
            If Me.ApakahAda(vTaxElement, vPrimaryGroup) = False Then
                Me.TambahTaxElemenTidakAda1(vTaxElement, vPrimaryGroup)
            End If
            vTaxElement = ClsTransaksiSalary_Item.EnumTaxElement.No10_IuranPensiun_THT_JHT
            If Me.ApakahAda(vTaxElement, vPrimaryGroup) = False Then
                Me.TambahTaxElemenTidakAda1(vTaxElement, vPrimaryGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahTaxElemenTidakAda1(ByVal vTaxElement As ClsTransaksiSalary_Item.EnumTaxElement, ByVal vPrimaryGroup As Double)
        Dim rowItemCetak1 As DseTransaksiSalary.TblCetakPajakRow

        Try
            rowItemCetak1 = Me.TBLCetak.NewTblCetakPajakRow
            With rowItemCetak1
                .Kode_Karyawan = vEmployeeCode
                .Nama_Karyawan = vEmployeeName
                .NIK = vNIK_Kary
                .NPWP = vNPWP
                .Period = vTransactionDate
                .Tax_Category = vStatusNPWP
                .Nationality = vKodeNegaraDomisili
                .Sex = vJenisKelamin
                .Inlay_Following_Section = True
                .Address = vAlamatPajak1
                If vAlamatPajak2 <> "" Then
                    .Address += vbCrLf + vAlamatPajak2
                End If
                .Departement = vDepartement
                .Job_Position = vPosition
                .Tax_Component_Desc = ""
                .Primary_Karyawan = vPrimaryEmployee
                .Primary_Group = vPrimaryGroup
                .Amount = 0
                .Amount_Gabungan = 0
                .Amount_Routine = 0
                .Item_Group = vTaxElement
                .Item_Group_Desc = Me.GetTaxElementString(vTaxElement)
                .Remarks_Gabungan = ""
                .Remarks_Routine = ""
                .Routine_NotRoutine = ""
            End With
            Me.TBLCetak.Rows.Add(rowItemCetak1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahAda(ByVal vTaxElement As ClsTransaksiSalary_Item.EnumTaxElement, ByVal vPrimaryGroup As Double) As Boolean
        Dim DummyFilter As String = ""
        Dim rowCetak1 As DseTransaksiSalary.TblCetakPajakRow
        Dim RetVal As Double = 0

        Try
            If Me.TBLCetak.DefaultView.Sort <> "[Primary Group]" Then
                Me.TBLCetak.DefaultView.Sort = "[Primary Group]"
            End If
            If Me.TBLCetak.DefaultView.RowFilter <> "" Then
                DummyFilter = Me.TBLCetak.DefaultView.RowFilter
                Me.TBLCetak.DefaultView.RowFilter = ""
            End If

            For ij As Double = Me.TBLCetak.DefaultView.Count - 1 To 0 Step -1
                rowCetak1 = Me.TBLCetak.DefaultView(ij).Row
                If rowCetak1.Primary_Group = vPrimaryGroup Then
                    If rowCetak1.Item_Group = CDbl(vTaxElement) Then
                        Return True
                    End If
                ElseIf rowCetak1.Primary_Group < vPrimaryGroup Then
                    Return False
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If DummyFilter <> "" Then
                Me.TBLCetak.DefaultView.RowFilter = DummyFilter
            End If
        End Try
    End Function

    Private Function GetPrimaryGroupMax() As Double
        Dim row1 As DseTransaksiSalary.TblCetakPajakRow
        Dim retVal As Double = Double.MinValue

        Try
            For Each row1 In Me.TBLCetak.Rows
                If retVal < row1.Primary_Group Then
                    retVal = row1.Primary_Group
                End If
            Next

            If retVal = Double.MinValue Then
                Return 0
            Else
                Return retVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function
End Class