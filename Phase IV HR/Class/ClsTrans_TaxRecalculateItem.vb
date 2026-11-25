Option Explicit On
Option Strict Off

Public Class ClsTrans_TaxRecalculateItem

    Dim ObjCd As Core_Module.ClsCekData

    Public Enum EnumTipeItemTax As Integer
        B4_Recalculate
        Aft_Recalculate
    End Enum

    Public Function IsiDGVItem(ByVal vDGVItem As Windows.Forms.DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTrans_TaxRecalculate, ByVal vTBLItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable) As Boolean
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculate_item1TableAdapter
        Dim TBL As New DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1DataTable
        Dim row As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1Row
        Dim DTAItemTax As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_taxrecalculate_itemtaxTableAdapter

        Dim rowDGV As Windows.Forms.DataGridViewRow
        Dim ObjCekData As New Core_Module.ClsCekData
        Dim DefPrec As Integer = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemTax.Connection.ConnectionString = Core_Module.GetConstring

            vTBLItemTax.Rows.Clear()
            DTAItemTax.FillBy_PrimaryTransaksi(vTBLItemTax, vPrimaryTransaksi)

            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            ''isikan(tabel)
            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.Col01Gaji.Index).Value = ObjCD.DigitGrouping(._01Gaji, DefPrec)
                        rowDGV.Cells(Frm.Col01GajiRutin.Index).Value = ObjCD.DigitGrouping(._01_Gaji_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col02NewTunjPPh.Index).Value = ObjCD.DigitGrouping(._02TunjanganPPh_new, DefPrec)
                        rowDGV.Cells(Frm.Col02NewTunjPPhDisetahunkan.Index).Value = ObjCD.DigitGrouping(._02_Tunjangan_PPH_Disetahunkan_new, DefPrec)
                        rowDGV.Cells(Frm.Col02NewTunjPPhDisetahunkanRutin.Index).Value = ObjCD.DigitGrouping(._02_Tunjangan_PPH_Disetahunkan_Rutin_new, DefPrec)
                        rowDGV.Cells(Frm.Col02NewTunjPPhNonRutin.Index).Value = ObjCD.DigitGrouping(._02_Tunjangan_PPH_Non_Rutin_new, DefPrec)
                        rowDGV.Cells(Frm.Col02NewTunjPPhRutin.Index).Value = ObjCD.DigitGrouping(._02_Tunjangan_PPH_Rutin_new, DefPrec)
                        rowDGV.Cells(Frm.Col02TunjPPh.Index).Value = ObjCD.DigitGrouping(._02TunjanganPPh, DefPrec)
                        rowDGV.Cells(Frm.Col02TunjPPhDisetahunkan.Index).Value = ObjCD.DigitGrouping(._02_Tunjangan_PPH_Disetahunkan, DefPrec)
                        rowDGV.Cells(Frm.Col02TunjPPhDisetahunkanRutin.Index).Value = ObjCD.DigitGrouping(._02_Tunjangan_PPH_Disetahunkan_Rutin, DefPrec)
                        rowDGV.Cells(Frm.Col02TunjPPHNonRutin.Index).Value = ObjCD.DigitGrouping(._02_Tunjangan_PPH_Non_Rutin, DefPrec)
                        rowDGV.Cells(Frm.Col02TunjPPhRutin.Index).Value = ObjCD.DigitGrouping(._02_Tunjangan_PPH_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col03TunjLainnya.Index).Value = ObjCD.DigitGrouping(._03TunjanganLainnya, DefPrec)
                        rowDGV.Cells(Frm.Col03TunjLainnyaRutin.Index).Value = ObjCD.DigitGrouping(._03_Tunjangan_Lainnya_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col04Honorarium.Index).Value = ObjCD.DigitGrouping(._04Honorarium, DefPrec)
                        rowDGV.Cells(Frm.Col04HonorariumRutin.Index).Value = ObjCD.DigitGrouping(._04_Honorarium_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col05PremiAsuransi.Index).Value = ObjCD.DigitGrouping(._05PremiAsuransi, DefPrec)
                        rowDGV.Cells(Frm.Col05PremiAsuransiRutin.Index).Value = ObjCD.DigitGrouping(._05_Premi_Asurasi_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col06Natura.Index).Value = ObjCD.DigitGrouping(._06Natura, DefPrec)
                        rowDGV.Cells(Frm.Col06NaturaRutin.Index).Value = ObjCD.DigitGrouping(._06_Natura_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col07Tantiem.Index).Value = ObjCD.DigitGrouping(._07Tantiem, DefPrec)
                        rowDGV.Cells(Frm.Col07TantiemRutin.Index).Value = ObjCD.DigitGrouping(._07_Tantiem_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col08NewPenghasilanBruto.Index).Value = ObjCD.DigitGrouping(._08PenghasilanBruto_new, DefPrec)
                        rowDGV.Cells(Frm.Col08NewPenghasilanBrutoRutin.Index).Value = ObjCD.DigitGrouping(._08_Penghasilan_Bruto_Rutin_new, DefPrec)
                        rowDGV.Cells(Frm.Col08PenghasilanBruto.Index).Value = ObjCD.DigitGrouping(._08PenghasilanBruto, DefPrec)
                        rowDGV.Cells(Frm.Col08PenghasilanBrutoRutin.Index).Value = ObjCD.DigitGrouping(._09_Biaya_Jabatan_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col09BiayaJabatan.Index).Value = ObjCD.DigitGrouping(._09BiayaJabatan, DefPrec)
                        rowDGV.Cells(Frm.Col09BiayaJabatanRutin.Index).Value = ObjCD.DigitGrouping(._09_Biaya_Jabatan_Rutin, DefPrec)
                        rowDGV.Cells(Frm.Col09NewBiayaJabatan.Index).Value = ObjCD.DigitGrouping(._09BiayaJabatan_new, DefPrec)
                        rowDGV.Cells(Frm.Col09NewBiayaJabatanRutin.Index).Value = ObjCD.DigitGrouping(._09_Biaya_Jabatan_Rutin_new, DefPrec)

                        rowDGV.Cells(Frm.Col10IuranPensiun.Index).Value = ObjCD.DigitGrouping(._10IuranPensiun, DefPrec)
                        rowDGV.Cells(Frm.Col10IuranPensiunRutin.Index).Value = ObjCD.DigitGrouping(._10_Iuran_Pensiun_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col11Jumlahengurang.Index).Value = ObjCD.DigitGrouping(._11JumlahPengurangan, DefPrec)
                        rowDGV.Cells(Frm.Col11JumlahPengurangRutin.Index).Value = ObjCD.DigitGrouping(._11_Jumlah_Pengurang_Rutin, DefPrec)
                        rowDGV.Cells(Frm.Col11NewJumlahPengurang.Index).Value = ObjCD.DigitGrouping(._11JumlahPengurangan_new, DefPrec)
                        rowDGV.Cells(Frm.Col11NewJumlahPengurangRutin.Index).Value = ObjCD.DigitGrouping(._11_Jumlah_Pengurang_Rutin_new, DefPrec)

                        rowDGV.Cells(Frm.Col14NewPenghasilanNettoUntPPh.Index).Value = ObjCD.DigitGrouping(._14JumlahPenghasilanNettoUtkPPh_new, DefPrec)
                        rowDGV.Cells(Frm.Col14NewPenghasilanNettoUntPPhRutin.Index).Value = ObjCD.DigitGrouping(._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin_new, DefPrec)
                        rowDGV.Cells(Frm.Col14PenghasilanNettoPPhRutin.Index).Value = ObjCD.DigitGrouping(._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin, DefPrec)
                        rowDGV.Cells(Frm.Col14PenghasilanNettoUntPPH.Index).Value = ObjCD.DigitGrouping(._14JumlahPenghasilanNettoUtkPPh, DefPrec)

                        rowDGV.Cells(Frm.Col15NewPTKP.Index).Value = ObjCD.DigitGrouping(._15PTKP_new, DefPrec)
                        rowDGV.Cells(Frm.Col15PTKP.Index).Value = ObjCD.DigitGrouping(._15PTKP_new, DefPrec)

                        rowDGV.Cells(Frm.Col16NewPKPSetahun.Index).Value = ObjCD.DigitGrouping(._16PKPSetahun_new, DefPrec)
                        rowDGV.Cells(Frm.Col16NewPKPSetahunRutin.Index).Value = ObjCD.DigitGrouping(._16_PKP_Setahun_Rutin_new, DefPrec)
                        rowDGV.Cells(Frm.Col16PKPSetahun.Index).Value = ObjCD.DigitGrouping(._16PKPSetahun, DefPrec)
                        rowDGV.Cells(Frm.Col16PKPSetahunRutin.Index).Value = ObjCD.DigitGrouping(._16_PKP_Setahun_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col17NewPPh21.Index).Value = ObjCD.DigitGrouping(._17PPh21_new, DefPrec)
                        rowDGV.Cells(Frm.Col17NewPPh21Rutin.Index).Value = ObjCD.DigitGrouping(._17_PPH_21_Rutin_new, DefPrec)
                        rowDGV.Cells(Frm.Col17PPH21.Index).Value = ObjCD.DigitGrouping(._17PPh21, DefPrec)
                        rowDGV.Cells(Frm.Col17PPH21Rutin.Index).Value = ObjCD.DigitGrouping(._17_PPH_21_Rutin, DefPrec)
                        rowDGV.Cells(Frm.Col18PPh21YgTelahDipotongMasaPjkSblmnya.Index).Value = ObjCD.DigitGrouping(._18PPh21YgTelahDipotongMasaPjkSblmnya, DefPrec)
                        rowDGV.Cells(Frm.Col18NewPPh21YgTelahDipotongMasaPjkSblmnya.Index).Value = ObjCD.DigitGrouping(._18PPh21YgTelahDipotongMasaPjkSblmnya_new, DefPrec)

                        rowDGV.Cells(Frm.Col19NewPPhTerutang.Index).Value = ObjCD.DigitGrouping(._19PPhTerhutang_new, DefPrec)
                        rowDGV.Cells(Frm.Col19NewPPhTerutangDisetahunkan.Index).Value = ObjCD.DigitGrouping(._19_PPH_Terhutang_Setahun_new, DefPrec)
                        rowDGV.Cells(Frm.Col19NewPPHTerutangDisetahunkanRutin.Index).Value = ObjCD.DigitGrouping(._19_PPH_Terhutang_Setahun_Rutin_new, DefPrec)
                        rowDGV.Cells(Frm.Col19NewPPhTerutangSebulanNonRutin.Index).Value = ObjCD.DigitGrouping(._19_PPH_Terhutang_Sebulan_Non_Rutin_new, DefPrec)
                        rowDGV.Cells(Frm.Col19NewPPhTerutangSebulanRutin.Index).Value = ObjCD.DigitGrouping(._19_PPH_Terutang_Sebulan_Rutin_new, DefPrec)
                        rowDGV.Cells(Frm.Col19PPHTerutang.Index).Value = ObjCD.DigitGrouping(._19PPhTerhutang, DefPrec)
                        rowDGV.Cells(Frm.Col19PPHTerutangDisetahunakan.Index).Value = ObjCD.DigitGrouping(._19_PPH_Terhutang_Setahun, DefPrec)
                        rowDGV.Cells(Frm.Col19PPhTerutangDisetahunkanRutin.Index).Value = ObjCD.DigitGrouping(._19_PPH_Terhutang_Setahun_Rutin, DefPrec)
                        rowDGV.Cells(Frm.Col19PPhTerutangSebulanRutin.Index).Value = ObjCD.DigitGrouping(._19_PPH_Terutang_Sebulan_Rutin, DefPrec)
                        rowDGV.Cells(Frm.Col19PPHTerutangSebulanNonRutin.Index).Value = ObjCD.DigitGrouping(._19_PPH_Terhutang_Sebulan_Non_Rutin, DefPrec)

                        rowDGV.Cells(Frm.Col21NewPPhPesangon.Index).Value = ObjCD.DigitGrouping(._21_PPH_Final_Pesangon_new, DefPrec)
                        rowDGV.Cells(Frm.Col21NewTarifPesangon.Index).Value = ObjCD.DigitGrouping(.Tarif_Pesangon_new, 2)
                        rowDGV.Cells(Frm.Col21NewTunjPesangon.Index).Value = ObjCD.DigitGrouping(._21_Tunj_PPH_FInal_Pesangon_new, DefPrec)
                        rowDGV.Cells(Frm.Col21Pesangon.Index).Value = ObjCD.DigitGrouping(._21_Pesangon, DefPrec)
                        rowDGV.Cells(Frm.Col21PPhPesangon.Index).Value = ObjCD.DigitGrouping(._21_PPH_Final_Pesangon, DefPrec)
                        rowDGV.Cells(Frm.Col21TarifPesagon.Index).Value = ObjCD.DigitGrouping(.Tarif_Pesangon, 2)
                        rowDGV.Cells(Frm.Col21TunjPPhPesangon.Index).Value = ObjCD.DigitGrouping(._21_Tunj_PPH_FInal_Pesangon, DefPrec)

                        rowDGV.Cells(Frm.ColPenghasilBruto.Index).Value = ObjCD.DigitGrouping(.Penghasilan_Bruto, DefPrec)
                        rowDGV.Cells(Frm.ColNewPenghasilanBruto.Index).Value = ObjCD.DigitGrouping(.Penghasilan_Bruto_new, DefPrec)

                        rowDGV.Cells(Frm.ColIncomeTax.Index).Value = ObjCD.DigitGrouping(.Income_Tax, DefPrec)
                        rowDGV.Cells(Frm.ColNewIncomeTax.Index).Value = ObjCD.DigitGrouping(.Income_Tax_New, DefPrec)
                        rowDGV.Cells(Frm.ColKurangBayar.Index).Value = ObjCD.DigitGrouping(.Kurang_Bayar, DefPrec)
                        rowDGV.Cells(Frm.ColNewKurangBayar.Index).Value = ObjCD.DigitGrouping(.Kurang_Bayar_New, DefPrec)

                        rowDGV.Cells(Frm.ColAmount.Index).Value = .Income_Tax
                        rowDGV.Cells(Frm.ColDate.Index).Value = .Tanggal_Transaksi
                        rowDGV.Cells(Frm.ColDifference.Index).Value = .Difference
                        rowDGV.Cells(Frm.ColDocNo.Index).Value = .No_Transaksi
                        rowDGV.Cells(Frm.ColEmployee.Index).Value = .Full_Name
                        rowDGV.Cells(Frm.ColNo.Index).Value = vDGVItem.RowCount
                        rowDGV.Cells(Frm.ColPrimaryEmployee.Index).Value = .primary_employee
                        rowDGV.Cells(Frm.ColPrimarySalary.Index).Value = .primary_salary
                        rowDGV.Cells(Frm.ColRecalculatedAmount.Index).Value = .Income_Tax_New
                        rowDGV.Cells(Frm.ColNPWP.Index).Value = .A_NPWP
                        rowDGV.Cells(Frm.ColNewNPWP.Index).Value = .A_NPWP_new
                        rowDGV.Cells(Frm.ColStatusNPWP.Index).Value = .A_Status_NPWP
                        rowDGV.Cells(Frm.ColNewStatusNPWP.Index).Value = .A_Status_NPWP_new
                        rowDGV.Cells(Frm.ColAlamat1.Index).Value = .A_Alamat1
                        rowDGV.Cells(Frm.ColAlamat2.Index).Value = .A_Alamat2
                        rowDGV.Cells(Frm.ColNewAlamat1.Index).Value = .A_Alamat1_new
                        rowDGV.Cells(Frm.ColNewAlamat2.Index).Value = .A_Alamat2_new
                        rowDGV.Cells(Frm.ColTipePajak.Index).Value = .Tipe_Tax
                    End With
                Next
            End If

            Frm.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Frm.ColRecalculatedAmount.DefaultCellStyle.Format = "N0"
            Frm.ColAmount.DefaultCellStyle.Format = "N0"
            Frm.ColDifference.DefaultCellStyle.Format = "N0"

            'set nomer
            'ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAItemTax.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Public Overridable Function SimpanItem(ByVal vPrimaryTransaksi As Double, ByVal vDGV As Windows.Forms.DataGridView, _
  ByVal Frm As FrmTrans_TaxRecalculate, ByVal vTBLItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable) As Boolean
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculate_itemTableAdapter
        Dim DTAItemTax As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_taxrecalculate_itemtaxTableAdapter
        Dim row As System.Windows.Forms.DataGridViewRow
        Dim row1 As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow
        Dim remark As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemTax.Connection.ConnectionString = Core_Module.GetConstring

            'insert item
            For Each row In vDGV.Rows
                DTA.Insert(vPrimaryTransaksi, row.Cells(Frm.ColPrimarySalary.Index).Value, row.Cells(Frm.ColPrimaryEmployee.Index).Value, row.Cells(Frm.Col01Gaji.Index).Value, _
                row.Cells(Frm.Col01GajiRutin.Index).Value, row.Cells(Frm.Col02TunjPPh.Index).Value, row.Cells(Frm.Col02TunjPPHNonRutin.Index).Value, row.Cells(Frm.Col02TunjPPhRutin.Index).Value, _
                row.Cells(Frm.Col02TunjPPhDisetahunkan.Index).Value, row.Cells(Frm.Col02TunjPPhDisetahunkanRutin.Index).Value, row.Cells(Frm.Col03TunjLainnya.Index).Value, _
                row.Cells(Frm.Col03TunjLainnyaRutin.Index).Value, row.Cells(Frm.Col04Honorarium.Index).Value, row.Cells(Frm.Col04HonorariumRutin.Index).Value, _
                row.Cells(Frm.Col05PremiAsuransi.Index).Value, row.Cells(Frm.Col05PremiAsuransiRutin.Index).Value, row.Cells(Frm.Col06Natura.Index).Value, _
                row.Cells(Frm.Col06NaturaRutin.Index).Value, row.Cells(Frm.Col07Tantiem.Index).Value, row.Cells(Frm.Col07TantiemRutin.Index).Value, _
                row.Cells(Frm.Col08PenghasilanBruto.Index).Value, row.Cells(Frm.Col08PenghasilanBrutoRutin.Index).Value, row.Cells(Frm.Col09BiayaJabatan.Index).Value, _
                row.Cells(Frm.Col09BiayaJabatanRutin.Index).Value, row.Cells(Frm.Col10IuranPensiun.Index).Value, row.Cells(Frm.Col10IuranPensiunRutin.Index).Value, _
                row.Cells(Frm.Col11Jumlahengurang.Index).Value, row.Cells(Frm.Col11JumlahPengurangRutin.Index).Value, row.Cells(Frm.Col14PenghasilanNettoUntPPH.Index).Value, _
                row.Cells(Frm.Col14PenghasilanNettoPPhRutin.Index).Value, row.Cells(Frm.Col15PTKP.Index).Value, row.Cells(Frm.Col16PKPSetahun.Index).Value, _
                row.Cells(Frm.Col16PKPSetahunRutin.Index).Value, row.Cells(Frm.Col17PPH21.Index).Value, row.Cells(Frm.Col17PPH21Rutin.Index).Value, _
                row.Cells(Frm.Col19PPHTerutang.Index).Value, row.Cells(Frm.Col19PPHTerutangDisetahunakan.Index).Value, row.Cells(Frm.Col19PPhTerutangDisetahunkanRutin.Index).Value, _
                row.Cells(Frm.Col19PPHTerutangSebulanNonRutin.Index).Value, row.Cells(Frm.Col19PPhTerutangSebulanRutin.Index).Value, row.Cells(Frm.ColNPWP.Index).Value, _
                row.Cells(Frm.ColAlamat1.Index).Value, row.Cells(Frm.ColAlamat2.Index).Value, row.Cells(Frm.ColStatusNPWP.Index).Value, row.Cells(Frm.Col21Pesangon.Index).Value, _
                row.Cells(Frm.Col21TunjPPhPesangon.Index).Value, row.Cells(Frm.Col21PPhPesangon.Index).Value, row.Cells(Frm.Col21TarifPesagon.Index).Value, row.Cells(Frm.ColPenghasilBruto.Index).Value, _
                row.Cells(Frm.Col02NewTunjPPh.Index).Value, row.Cells(Frm.Col02NewTunjPPhNonRutin.Index).Value, row.Cells(Frm.Col02NewTunjPPhRutin.Index).Value, _
                row.Cells(Frm.Col02TunjPPhDisetahunkan.Index).Value, row.Cells(Frm.Col02TunjPPhDisetahunkanRutin.Index).Value, row.Cells(Frm.Col08NewPenghasilanBruto.Index).Value, _
                row.Cells(Frm.Col08PenghasilanBrutoRutin.Index).Value, row.Cells(Frm.Col09NewBiayaJabatan.Index).Value, row.Cells(Frm.Col09NewBiayaJabatanRutin.Index).Value, _
                row.Cells(Frm.Col11NewJumlahPengurang.Index).Value, row.Cells(Frm.Col11NewJumlahPengurangRutin.Index).Value, row.Cells(Frm.Col14NewPenghasilanNettoUntPPh.Index).Value, _
                row.Cells(Frm.Col14NewPenghasilanNettoUntPPhRutin.Index).Value, row.Cells(Frm.Col15NewPTKP.Index).Value, row.Cells(Frm.Col16NewPKPSetahun.Index).Value, _
                row.Cells(Frm.Col16NewPKPSetahunRutin.Index).Value, row.Cells(Frm.Col17NewPPh21.Index).Value, row.Cells(Frm.Col17NewPPh21Rutin.Index).Value, row.Cells(Frm.Col19NewPPhTerutang.Index).Value, _
                row.Cells(Frm.Col19NewPPhTerutangDisetahunkan.Index).Value, row.Cells(Frm.Col19NewPPHTerutangDisetahunkanRutin.Index).Value, row.Cells(Frm.Col19NewPPhTerutangSebulanNonRutin.Index).Value, _
                row.Cells(Frm.Col19NewPPhTerutangSebulanRutin.Index).Value, row.Cells(Frm.ColNewNPWP.Index).Value, row.Cells(Frm.ColNewAlamat1.Index).Value, row.Cells(Frm.ColNewAlamat2.Index).Value, _
                row.Cells(Frm.ColNewStatusNPWP.Index).Value, row.Cells(Frm.ColTipePajak.Index).Value, row.Cells(Frm.Col21NewTunjPesangon.Index).Value, _
                row.Cells(Frm.Col21NewPPhPesangon.Index).Value, row.Cells(Frm.ColDifference.Index).Value, row.Cells(Frm.Col21NewTarifPesangon.Index).Value, _
                row.Cells(Frm.ColNewPenghasilanBruto.Index).Value, row.Cells(Frm.Col18PPh21YgTelahDipotongMasaPjkSblmnya.Index).Value, row.Cells(Frm.Col18NewPPh21YgTelahDipotongMasaPjkSblmnya.Index).Value, _
                row.Cells(Frm.ColIncomeTax.Index).Value, row.Cells(Frm.ColNewIncomeTax.Index).Value, row.Cells(Frm.ColKurangBayar.Index).Value, row.Cells(Frm.ColNewKurangBayar.Index).Value)
            Next

            'insert item tax
            For Each row1 In vTBLItemTax.Rows
                With row1
                    DTAItemTax.Insert(vPrimaryTransaksi, .primary_salary, .tipe_tax, .identifier_salary_component, .is_rutin, .amount_month, .amount_gabungan_disetahunkan, .remarks_gabungan, _
                    .amount_rutin_disetahunkan, .remarks_rutin, .tipe_item_tax)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculate_itemTableAdapter
        Dim DTA1 As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_taxrecalculate_itemtaxTableAdapter

        Try
            'hapus data item
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Delete(vPrimaryTransaksi)
            DTA.Delete(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
