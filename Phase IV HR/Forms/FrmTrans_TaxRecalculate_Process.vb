Public Class FrmTrans_TaxRecalculate_Process
    Dim Selesai As Boolean = False
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjKrywn As New Core_Module.ClsMasterKaryawan
    Dim Obj As New ClsTrans_TaxRecalculate
    Dim ObjSalary As New ClsTransaksiSalary
    Dim ObjPTKP As New HR.ClsMasterSettingPajak
    Dim DefPrec As Integer = 0
    Dim Step1 As Integer = 0
    Dim Step2 As Integer = 5
    Dim TipeTaxCalculation As ClsTransaksiSalary.EnumTax = ClsTransaksiSalary.EnumTax.Gross
    Dim DTATarif As New DseTransaksiSalaryTableAdapters.tblmaster_tarif_pphTableAdapter
    Dim vTBLTarif As New DseTransaksiSalary.tblmaster_tarif_pphDataTable
    Dim DTATarifPPHFinal As New DseTransaksiSalaryTableAdapters.tblmaster_tarif_pph_final_pesangonTableAdapter
    Dim vTBLTarifPPHFinal As New DseTransaksiSalary.tblmaster_tarif_pph_final_pesangonDataTable
    Dim TBLItemIterasi As New DseTransaksiSalary.tbltrans_salary_item_iterationDataTable
    Dim TBLItemIterasiPPHFinal As New DseTransaksiSalary.tbltrans_salary_item_iteration1DataTable
    Dim Perct_BiayaJabatan As Decimal = 0.05
    Dim BiayaJabatanMaxValue As Decimal = 6000000

    Public vTBLItem As New DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1DataTable
    Public vTBLItemTax As New DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable
    Public vPeriod As Date = #1/1/1950#

    Private Sub FrmTrans_TaxRecalculate_Process_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            DTATarif.Connection.ConnectionString = Core_Module.GetConstring
            Me.vTBLTarif.Rows.Clear()
            DTATarif.Fill(Me.vTBLTarif)

            DTATarifPPHFinal.Connection.ConnectionString = Core_Module.GetConstring
            Me.vTBLTarifPPHFinal.Rows.Clear()
            DTATarifPPHFinal.Fill(Me.vTBLTarifPPHFinal)

            Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'Dim rowItem As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1Row
        'Dim rowItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow
        Dim DTASalary As New DseTrans_Tax_RecalculateTableAdapters.DataTable1TableAdapter
        Dim TBLSalary As New DseTrans_Tax_Recalculate.DataTable1DataTable
        Dim DTAKary As New DseTrans_Tax_RecalculateTableAdapters.tblmaster_karyawanTableAdapter
        Dim TBLKary As New DseTrans_Tax_Recalculate.tblmaster_karyawanDataTable
        'Dim rowKary As DseTrans_Tax_Recalculate.tblmaster_karyawanRow
        Dim DTALama As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculate_item1TableAdapter
        Dim TBLLama As New DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1DataTable
        'Dim rowLama As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1Row
        Dim DTAItemTaxLama As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_taxrecalculate_itemtaxTableAdapter
        Dim TBLItemTaxLama As New DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable
        'Dim rowItemTaxLama As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow
        Dim vPrimaryTaxRecalculate As Double = -1
        Dim DTASalaryItemTax As New DseTrans_Tax_RecalculateTableAdapters.tbltrans_salary_itemtaxTableAdapter
        Dim TBLSalaryItemTax As New DseTrans_Tax_Recalculate.tbltrans_salary_itemtaxDataTable
        'Dim DV As DataRowView

        Try
            DTASalary.Connection.ConnectionString = Core_Module.GetConstring
            DTAKary.Connection.ConnectionString = Core_Module.GetConstring
            DTALama.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemTaxLama.Connection.ConnectionString = Core_Module.GetConstring
            DTASalaryItemTax.Connection.ConnectionString = Core_Module.GetConstring

            ObjCD.UpdateThreading()
            Me.ProgressBar1.Value = 0
            Me.vTBLItem.Rows.Clear()
            Me.vTBLItemTax.Rows.Clear()

            'kumpulkan data karyawan
            Me.Step1 += 1
            Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + " : Loading employee data"
            DTAKary.Fill(TBLKary)

            'kumpulkan semua salary yang ada sesuai dengan periode kompensasi
            Me.KumpulkanDataSalary(DTASalary, TBLSalary, TBLKary)

            'ganti dengan data tax recalculate jika hitung ulang
            vPrimaryTaxRecalculate = Obj.GetPrimaryByPeriod(Me.vPeriod)
            Me.UpdateDataSalary(vPrimaryTaxRecalculate, DTALama, TBLLama, TBLSalary)

            'masukkan data item tax, dari data lama + dari item salary
            Me.KumpulkanDataItemTax(DTASalaryItemTax, TBLSalaryItemTax, DTAItemTaxLama, TBLItemTaxLama, _
            vPrimaryTaxRecalculate, TBLSalary)

            'hitung ulang pajak dan perhitungkan differencenya
            Me.Step1 += 1
            Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + " : Recalculating Tax"
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = Me.vTBLItem.Rows.Count
            For ij As Integer = 0 To Me.vTBLItem.Rows.Count - 1
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 500 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 25) = 0 Then
                        Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                        " : Recalculating Tax (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                    End If
                Else
                    Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                    " : Recalculating Tax (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                End If

                Me.HitungPajak1(ij)
            Next

            Me.Selesai = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASalary.Dispose()
            TBLSalary.Dispose()
            DTAKary.Dispose()
            TBLKary.Dispose()
            DTALama.Dispose()
            TBLLama.Dispose()
            DTAItemTaxLama.Dispose()
            TBLItemTaxLama.Dispose()
            DTASalaryItemTax.Dispose()
            TBLSalaryItemTax.Dispose()
        End Try
    End Sub

    Private Sub HitungPajak1(ByVal vIndex As Integer)
        'untuk sementara ditutup dulu, karena update cara perhitungan III dan terkait subjective tax obligation
        '2019-01-05
        MsgBox("Logic error: Incompatible parameter. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
        'Dim rowRef As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1Row
        ''Dim rowDGV As Windows.Forms.DataGridViewRow
        'Dim rowItemIterasi, rowItemIterasiSebelum As DseTransaksiSalary.tbltrans_salary_item_iterationRow
        ''Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        'Dim TipeSalaryKomponen As Core_Module.FrmMasterSalaryFormula_Component.EnumType = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income
        ''Dim TipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement
        'Dim JmlIterasi As Double = 0
        'Dim Selisih1, Selisih2 As Decimal
        'Dim UseNPWP As Core_Module.ClsMasterKaryawan.EnumIsNPWP = Core_Module.ClsMasterKaryawan.EnumIsNPWP.Yes
        'Dim FaktorKaliDisetahunkan As Integer = 12
        'Dim StartWorkingDate As Date = Today
        'Dim ResignDate As Date = #1/1/1950#
        'Dim vResign As Boolean = False
        'Dim vTBLOtherPayment1 As New DseTransaksiSalary.tbltrans_salary_item_other1DataTable
        'Dim NPWPCategory As Core_Module.ClsMasterKaryawan.GolonganPajak = Core_Module.ClsMasterKaryawan.GolonganPajak.K0
        'Dim TaxCategoryDesc As String = ""
        'Dim Txt12PjkJmlPenghasilanNetto As Decimal = 0
        'Dim Txt13PjkJmlHasilMsPjkSblm As Decimal = 0
        'Dim Txt20 As Decimal = 0
        'Dim TBLItemTaxDummy As New DseTransaksiSalary.tbltrans_salary_itemtaxDataTable
        'Dim rowItemTaxDummy As DseTransaksiSalary.tbltrans_salary_itemtaxRow

        'Try
        '    rowRef = vTBLItem.Rows(vIndex)
        '    Me.SalinKeTabelItemTaxDummy(rowRef.primary_salary, TBLItemTaxDummy, rowRef.No_Transaksi, rowRef.primary_employee, rowRef.Tanggal_Transaksi)

        '    ResignDate = Me.ObjKrywn.GetResignDate(rowRef.primary_employee)

        '    rowRef.Penghasilan_Bruto_new = 0
        '    For Each rowItemTaxDummy In TBLItemTaxDummy.Rows
        '        'klo disini tunjangan PPH tidak masuk, beda dengan form salary
        '        'tunjangan pph, nanti diperhitungkan dibawah
        '        If rowItemTaxDummy.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT Or _
        '        rowItemTaxDummy.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb Or _
        '        rowItemTaxDummy.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya Or _
        '        rowItemTaxDummy.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja Or _
        '        rowItemTaxDummy.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH Or _
        '        rowItemTaxDummy.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi Then
        '            rowRef.Penghasilan_Bruto_new += rowItemTaxDummy.amount_month
        '        End If
        '    Next

        '    StartWorkingDate = Me.ObjKrywn.GetFirstTaxPeriod(rowRef.primary_employee)

        '    'If rowRef.primary_salary = 82 Then
        '    '    MsgBox("")
        '    'End If
        '    If ResignDate >= rowRef.Start_Date.Date And ResignDate <= rowRef.End_Date Then
        '        vResign = True
        '    Else
        '        vResign = False
        '    End If

        '    If StartWorkingDate.Date.Year < Me.ObjCD.GetFirstDateofTheMonth(rowRef.End_Date.Date).Date.Year Then
        '        If ResignDate = #1/1/1950# Then
        '            FaktorKaliDisetahunkan = 12
        '        Else
        '            If ResignDate.Date <= rowRef.End_Date.Date And ResignDate >= rowRef.Start_Date.Date Then
        '                FaktorKaliDisetahunkan = ResignDate.Month
        '            Else
        '                FaktorKaliDisetahunkan = 12
        '            End If
        '        End If
        '    ElseIf StartWorkingDate.Date.Year = Me.ObjCD.GetFirstDateofTheMonth(rowRef.End_Date.Date).Date.Year Then
        '        If ResignDate = #1/1/1950# Then
        '            If StartWorkingDate.Month = 1 Then
        '                FaktorKaliDisetahunkan = 12
        '            ElseIf StartWorkingDate.Month = 2 Then
        '                FaktorKaliDisetahunkan = 11
        '            ElseIf StartWorkingDate.Month = 3 Then
        '                FaktorKaliDisetahunkan = 10
        '            ElseIf StartWorkingDate.Month = 4 Then
        '                FaktorKaliDisetahunkan = 9
        '            ElseIf StartWorkingDate.Month = 5 Then
        '                FaktorKaliDisetahunkan = 8
        '            ElseIf StartWorkingDate.Month = 6 Then
        '                FaktorKaliDisetahunkan = 7
        '            ElseIf StartWorkingDate.Month = 7 Then
        '                FaktorKaliDisetahunkan = 6
        '            ElseIf StartWorkingDate.Month = 8 Then
        '                FaktorKaliDisetahunkan = 5
        '            ElseIf StartWorkingDate.Month = 9 Then
        '                FaktorKaliDisetahunkan = 4
        '            ElseIf StartWorkingDate.Month = 10 Then
        '                FaktorKaliDisetahunkan = 3
        '            ElseIf StartWorkingDate.Month = 11 Then
        '                FaktorKaliDisetahunkan = 2
        '            ElseIf StartWorkingDate.Month = 12 Then
        '                FaktorKaliDisetahunkan = 1
        '            End If
        '        Else
        '            If ResignDate.Date <= rowRef.End_Date.Date And ResignDate >= rowRef.Start_Date.Date Then
        '                FaktorKaliDisetahunkan = ResignDate.Month - StartWorkingDate.Month + 1
        '            Else
        '                If StartWorkingDate.Month = 1 Then
        '                    FaktorKaliDisetahunkan = 12
        '                ElseIf StartWorkingDate.Month = 2 Then
        '                    FaktorKaliDisetahunkan = 11
        '                ElseIf StartWorkingDate.Month = 3 Then
        '                    FaktorKaliDisetahunkan = 10
        '                ElseIf StartWorkingDate.Month = 4 Then
        '                    FaktorKaliDisetahunkan = 9
        '                ElseIf StartWorkingDate.Month = 5 Then
        '                    FaktorKaliDisetahunkan = 8
        '                ElseIf StartWorkingDate.Month = 6 Then
        '                    FaktorKaliDisetahunkan = 7
        '                ElseIf StartWorkingDate.Month = 7 Then
        '                    FaktorKaliDisetahunkan = 6
        '                ElseIf StartWorkingDate.Month = 8 Then
        '                    FaktorKaliDisetahunkan = 5
        '                ElseIf StartWorkingDate.Month = 9 Then
        '                    FaktorKaliDisetahunkan = 4
        '                ElseIf StartWorkingDate.Month = 10 Then
        '                    FaktorKaliDisetahunkan = 3
        '                ElseIf StartWorkingDate.Month = 11 Then
        '                    FaktorKaliDisetahunkan = 2
        '                ElseIf StartWorkingDate.Month = 12 Then
        '                    FaktorKaliDisetahunkan = 1
        '                End If
        '            End If
        '        End If
        '    Else
        '        FaktorKaliDisetahunkan = 12
        '    End If

        '    'kosongkan item tax detail dan item iterasi
        '    Me.TBLItemIterasi.Clear()

        '    'hitung detail pajak untuk iterasi 0
        '    UseNPWP = Me.ObjKrywn.GetIsNPWP(rowRef.primary_employee)
        '    NPWPCategory = Me.ObjKrywn.GetNPWPCategory(rowRef.primary_employee)
        '    TaxCategoryDesc = Me.ObjPTKP.GetDesc(NPWPCategory)
        '    rowRef._15PTKP_new = Me.ObjPTKP.GetPTKPPerTahun(NPWPCategory)
        '    'MsgBox(rowRef._15PTKP_new)
        '    If rowRef.End_Date.Month = 12 Or vResign = True Then
        '        If Format(StartWorkingDate, "MMM yyyy") <> Format(ResignDate, "MMM yyyy") Then
        '            If ResignDate = #1/1/1950# Then
        '                ResignDate = rowRef.End_Date
        '            End If
        '            rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new = Me.ObjSalary.GetPPHTerutang_Rutin_NonRutin(rowRef.primary_employee, StartWorkingDate, _
        '            ObjCD.GetFirstDateofTheMonth(ResignDate) - TimeSpan.FromDays(1))
        '            'If rowRef.primary_salary = 82 Then
        '            '    MsgBox(rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new)
        '            '    MsgBox(StartWorkingDate.ToLongDateString)
        '            '    MsgBox((ObjCD.GetFirstDateofTheMonth(ResignDate) - TimeSpan.FromDays(1)).ToLongDateString)
        '            '    MsgBox(rowRef.primary_employee)
        '            'End If
        '        Else
        '            rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new = 0
        '        End If
        '    End If

        '    If rowRef.End_Date.Month = 12 Or vResign = True Then
        '        Me.ObjSalary.HitungPajak1_Aktual_EndOfYear(0, 0, UseNPWP, Me.vTBLTarif, 0, 0, StartWorkingDate, ResignDate, FaktorKaliDisetahunkan, _
        '        TBLItemTaxDummy, Me.Perct_BiayaJabatan, Me.BiayaJabatanMaxValue, Me.ObjCD, NPWPCategory, TaxCategoryDesc, rowRef.primary_salary, _
        '        rowRef.primary_employee, rowRef.Tanggal_Transaksi, rowRef.No_Transaksi, rowRef.End_Date, rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new, _
        '        rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new, rowRef._02_Tunjangan_PPH_Disetahunkan_new, rowRef._02_Tunjangan_PPH_Rutin_new, _
        '        rowRef._02_Tunjangan_PPH_Non_Rutin_new, rowRef._02TunjanganPPh_new, rowRef._01_Gaji_Rutin, rowRef._01Gaji, _
        '        rowRef._03TunjanganLainnya, rowRef._03_Tunjangan_Lainnya_Rutin, rowRef._04Honorarium, rowRef._04_Honorarium_Rutin, _
        '        rowRef._05PremiAsuransi, rowRef._05_Premi_Asurasi_Rutin, rowRef._06Natura, rowRef._06_Natura_Rutin, _
        '        rowRef._07Tantiem, rowRef._07_Tantiem_Rutin, rowRef._08PenghasilanBruto_new, rowRef._08_Penghasilan_Bruto_Rutin_new, _
        '        rowRef._09BiayaJabatan_new, rowRef._09_Biaya_Jabatan_Rutin_new, rowRef._10IuranPensiun, rowRef._10_Iuran_Pensiun_Rutin, _
        '        rowRef._11JumlahPengurangan_new, rowRef._11_Jumlah_Pengurang_Rutin_new, Txt12PjkJmlPenghasilanNetto, Txt13PjkJmlHasilMsPjkSblm, _
        '        rowRef._14JumlahPenghasilanNettoUtkPPh_new, rowRef._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin_new, rowRef._15PTKP_new, rowRef._16PKPSetahun_new, _
        '        rowRef._16_PKP_Setahun_Rutin_new, rowRef._17PPh21_new, rowRef._17_PPH_21_Rutin_new, rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new, rowRef._19_PPH_Terhutang_Setahun_new, _
        '        rowRef._19_PPH_Terhutang_Setahun_Rutin_new, rowRef._19_PPH_Terhutang_Sebulan_Non_Rutin_new, rowRef._19_PPH_Terutang_Sebulan_Rutin_new, _
        '        rowRef._19PPhTerhutang_new, rowRef.Income_Tax_New, rowRef.Kurang_Bayar_New, Txt20)
        '        rowItemIterasi = Me.TBLItemIterasi.Newtbltrans_salary_item_iterationRow
        '        With rowItemIterasi
        '            .PPH_Terutang_Non_Rutin = CDec(rowRef._19_PPH_Terhutang_Setahun_new) - CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new)
        '            .PPH_Terutang_Rutin = CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new)
        '            .primary = Me.GetPrimaryMax(Me.TBLItemIterasi)
        '            .primary_transaksi = rowRef.primary_salary
        '            .Tunjangan_PPH_Rutin = CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new)
        '            ._Tunjangan_PPH_Non_Rutin = CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_new) - CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new)
        '        End With
        '        Me.TBLItemIterasi.Rows.Add(rowItemIterasi)
        '    Else
        '        Me.ObjSalary.HitungPajak1_Aktual(UseNPWP, Me.vTBLTarif, FaktorKaliDisetahunkan, _
        '        TBLItemTaxDummy, Me.Perct_BiayaJabatan, Me.BiayaJabatanMaxValue, Me.ObjCD, NPWPCategory, TaxCategoryDesc, _
        '       rowRef.primary_salary, rowRef.primary_employee, rowRef.Tanggal_Transaksi, rowRef.No_Transaksi, 0, _
        '        0, rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new, rowRef._02_Tunjangan_PPH_Disetahunkan_new, rowRef._02_Tunjangan_PPH_Rutin_new, _
        '        rowRef._02_Tunjangan_PPH_Non_Rutin_new, rowRef._02TunjanganPPh_new, rowRef._01_Gaji_Rutin, rowRef._01Gaji, _
        '        rowRef._03TunjanganLainnya, rowRef._03_Tunjangan_Lainnya_Rutin, rowRef._04Honorarium, rowRef._04_Honorarium_Rutin, _
        '        rowRef._05PremiAsuransi, rowRef._05_Premi_Asurasi_Rutin, rowRef._06Natura, rowRef._06_Natura_Rutin, _
        '        rowRef._07Tantiem, rowRef._07_Tantiem_Rutin, rowRef._08PenghasilanBruto_new, rowRef._08_Penghasilan_Bruto_Rutin_new, _
        '        rowRef._09BiayaJabatan_new, rowRef._09_Biaya_Jabatan_Rutin_new, rowRef._10IuranPensiun, rowRef._10_Iuran_Pensiun_Rutin, _
        '        rowRef._11JumlahPengurangan_new, rowRef._11_Jumlah_Pengurang_Rutin_new, Txt12PjkJmlPenghasilanNetto, Txt13PjkJmlHasilMsPjkSblm, _
        '        rowRef._14JumlahPenghasilanNettoUtkPPh_new, rowRef._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin_new, rowRef._15PTKP_new, rowRef._16PKPSetahun_new, _
        '        rowRef._16_PKP_Setahun_Rutin_new, rowRef._17PPh21_new, rowRef._17_PPH_21_Rutin_new, rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new, rowRef._19_PPH_Terhutang_Setahun_new, _
        '        rowRef._19_PPH_Terhutang_Setahun_Rutin_new, rowRef._19_PPH_Terhutang_Sebulan_Non_Rutin_new, rowRef._19_PPH_Terutang_Sebulan_Rutin_new, _
        '        rowRef._19PPhTerhutang_new, rowRef.Income_Tax_New, rowRef.Kurang_Bayar_New, Txt20)

        '        rowItemIterasi = Me.TBLItemIterasi.Newtbltrans_salary_item_iterationRow
        '        With rowItemIterasi
        '            .PPH_Terutang_Non_Rutin = CDec(rowRef._19_PPH_Terhutang_Setahun_new) - CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new)
        '            .PPH_Terutang_Rutin = CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new)
        '            .primary = Me.GetPrimaryMax(Me.TBLItemIterasi)
        '            .primary_transaksi = rowRef.primary_salary
        '            .Tunjangan_PPH_Rutin = CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new)
        '            ._Tunjangan_PPH_Non_Rutin = CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_new) - CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new)
        '        End With
        '        Me.TBLItemIterasi.Rows.Add(rowItemIterasi)
        '    End If

        '    'jika gross up, maka hitung pajak untuk iterasi selanjutnya sampai dengan tunjangan PPH = PPH terhutang
        '    If rowRef.Tipe_Tax = CInt(ClsTransaksiSalary.EnumTax.Gross_Up) Then
        '        'While CDec(rowref._02TunjanganPPh_new ) <> CDec(rowref._19PPhTerhutang_new)
        '        While CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_new) <> CDec(rowRef._19_PPH_Terhutang_Setahun_new) Or _
        '        CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new) <> CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new)
        '            If JmlIterasi > 10000 Then
        '                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
        '                Exit Sub
        '            End If

        '            rowItemIterasiSebelum = Me.TBLItemIterasi.Rows(Me.TBLItemIterasi.Rows.Count - 1)
        '            Selisih1 = rowItemIterasiSebelum.PPH_Terutang_Non_Rutin + rowItemIterasiSebelum.PPH_Terutang_Rutin - rowItemIterasiSebelum._Tunjangan_PPH_Non_Rutin - rowItemIterasiSebelum.Tunjangan_PPH_Rutin
        '            Selisih2 = rowItemIterasiSebelum.PPH_Terutang_Rutin - rowItemIterasiSebelum.Tunjangan_PPH_Rutin
        '            If Selisih1 = 1 Then
        '                Selisih1 = 0
        '            End If
        '            If Selisih2 = 1 Then
        '                Selisih2 = 0
        '            End If
        '            'End If
        '            If rowRef.End_Date.Month <> 12 And vResign = False Then
        '                Me.ObjSalary.HitungPajak1_Aktual(UseNPWP, Me.vTBLTarif, FaktorKaliDisetahunkan, _
        '                TBLItemTaxDummy, Me.Perct_BiayaJabatan, Me.BiayaJabatanMaxValue, Me.ObjCD, NPWPCategory, TaxCategoryDesc, _
        '               rowRef.primary_salary, rowRef.primary_employee, rowRef.Tanggal_Transaksi, rowRef.No_Transaksi, CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new) + Me.ObjCD.RoundDown1(Selisih2 / 2, 0), _
        '                CDec(rowRef._19_PPH_Terhutang_Setahun_new) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new, rowRef._02_Tunjangan_PPH_Disetahunkan_new, rowRef._02_Tunjangan_PPH_Rutin_new, _
        '                rowRef._02_Tunjangan_PPH_Non_Rutin_new, rowRef._02TunjanganPPh_new, rowRef._01_Gaji_Rutin, rowRef._01Gaji, _
        '                rowRef._03TunjanganLainnya, rowRef._03_Tunjangan_Lainnya_Rutin, rowRef._04Honorarium, rowRef._04_Honorarium_Rutin, _
        '                rowRef._05PremiAsuransi, rowRef._05_Premi_Asurasi_Rutin, rowRef._06Natura, rowRef._06_Natura_Rutin, _
        '                rowRef._07Tantiem, rowRef._07_Tantiem_Rutin, rowRef._08PenghasilanBruto_new, rowRef._08_Penghasilan_Bruto_Rutin_new, _
        '                rowRef._09BiayaJabatan_new, rowRef._09_Biaya_Jabatan_Rutin_new, rowRef._10IuranPensiun, rowRef._10_Iuran_Pensiun_Rutin, _
        '                rowRef._11JumlahPengurangan_new, rowRef._11_Jumlah_Pengurang_Rutin_new, Txt12PjkJmlPenghasilanNetto, Txt13PjkJmlHasilMsPjkSblm, _
        '                rowRef._14JumlahPenghasilanNettoUtkPPh_new, rowRef._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin_new, rowRef._15PTKP_new, rowRef._16PKPSetahun_new, _
        '                rowRef._16_PKP_Setahun_Rutin_new, rowRef._17PPh21_new, rowRef._17_PPH_21_Rutin_new, rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new, rowRef._19_PPH_Terhutang_Setahun_new, _
        '                rowRef._19_PPH_Terhutang_Setahun_Rutin_new, rowRef._19_PPH_Terhutang_Sebulan_Non_Rutin_new, rowRef._19_PPH_Terutang_Sebulan_Rutin_new, _
        '                rowRef._19PPhTerhutang_new, rowRef.Income_Tax_New, rowRef.Kurang_Bayar_New, Txt20)

        '                rowItemIterasi = Me.TBLItemIterasi.Newtbltrans_salary_item_iterationRow

        '                With rowItemIterasi
        '                    .PPH_Terutang_Non_Rutin = CDec(rowRef._19_PPH_Terhutang_Setahun_new) - CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new)
        '                    .PPH_Terutang_Rutin = CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new)
        '                    .primary = Me.GetPrimaryMax(Me.TBLItemIterasi)
        '                    .primary_transaksi = rowRef.primary_salary
        '                    .Tunjangan_PPH_Rutin = CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new)
        '                    ._Tunjangan_PPH_Non_Rutin = CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_new) - CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new)
        '                End With
        '                Me.TBLItemIterasi.Rows.Add(rowItemIterasi)
        '            Else
        '                Me.ObjSalary.HitungPajak1_Aktual_EndOfYear(CDec(rowRef._19_PPH_Terutang_Sebulan_Rutin_new) + Me.ObjCD.RoundDown1(Selisih2 / 2, 0), _
        '                CDec(rowRef._19_PPH_Terhutang_Sebulan_Non_Rutin_new) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), _
        '                UseNPWP, Me.vTBLTarif, _
        '                 CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new) + Me.ObjCD.RoundDown1(Selisih2 / 2, 0), _
        '                CDec(rowRef._19_PPH_Terhutang_Setahun_new) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), StartWorkingDate, ResignDate, FaktorKaliDisetahunkan, _
        '                TBLItemTaxDummy, Me.Perct_BiayaJabatan, Me.BiayaJabatanMaxValue, Me.ObjCD, NPWPCategory, TaxCategoryDesc, rowRef.primary_salary, _
        '                rowRef.primary_employee, rowRef.Tanggal_Transaksi, rowRef.No_Transaksi, rowRef.End_Date, rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new, _
        '                rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new, rowRef._02_Tunjangan_PPH_Disetahunkan_new, rowRef._02_Tunjangan_PPH_Rutin_new, _
        '                rowRef._02_Tunjangan_PPH_Non_Rutin_new, rowRef._02TunjanganPPh_new, rowRef._01_Gaji_Rutin, rowRef._01Gaji, _
        '                rowRef._03TunjanganLainnya, rowRef._03_Tunjangan_Lainnya_Rutin, rowRef._04Honorarium, rowRef._04_Honorarium_Rutin, _
        '                rowRef._05PremiAsuransi, rowRef._05_Premi_Asurasi_Rutin, rowRef._06Natura, rowRef._06_Natura_Rutin, _
        '                rowRef._07Tantiem, rowRef._07_Tantiem_Rutin, rowRef._08PenghasilanBruto_new, rowRef._08_Penghasilan_Bruto_Rutin_new, _
        '                rowRef._09BiayaJabatan_new, rowRef._09_Biaya_Jabatan_Rutin_new, rowRef._10IuranPensiun, rowRef._10_Iuran_Pensiun_Rutin, _
        '                rowRef._11JumlahPengurangan_new, rowRef._11_Jumlah_Pengurang_Rutin_new, Txt12PjkJmlPenghasilanNetto, Txt13PjkJmlHasilMsPjkSblm, _
        '                rowRef._14JumlahPenghasilanNettoUtkPPh_new, rowRef._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin_new, rowRef._15PTKP_new, rowRef._16PKPSetahun_new, _
        '                rowRef._16_PKP_Setahun_Rutin_new, rowRef._17PPh21_new, rowRef._17_PPH_21_Rutin_new, rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new, rowRef._19_PPH_Terhutang_Setahun_new, _
        '                rowRef._19_PPH_Terhutang_Setahun_Rutin_new, rowRef._19_PPH_Terhutang_Sebulan_Non_Rutin_new, rowRef._19_PPH_Terutang_Sebulan_Rutin_new, _
        '                rowRef._19PPhTerhutang_new, rowRef.Income_Tax_New, rowRef.Kurang_Bayar_New, Txt20)

        '                rowItemIterasi = Me.TBLItemIterasi.Newtbltrans_salary_item_iterationRow
        '                With rowItemIterasi
        '                    .PPH_Terutang_Non_Rutin = CDec(rowRef._19_PPH_Terhutang_Setahun_new) - CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new)
        '                    .PPH_Terutang_Rutin = CDec(rowRef._19_PPH_Terhutang_Setahun_Rutin_new)
        '                    .primary = Me.GetPrimaryMax(Me.TBLItemIterasi)
        '                    .primary_transaksi = rowRef.primary_salary
        '                    .Tunjangan_PPH_Rutin = CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new)
        '                    ._Tunjangan_PPH_Non_Rutin = CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_new) - CDec(rowRef._02_Tunjangan_PPH_Disetahunkan_Rutin_new)
        '                End With
        '                Me.TBLItemIterasi.Rows.Add(rowItemIterasi)
        '            End If
        '        End While
        '    End If

        '    Me.HitungPPhFinalPesangon(vIndex, UseNPWP, TBLItemTaxDummy)
        '    rowRef.Penghasilan_Bruto_new += CDec(rowRef._02_Tunjangan_PPH_Non_Rutin_new) + CDec(rowRef._02_Tunjangan_PPH_Rutin_new) - CDec(rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new)
        '    For Each rowItemTaxDummy In TBLItemTaxDummy.Rows
        '        If rowItemTaxDummy.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No8_Jumlah_Penghasilan_Bruto Then
        '            rowItemTaxDummy.amount_month = rowRef.Penghasilan_Bruto_new
        '            Exit For
        '        End If
        '    Next

        '    rowRef.Kurang_Bayar_New = Me.ObjCD.DigitGrouping(rowRef.Kurang_Bayar_New, 0)
        '    rowRef.Income_Tax_New = Me.ObjCD.DigitGrouping(rowRef.Income_Tax_New, 0)
        '    'If rowRef.primary_salary = 38 Then
        '    '    MsgBox(rowRef.Penghasilan_Bruto_new.ToString)
        '    'End If
        '    'MsgBox(rowRef.Kurang_Bayar_New.ToString)
        '    Me.SalinDariTabelItemTaxDummy(rowRef.primary_salary, TBLItemTaxDummy)
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    vTBLOtherPayment1.Dispose()
        '    TBLItemTaxDummy = Nothing
        'End Try
    End Sub

    Private Sub HitungPPhFinalPesangon(ByVal vIndex As Integer, ByVal UseNPWP As Core_Module.ClsMasterKaryawan.EnumIsNPWP, ByVal TBLItemTaxDummy As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable)
        'untuk sementara ditutup dulu, karena update cara perhitungan III dan terkait subjective tax obligation
        '2019-01-05
        MsgBox("Logic error: Incompatible parameter. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")

        'Dim rowRef As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1Row
        ''Dim rowDGV As Windows.Forms.DataGridViewRow
        'Dim rowItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        'Dim rowItemIterasi, rowItemIterasiSebelum As DseTransaksiSalary.tbltrans_salary_item_iteration1Row
        ''Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        'Dim TipeSalaryKomponen As Core_Module.FrmMasterSalaryFormula_Component.EnumType = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income
        ''Dim TipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement
        'Dim JmlIterasi As Double = 0
        'Dim Selisih1 As Decimal
        'Dim JmlPesangon As Decimal = 0
        'Dim Txt20 As Decimal = 0

        'Try
        '    rowRef = vTBLItem.Rows(vIndex)

        '    'kosongkan item tax detail dan item iterasi
        '    Me.TBLItemIterasiPPHFinal.Clear()
        '    For Each rowItemTax In TBLItemTaxDummy.Rows
        '        If rowItemTax.tipe_tax = ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon Then
        '            JmlPesangon += rowItemTax.amount_month
        '        End If
        '    Next

        '    'hitung detail pajak untuk iterasi 0
        '    Me.ObjSalary.HitungPPhFinalPesangon_Aktual(0, UseNPWP, Me.vTBLTarifPPHFinal, JmlPesangon, rowRef.primary_salary, rowRef.primary_employee, _
        '    rowRef.Tanggal_Transaksi, rowRef.No_Transaksi, TBLItemTaxDummy, Me.ObjCD, rowRef.End_Date, rowRef._21_Pesangon, _
        '    rowRef._21_Tunj_PPH_FInal_Pesangon_new, rowRef._21_PPH_Final_Pesangon_new, rowRef.Tarif_Pesangon_new, rowRef._19PPhTerhutang_new, _
        '    rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new, rowRef.Income_Tax_New, rowRef.Kurang_Bayar_New, Txt20)
        '    rowItemIterasi = Me.TBLItemIterasiPPHFinal.Newtbltrans_salary_item_iteration1Row
        '    With rowItemIterasi
        '        .PPH_Terutang = CDec(rowRef._21_PPH_Final_Pesangon_new)
        '        .Tunjangan_PPH = CDec(rowRef._21_Tunj_PPH_FInal_Pesangon_new)
        '        .primary = Me.GetPrimaryMax(Me.TBLItemIterasiPPHFinal)
        '        .primary_transaksi = -1
        '    End With
        '    Me.TBLItemIterasiPPHFinal.Rows.Add(rowItemIterasi)

        '    'jika gross up, maka hitung pajak untuk iterasi selanjutnya sampai dengan tunjangan PPH = PPH terhutang
        '    If rowRef.Tipe_Tax = CInt(ClsTransaksiSalary.EnumTax.Gross_Up) Then
        '        While CDec(rowRef._21_PPH_Final_Pesangon_new) <> CDec(rowRef._21_Tunj_PPH_FInal_Pesangon_new)
        '            If JmlIterasi > 10000 Then
        '                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
        '                Exit Sub
        '            End If
        '            'If JmlIterasi >= 1 Then
        '            '    Exit Sub
        '            'End If
        '            JmlIterasi += 1
        '            rowItemIterasiSebelum = Me.TBLItemIterasiPPHFinal.Rows(Me.TBLItemIterasiPPHFinal.Rows.Count - 1)
        '            Selisih1 = rowItemIterasiSebelum.PPH_Terutang - rowItemIterasiSebelum.Tunjangan_PPH
        '            If Selisih1 = 1 Then
        '                Selisih1 = 0
        '            End If

        '            Me.ObjSalary.HitungPPhFinalPesangon_Aktual(CDec(rowRef._21_PPH_Final_Pesangon_new) + Me.ObjCD.RoundDown1(Selisih1 / 2, 0), _
        '            UseNPWP, Me.vTBLTarifPPHFinal, JmlPesangon, rowRef.primary_salary, rowRef.primary_employee, _
        '            rowRef.Tanggal_Transaksi, rowRef.No_Transaksi, TBLItemTaxDummy, Me.ObjCD, rowRef.End_Date, rowRef._21_Pesangon, _
        '            rowRef._21_Tunj_PPH_FInal_Pesangon_new, rowRef._21_PPH_Final_Pesangon_new, rowRef.Tarif_Pesangon_new, rowRef._19PPhTerhutang_new, _
        '            rowRef._18PPh21YgTelahDipotongMasaPjkSblmnya_new, rowRef.Income_Tax_New, rowRef.Kurang_Bayar_New, Txt20)

        '            rowItemIterasi = Me.TBLItemIterasiPPHFinal.Newtbltrans_salary_item_iteration1Row
        '            With rowItemIterasi
        '                .PPH_Terutang = CDec(rowRef._21_PPH_Final_Pesangon_new)
        '                .Tunjangan_PPH = CDec(rowRef._21_Tunj_PPH_FInal_Pesangon_new)
        '                .primary = Me.GetPrimaryMax(Me.TBLItemIterasiPPHFinal)
        '                .primary_transaksi = -1
        '            End With
        '            Me.TBLItemIterasiPPHFinal.Rows.Add(rowItemIterasi)
        '        End While
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Function GetPrimaryMax(ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iteration1DataTable) As Double
        Dim rowItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iteration1Row
        Dim RetVal As Double = Double.MinValue

        Try
            For Each rowItemIterasi In vTBLItemIterasi.Rows
                If RetVal < rowItemIterasi.primary Then
                    RetVal = rowItemIterasi.primary
                End If
            Next

            If RetVal = Double.MinValue Then
                Return 0
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeTabelItemTaxDummy(ByVal vPrimarySalary As Double, ByVal vTBLItemTaxDummy As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, _
    ByVal NoTrans As String, ByVal vPrimaryEmployee As Double, ByVal vTglTransaksi As Date)
        Dim rowItemTaxDummy As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim rowItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow
        Dim DV As DataRowView

        Try
            'kosongkan data tabel
            vTBLItemTaxDummy.Rows.Clear()

            'salin dari item b4 recalculate
            Me.vTBLItemTax.DefaultView.RowFilter = "[primary salary] = " + vPrimarySalary.ToString + _
            " AND [tipe item tax] = " + CInt(ClsTrans_TaxRecalculateItem.EnumTipeItemTax.B4_Recalculate).ToString
            For Each DV In Me.vTBLItemTax.DefaultView
                rowItemTax = DV.Row
                rowItemTaxDummy = vTBLItemTaxDummy.Newtbltrans_salary_itemtaxRow
                With rowItemTaxDummy
                    .amount_gabungan_disetahunkan = rowItemTax.amount_gabungan_disetahunkan
                    .amount_month = rowItemTax.amount_month
                    .amount_rutin_disetahunkan = rowItemTax.amount_rutin_disetahunkan
                    .identifier_salary_component = rowItemTax.identifier_salary_component
                    .is_rutin = rowItemTax.is_rutin
                    .No_Transaksi = NoTrans
                    .primary = vTBLItemTaxDummy.Rows.Count + 1
                    .Primary_Employee = vPrimaryEmployee
                    .primary_transaksi = rowItemTax.primary_salary
                    .remarks_gabungan = rowItemTax.remarks_gabungan
                    .remarks_rutin = rowItemTax.remarks_rutin
                    .Tanggal_Transaksi = vTglTransaksi
                    .tipe_tax = rowItemTax.tipe_tax
                End With
                vTBLItemTaxDummy.Rows.Add(rowItemTaxDummy)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinDariTabelItemTaxDummy(ByVal vPrimarySalary As Double, ByVal vTBLItemTaxDummy As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable)
        Dim rowItemTaxDummy As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim rowItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow
        Dim DV As DataRowView

        Try
            'hapuskan item aft recalculte
            Me.vTBLItemTax.DefaultView.RowFilter = "[primary salary] = " + vPrimarySalary.ToString + _
            " AND [tipe item tax] = " + CInt(ClsTrans_TaxRecalculateItem.EnumTipeItemTax.Aft_Recalculate).ToString
            While Me.vTBLItemTax.DefaultView.Count > 0
                Me.vTBLItemTax.Rows.Remove(Me.vTBLItemTax.DefaultView(0).Row)
            End While

            'salin data item tax dan tandain bahwa aft recalculate
            For Each rowItemTaxDummy In vTBLItemTaxDummy.Rows
                rowItemTax = Me.vTBLItemTax.Newtbltrans_taxrecalculate_itemtaxRow
                With rowItemTax
                    .amount_gabungan_disetahunkan = rowItemTaxDummy.amount_gabungan_disetahunkan
                    .amount_month = rowItemTaxDummy.amount_month
                    .amount_rutin_disetahunkan = rowItemTaxDummy.amount_rutin_disetahunkan
                    .identifier_salary_component = rowItemTaxDummy.identifier_salary_component
                    .is_rutin = rowItemTaxDummy.is_rutin
                    .primary = Me.vTBLItemTax.Rows.Count + 1
                    .primary_salary = rowItemTaxDummy.primary_transaksi
                    .primary_transaksi = -1
                    .remarks_gabungan = rowItemTaxDummy.remarks_gabungan
                    .remarks_rutin = rowItemTaxDummy.remarks_rutin
                    .tipe_item_tax = CInt(ClsTrans_TaxRecalculateItem.EnumTipeItemTax.Aft_Recalculate)
                    .tipe_tax = rowItemTaxDummy.tipe_tax
                End With
                Me.vTBLItemTax.Rows.Add(rowItemTax)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KumpulkanDataItemTax(ByVal DTASalaryItemTax As DseTrans_Tax_RecalculateTableAdapters.tbltrans_salary_itemtaxTableAdapter, _
    ByVal TBLSalaryItemTax As DseTrans_Tax_Recalculate.tbltrans_salary_itemtaxDataTable, _
    ByVal DTAItemTaxLama As DseTrans_Tax_RecalculateTableAdapters.tbltrans_taxrecalculate_itemtaxTableAdapter, _
    ByVal TBLItemTaxLama As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxDataTable, _
    ByVal vPrimaryTaxRecalculate As Double, ByVal TBLSalary As DseTrans_Tax_Recalculate.DataTable1DataTable)
        Dim rowItem As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1Row
        Dim rowSalaryItemTax As DseTrans_Tax_Recalculate.tbltrans_salary_itemtaxRow
        Dim rowItemTax As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow
        Dim rowItemTaxLama As DseTrans_Tax_Recalculate.tbltrans_taxrecalculate_itemtaxRow
        Dim DV As DataRowView

        Try
            Me.Step1 += 1
            Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + " : Collecting tax item data"
            DTAItemTaxLama.FillBy_PrimaryTrans_Tipe(TBLItemTaxLama, vPrimaryTaxRecalculate, ClsTrans_TaxRecalculateItem.EnumTipeItemTax.B4_Recalculate)
            DTASalaryItemTax.Fill(TBLSalaryItemTax, Me.ObjCD.GetLasDateOfTheMonth(Me.vPeriod), Me.ObjCD.GetFirstDateofTheMonth(Me.vPeriod))
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = vTBLItem.Rows.Count
            For Each rowItem In vTBLItem.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 500 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 25) = 0 Then
                        Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                        " : Collecting tax item data (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                    End If
                Else
                    Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                    " : Collecting tax item data (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                End If

                TBLItemTaxLama.DefaultView.RowFilter = "[primary salary] = " + rowItem.primary_salary.ToString
                If TBLItemTaxLama.DefaultView.Count = 0 Then
                    'ambil lansung dari item salary
                    TBLSalaryItemTax.DefaultView.RowFilter = "[primary transaksi] = " + rowItem.primary_salary.ToString
                    For Each DV In TBLSalaryItemTax.DefaultView
                        rowSalaryItemTax = DV.Row
                        rowItemTax = Me.vTBLItemTax.Newtbltrans_taxrecalculate_itemtaxRow
                        With rowItemTax
                            .amount_gabungan_disetahunkan = rowSalaryItemTax.amount_gabungan_disetahunkan
                            .amount_month = rowSalaryItemTax.amount_month
                            .amount_rutin_disetahunkan = rowSalaryItemTax.amount_rutin_disetahunkan
                            .identifier_salary_component = rowSalaryItemTax.identifier_salary_component
                            .is_rutin = rowSalaryItemTax.is_rutin
                            .primary = Me.vTBLItemTax.Rows.Count + 1
                            .primary_salary = rowSalaryItemTax.primary_transaksi
                            .primary_transaksi = -1
                            .remarks_gabungan = rowSalaryItemTax.remarks_gabungan
                            .remarks_rutin = rowSalaryItemTax.remarks_rutin
                            .tipe_item_tax = ClsTrans_TaxRecalculateItem.EnumTipeItemTax.B4_Recalculate
                            .tipe_tax = rowSalaryItemTax.tipe_tax
                        End With
                        Me.vTBLItemTax.Rows.Add(rowItemTax)
                    Next
                Else
                    'ambil dr item tax calculate lama karena yg di tbl trans salary sudah di update ke data baru
                    'tidak bisa digunakan sebagai dasar perhitungan
                    For Each DV In TBLItemTaxLama.DefaultView
                        rowItemTaxLama = DV.Row
                        rowItemTax = Me.vTBLItemTax.Newtbltrans_taxrecalculate_itemtaxRow
                        With rowItemTax
                            .amount_gabungan_disetahunkan = rowItemTaxLama.amount_gabungan_disetahunkan
                            .amount_month = rowItemTaxLama.amount_month
                            .amount_rutin_disetahunkan = rowItemTaxLama.amount_rutin_disetahunkan
                            .identifier_salary_component = rowItemTaxLama.identifier_salary_component
                            .is_rutin = rowItemTaxLama.is_rutin
                            .primary = Me.vTBLItemTax.Rows.Count + 1
                            .primary_salary = rowItemTaxLama.primary_salary
                            .primary_transaksi = -1
                            .remarks_gabungan = rowItemTaxLama.remarks_gabungan
                            .remarks_rutin = rowItemTaxLama.remarks_rutin
                            .tipe_item_tax = ClsTrans_TaxRecalculateItem.EnumTipeItemTax.B4_Recalculate
                            .tipe_tax = rowItemTaxLama.tipe_tax
                        End With
                        Me.vTBLItemTax.Rows.Add(rowItemTax)
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KumpulkanDataSalary(ByVal DTASalary As DseTrans_Tax_RecalculateTableAdapters.DataTable1TableAdapter, _
    ByVal TBLSalary As DseTrans_Tax_Recalculate.DataTable1DataTable, ByVal TBLKary As DseTrans_Tax_Recalculate.tblmaster_karyawanDataTable)
        Dim rowItem As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1Row
        Dim rowKary As DseTrans_Tax_Recalculate.tblmaster_karyawanRow
        Dim rowSalary As DseTrans_Tax_Recalculate.DataTable1Row

        Try
            Me.Step1 += 1
            Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + " : Loading salary data"
            DTASalary.Fill(TBLSalary, Me.ObjCD.GetLasDateOfTheMonth(Me.vPeriod), _
            Me.ObjCD.GetFirstDateofTheMonth(Me.vPeriod), CInt(ClsTransaksiSalary.EnumTax.Manual_Tax_Calculation))
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBLSalary.Rows.Count
            For Each rowSalary In TBLSalary.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 500 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 25) = 0 Then
                        Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                        " : Loading salary data (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                    End If
                Else
                    Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                    " : Loading salary data (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                End If

                If rowSalary.Tax <> ClsTransaksiSalary.EnumTax.Manual_Tax_Calculation Then
                    rowItem = vTBLItem.Newtbltrans_tax_recalculate_item1Row
                    With rowItem
                        .Primary = vTBLItem.Rows.Count + 1

                        'isikan data lama
                        .A_Alamat1 = rowSalary.A_Alamat1
                        .A_Alamat2 = rowSalary.A_Alamat2
                        .A_NPWP = rowSalary.A_NPWP
                        .A_Status_NPWP = rowSalary.A_Status_NPWP
                        .No_Transaksi = rowSalary.No_Transaksi
                        .primary_employee = rowSalary.Primary_Employee
                        .primary_salary = rowSalary.Primary
                        .primary_transaksi = -1
                        .Tanggal_Transaksi = rowSalary.Tanggal_Transaksi
                        ._01_Gaji_Rutin = rowSalary._01_Gaji_Rutin
                        ._01Gaji = rowSalary._01Gaji
                        ._02_Tunjangan_PPH_Disetahunkan = rowSalary._02_Tunjangan_PPH_Disetahunkan
                        ._02_Tunjangan_PPH_Disetahunkan_Rutin = rowSalary._02_Tunjangan_PPH_Disetahunkan_Rutin
                        ._02_Tunjangan_PPH_Non_Rutin = rowSalary._02_Tunjangan_PPH_Non_Rutin
                        ._02_Tunjangan_PPH_Rutin = rowSalary._02_Tunjangan_PPH_Rutin
                        ._02TunjanganPPh = rowSalary._02TunjanganPPh
                        ._03_Tunjangan_Lainnya_Rutin = rowSalary._03_Tunjangan_Lainnya_Rutin
                        ._03TunjanganLainnya = rowSalary._03TunjanganLainnya
                        ._04_Honorarium_Rutin = rowSalary._04_Honorarium_Rutin
                        ._04Honorarium = rowSalary._04Honorarium
                        ._05_Premi_Asurasi_Rutin = rowSalary._05_Premi_Asurasi_Rutin
                        ._05PremiAsuransi = rowSalary._05PremiAsuransi
                        ._06_Natura_Rutin = rowSalary._06_Natura_Rutin
                        ._06Natura = rowSalary._06Natura
                        ._07_Tantiem_Rutin = rowSalary._07_Tantiem_Rutin
                        ._07Tantiem = rowSalary._07Tantiem
                        ._08_Penghasilan_Bruto_Rutin = rowSalary._08_Penghasilan_Bruto_Rutin
                        ._08PenghasilanBruto = rowSalary._08PenghasilanBruto
                        ._09_Biaya_Jabatan_Rutin = rowSalary._09_Biaya_Jabatan_Rutin
                        ._09BiayaJabatan = rowSalary._09BiayaJabatan
                        ._10_Iuran_Pensiun_Rutin = rowSalary._10_Iuran_Pensiun_Rutin
                        ._10IuranPensiun = rowSalary._10IuranPensiun
                        ._11_Jumlah_Pengurang_Rutin = rowSalary._11_Jumlah_Pengurang_Rutin
                        ._11JumlahPengurangan = rowSalary._11JumlahPengurangan
                        ._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin = rowSalary._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin
                        ._14JumlahPenghasilanNettoUtkPPh = rowSalary._14JumlahPenghasilanNettoUtkPPh
                        ._15PTKP = rowSalary._15PTKP
                        ._16_PKP_Setahun_Rutin = rowSalary._16_PKP_Setahun_Rutin
                        ._16PKPSetahun = rowSalary._16PKPSetahun
                        ._17_PPH_21_Rutin = rowSalary._17_PPH_21_Rutin
                        ._17PPh21 = rowSalary._17PPh21
                        ._18PPh21YgTelahDipotongMasaPjkSblmnya = rowSalary._18PPh21YgTelahDipotongMasaPjkSblmnya
                        ._19_PPH_Terhutang_Sebulan_Non_Rutin = rowSalary._19_PPH_Terhutang_Sebulan_Non_Rutin
                        ._19_PPH_Terhutang_Setahun = rowSalary._19_PPH_Terhutang_Setahun
                        ._19_PPH_Terhutang_Setahun_Rutin = rowSalary._19_PPH_Terhutang_Setahun_Rutin
                        ._19_PPH_Terutang_Sebulan_Rutin = rowSalary._19_PPH_Terutang_Sebulan_Rutin
                        ._19PPhTerhutang = rowSalary._19PPhTerhutang
                        ._21_Pesangon = rowSalary._21_Pesangon
                        ._21_PPH_Final_Pesangon = rowSalary._21_PPH_Final_Pesangon
                        ._21_Tunj_PPH_FInal_Pesangon = rowSalary._21_Tunj_PPH_FInal_Pesangon
                        .Tarif_Pesangon = rowSalary.Tarif_Pesangon
                        .Penghasilan_Bruto = rowSalary.Penghasilan_Bruto
                        .Tipe_Tax = rowSalary.Tax
                        .Start_Date = rowSalary.Start_Date
                        .End_Date = rowSalary.End_Date
                        .Income_Tax = rowSalary.Income_Tax
                        .Kurang_Bayar = rowSalary.Kurang_Bayar_Tax

                        'isikan data baru
                        TBLKary.DefaultView.RowFilter = "primary = " + .primary_employee.ToString
                        rowKary = TBLKary.DefaultView(0).Row
                        .A_Alamat1_new = rowKary.npwp_address
                        .A_Alamat2_new = ""
                        .A_NPWP_new = rowKary.npwp_no
                        .A_Status_NPWP_new = rowKary.npwp_category
                        .Full_Name = rowKary.Full_Name

                        ._02_Tunjangan_PPH_Disetahunkan_new = 0
                        ._02_Tunjangan_PPH_Disetahunkan_Rutin_new = 0
                        ._02_Tunjangan_PPH_Rutin_new = 0
                        ._02_Tunjangan_PPH_Non_Rutin_new = 0
                        ._02TunjanganPPh_new = 0
                        ._08_Penghasilan_Bruto_Rutin_new = 0
                        ._08PenghasilanBruto_new = 0
                        ._09_Biaya_Jabatan_Rutin_new = 0
                        ._09BiayaJabatan_new = 0
                        ._11_Jumlah_Pengurang_Rutin_new = 0
                        ._11JumlahPengurangan_new = 0
                        ._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin_new = 0
                        ._14JumlahPenghasilanNettoUtkPPh_new = 0
                        ._15PTKP_new = 0
                        ._16_PKP_Setahun_Rutin_new = 0
                        ._16PKPSetahun_new = 0
                        ._17_PPH_21_Rutin_new = 0
                        ._17PPh21_new = 0
                        ._18PPh21YgTelahDipotongMasaPjkSblmnya_new = 0
                        ._19_PPH_Terhutang_Sebulan_Non_Rutin_new = 0
                        ._19_PPH_Terhutang_Setahun_new = 0
                        ._19_PPH_Terhutang_Setahun_Rutin_new = 0
                        ._19_PPH_Terutang_Sebulan_Rutin_new = 0
                        ._19PPhTerhutang_new = 0
                        ._21_PPH_Final_Pesangon_new = 0
                        ._21_Tunj_PPH_FInal_Pesangon_new = 0
                        .Tarif_Pesangon_new = 0
                        .Penghasilan_Bruto_new = 0
                        .Income_Tax_New = 0
                        .Kurang_Bayar_New = 0
                        .Difference = 0
                    End With
                    vTBLItem.Rows.Add(rowItem)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateDataSalary(ByVal vPrimaryTaxRecalculate As Double, ByVal DTALama As DseTrans_Tax_RecalculateTableAdapters.tbltrans_tax_recalculate_item1TableAdapter, _
    ByVal TBLLama As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1DataTable, _
    ByVal TBLSalary As DseTrans_Tax_Recalculate.DataTable1DataTable)
        Dim rowItem As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1Row
        Dim rowLama As DseTrans_Tax_Recalculate.tbltrans_tax_recalculate_item1Row

        Try
            Me.Step1 += 1
            Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + " : Updating salary data"
            DTALama.Fill(TBLLama, vPrimaryTaxRecalculate)
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBLSalary.Rows.Count
            For Each rowLama In TBLLama.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 500 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 25) = 0 Then
                        Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                        " : Updating salary data (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                    End If
                Else
                    Me.LblProgress.Text = "Step " + Me.Step1.ToString + " of " + Me.Step2.ToString + _
                    " : Updating salary data (Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + ")"
                End If

                Me.vTBLItem.DefaultView.RowFilter = "[primary salary] = " + rowLama.primary_salary.ToString
                If Me.vTBLItem.DefaultView.Count > 0 Then
                    rowItem = vTBLItem.DefaultView(0).Row
                    With rowItem
                        .A_Alamat1 = rowLama.A_Alamat1
                        .A_Alamat2 = rowLama.A_Alamat2
                        .A_NPWP = rowLama.A_NPWP
                        .A_Status_NPWP = rowLama.A_Status_NPWP
                        ._01_Gaji_Rutin = rowLama._01_Gaji_Rutin
                        ._01Gaji = rowLama._01Gaji
                        ._02_Tunjangan_PPH_Disetahunkan = rowLama._02_Tunjangan_PPH_Disetahunkan
                        ._02_Tunjangan_PPH_Disetahunkan_Rutin = rowLama._02_Tunjangan_PPH_Disetahunkan_Rutin
                        ._02_Tunjangan_PPH_Non_Rutin = rowLama._02_Tunjangan_PPH_Non_Rutin
                        ._02_Tunjangan_PPH_Rutin = rowLama._02_Tunjangan_PPH_Rutin
                        ._02TunjanganPPh = rowLama._02TunjanganPPh
                        ._03_Tunjangan_Lainnya_Rutin = rowLama._03_Tunjangan_Lainnya_Rutin
                        ._03TunjanganLainnya = rowLama._03TunjanganLainnya
                        ._04_Honorarium_Rutin = rowLama._04_Honorarium_Rutin
                        ._04Honorarium = rowLama._04Honorarium
                        ._05_Premi_Asurasi_Rutin = rowLama._05_Premi_Asurasi_Rutin
                        ._05PremiAsuransi = rowLama._05PremiAsuransi
                        ._06_Natura_Rutin = rowLama._06_Natura_Rutin
                        ._06Natura = rowLama._06Natura
                        ._07_Tantiem_Rutin = rowLama._07_Tantiem_Rutin
                        ._07Tantiem = rowLama._07Tantiem
                        ._08_Penghasilan_Bruto_Rutin = rowLama._08_Penghasilan_Bruto_Rutin
                        ._08PenghasilanBruto = rowLama._08PenghasilanBruto
                        ._09_Biaya_Jabatan_Rutin = rowLama._09_Biaya_Jabatan_Rutin
                        ._09BiayaJabatan = rowLama._09BiayaJabatan
                        ._10_Iuran_Pensiun_Rutin = rowLama._10_Iuran_Pensiun_Rutin
                        ._10IuranPensiun = rowLama._10IuranPensiun
                        ._11_Jumlah_Pengurang_Rutin = rowLama._11_Jumlah_Pengurang_Rutin
                        ._11JumlahPengurangan = rowLama._11JumlahPengurangan
                        ._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin = rowLama._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin
                        ._14JumlahPenghasilanNettoUtkPPh = rowLama._14JumlahPenghasilanNettoUtkPPh
                        ._15PTKP = rowLama._15PTKP
                        ._16_PKP_Setahun_Rutin = rowLama._16_PKP_Setahun_Rutin
                        ._16PKPSetahun = rowLama._16PKPSetahun
                        ._17_PPH_21_Rutin = rowLama._17_PPH_21_Rutin
                        ._17PPh21 = rowLama._17PPh21
                        ._18PPh21YgTelahDipotongMasaPjkSblmnya = rowLama._18PPh21YgTelahDipotongMasaPjkSblmnya
                        ._19_PPH_Terhutang_Sebulan_Non_Rutin = rowLama._19_PPH_Terhutang_Sebulan_Non_Rutin
                        ._19_PPH_Terhutang_Setahun = rowLama._19_PPH_Terhutang_Setahun
                        ._19_PPH_Terhutang_Setahun_Rutin = rowLama._19_PPH_Terhutang_Setahun_Rutin
                        ._19_PPH_Terutang_Sebulan_Rutin = rowLama._19_PPH_Terutang_Sebulan_Rutin
                        ._19PPhTerhutang = rowLama._19PPhTerhutang
                        ._21_Pesangon = rowLama._21_Pesangon
                        ._21_PPH_Final_Pesangon = rowLama._21_PPH_Final_Pesangon
                        ._21_Tunj_PPH_FInal_Pesangon = rowLama._21_Tunj_PPH_FInal_Pesangon
                        .Tarif_Pesangon = rowLama.Tarif_Pesangon
                        .Penghasilan_Bruto = rowLama.Penghasilan_Bruto
                        .Tipe_Tax = rowLama.Tipe_Tax
                        .Income_Tax = rowLama.Income_Tax
                        .Kurang_Bayar = rowLama.Kurang_Bayar
                    End With
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.Selesai = True Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
            Else
                MsgBox("Process interupted", MsgBoxStyle.Information, "Information")
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
            End If
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax(ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationDataTable) As Double
        Dim rowItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationRow
        Dim RetVal As Double = Double.MinValue

        Try
            For Each rowItemIterasi In vTBLItemIterasi.Rows
                If RetVal < rowItemIterasi.primary Then
                    RetVal = rowItemIterasi.primary
                End If
            Next

            If RetVal = Double.MinValue Then
                Return 0
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class