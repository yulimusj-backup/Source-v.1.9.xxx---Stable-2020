Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmLaporanIncomeTaxStaff_Recap
    Public vTblSalary As DseMonthlyTax.DataTable1DataTable
    Public RefPeriod As Date

    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjCD As New Core_Module.ClsCekData
    Dim BackColorAll As System.Drawing.Color = Core_Module.GetSMESSBackColorAll
    Dim DefPrec1 As Integer = 0

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub SortTBLItem(ByVal TBLItem As DseMonthlyTax.tbltrans_salary_itemDataTable)
        Dim TBLItemD As New DseMonthlyTax.tbltrans_salary_itemDataTable
        Dim rowItemD, rowItem As DseMonthlyTax.tbltrans_salary_itemRow
        Dim DV As DataRowView

        Try
            TBLItem.DefaultView.Sort = "[Primary Transaksi], [Tax Component]"
            For Each DV In TBLItem.DefaultView
                rowItem = DV.Row

                rowItemD = TBLItemD.Newtbltrans_salary_itemRow

                For ij As Integer = 0 To TBLItem.Columns.Count - 1
                    rowItemD.Item(ij) = rowItem.Item(ij)
                Next
                TBLItemD.Rows.Add(rowItemD)
            Next

            TBLItem.Rows.Clear()
            While TBLItemD.Rows.Count > 0
                rowItem = TBLItem.Newtbltrans_salary_itemRow
                For ij As Integer = 0 To TBLItem.Rows.Count - 1
                    rowItem.Item(ij) = rowItemD.Item(ij)
                Next
                TBLItem.Rows.Add(rowItem)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLItemD.Rows.Clear()
            TBLItemD.Dispose()
        End Try
    End Sub

    Private Sub LoadData()
        Dim rowSalary As DseMonthlyTax.DataTable1Row
        Dim rowDGV As Windows.Forms.DataGridViewRow
        Dim DefPrec As Integer 
        Dim TaxType As ClsTransaksiSalary.EnumTax = ClsTransaksiSalary.EnumTax.Gross
        Dim vStatus_NPWP As Core_Module.ClsMasterKaryawan.GolonganPajak = Core_Module.ClsMasterKaryawan.GolonganPajak.TK0
        Dim JK As ClsTransaksiSalary.EnumJenisKelamin = ClsTransaksiSalary.EnumJenisKelamin.Pria
        Dim DTAItem As New DseMonthlyTaxTableAdapters.tbltrans_salary_itemTableAdapter
        Dim TBLItem As New DseMonthlyTax.tbltrans_salary_itemDataTable
        Dim rowItem As DseMonthlyTax.tbltrans_salary_itemRow
        Dim DTAItemOther As New DseMonthlyTaxTableAdapters.tbltrans_salary_item6TableAdapter
        Dim TBLItemOther As New DseMonthlyTax.tbltrans_salary_item6DataTable
        Dim rowItemOther As DseMonthlyTax.tbltrans_salary_item6Row
        Dim Ketemu As Boolean = False
        Dim PerluSort As Boolean = False
        Dim TaxComponent_Pesangon As Core_Module.ClsMasterSalaryFormula_item.EnumTax2 = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.pesangon

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemOther.Connection.ConnectionString = Core_Module.GetConstring

            'ini khusus untuk yang akhir tahun atau yg resign, karena logika hitungannya beda dengan yang bulanan biasa
            If Me.RefPeriod.Month = 12 Then
                DTAItem.Fill(TBLItem, Me.ObjCD.GetLasDateOfTheMonth(Me.RefPeriod).Date, Me.ObjCD.GetFirstDateofTheMonth(Me.RefPeriod).Date, TaxComponent_Pesangon)
                DTAItemOther.Fill(TBLItemOther, Me.ObjCD.GetLasDateOfTheMonth(Me.RefPeriod).Date, Me.ObjCD.GetFirstDateofTheMonth(Me.RefPeriod).Date, TaxComponent_Pesangon)
            Else
                DTAItem.FillBy_TglSampaiDari_ResignOnly(TBLItem, Me.ObjCD.GetLasDateOfTheMonth(Me.RefPeriod).Date, Me.ObjCD.GetFirstDateofTheMonth(Me.RefPeriod).Date, 1, TaxComponent_Pesangon)
                DTAItemOther.FillBy_ResignOnly(TBLItemOther, Me.ObjCD.GetLasDateOfTheMonth(Me.RefPeriod).Date, Me.ObjCD.GetFirstDateofTheMonth(Me.RefPeriod).Date, 1, TaxComponent_Pesangon)
            End If

            'tambahkan dengan item other payment
            While TBLItemOther.Rows.Count > 0
                Ketemu = False
                rowItemOther = TBLItemOther.Rows(0)

                For Each rowItem In TBLItem.Rows
                    If rowItem.Primary_Transaksi = rowItemOther.Primary_Transaksi And _
                    rowItem.Tax_Component = rowItemOther.Tax_Component And _
                    rowItem.Type = rowItemOther.Type And rowItem.Is_Rutin = rowItemOther.Is_Rutin Then
                        rowItem.Amount += rowItemOther.Amount
                        Ketemu = True
                        Exit For
                    End If
                Next

                If Ketemu = False Then
                    PerluSort = True
                    'tambahkan baris baru
                    rowItem = TBLItem.Newtbltrans_salary_itemRow

                    rowItem.Primary_Transaksi = rowItemOther.Primary_Transaksi
                    rowItem.Tax_Component = rowItemOther.Tax_Component
                    rowItem.Type = rowItemOther.Type
                    rowItem.Is_Rutin = rowItemOther.Is_Rutin
                    rowItem.Amount = rowItemOther.Amount

                    TBLItem.Rows.Add(rowItem)
                End If

                TBLItemOther.Rows.Remove(rowItemOther)
            End While
            If PerluSort = True Then
                Me.SortTBLItem(TBLItem)
            End If

            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            DefPrec1 = DefPrec
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
            Me.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = Windows.Forms.DataGridViewAutoSizeRowsMode.None
            Me.dgvItem.Rows.Clear()

            For Each rowSalary In Me.vTblSalary.Rows
                With rowSalary 
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    rowDGV.Cells(Me.ColGaji_BiayaJabatan_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._09BiayaJabatan, DefPrec)
                    rowDGV.Cells(Me.ColGaji_BiayaJabatan_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._09_Biaya_Jabatan_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_Code.Index).Value = .Code
                    rowDGV.Cells(Me.ColGaji_DeptCode.Index).Value = .Dept_Code
                    rowDGV.Cells(Me.ColGaji_DeptName.Index).Value = .Dept_Desc
                    rowDGV.Cells(Me.ColGaji_EndPeriod.Index).Value = .Period_End
                    rowDGV.Cells(Me.ColGaji_Gaji_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._01Gaji, DefPrec)
                    rowDGV.Cells(Me.ColGaji_Gaji_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._01_Gaji_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_Honorarium_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._04Honorarium, DefPrec)
                    rowDGV.Cells(Me.ColGaji_Honorarium_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._04_Honorarium_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_IuranPensiun_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._10IuranPensiun, DefPrec)
                    rowDGV.Cells(Me.ColGaji_IuranPensiun_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._10_Iuran_Pensiun_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_JobPositionCode.Index).Value = .Job_Position_Code
                    rowDGV.Cells(Me.ColGaji_JobPositionDesc.Index).Value = .Job_Position_Desc
                    rowDGV.Cells(Me.ColGaji_JumlahPengurang_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._11JumlahPengurangan, DefPrec)
                    rowDGV.Cells(Me.ColGaji_JumlahPengurang_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._11_Jumlah_Pengurang_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_Name.Index).Value = .Full_Name
                    rowDGV.Cells(Me.ColGaji_Nationality.Index).Value = .A_Kode_Negara_Domisili
                    rowDGV.Cells(Me.ColGaji_Natura_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._06Natura, DefPrec)
                    rowDGV.Cells(Me.ColGaji_Natura_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._06_Natura_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_NIK.Index).Value = .A_NIK
                    rowDGV.Cells(Me.ColGaji_No.Index).Value = Me.dgvItem.Rows.Count
                    rowDGV.Cells(Me.ColGaji_NPWP.Index).Value = .A_NPWP
                    rowDGV.Cells(Me.ColGaji_PenghasilBruto_Mix_Value.Index).Value = Me.ObjCD.DigitGrouping(.Penghasilan_Bruto, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PenghasilBruto_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(._08PenghasilanBruto, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PenghasilBruto_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(._08_Penghasilan_Bruto_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PenghNetMsPjkSblm_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(._13JumlahPenghasilanMasaPjkSblmnya, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PenghNetto_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(._14JumlahPenghasilanNettoUtkPPh, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PenghNetto_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PeriodCount.Index).Value = .Period_Count
                    rowDGV.Cells(Me.ColGaji_PKP_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(._16PKPSetahun, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PKP_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(._16_PKP_Setahun_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PPhKurangBayar.Index).Value = Me.ObjCD.DigitGrouping(._20PPhPs21DanPPhPs26YgTelahDipotongNLunas - rowSalary._21_Tunj_PPH_FInal_Pesangon, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PPhSudahBayar.Index).Value = Me.ObjCD.DigitGrouping(._18PPh21YgTelahDipotongMasaPjkSblmnya, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PPhSudahBayarMasaSblm.Index).Value = Me.ObjCD.DigitGrouping(._18PPh21YgTelahDipotongMasaPjkSblmnya1, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PPhTerutang_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._19_PPH_Terhutang_Setahun, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PPhTerutang_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._19_PPH_Terhutang_Setahun_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PPhTerutang_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._19_PPH_Terhutang_Sebulan_Non_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PPhTerutang_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._19_PPH_Terutang_Sebulan_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PPhTerutang_Total.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._19_PPH_Terutang_Sebulan_Rutin + rowSalary._19_PPH_Terhutang_Sebulan_Non_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PremiAsuransi_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._05PremiAsuransi, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PremiAsuransi_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._05_Premi_Asurasi_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PrimaryTransaksi.Index).Value = rowSalary.Primary
                    rowDGV.Cells(Me.ColGaji_PrimaryEmployee.Index).Value = rowSalary.Primary_Employee
                    rowDGV.Cells(Me.ColGaji_PTKP.Index).Value = Me.ObjCD.DigitGrouping(._15PTKP, DefPrec)
                    JK = .A_Jenis_Kelamin
                    If JK = ClsTransaksiSalary.EnumJenisKelamin.Pria Then
                        rowDGV.Cells(Me.ColGaji_Sex.Index).Value = "LAKI-LAKI"
                    Else
                        rowDGV.Cells(Me.ColGaji_Sex.Index).Value = "PEREMPUAN"
                    End If
                    rowDGV.Cells(Me.ColGaji_StartPeriod.Index).Value = .Period_Start
                    rowDGV.Cells(Me.ColGaji_Tantiem_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._07Tantiem, DefPrec)
                    rowDGV.Cells(Me.ColGaji_Tantiem_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._07_Tantiem_Rutin, DefPrec)
                    TaxType = .Tax
                    If TaxType = ClsTransaksiSalary.EnumTax.Gross_Up Then
                        rowDGV.Cells(Me.ColGaji_TaxCalType.Index).Value = "Gross Up"
                    ElseIf TaxType = ClsTransaksiSalary.EnumTax.Gross Then
                        rowDGV.Cells(Me.ColGaji_TaxCalType.Index).Value = "Gross"
                    ElseIf TaxType = ClsTransaksiSalary.EnumTax.Manual_Tax_Calculation Then
                        rowDGV.Cells(Me.ColGaji_TaxCalType.Index).Value = "-"
                    ElseIf TaxType = ClsTransaksiSalary.EnumTax.Netto Then
                        rowDGV.Cells(Me.ColGaji_TaxCalType.Index).Value = "Netto"
                    End If
                    vStatus_NPWP = .A_Status_NPWP
                    If vStatus_NPWP = ClsMasterKaryawan.GolonganPajak.TK0 Then
                        rowDGV.Cells(Me.ColGaji_TaxCategory.Index).Value = "TK0"
                    ElseIf vStatus_NPWP = ClsMasterKaryawan.GolonganPajak.TK1 Then
                        rowDGV.Cells(Me.ColGaji_TaxCategory.Index).Value = "TK1"
                    ElseIf vStatus_NPWP = ClsMasterKaryawan.GolonganPajak.TK2 Then
                        rowDGV.Cells(Me.ColGaji_TaxCategory.Index).Value = "TK2"
                    ElseIf vStatus_NPWP = ClsMasterKaryawan.GolonganPajak.TK3 Then
                        rowDGV.Cells(Me.ColGaji_TaxCategory.Index).Value = "TK3"
                    ElseIf vStatus_NPWP = ClsMasterKaryawan.GolonganPajak.K0 Then
                        rowDGV.Cells(Me.ColGaji_TaxCategory.Index).Value = "K0"
                    ElseIf vStatus_NPWP = ClsMasterKaryawan.GolonganPajak.K1 Then
                        rowDGV.Cells(Me.ColGaji_TaxCategory.Index).Value = "K1"
                    ElseIf vStatus_NPWP = ClsMasterKaryawan.GolonganPajak.K2 Then
                        rowDGV.Cells(Me.ColGaji_TaxCategory.Index).Value = "K2"
                    ElseIf vStatus_NPWP = ClsMasterKaryawan.GolonganPajak.K3 Then
                        rowDGV.Cells(Me.ColGaji_TaxCategory.Index).Value = "K3"
                    End If
                    rowDGV.Cells(Me.ColGaji_TunjLain_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._03TunjanganLainnya, DefPrec)
                    rowDGV.Cells(Me.ColGaji_TunjLain_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._03_Tunjangan_Lainnya_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_TunjPPh_Mix_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._02_Tunjangan_PPH_Disetahunkan, DefPrec)
                    rowDGV.Cells(Me.ColGaji_TunjPPh_Routine_Annual.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._02_Tunjangan_PPH_Disetahunkan_Rutin, DefPrec)
                    rowDGV.Cells(Me.ColGaji_TunjPPh_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._02_Tunjangan_PPH_Non_Rutin, DefPrec)
                    If rowSalary.Is_Resign = 1 Or Me.RefPeriod.Month = 12 Then
                        rowDGV.Cells(Me.ColGaji_TunjPPh_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._02_Tunjangan_PPH_Rutin - rowSalary._21_Tunj_PPH_FInal_Pesangon, DefPrec)
                        rowDGV.Cells(Me.ColGaji_TunjPPh_Total.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._02_Tunjangan_PPH_Rutin + rowSalary._02_Tunjangan_PPH_Non_Rutin - rowSalary._21_Tunj_PPH_FInal_Pesangon, DefPrec)
                    Else
                        rowDGV.Cells(Me.ColGaji_TunjPPh_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._02_Tunjangan_PPH_Rutin, DefPrec)
                        rowDGV.Cells(Me.ColGaji_TunjPPh_Total.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._02_Tunjangan_PPH_Rutin + rowSalary._02_Tunjangan_PPH_Non_Rutin, DefPrec)
                    End If

                    rowDGV.Cells(Me.ColGaji_Pesangon.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._21_Pesangon, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PesangonPPH.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._21_PPH_Final_Pesangon, DefPrec)
                    rowDGV.Cells(Me.ColGaji_PesangonTunjPPH.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._21_Tunj_PPH_FInal_Pesangon, DefPrec)

                    If rowSalary.Is_Resign = 1 Or Me.RefPeriod.Month = 12 Then
                        rowDGV.Cells(Me.ColGaji_Gaji_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun, _
                        0, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_Gaji_Rutin_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun, _
                        1, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_Honorarium_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.honorarium, _
                        0, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_Honorarium_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun, _
                        1, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_IuranPensiun_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.iuran_pensiun, _
                        0, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_IuranPensiun_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.iuran_pensiun, _
                        1, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_Natura_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura, _
                        0, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_Natura_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura, _
                        1, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_PremiAsuransi_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi, _
                        0, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_PremiAsuransi_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi, _
                        1, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_Tantiem_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr, _
                        0, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_Tantiem_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr, _
                        1, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_TunjLain_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya, _
                        0, TBLItem, rowSalary.Primary), DefPrec)
                        rowDGV.Cells(Me.ColGaji_TunjLain_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(Me.GetValue(ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya, _
                        1, TBLItem, rowSalary.Primary), DefPrec)
                    Else
                        rowDGV.Cells(Me.ColGaji_Gaji_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._01Gaji - rowSalary._01_Gaji_Rutin, DefPrec)
                        rowDGV.Cells(Me.ColGaji_Gaji_Rutin_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._01_Gaji_Rutin / .Period_Count, DefPrec)
                        rowDGV.Cells(Me.ColGaji_Honorarium_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._04Honorarium - rowSalary._04_Honorarium_Rutin, DefPrec)
                        rowDGV.Cells(Me.ColGaji_Honorarium_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._04_Honorarium_Rutin / .Period_Count, DefPrec)
                        rowDGV.Cells(Me.ColGaji_IuranPensiun_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._10IuranPensiun - rowSalary._10_Iuran_Pensiun_Rutin, DefPrec)
                        rowDGV.Cells(Me.ColGaji_IuranPensiun_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._10_Iuran_Pensiun_Rutin / .Period_Count, DefPrec)
                        rowDGV.Cells(Me.ColGaji_Natura_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._06Natura - rowSalary._06_Natura_Rutin, DefPrec)
                        rowDGV.Cells(Me.ColGaji_Natura_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._06_Natura_Rutin / .Period_Count, DefPrec)
                        rowDGV.Cells(Me.ColGaji_PremiAsuransi_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._05PremiAsuransi - rowSalary._05_Premi_Asurasi_Rutin, DefPrec)
                        rowDGV.Cells(Me.ColGaji_PremiAsuransi_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._05_Premi_Asurasi_Rutin / .Period_Count, DefPrec)
                        rowDGV.Cells(Me.ColGaji_Tantiem_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._07Tantiem - rowSalary._07_Tantiem_Rutin, DefPrec)
                        rowDGV.Cells(Me.ColGaji_Tantiem_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._07_Tantiem_Rutin / .Period_Count, DefPrec)
                        rowDGV.Cells(Me.ColGaji_TunjLain_NotRoutine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._03TunjanganLainnya - rowSalary._03_Tunjangan_Lainnya_Rutin, DefPrec)
                        rowDGV.Cells(Me.ColGaji_TunjLain_Routine_Value.Index).Value = Me.ObjCD.DigitGrouping(rowSalary._03_Tunjangan_Lainnya_Rutin / .Period_Count, DefPrec)
                    End If
                End With
            Next 

            Me.IsikanSubTotal()
            Me.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
            Me.AturSizeKolom()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            DTAItemOther.Dispose()
            TBLItemOther.Dispose()
        End Try
    End Sub

    Private Function GetValue(ByVal vTaxComponent As ClsMasterSalaryFormula_item.EnumTax2, ByVal vIsRutin As Integer, _
    ByVal vTBLitem As DseMonthlyTax.tbltrans_salary_itemDataTable, ByVal vPrimaryTransaksi As Double) As Decimal
        Dim rowItem As DseMonthlyTax.tbltrans_salary_itemRow
        Dim Idx As Integer = 0
        Dim RetVal As Decimal = 0
        Dim TipeKomponen As Core_Module.FrmMasterSalaryFormula_Component.EnumType = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income

        Try
            While Idx <= vTBLitem.Rows.Count - 1
                rowItem = vTBLitem.Rows(Idx)

                If rowItem.Primary_Transaksi = vPrimaryTransaksi Then
                    If rowItem.Tax_Component = CInt(vTaxComponent) And rowItem.Is_Rutin = vIsRutin Then
                        TipeKomponen = rowItem.Type

                        If vTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun Or _
                        vTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.honorarium Or _
                        vTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura Or _
                        vTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr Or _
                        vTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi Or _
                        vTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya Then
                            If TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income Or _
                            TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                                RetVal += rowItem.Amount
                            ElseIf TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                            TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                                RetVal -= rowItem.Amount
                            Else
                                MsgBox("Logic error: component tipe not defined. Please contact your provider if you see this message.", MsgBoxStyle.Information, "Information")
                                Exit Function
                            End If
                        ElseIf vTaxComponent = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.iuran_pensiun Then
                            If TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income Or _
                            TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.EmployeeExpense Then
                                RetVal += rowItem.Amount
                            ElseIf TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Deduction Or _
                            TipeKomponen = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense Then
                                RetVal -= rowItem.Amount
                            Else
                                MsgBox("Logic error: component tipe not defined. Please contact your provider if you see this message.", MsgBoxStyle.Information, "Information")
                            End If
                        Else
                            MsgBox("Logic error: tax component not defined. Please contact your provider if you see this message.", MsgBoxStyle.Information, "Information")
                        End If 
                        vTBLitem.Rows.Remove(rowItem)
                    Else
                        Idx += 1
                    End If
                ElseIf rowItem.Primary_Transaksi > vPrimaryTransaksi Then
                    Return RetVal
                Else
                    Idx += 1
                End If

            End While

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsikanSubTotal()
        Try
            Me.dgvItem.Rows.Add()
            Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Cells(Me.ColGaji_Name.Index).Value = "Total"
            Me.SetSubTotal(Me.ColGaji_Gaji_NotRoutine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_Gaji_Rutin_Value.Index)
            Me.SetSubTotal(Me.ColGaji_Honorarium_NotRoutine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_Honorarium_Routine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_IuranPensiun_NotRoutine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_IuranPensiun_Routine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_Natura_NotRoutine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_Natura_Routine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_PenghasilBruto_Mix_Value.Index)
            Me.SetSubTotal(Me.ColGaji_PPhKurangBayar.Index)
            Me.SetSubTotal(Me.ColGaji_PPhSudahBayar.Index)
            Me.SetSubTotal(Me.ColGaji_PPhTerutang_NotRoutine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_PPhTerutang_Routine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_PPhTerutang_Total.Index)
            Me.SetSubTotal(Me.ColGaji_PremiAsuransi_NotRoutine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_PremiAsuransi_Routine_Value.Index) 
            Me.SetSubTotal(Me.ColGaji_Tantiem_NotRoutine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_Tantiem_Routine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_TunjLain_NotRoutine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_TunjLain_Routine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_TunjPPh_NotRoutine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_TunjPPh_Routine_Value.Index)
            Me.SetSubTotal(Me.ColGaji_TunjPPh_Total.Index)
            Me.SetSubTotal(Me.ColGaji_PPhSudahBayarMasaSblm.Index)
            Me.SetSubTotal(Me.ColGaji_Pesangon.Index)
            Me.SetSubTotal(Me.ColGaji_PesangonPPH.Index)
            Me.SetSubTotal(Me.ColGaji_PesangonTunjPPH.Index)

            If Me.RefPeriod.Month = 12 Then
                Me.SetSubTotal(Me.ColGaji_Gaji_Mix_Annual.Index) 
                Me.SetSubTotal(Me.ColGaji_Gaji_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_Honorarium_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_Honorarium_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_IuranPensiun_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_IuranPensiun_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_JumlahPengurang_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_JumlahPengurang_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_Natura_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_Natura_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PenghasilBruto_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PenghasilBruto_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PenghNetto_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PenghNetto_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PenghNetMsPjkSblm_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PKP_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PPhTerutang_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PPhTerutang_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PKP_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PremiAsuransi_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_PremiAsuransi_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_Tantiem_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_Tantiem_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_TunjLain_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_TunjLain_Routine_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_TunjPPh_Mix_Annual.Index)
                Me.SetSubTotal(Me.ColGaji_TunjPPh_Routine_Annual.Index)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanIncomeTaxStaff_Recap_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Col As Windows.Forms.DataGridViewColumn

        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AllowUserToResizeRows = False

            For Each Col In Me.dgvH1.Columns
                Col.SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Next
            For Each Col In Me.dgvH2.Columns
                Col.SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Next
            For Each Col In Me.dgvItem.Columns
                Col.SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Next

            Me.AturSizeKolom()
            Me.LoadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim DummyString As String = "'"
        Dim ListIdxExclude As New List(Of Integer) 

        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColGaji_Gaji_NotRoutine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun, 0, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Gaji_Rutin_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun, 1, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_TunjLain_NotRoutine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya, 0, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_TunjLain_Routine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya, 1, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Honorarium_NotRoutine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.honorarium, 0, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Honorarium_Routine_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.honorarium, 1, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_PremiAsuransi_NotRoutine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi, 0, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_PremiAsuransi_Routine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi, 1, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Natura_NotRoutine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura, 0, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Natura_Routine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura, 1, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Tantiem_NotRoutine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr, 0, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Tantiem_Routine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr, 1, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_IuranPensiun_NotRoutine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.iuran_pensiun, 0, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_IuranPensiun_Routine_Value.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        Me.ExplainMonthly(Me.dgvItem.Item(Me.ColGaji_PrimaryTransaksi.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.iuran_pensiun, 1, e.RowIndex, e.ColumnIndex)
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Gaji_Mix_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then 
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun, False, _
                            Me.ColGaji_Gaji_Mix_Annual.Index, Me.ColGaji_Gaji_NotRoutine_Value.Index, Me.ColGaji_Gaji_Routine_Annual.Index, Me.ColGaji_Gaji_Rutin_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountMix(e.RowIndex, Me.ColGaji_Gaji_Rutin_Value.Index, Me.ColGaji_Gaji_NotRoutine_Value.Index, _
                            Me.ColGaji_Gaji_Mix_Annual.Index, Me.ColGaji_Gaji_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Gaji_Routine_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun, True, _
                            Me.ColGaji_Gaji_Mix_Annual.Index, Me.ColGaji_Gaji_NotRoutine_Value.Index, Me.ColGaji_Gaji_Routine_Annual.Index, Me.ColGaji_Gaji_Rutin_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountRoutine(e.RowIndex, Me.ColGaji_Gaji_Rutin_Value.Index, Me.ColGaji_Gaji_NotRoutine_Value.Index, _
                            Me.ColGaji_Gaji_Mix_Annual.Index, Me.ColGaji_Gaji_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_TunjLain_Mix_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya, False, _
                            Me.ColGaji_TunjLain_Mix_Annual.Index, Me.ColGaji_TunjLain_NotRoutine_Value.Index, Me.ColGaji_TunjLain_Routine_Annual.Index, Me.ColGaji_TunjLain_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountMix(e.RowIndex, Me.ColGaji_TunjLain_Routine_Value.Index, Me.ColGaji_TunjLain_NotRoutine_Value.Index, _
                            Me.ColGaji_TunjLain_Mix_Annual.Index, Me.ColGaji_TunjLain_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_TunjLain_Routine_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya, True, _
                            Me.ColGaji_TunjLain_Mix_Annual.Index, Me.ColGaji_TunjLain_NotRoutine_Value.Index, Me.ColGaji_TunjLain_Routine_Annual.Index, Me.ColGaji_TunjLain_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountRoutine(e.RowIndex, Me.ColGaji_TunjLain_Routine_Value.Index, Me.ColGaji_TunjLain_NotRoutine_Value.Index, _
                            Me.ColGaji_TunjLain_Mix_Annual.Index, Me.ColGaji_TunjLain_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Honorarium_Mix_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.honorarium, False, _
                            Me.ColGaji_Honorarium_Mix_Annual.Index, Me.ColGaji_Honorarium_NotRoutine_Value.Index, Me.ColGaji_Honorarium_Routine_Annual.Index, Me.ColGaji_Honorarium_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountMix(e.RowIndex, Me.ColGaji_Honorarium_Routine_Value.Index, Me.ColGaji_Honorarium_NotRoutine_Value.Index, _
                            Me.ColGaji_Honorarium_Mix_Annual.Index, Me.ColGaji_Honorarium_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Honorarium_Routine_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.honorarium, True, _
                            Me.ColGaji_Honorarium_Mix_Annual.Index, Me.ColGaji_Honorarium_NotRoutine_Value.Index, Me.ColGaji_Honorarium_Routine_Annual.Index, Me.ColGaji_Honorarium_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountRoutine(e.RowIndex, Me.ColGaji_Honorarium_Routine_Value.Index, Me.ColGaji_Honorarium_NotRoutine_Value.Index, _
                            Me.ColGaji_Honorarium_Mix_Annual.Index, Me.ColGaji_Honorarium_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_PremiAsuransi_Mix_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi, False, _
                            Me.ColGaji_PremiAsuransi_Mix_Annual.Index, Me.ColGaji_PremiAsuransi_NotRoutine_Value.Index, Me.ColGaji_PremiAsuransi_Routine_Annual.Index, Me.ColGaji_PremiAsuransi_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountMix(e.RowIndex, Me.ColGaji_PremiAsuransi_Routine_Value.Index, Me.ColGaji_PremiAsuransi_NotRoutine_Value.Index, _
                            Me.ColGaji_PremiAsuransi_Mix_Annual.Index, Me.ColGaji_PremiAsuransi_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_PremiAsuransi_Routine_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi, True, _
                            Me.ColGaji_PremiAsuransi_Mix_Annual.Index, Me.ColGaji_PremiAsuransi_NotRoutine_Value.Index, Me.ColGaji_PremiAsuransi_Routine_Annual.Index, Me.ColGaji_PremiAsuransi_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountRoutine(e.RowIndex, Me.ColGaji_PremiAsuransi_Routine_Value.Index, Me.ColGaji_PremiAsuransi_NotRoutine_Value.Index, _
                            Me.ColGaji_PremiAsuransi_Mix_Annual.Index, Me.ColGaji_PremiAsuransi_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Natura_Mix_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura, False, _
                            Me.ColGaji_Natura_Mix_Annual.Index, Me.ColGaji_Natura_NotRoutine_Value.Index, Me.ColGaji_Natura_Routine_Annual.Index, Me.ColGaji_Natura_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountMix(e.RowIndex, Me.ColGaji_Natura_Routine_Value.Index, Me.ColGaji_Natura_NotRoutine_Value.Index, _
                            Me.ColGaji_Natura_Mix_Annual.Index, Me.ColGaji_Natura_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Natura_Routine_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura, True, _
                            Me.ColGaji_Natura_Mix_Annual.Index, Me.ColGaji_Natura_NotRoutine_Value.Index, Me.ColGaji_Natura_Routine_Annual.Index, Me.ColGaji_Natura_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountRoutine(e.RowIndex, Me.ColGaji_Natura_Routine_Value.Index, Me.ColGaji_Natura_NotRoutine_Value.Index, _
                            Me.ColGaji_Natura_Mix_Annual.Index, Me.ColGaji_Natura_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Tantiem_Mix_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr, False, _
                            Me.ColGaji_Tantiem_Mix_Annual.Index, Me.ColGaji_Tantiem_NotRoutine_Value.Index, Me.ColGaji_Tantiem_Routine_Annual.Index, Me.ColGaji_Tantiem_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountMix(e.RowIndex, Me.ColGaji_Tantiem_Routine_Value.Index, Me.ColGaji_Tantiem_NotRoutine_Value.Index, _
                            Me.ColGaji_Tantiem_Mix_Annual.Index, Me.ColGaji_Tantiem_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_Tantiem_Routine_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr, True, _
                            Me.ColGaji_Tantiem_Mix_Annual.Index, Me.ColGaji_Tantiem_NotRoutine_Value.Index, Me.ColGaji_Tantiem_Routine_Annual.Index, Me.ColGaji_Tantiem_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountRoutine(e.RowIndex, Me.ColGaji_Tantiem_Routine_Value.Index, Me.ColGaji_Tantiem_NotRoutine_Value.Index, _
                            Me.ColGaji_Tantiem_Mix_Annual.Index, Me.ColGaji_Tantiem_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_IuranPensiun_Mix_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.iuran_pensiun, False, _
                            Me.ColGaji_IuranPensiun_Mix_Annual.Index, Me.ColGaji_IuranPensiun_NotRoutine_Value.Index, Me.ColGaji_IuranPensiun_Routine_Annual.Index, Me.ColGaji_IuranPensiun_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountMix(e.RowIndex, Me.ColGaji_IuranPensiun_Routine_Value.Index, Me.ColGaji_IuranPensiun_NotRoutine_Value.Index, _
                            Me.ColGaji_IuranPensiun_Mix_Annual.Index, Me.ColGaji_IuranPensiun_Routine_Annual.Index)
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGaji_IuranPensiun_Routine_Annual.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> 0 Then
                        If CInt(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value) = CInt(Me.RefPeriod.Month) Then
                            Me.ExplainYearly1(Me.dgvItem.Item(Me.ColGaji_PrimaryEmployee.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, e.RowIndex).Value, ClsMasterSalaryFormula_item.EnumTax2.iuran_pensiun, True, _
                            Me.ColGaji_IuranPensiun_Mix_Annual.Index, Me.ColGaji_IuranPensiun_NotRoutine_Value.Index, Me.ColGaji_IuranPensiun_Routine_Annual.Index, Me.ColGaji_IuranPensiun_Routine_Value.Index, e.RowIndex)
                        Else
                            Me.ExplainAnnualAmountRoutine(e.RowIndex, Me.ColGaji_IuranPensiun_Routine_Value.Index, Me.ColGaji_IuranPensiun_NotRoutine_Value.Index, _
                            Me.ColGaji_IuranPensiun_Mix_Annual.Index, Me.ColGaji_IuranPensiun_Routine_Annual.Index)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListIdxExclude = Nothing 
        End Try
    End Sub

    Private Sub ExplainAnnualAmountRoutine(ByVal vRowIdx As Integer, ByVal vColMonthlyValue_RoutineIdx As Integer, _
    ByVal vColMonthlyValue_NotRoutineIdx As Integer, ByVal vColAnnualMixValueIdx As Integer, ByVal vColAnnualRoutineValueIdx As Integer)
        Dim DummyString As String = ""
        Dim ListIdxExclude As New List(Of Integer)

        Try
            DummyString = "= Monthly Amount * Period Count"
            DummyString = DummyString + vbCrLf + "= " + Me.dgvItem.Item(vColMonthlyValue_RoutineIdx, vRowIdx).Value.ToString + " * " + Me.dgvItem.Item(Me.ColGaji_PeriodCount.Index, vRowIdx).Value.ToString + ""
            DummyString = DummyString + vbCrLf + "= " + Me.dgvItem.Item(vColAnnualRoutineValueIdx, vRowIdx).Value.ToString
            DummyString = DummyString + vbCrLf + vbCrLf + "Start Period = " + Format(CDate(Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, vRowIdx).Value.ToString + "/1/" + Me.RefPeriod.Year.ToString), "MMM yyyy")
            DummyString = DummyString + vbCrLf + "End Period = " + Format(CDate(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, vRowIdx).Value.ToString + "/1/" + Me.RefPeriod.Year.ToString), "MMM yyyy")
            DummyString = DummyString + vbCrLf + "Period Count= " + Me.dgvItem.Item(Me.ColGaji_PeriodCount.Index, vRowIdx).Value.ToString


            ListIdxExclude.Add(vColMonthlyValue_NotRoutineIdx)
            ListIdxExclude.Add(vColMonthlyValue_RoutineIdx)
            ListIdxExclude.Add(vColAnnualRoutineValueIdx)
            ListIdxExclude.Add(vColAnnualMixValueIdx)
            Me.ExplainByText(DummyString, vRowIdx, vColAnnualRoutineValueIdx, ListIdxExclude, "Explain Annual Amount (Routine)")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListIdxExclude = Nothing
        End Try
    End Sub

    Private Sub ExplainAnnualAmountMix(ByVal vRowIdx As Integer, ByVal vColMonthlyValue_RoutineIdx As Integer, _
    ByVal vColMonthlyValue_NotRoutineIdx As Integer, ByVal vColAnnualMixValueIdx As Integer, ByVal vColAnnualRoutineValueIdx As Integer)
        Dim DummyString As String = ""
        Dim ListIdxExclude As New List(Of Integer)

        Try
            DummyString = "= Routine Annual Value + Monthly Not Routine Value"
            DummyString = DummyString + vbCrLf + "= (" + Me.dgvItem.Item(vColMonthlyValue_RoutineIdx, vRowIdx).Value.ToString + " * " + Me.dgvItem.Item(Me.ColGaji_PeriodCount.Index, vRowIdx).Value.ToString + ")" + _
            " + " + Me.dgvItem.Item(vColMonthlyValue_NotRoutineIdx, vRowIdx).Value.ToString + ""
            DummyString = DummyString + vbCrLf + "= " + Me.dgvItem.Item(vColAnnualMixValueIdx, vRowIdx).Value.ToString
            DummyString = DummyString + vbCrLf + vbCrLf + "Start Period = " + Format(CDate(Me.dgvItem.Item(Me.ColGaji_StartPeriod.Index, vRowIdx).Value.ToString + "/1/" + Me.RefPeriod.Year.ToString), "MMM yyyy")
            DummyString = DummyString + vbCrLf + "End Period = " + Format(CDate(Me.dgvItem.Item(Me.ColGaji_EndPeriod.Index, vRowIdx).Value.ToString + "/1/" + Me.RefPeriod.Year.ToString), "MMM yyyy")
            DummyString = DummyString + vbCrLf + "Period Count= " + Me.dgvItem.Item(Me.ColGaji_PeriodCount.Index, vRowIdx).Value.ToString

            ListIdxExclude.Add(vColMonthlyValue_NotRoutineIdx)
            ListIdxExclude.Add(vColMonthlyValue_RoutineIdx)
            ListIdxExclude.Add(vColAnnualRoutineValueIdx)
            ListIdxExclude.Add(vColAnnualMixValueIdx)
            Me.ExplainByText(DummyString, vRowIdx, vColAnnualMixValueIdx, ListIdxExclude, "Explain Annual Amount (Routine + Not Routine)")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListIdxExclude = Nothing
        End Try
    End Sub

    Private Sub dgvItem_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles dgvItem.Scroll
        Try
            Me.dgvH1.HorizontalScrollingOffset = Me.dgvItem.HorizontalScrollingOffset
            Me.dgvH2.HorizontalScrollingOffset = Me.dgvItem.HorizontalScrollingOffset
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SizeChanged
        Try
            Me.AturSizeKolom()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturSizeKolom()
        Try
            Me.dgvH1.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            Me.dgvH1.AutoSizeRowsMode = Windows.Forms.DataGridViewAutoSizeRowsMode.None
            Me.dgvH2.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            Me.dgvH2.AutoSizeRowsMode = Windows.Forms.DataGridViewAutoSizeRowsMode.None
            Me.dgvH2.ColumnHeadersHeight = 75

            Me.ColH1_0.Width = Me.ColGaji_Code.Width + Me.ColGaji_No.Width + Me.ColGaji_Name.Width
            Me.ColH2_0.Width = Me.ColH1_0.Width

            Me.ColH1_1.Width = Me.ColGaji_DeptCode.Width + Me.ColGaji_DeptName.Width + _
            Me.ColGaji_EndPeriod.Width + Me.ColGaji_JobPositionCode.Width + Me.ColGaji_JobPositionDesc.Width + _
            Me.ColGaji_NPWP.Width + Me.ColGaji_Sex.Width + Me.ColGaji_StartPeriod.Width + _
            Me.ColGaji_TaxCategory.Width + Me.ColGaji_NIK.Width + Me.ColGaji_Nationality.Width + _
            ColGaji_TaxCalType.Width + Me.ColGaji_PeriodCount.Width

            Me.ColH2_1.Width = Me.ColH1_1.Width

            Me.ColH1_ColGaji_Value.Width = Me.ColGaji_Gaji_NotRoutine_Value.Width + Me.ColGaji_Gaji_Rutin_Value.Width
            Me.ColH1_Gaji_Annual.Width = Me.ColGaji_Gaji_Mix_Annual.Width + Me.ColGaji_Gaji_Routine_Annual.Width
            Me.ColH2_Gaji.Width = Me.ColH1_Gaji_Annual.Width + Me.ColH1_ColGaji_Value.Width

            Me.ColH1_TunjPPh_Annual.Width = Me.ColGaji_TunjPPh_Mix_Annual.Width + Me.ColGaji_TunjPPh_Routine_Annual.Width
            Me.ColH1_TunjPPh_Value.Width = Me.ColGaji_TunjPPh_NotRoutine_Value.Width + Me.ColGaji_TunjPPh_Routine_Value.Width + Me.ColGaji_TunjPPh_Total.Width
            Me.ColH2_TunjPPh21.Width = Me.ColH1_TunjPPh_Value.Width + Me.ColH1_TunjPPh_Annual.Width

            Me.ColH1_TunjLain_Value.Width = Me.ColGaji_TunjLain_NotRoutine_Value.Width + Me.ColGaji_TunjLain_Routine_Value.Width
            Me.ColH1_TunjLain_Annual.Width = Me.ColGaji_TunjLain_Mix_Annual.Width + Me.ColGaji_TunjLain_Routine_Annual.Width
            Me.ColH2_TunjLain.Width = Me.ColH1_TunjLain_Value.Width + Me.ColH1_TunjLain_Annual.Width

            Me.ColH1_Honorarium_Value.Width = Me.ColGaji_Honorarium_NotRoutine_Value.Width + Me.ColGaji_Honorarium_Routine_Value.Width
            Me.ColH1_Honorarium_Annual.Width = Me.ColGaji_Honorarium_Routine_Annual.Width + Me.ColGaji_Honorarium_Mix_Annual.Width
            Me.ColH2_Honorarium.Width = Me.ColH1_Honorarium_Value.Width + Me.ColH1_Honorarium_Annual.Width

            Me.ColH1_PremiAsuransi_Value.Width = Me.ColGaji_PremiAsuransi_NotRoutine_Value.Width + Me.ColGaji_PremiAsuransi_Routine_Value.Width
            Me.ColH1_PremiAsuransi_Annual.Width = Me.ColGaji_PremiAsuransi_Mix_Annual.Width + Me.ColGaji_PremiAsuransi_Routine_Annual.Width
            Me.ColH2_PremiAsuransi.Width = Me.ColH1_PremiAsuransi_Value.Width + Me.ColH1_PremiAsuransi_Annual.Width

            Me.ColH1_Natura_Value.Width = Me.ColGaji_Natura_NotRoutine_Value.Width + Me.ColGaji_Natura_Routine_Value.Width
            Me.ColH1_Natura_Annual.Width = Me.ColGaji_Natura_Mix_Annual.Width + Me.ColGaji_Natura_Routine_Annual.Width
            Me.ColH2_Natura.Width = Me.ColH1_Natura_Value.Width + Me.ColH1_Natura_Annual.Width

            Me.ColH1_Tantiem_Value.Width = Me.ColGaji_Tantiem_NotRoutine_Value.Width + Me.ColGaji_Tantiem_Routine_Value.Width
            Me.ColH1_Tantiem_Annual.Width = Me.ColGaji_Tantiem_Mix_Annual.Width + Me.ColGaji_Tantiem_Routine_Annual.Width
            Me.ColH2_Tantiem.Width = Me.ColH1_Tantiem_Value.Width + Me.ColH1_Tantiem_Annual.Width

            Me.ColH1_PenghasilBruto_Value.Width = Me.ColGaji_PenghasilBruto_Mix_Value.Width
            Me.ColH1_PenghasilBruto_Annual.Width = Me.ColGaji_PenghasilBruto_Mix_Value.Width + Me.ColGaji_PenghasilBruto_Routine_Annual.Width
            Me.ColH2_PenghasilBruto.Width = Me.ColH1_PenghasilBruto_Value.Width + Me.ColH1_PenghasilBruto_Annual.Width

            Me.ColH1_BiayaJabatan_Annual.Width = Me.ColGaji_BiayaJabatan_Mix_Annual.Width + Me.ColGaji_BiayaJabatan_Routine_Annual.Width
            Me.ColH2_BiayaJabatan.Width = Me.ColH1_BiayaJabatan_Annual.Width

            Me.ColH1_IuranPensiun_Value.Width = Me.ColGaji_IuranPensiun_NotRoutine_Value.Width + Me.ColGaji_IuranPensiun_Routine_Value.Width
            Me.ColH1_IuranPensiun_Annual.Width = Me.ColGaji_IuranPensiun_Routine_Annual.Width + Me.ColGaji_IuranPensiun_Mix_Annual.Width
            Me.ColH2_IuranPensiun.Width = Me.ColH1_IuranPensiun_Value.Width + Me.ColH1_IuranPensiun_Annual.Width

            Me.ColH1_JumlahPengurang_Annual.Width = Me.ColGaji_JumlahPengurang_Mix_Annual.Width + Me.ColGaji_JumlahPengurang_Routine_Annual.Width
            Me.ColH2_JumlahPengurang.Width = Me.ColH1_JumlahPengurang_Annual.Width

            Me.ColH1_PenghNetMsPjkSblm_Annual.Width = Me.ColGaji_PenghNetMsPjkSblm_Mix_Annual.Width
            Me.ColH2_PenghNetMsPjkSblm.Width = Me.ColH1_PenghNetMsPjkSblm_Annual.Width

            Me.ColH1_PenghNetto_Annual.Width = Me.ColGaji_PenghNetto_Mix_Annual.Width + Me.ColGaji_PenghNetto_Routine_Annual.Width
            Me.ColH2_PenghNetto.Width = Me.ColH1_PenghNetto_Annual.Width

            Me.ColH1_PTKP.Width = Me.ColGaji_PTKP.Width
            Me.ColH2_PTKP.Width = Me.ColH1_PTKP.Width

            Me.ColH1_PKP_Annual.Width = Me.ColGaji_PKP_Mix_Annual.Width + Me.ColGaji_PKP_Routine_Annual.Width
            Me.ColH2_PKP.Width = Me.ColH1_PKP_Annual.Width

            Me.ColH1_PPhTerutang_Value.Width = Me.ColGaji_PPhTerutang_NotRoutine_Value.Width + Me.ColGaji_PPhTerutang_Routine_Value.Width + Me.ColGaji_PPhTerutang_Total.Width
            Me.ColH1_PPhTerutang_Annual.Width = Me.ColGaji_PPhTerutang_Routine_Annual.Width + Me.ColGaji_PPhTerutang_Mix_Annual.Width
            Me.ColH2_PPhTerutang.Width = Me.ColH1_PPhTerutang_Value.Width + Me.ColH1_PPhTerutang_Annual.Width

            Me.ColH1_PPhSudahBayar.Width = Me.ColGaji_PPhSudahBayar.Width + Me.ColGaji_PPhSudahBayarMasaSblm.Width
            Me.ColH2_PPhSudahBayar.Width = Me.ColH1_PPhSudahBayar.Width

            Me.ColH1_PPhKurangBayar.Width = Me.ColGaji_PPhKurangBayar.Width
            Me.ColH2_PPhKurangBayar.Width = Me.ColH1_PPhKurangBayar.Width

            Me.ColH1_Pesangon.Width = Me.ColGaji_Pesangon.Width + Me.ColGaji_PesangonPPH.Width + Me.ColGaji_PesangonTunjPPH.Width
            Me.ColH2_Pesangon.Width = Me.ColH1_Pesangon.Width

            Me.dgvH2.Left = Me.dgvItem.Left
            Me.dgvH1.Left = Me.dgvItem.Left
            Me.dgvH1.Top = Me.dgvH2.Top + Me.dgvH2.ColumnHeadersHeight - 1
            Me.dgvItem.Top = Me.dgvH1.Top + Me.dgvH1.ColumnHeadersHeight - 1
            Me.dgvItem.Height = Me.btnClose.Top - 5 - Me.dgvItem.Top
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetSubTotal(ByVal vIndex As Integer)
        Dim rowDGV As Windows.Forms.DataGridViewRow
        Dim Total As Decimal = 0
        Dim DefPrec As Integer

        Try
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency

            For Each rowDGV In Me.dgvItem.Rows
                If IsNumeric(rowDGV.Cells(vIndex).Value) = True Then
                    Total += CDec(rowDGV.Cells(vIndex).Value)
                End If
            Next

            rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
            rowDGV.Cells(vIndex).Value = Me.ObjCD.DigitGrouping(Total, DefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExplainMonthly(ByVal vPrimaryTransaksi As Double, ByVal vEnumTaxComponent As ClsMasterSalaryFormula_item.EnumTax2, ByVal vIsRutin As Integer, _
    ByVal vRowIdx As Integer, ByVal vColIdx As Integer)
        Dim Frm1 As New FrmLaporanIncomeTaxStaff_Recap_ExplainMonthly

        Try
            If vPrimaryTransaksi <> -1 And vPrimaryTransaksi <> 0 Then
                Frm1.mPrimaryTransaksi = vPrimaryTransaksi
                Frm1.mTaxComponent = vEnumTaxComponent
                Frm1.mDefPrec = Me.DefPrec1
                Frm1.mIsRutin = vIsRutin

                Me.dgvItem.ClearSelection()
                Me.dgvItem.Rows(vRowIdx).Selected = True
                Me.dgvItem.Item(vColIdx, vRowIdx).Selected = False
                Frm1.ShowDialog()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm1.Dispose()
        End Try
    End Sub

    Private Sub ExplainByText(ByVal vExplainString As String, _
    ByVal vRowIdx As Integer, ByVal vColIdx As Integer, ByVal vListIdxExclude As List(Of Integer), ByVal vFrmText As String)
        Dim Frm1 As New FrmLaporanIncomeTaxStaff_Recap_ExplainYearly
        Dim DummyInt As Integer

        Try
            Frm1.TxtExplain.Text = vExplainString
            Frm1.Text = vFrmText
            Me.dgvItem.ClearSelection()
            Me.dgvItem.Rows(vRowIdx).Selected = True
            Me.dgvItem.Item(vColIdx, vRowIdx).Selected = False
            For Each DummyInt In vListIdxExclude
                Me.dgvItem.Item(DummyInt, vRowIdx).Selected = False
            Next
            Frm1.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm1.Dispose()
        End Try
    End Sub

    Private Sub ExplainYearly1(ByVal vPrimaryEmployee As Double, ByVal vStartPeriod As Integer, ByVal vEndPeriod As Integer, _
    ByVal vComponentType As Core_Module.ClsMasterSalaryFormula_item.EnumTax2, ByVal vIsRoutineOnly As Boolean, _
    ByVal ExcludeIdx1 As Integer, ByVal ExcludeIdx2 As Integer, ByVal ExcludeIdx3 As Integer, ByVal ExcludeIdx4 As Integer, ByVal vRowIdx As Integer)
        Dim Frm1 As New FrmLaporanIncomeTaxStaff_Recap_ExplainYearly1

        Try
            Frm1.mPrimaryEmployee = vPrimaryEmployee
            Frm1.mRefPeriod = Me.RefPeriod
            Frm1.mStartPeriod = vStartPeriod
            Frm1.mEndPeriod = vEndPeriod
            Frm1.mComponentType = vComponentType
            Frm1.mDefPrec = Me.DefPrec1
            Frm1.mIsRoutineOnly = vIsRoutineOnly

            Me.dgvItem.ClearSelection()
            Me.dgvItem.Rows(vRowIdx).Selected = True
            Me.dgvItem.Item(ExcludeIdx1, vRowIdx).Selected = False
            Me.dgvItem.Item(ExcludeIdx2, vRowIdx).Selected = False
            Me.dgvItem.Item(ExcludeIdx3, vRowIdx).Selected = False
            Me.dgvItem.Item(ExcludeIdx4, vRowIdx).Selected = False


            Frm1.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class