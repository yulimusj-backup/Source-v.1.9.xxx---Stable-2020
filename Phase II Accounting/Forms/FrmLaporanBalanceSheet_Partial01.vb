Option Explicit On
Option Strict Off

Partial Public Class FrmLaporanBalanceSheet
    Private Sub CetakBalanceSheet_6Months()
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL3 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL4 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL5 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL6 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim Bulan1, Bulan2, Bulan3, Bulan4, Bulan5, Bulan6 As Date

        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim Tipe As Double = 0
        Dim StartDate As Date = Me.dtpFrom_I.Value
        Dim StartDate1 As Date = StartDate
        Dim TotalAset, TotalHutang, TotalModal As Decimal
        Dim TotalAset1, TotalHutang1, TotalModal1 As Decimal
        Dim TotalAset2, TotalHutang2, TotalModal2 As Decimal
        Dim TotalAset3, TotalHutang3, TotalModal3 As Decimal
        Dim TotalAset4, TotalHutang4, TotalModal4 As Decimal
        Dim TotalAset5, TotalHutang5, TotalModal5 As Decimal

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            StartDate = Me.dtpFrom_I.Value
            Bulan1 = StartDate
            Me.IsiTabel(StartDate, TBL)
            'bulan 2
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Bulan2 = StartDate
            Me.IsiTabel(StartDate, TBL2)
            'bulan 3
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Bulan3 = StartDate
            Me.IsiTabel(StartDate, TBL3)
            'bulan 4
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Bulan4 = StartDate
            Me.IsiTabel(StartDate, TBL4)
            'bulan 5
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Bulan5 = StartDate
            Me.IsiTabel(StartDate, TBL5)
            'bulan 6
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Bulan6 = StartDate
            Me.IsiTabel(StartDate, TBL6)

            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL)
            End If
            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL2)
            End If

            '2. Tampilkan pada cetakan
            'generate report
            DS.TblCetak.Rows.Clear()
            For i = 0 To TBL.Rows.Count - 1
                row = TBL.Rows(i)
                rowCetak = DS.TblCetak.NewTblCetakRow()

                With rowCetak
                    'pindahkan data periode I
                    .Akun = row.Akun
                    .Header = row.Header
                    If row.Saldo_Normal = 0 Then
                        .Saldo_Period_I = row.Debet
                    Else
                        .Saldo_Period_I = row.Credit
                    End If
                    'TR//Perubahan posisi akun dibuat rata kiri sama misahin antara No / Name Account
                    If Me.chkShowAccountCode.Checked = False Then
                        .Akun = row.Akun.Trim
                    Else
                        If row.Total = 1 Then
                            .Akun_Name = row.Akun.Trim
                        Else
                            .Akun_No = row.Account_Code
                            .Akun_Name = row.Description
                            If .Akun_Name = "" Then
                                .Akun_No = ""
                                .Akun_Name = row.Account_Code
                            End If
                        End If
                    End If
                    .Total = row.Total
                    .Aktif = row.Aktif
                    .Saldo_Normal = row.Saldo_Normal
                    .Level_Perkiraan = row.Level_Perkiraan
                    .Bold = row.Bold
                    .Urutan = row.Tipe_Perkiraan

                    'pindahkan data periode II
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL2.DefaultView.RowFilter = filter
                    If TBL2.DefaultView.Count > 0 Then
                        row = TBL2.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_II = row.Debet
                        Else
                            .Saldo_Period_II = row.Credit
                        End If
                    End If

                    'pindahkan data periode III
                    TBL3.DefaultView.RowFilter = filter
                    If TBL3.DefaultView.Count > 0 Then
                        row = TBL3.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_III = row.Debet
                        Else
                            .Saldo_Period_III = row.Credit
                        End If
                    End If

                    'pindahkan data periode IV
                    TBL4.DefaultView.RowFilter = filter
                    If TBL4.DefaultView.Count > 0 Then
                        row = TBL4.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_IV = row.Debet
                        Else
                            .Saldo_Period_IV = row.Credit
                        End If
                    End If

                    'pindahkan data periode V
                    TBL5.DefaultView.RowFilter = filter
                    If TBL5.DefaultView.Count > 0 Then
                        row = TBL5.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_V = row.Debet
                        Else
                            .Saldo_Period_V = row.Credit
                        End If
                    End If

                    'pindahkan data periode VI
                    TBL6.DefaultView.RowFilter = filter
                    If TBL6.DefaultView.Count > 0 Then
                        row = TBL6.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_VI = row.Debet
                        Else
                            .Saldo_Period_VI = row.Credit
                        End If
                    End If

                    'hitung variance
                    .Variance = 0
                    .Variance_Persen = 0
                End With

                If row.Primary_Akun = 1 Then
                    TotalAset = rowCetak.Saldo_Period_I
                    TotalAset1 = rowCetak.Saldo_Period_II
                    TotalAset2 = rowCetak.Saldo_Period_III
                    TotalAset3 = rowCetak.Saldo_Period_IV
                    TotalAset4 = rowCetak.Saldo_Period_V
                    TotalAset5 = rowCetak.Saldo_Period_VI
                ElseIf row.Primary_Akun = 2 Then
                    TotalHutang = rowCetak.Saldo_Period_I
                    TotalHutang1 = rowCetak.Saldo_Period_II
                    TotalHutang2 = rowCetak.Saldo_Period_III
                    TotalHutang3 = rowCetak.Saldo_Period_IV
                    TotalHutang4 = rowCetak.Saldo_Period_V
                    TotalHutang5 = rowCetak.Saldo_Period_VI
                ElseIf row.Primary_Akun = 3 Then
                    TotalModal = rowCetak.Saldo_Period_I 
                    TotalModal1 = rowCetak.Saldo_Period_II
                    TotalModal2 = rowCetak.Saldo_Period_III
                    TotalModal3 = rowCetak.Saldo_Period_IV
                    TotalModal4 = rowCetak.Saldo_Period_V
                    TotalModal5 = rowCetak.Saldo_Period_VI
                End If

                DS.TblCetak.Rows.Add(rowCetak)
            Next

            'Buang semua akun yang nilainya 0 dan akun yang tidak aktif, tergantung pilihan user
            If Me.chkHideZero.Checked = True Then
                i = 0
                While i < DS.TblCetak.Rows.Count
                    rowCetak = DS.TblCetak.Rows(i)
                    If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                    rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 And rowCetak.Level_Perkiraan <> 1 Then
                        DS.TblCetak.Rows.Remove(rowCetak)
                    Else
                        i += 1
                    End If
                End While
            Else
                If Me.chkHideZeroInactiveAccount.Checked = True Then
                    i = 0
                    While i < DS.TblCetak.Rows.Count
                        rowCetak = DS.TblCetak.Rows(i)
                        If rowCetak.Aktif = 0 Then
                            If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                            rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 = 0 Then
                                DS.TblCetak.Rows.Remove(rowCetak)
                            Else
                                i += 1
                            End If
                        Else
                            i += 1
                        End If
                    End While
                End If
            End If


            Dim Akun As String = ""
            For ij As Integer = 0 To DS.TblCetak.Rows.Count - 1
                rowCetak = DS.TblCetak.Rows(ij)
                If Akun <> rowCetak.Akun.Substring(0, 1).ToString And rowCetak.Level_Perkiraan = 1 Then
                    Akun = rowCetak.Akun.Substring(0, 1).ToString
                    If ij = 0 Then
                        rowCetak.Group = 1
                    Else
                        rowCetak.Group = 0
                    End If
                Else
                    rowCetak.Group = 1
                End If
            Next

            Me.dtpFrom_I.Value = StartDate1
            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value)
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)

            Me.report_6Months.SetDataSource(DS)
            report_6Months.SetParameterValue(report_6Months.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            report_6Months.SetParameterValue(report_6Months.Parameter_TanggalDari_1.ParameterFieldName, vDtpFrom_1)
            report_6Months.SetParameterValue(report_6Months.Parameter_TanggalDari_2.ParameterFieldName, vDtpFrom_2)
            report_6Months.SetParameterValue(report_6Months.Parameter_Tipe.ParameterFieldName, Tipe)
            report_6Months.SetParameterValue(report_6Months.Parameter_Period_I.ParameterFieldName, "Period I")
            report_6Months.SetParameterValue(report_6Months.Parameter_Period_II.ParameterFieldName, "Period II")
            report_6Months.SetParameterValue(report_6Months.Parameter_bulan1.ParameterFieldName, Bulan1)
            report_6Months.SetParameterValue(report_6Months.Parameter_bulan2.ParameterFieldName, Bulan2)
            report_6Months.SetParameterValue(report_6Months.Parameter_bulan3.ParameterFieldName, Bulan3)
            report_6Months.SetParameterValue(report_6Months.Parameter_bulan4.ParameterFieldName, Bulan4)
            report_6Months.SetParameterValue(report_6Months.Parameter_bulan5.ParameterFieldName, Bulan5)
            report_6Months.SetParameterValue(report_6Months.Parameter_bulan6.ParameterFieldName, Bulan6)

            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report_6Months.SetParameterValue(report_6Months.Parameter_Negatif.ParameterFieldName, Me.Negatif)
            If Me.chkShowAccountCode.Checked = True Then
                Me.report_6Months.SetParameterValue(report_6Months.Parameter_ShowNoAccount.ParameterFieldName, 1)
            Else
                Me.report_6Months.SetParameterValue(report_6Months.Parameter_ShowNoAccount.ParameterFieldName, 0)
            End If
            Me.crv1.ReportSource = Me.report_6Months
            Me.crv1.Refresh()
            Me.crv1.Show()

            If Math.Round(TotalAset, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.dtpFrom_I.Value, "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset1, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang1, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal1, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32), "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset1, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang1, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal1, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset2, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang2, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal2, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + 31), "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset2, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang2, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal2, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset3, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang3, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal3, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + (31 * 2)), "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset3, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang3, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal3, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset4, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang4, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal4, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + (31 * 3)), "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset4, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang4, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal4, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset5, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang5, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal5, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + (31 * 4)), "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset5, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang5, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal5, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
            TBL3.Dispose()
            TBL4.Dispose()
            TBL5.Dispose()
            TBL6.Dispose()
        End Try
    End Sub

    Private Sub CetakBalanceSheet_6Months_WithBudget()
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL3 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL4 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL5 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL6 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim Bulan1, Bulan2, Bulan3, Bulan4, Bulan5, Bulan6 As Date

        Dim TBLB As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL2B As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL3B As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL4B As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL5B As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL6B As New DseLaporanTrialBalance.Trial_BalanceDataTable

        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim Tipe As Double = 0
        Dim StartDate As Date = Me.dtpFrom_I.Value
        Dim StartDate1 As Date = StartDate

        Dim TotalAset, TotalHutang, TotalModal As Decimal
        Dim TotalAset1, TotalHutang1, TotalModal1 As Decimal
        Dim TotalAset2, TotalHutang2, TotalModal2 As Decimal
        Dim TotalAset3, TotalHutang3, TotalModal3 As Decimal
        Dim TotalAset4, TotalHutang4, TotalModal4 As Decimal
        Dim TotalAset5, TotalHutang5, TotalModal5 As Decimal
        Dim BudgetAset, BudgetHutang, BudgetModal As Decimal
        Dim BudgetAset1, BudgetHutang1, BudgetModal1 As Decimal
        Dim BudgetAset2, BudgetHutang2, BudgetModal2 As Decimal
        Dim BudgetAset3, BudgetHutang3, BudgetModal3 As Decimal
        Dim BudgetAset4, BudgetHutang4, BudgetModal4 As Decimal
        Dim BudgetAset5, BudgetHutang5, BudgetModal5 As Decimal

        Try

            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            StartDate = Me.dtpFrom_I.Value
            Bulan1 = StartDate
            Me.IsiTabel(StartDate, TBL)
            Me.ObjBudget.GetBudget_BS(Me.ObjCD.GetLasDateOfTheMonth(StartDate), TBLB)
            'bulan 2
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Bulan2 = StartDate
            Me.IsiTabel(StartDate, TBL2)
            Me.ObjBudget.GetBudget_BS(Me.ObjCD.GetLasDateOfTheMonth(StartDate), TBL2B)
            'bulan 3
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Bulan3 = StartDate
            Me.IsiTabel(StartDate, TBL3)
            Me.ObjBudget.GetBudget_BS(Me.ObjCD.GetLasDateOfTheMonth(StartDate), TBL3B)
            'bulan 4
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Bulan4 = StartDate
            Me.IsiTabel(StartDate, TBL4)
            Me.ObjBudget.GetBudget_BS(Me.ObjCD.GetLasDateOfTheMonth(StartDate), TBL4B)
            'bulan 5
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Bulan5 = StartDate
            Me.IsiTabel(StartDate, TBL5)
            Me.ObjBudget.GetBudget_BS(Me.ObjCD.GetLasDateOfTheMonth(StartDate), TBL5B)
            'bulan 6
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Bulan6 = StartDate
            Me.IsiTabel(StartDate, TBL6)
            Me.ObjBudget.GetBudget_BS(Me.ObjCD.GetLasDateOfTheMonth(StartDate), TBL6B)

            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL)
            End If
            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL2)
            End If

            '2. Tampilkan pada cetakan
            'generate report
            DS.TblCetak.Rows.Clear()
            For i = 0 To TBL.Rows.Count - 1
                row = TBL.Rows(i)
                rowCetak = DS.TblCetak.NewTblCetakRow()

                With rowCetak
                    'pindahkan data periode I
                    .Akun = row.Akun
                    .Header = row.Header
                    If row.Saldo_Normal = 0 Then
                        .Saldo_Period_I = row.Debet
                    Else
                        .Saldo_Period_I = row.Credit
                    End If
                    'TR//Perubahan posisi akun dibuat rata kiri sama misahin antara No / Name Account
                    If Me.chkShowAccountCode.Checked = False Then
                        .Akun = row.Akun.Trim
                    Else
                        If row.Total = 1 Then
                            .Akun_Name = row.Akun.Trim
                        Else
                            .Akun_No = row.Account_Code
                            .Akun_Name = row.Description
                            If .Akun_Name = "" Then
                                .Akun_No = ""
                                .Akun_Name = row.Account_Code
                            End If
                        End If
                    End If
                    .Total = row.Total
                    .Aktif = row.Aktif
                    .Saldo_Normal = row.Saldo_Normal
                    .Level_Perkiraan = row.Level_Perkiraan
                    .Bold = row.Bold
                    .Urutan = row.Tipe_Perkiraan

                    'pindahkan data Budget I
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBLB.DefaultView.RowFilter = filter
                    If TBLB.DefaultView.Count > 0 Then
                        row = TBLB.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Budget_Period_I = row.Debet
                        Else
                            .Budget_Period_I = row.Credit
                        End If
                    Else
                        .Budget_Period_I = 0
                    End If

                    'pindahkan data periode II
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL2.DefaultView.RowFilter = filter
                    If TBL2.DefaultView.Count > 0 Then
                        row = TBL2.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_II = row.Debet
                        Else
                            .Saldo_Period_II = row.Credit
                        End If
                    End If

                    'pindahkan data Budget II
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL2B.DefaultView.RowFilter = filter
                    If TBL2B.DefaultView.Count > 0 Then
                        row = TBL2B.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Budget_Period_II = row.Debet
                        Else
                            .Budget_Period_II = row.Credit
                        End If
                    Else
                        .Budget_Period_II = 0
                    End If

                    'pindahkan data periode III
                    TBL3.DefaultView.RowFilter = filter
                    If TBL3.DefaultView.Count > 0 Then
                        row = TBL3.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_III = row.Debet
                        Else
                            .Saldo_Period_III = row.Credit
                        End If
                    End If

                    'pindahkan data Budget III
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL3B.DefaultView.RowFilter = filter
                    If TBL3B.DefaultView.Count > 0 Then
                        row = TBL3B.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Budget_Period_III = row.Debet
                        Else
                            .Budget_Period_III = row.Credit
                        End If
                    Else
                        .Budget_Period_III = 0
                    End If

                    'pindahkan data periode IV
                    TBL4.DefaultView.RowFilter = filter
                    If TBL4.DefaultView.Count > 0 Then
                        row = TBL4.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_IV = row.Debet
                        Else
                            .Saldo_Period_IV = row.Credit
                        End If
                    End If

                    'pindahkan data Budget IV
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL4B.DefaultView.RowFilter = filter
                    If TBL4B.DefaultView.Count > 0 Then
                        row = TBL4B.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Budget_Period_IV = row.Debet
                        Else
                            .Budget_Period_IV = row.Credit
                        End If
                    Else
                        .Budget_Period_IV = 0
                    End If

                    'pindahkan data periode V
                    TBL5.DefaultView.RowFilter = filter
                    If TBL5.DefaultView.Count > 0 Then
                        row = TBL5.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_V = row.Debet
                        Else
                            .Saldo_Period_V = row.Credit
                        End If
                    End If

                    'pindahkan data Budget V
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL5B.DefaultView.RowFilter = filter
                    If TBL5B.DefaultView.Count > 0 Then
                        row = TBL5B.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Budget_Period_V = row.Debet
                        Else
                            .Budget_Period_V = row.Credit
                        End If
                    Else
                        .Budget_Period_V = 0
                    End If

                    'pindahkan data periode VI
                    TBL6.DefaultView.RowFilter = filter
                    If TBL6.DefaultView.Count > 0 Then
                        row = TBL6.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_VI = row.Debet
                        Else
                            .Saldo_Period_VI = row.Credit
                        End If
                    End If

                    'pindahkan data Budget VI
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL6B.DefaultView.RowFilter = filter
                    If TBL6B.DefaultView.Count > 0 Then
                        row = TBL6B.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Budget_Period_VI = row.Debet
                        Else
                            .Budget_Period_VI = row.Credit
                        End If
                    Else
                        .Budget_Period_VI = 0
                    End If

                    'hitung variance
                    .Variance = 0
                    .Variance_Persen = 0
                End With

                If row.Primary_Akun = 1 Then
                    TotalAset = rowCetak.Saldo_Period_I
                    TotalAset1 = rowCetak.Saldo_Period_II
                    TotalAset2 = rowCetak.Saldo_Period_III
                    TotalAset3 = rowCetak.Saldo_Period_IV
                    TotalAset4 = rowCetak.Saldo_Period_V
                    TotalAset5 = rowCetak.Saldo_Period_VI
                    BudgetAset = rowCetak.Budget_Period_I
                    BudgetAset1 = rowCetak.Budget_Period_II
                    BudgetAset2 = rowCetak.Budget_Period_III
                    BudgetAset3 = rowCetak.Budget_Period_IV
                    BudgetAset4 = rowCetak.Budget_Period_V
                    BudgetAset5 = rowCetak.Budget_Period_VI
                ElseIf row.Primary_Akun = 2 Then
                    TotalHutang = rowCetak.Saldo_Period_I
                    TotalHutang1 = rowCetak.Saldo_Period_II
                    TotalHutang2 = rowCetak.Saldo_Period_III
                    TotalHutang3 = rowCetak.Saldo_Period_IV
                    TotalHutang4 = rowCetak.Saldo_Period_V
                    TotalHutang5 = rowCetak.Saldo_Period_VI
                    BudgetHutang = rowCetak.Budget_Period_I
                    BudgetHutang1 = rowCetak.Budget_Period_II
                    BudgetHutang2 = rowCetak.Budget_Period_III
                    BudgetHutang3 = rowCetak.Budget_Period_IV
                    BudgetHutang4 = rowCetak.Budget_Period_V
                    BudgetHutang5 = rowCetak.Budget_Period_VI
                ElseIf row.Primary_Akun = 3 Then
                    TotalModal = rowCetak.Saldo_Period_I
                    TotalModal1 = rowCetak.Saldo_Period_II
                    TotalModal2 = rowCetak.Saldo_Period_III
                    TotalModal3 = rowCetak.Saldo_Period_IV
                    TotalModal4 = rowCetak.Saldo_Period_V
                    TotalModal5 = rowCetak.Saldo_Period_VI
                    BudgetModal = rowCetak.Budget_Period_I
                    BudgetModal1 = rowCetak.Budget_Period_II
                    BudgetModal2 = rowCetak.Budget_Period_III
                    BudgetModal3 = rowCetak.Budget_Period_IV
                    BudgetModal4 = rowCetak.Budget_Period_V
                    BudgetModal5 = rowCetak.Budget_Period_VI
                End If

                DS.TblCetak.Rows.Add(rowCetak)
            Next

            'Buang semua akun yang nilainya 0 dan akun yang tidak aktif, tergantung pilihan user
            If Me.chkHideZero.Checked = True Then
                i = 0
                While i < DS.TblCetak.Rows.Count
                    rowCetak = DS.TblCetak.Rows(i)
                    If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                    rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 And rowCetak.Level_Perkiraan <> 1 Then
                        DS.TblCetak.Rows.Remove(rowCetak)
                    Else
                        i += 1
                    End If
                End While
            Else
                If Me.chkHideZeroInactiveAccount.Checked = True Then
                    i = 0
                    While i < DS.TblCetak.Rows.Count
                        rowCetak = DS.TblCetak.Rows(i)
                        If rowCetak.Aktif = 0 Then
                            If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                            rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 = 0 Then
                                DS.TblCetak.Rows.Remove(rowCetak)
                            Else
                                i += 1
                            End If
                        Else
                            i += 1
                        End If
                    End While
                End If
            End If


            Dim Akun As String = ""
            For ij As Integer = 0 To DS.TblCetak.Rows.Count - 1
                rowCetak = DS.TblCetak.Rows(ij)
                If Akun <> rowCetak.Akun.Substring(0, 1).ToString And rowCetak.Level_Perkiraan = 1 Then
                    Akun = rowCetak.Akun.Substring(0, 1).ToString
                    If ij = 0 Then
                        rowCetak.Group = 1
                    Else
                        rowCetak.Group = 0
                    End If
                Else
                    rowCetak.Group = 1
                End If
            Next

            Me.dtpFrom_I.Value = StartDate1
            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value)
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)

            Me.report_6Months_w_Budget.SetDataSource(DS)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_TanggalDari_1.ParameterFieldName, vDtpFrom_1)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_TanggalDari_2.ParameterFieldName, vDtpFrom_2)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_Tipe.ParameterFieldName, Tipe)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_Period_I.ParameterFieldName, "Period I")
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_Period_II.ParameterFieldName, "Period II")
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_bulan1.ParameterFieldName, Bulan1)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_bulan2.ParameterFieldName, Bulan2)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_bulan3.ParameterFieldName, Bulan3)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_bulan4.ParameterFieldName, Bulan4)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_bulan5.ParameterFieldName, Bulan5)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_bulan6.ParameterFieldName, Bulan6)
            report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_Tipe.ParameterFieldName, 1)

            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_Negatif.ParameterFieldName, Me.Negatif)
            If Me.chkShowAccountCode.Checked = True Then
                Me.report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_ShowNoAccount.ParameterFieldName, 1)
            Else
                Me.report_6Months_w_Budget.SetParameterValue(report_6Months_w_Budget.Parameter_ShowNoAccount.ParameterFieldName, 0)
            End If
            Me.crv1.ReportSource = Me.report_6Months_w_Budget
            Me.crv1.Refresh()
            Me.crv1.Show()

            If Math.Round(TotalAset, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.dtpFrom_I.Value, "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset1, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang1, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal1, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32), "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset1, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang1, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal1, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset2, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang2, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal2, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + 31), "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset2, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang2, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal2, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset3, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang3, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal3, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + (31 * 2)), "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset3, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang3, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal3, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset4, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang4, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal4, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + (31 * 3)), "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset4, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang4, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal4, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset5, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang5, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal5, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + (31 * 4)), "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset5, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang5, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal5, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(BudgetAset, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(BudgetHutang, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(BudgetModal, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check budget data on " + Format(Me.dtpFrom_I.Value, "MMM yyyy") + vbCrLf + _
                "Budget Asset: " + Me.ObjCD.DigitGrouping(BudgetAset, Me.Def_Prec) + vbCrLf + _
                "Budget Liabilities: " + Me.ObjCD.DigitGrouping(BudgetHutang, Me.Def_Prec) + vbCrLf + _
                "Budget Equity: " + Me.ObjCD.DigitGrouping(BudgetModal, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(BudgetAset1, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(BudgetHutang1, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(BudgetModal1, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check budget data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32), "MMM yyyy") + vbCrLf + _
                "Budget Asset: " + Me.ObjCD.DigitGrouping(BudgetAset1, Me.Def_Prec) + vbCrLf + _
                "Budget Liabilities: " + Me.ObjCD.DigitGrouping(BudgetHutang1, Me.Def_Prec) + vbCrLf + _
                "Budget Equity: " + Me.ObjCD.DigitGrouping(BudgetModal1, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(BudgetAset2, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(BudgetHutang2, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(BudgetModal2, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check budget data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + 31), "MMM yyyy") + vbCrLf + _
                "Budget Asset: " + Me.ObjCD.DigitGrouping(BudgetAset2, Me.Def_Prec) + vbCrLf + _
                "Budget Liabilities: " + Me.ObjCD.DigitGrouping(BudgetHutang2, Me.Def_Prec) + vbCrLf + _
                "Budget Equity: " + Me.ObjCD.DigitGrouping(BudgetModal2, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(BudgetAset3, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(BudgetHutang3, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(BudgetModal3, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check budget data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + (31 * 2)), "MMM yyyy") + vbCrLf + _
                "Budget Asset: " + Me.ObjCD.DigitGrouping(BudgetAset3, Me.Def_Prec) + vbCrLf + _
                "Budget Liabilities: " + Me.ObjCD.DigitGrouping(BudgetHutang3, Me.Def_Prec) + vbCrLf + _
                "Budget Equity: " + Me.ObjCD.DigitGrouping(BudgetModal3, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(BudgetAset4, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(BudgetHutang4, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(BudgetModal4, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check budget data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + (31 * 3)), "MMM yyyy") + vbCrLf + _
                "Budget Asset: " + Me.ObjCD.DigitGrouping(BudgetAset4, Me.Def_Prec) + vbCrLf + _
                "Budget Liabilities: " + Me.ObjCD.DigitGrouping(BudgetHutang4, Me.Def_Prec) + vbCrLf + _
                "Budget Equity: " + Me.ObjCD.DigitGrouping(BudgetModal4, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(BudgetAset5, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(BudgetHutang5, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(BudgetModal5, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check budget data on " + Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value) + TimeSpan.FromDays(32 + (31 * 4)), "MMM yyyy") + vbCrLf + _
                "Budget Asset: " + Me.ObjCD.DigitGrouping(BudgetAset5, Me.Def_Prec) + vbCrLf + _
                "Budget Liabilities: " + Me.ObjCD.DigitGrouping(BudgetHutang5, Me.Def_Prec) + vbCrLf + _
                "Budget Equity: " + Me.ObjCD.DigitGrouping(BudgetModal5, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
            TBL3.Dispose()
            TBL4.Dispose()
            TBL5.Dispose()
            TBL6.Dispose()
        End Try
    End Sub

    Private Sub CetakBalanceSheet_12Months()
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL3 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL4 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL5 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL6 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL7 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL8 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL9 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL10 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL11 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL12 As New DseLaporanTrialBalance.Trial_BalanceDataTable

        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim Tipe As Double = 0
        Dim StartDate As Date = Me.dtpFrom_I.Value
        Dim StartDate1 As Date = StartDate

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            StartDate = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpFrom_I.Value)
            Me.IsiTabel(StartDate, TBL)
            'bulan 2
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL2)
            'bulan 3
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL3)
            'bulan 4
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL4)
            'bulan 5
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL5)
            'bulan 6
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL6)
            'bulan 7
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL7)
            'bulan 8
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL8)
            'bulan 9
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL9)
            'bulan 10
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL10)
            'bulan 11
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL11)
            'bulan 12
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL12)

            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL)
            End If
            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL2)
            End If

            '2. Tampilkan pada cetakan
            'generate report
            DS.TblCetak.Rows.Clear()
            For i = 0 To TBL.Rows.Count - 1
                row = TBL.Rows(i)
                rowCetak = DS.TblCetak.NewTblCetakRow()

                With rowCetak
                    'pindahkan data periode I
                    .Akun = row.Akun
                    .Header = row.Header
                    If row.Saldo_Normal = 0 Then
                        .Saldo_Period_I = row.Debet
                    Else
                        .Saldo_Period_I = row.Credit
                    End If
                    .Total = row.Total
                    .Aktif = row.Aktif
                    .Saldo_Normal = row.Saldo_Normal
                    .Level_Perkiraan = row.Level_Perkiraan
                    .Bold = row.Bold
                    .Urutan = row.Tipe_Perkiraan

                    'pindahkan data periode II
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL2.DefaultView.RowFilter = filter
                    If TBL2.DefaultView.Count > 0 Then
                        row = TBL2.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_II = row.Debet
                        Else
                            .Saldo_Period_II = row.Credit
                        End If
                    End If

                    'pindahkan data periode III
                    TBL3.DefaultView.RowFilter = filter
                    If TBL3.DefaultView.Count > 0 Then
                        row = TBL3.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_III = row.Debet
                        Else
                            .Saldo_Period_III = row.Credit
                        End If
                    End If

                    'pindahkan data periode IV
                    TBL4.DefaultView.RowFilter = filter
                    If TBL4.DefaultView.Count > 0 Then
                        row = TBL4.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_IV = row.Debet
                        Else
                            .Saldo_Period_IV = row.Credit
                        End If
                    End If

                    'pindahkan data periode V
                    TBL5.DefaultView.RowFilter = filter
                    If TBL5.DefaultView.Count > 0 Then
                        row = TBL5.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_V = row.Debet
                        Else
                            .Saldo_Period_V = row.Credit
                        End If
                    End If

                    'pindahkan data periode VI
                    TBL6.DefaultView.RowFilter = filter
                    If TBL6.DefaultView.Count > 0 Then
                        row = TBL6.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_VI = row.Debet
                        Else
                            .Saldo_Period_VI = row.Credit
                        End If
                    End If

                    'pindahkan data periode VII
                    TBL7.DefaultView.RowFilter = filter
                    If TBL7.DefaultView.Count > 0 Then
                        row = TBL7.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_VII = row.Debet
                        Else
                            .Saldo_Period_VII = row.Credit
                        End If
                    End If

                    'pindahkan data periode VIII
                    TBL8.DefaultView.RowFilter = filter
                    If TBL8.DefaultView.Count > 0 Then
                        row = TBL8.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_VIII = row.Debet
                        Else
                            .Saldo_Period_VIII = row.Credit
                        End If
                    End If

                    'pindahkan data periode IX
                    TBL9.DefaultView.RowFilter = filter
                    If TBL9.DefaultView.Count > 0 Then
                        row = TBL9.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_IX = row.Debet
                        Else
                            .Saldo_Period_IX = row.Credit
                        End If
                    End If

                    'pindahkan data periode X
                    TBL10.DefaultView.RowFilter = filter
                    If TBL10.DefaultView.Count > 0 Then
                        row = TBL10.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_X = row.Debet
                        Else
                            .Saldo_Period_X = row.Credit
                        End If
                    End If

                    'pindahkan data periode XI
                    TBL11.DefaultView.RowFilter = filter
                    If TBL11.DefaultView.Count > 0 Then
                        row = TBL11.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_XI = row.Debet
                        Else
                            .Saldo_Period_XI = row.Credit
                        End If
                    End If

                    'pindahkan data periode XII
                    TBL12.DefaultView.RowFilter = filter
                    If TBL12.DefaultView.Count > 0 Then
                        row = TBL12.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_XII = row.Debet
                        Else
                            .Saldo_Period_XII = row.Credit
                        End If
                    End If

                    'hitung variance
                    .Variance = 0
                    .Variance_Persen = 0
                End With

                DS.TblCetak.Rows.Add(rowCetak)
            Next

            'Buang semua akun yang nilainya 0 dan akun yang tidak aktif, tergantung pilihan user
            If Me.chkHideZero.Checked = True Then
                i = 0
                While i < DS.TblCetak.Rows.Count
                    rowCetak = DS.TblCetak.Rows(i)
                    If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                    rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 And _
                    rowCetak.Saldo_Period_VII = 0 And rowCetak.Saldo_Period_VIII = 0 And rowCetak.Saldo_Period_IX = 0 And _
                    rowCetak.Saldo_Period_X = 0 And rowCetak.Saldo_Period_XI = 0 And rowCetak.Saldo_Period_XII = 0 And _
                    rowCetak.Level_Perkiraan <> 1 Then
                        DS.TblCetak.Rows.Remove(rowCetak)
                    Else
                        i += 1
                    End If
                End While
            Else
                If Me.chkHideZeroInactiveAccount.Checked = True Then
                    i = 0
                    While i < DS.TblCetak.Rows.Count
                        rowCetak = DS.TblCetak.Rows(i)
                        If rowCetak.Aktif = 0 Then
                            If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                            rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 = 0 And _
                            rowCetak.Saldo_Period_VII = 0 And rowCetak.Saldo_Period_VIII = 0 And rowCetak.Saldo_Period_IX = 0 And _
                            rowCetak.Saldo_Period_X = 0 And rowCetak.Saldo_Period_XI = 0 And rowCetak.Saldo_Period_XII = 0 Then
                                DS.TblCetak.Rows.Remove(rowCetak)
                            Else
                                i += 1
                            End If
                        Else
                            i += 1
                        End If
                    End While
                End If
            End If


            Dim Akun As String = ""
            For ij As Integer = 0 To DS.TblCetak.Rows.Count - 1
                rowCetak = DS.TblCetak.Rows(ij)
                If Akun <> rowCetak.Akun.Substring(0, 1).ToString And rowCetak.Level_Perkiraan = 1 Then
                    Akun = rowCetak.Akun.Substring(0, 1).ToString
                    If ij = 0 Then
                        rowCetak.Group = 1
                    Else
                        rowCetak.Group = 0
                    End If
                Else
                    rowCetak.Group = 1
                End If
            Next

            Me.dtpFrom_I.Value = StartDate1
            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value)
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)

            Me.report_12Months.SetDataSource(DS)
            report_12Months.SetParameterValue(report_12Months.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            report_12Months.SetParameterValue(report_12Months.Parameter_TanggalDari_1.ParameterFieldName, vDtpFrom_1)
            report_12Months.SetParameterValue(report_12Months.Parameter_TanggalDari_2.ParameterFieldName, vDtpFrom_2)
            report_12Months.SetParameterValue(report_12Months.Parameter_Tipe.ParameterFieldName, Tipe)
            report_12Months.SetParameterValue(report_12Months.Parameter_Period_I.ParameterFieldName, "Period I")
            report_12Months.SetParameterValue(report_12Months.Parameter_Period_II.ParameterFieldName, "Period II")

            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report_12Months.SetParameterValue(report_12Months.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.report_12Months
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
            TBL3.Dispose()
            TBL4.Dispose()
            TBL5.Dispose()
            TBL6.Dispose()
        End Try
    End Sub

    Private Sub CetakBalanceSheet_12MonthsWithBudget()
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL3 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL4 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL5 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL6 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL7 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL8 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL9 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL10 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL11 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL12 As New DseLaporanTrialBalance.Trial_BalanceDataTable

        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim Tipe As Double = 0
        Dim StartDate As Date = Me.dtpFrom_I.Value
        Dim StartDate1 As Date = StartDate

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            StartDate = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpFrom_I.Value)
            Me.IsiTabel(StartDate, TBL)
            'bulan 2
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL2)
            'bulan 3
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL3)
            'bulan 4
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL4)
            'bulan 5
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL5)
            'bulan 6
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL6)
            'bulan 7
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL7)
            'bulan 8
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL8)
            'bulan 9
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL9)
            'bulan 10
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL10)
            'bulan 11
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL11)
            'bulan 12
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            Me.IsiTabel(StartDate, TBL12)

            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL)
            End If
            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL2)
            End If

            '2. Tampilkan pada cetakan
            'generate report
            DS.TblCetak.Rows.Clear()
            For i = 0 To TBL.Rows.Count - 1
                row = TBL.Rows(i)
                rowCetak = DS.TblCetak.NewTblCetakRow()

                With rowCetak
                    'pindahkan data periode I
                    .Akun = row.Akun
                    .Header = row.Header
                    If row.Saldo_Normal = 0 Then
                        .Saldo_Period_I = row.Debet
                    Else
                        .Saldo_Period_I = row.Credit
                    End If
                    .Total = row.Total
                    .Aktif = row.Aktif
                    .Saldo_Normal = row.Saldo_Normal
                    .Level_Perkiraan = row.Level_Perkiraan
                    .Bold = row.Bold
                    .Urutan = row.Tipe_Perkiraan

                    'pindahkan data periode II
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL2.DefaultView.RowFilter = filter
                    If TBL2.DefaultView.Count > 0 Then
                        row = TBL2.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_II = row.Debet
                        Else
                            .Saldo_Period_II = row.Credit
                        End If
                    End If

                    'pindahkan data periode III
                    TBL3.DefaultView.RowFilter = filter
                    If TBL3.DefaultView.Count > 0 Then
                        row = TBL3.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_III = row.Debet
                        Else
                            .Saldo_Period_III = row.Credit
                        End If
                    End If

                    'pindahkan data periode IV
                    TBL4.DefaultView.RowFilter = filter
                    If TBL4.DefaultView.Count > 0 Then
                        row = TBL4.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_IV = row.Debet
                        Else
                            .Saldo_Period_IV = row.Credit
                        End If
                    End If

                    'pindahkan data periode V
                    TBL5.DefaultView.RowFilter = filter
                    If TBL5.DefaultView.Count > 0 Then
                        row = TBL5.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_V = row.Debet
                        Else
                            .Saldo_Period_V = row.Credit
                        End If
                    End If

                    'pindahkan data periode VI
                    TBL6.DefaultView.RowFilter = filter
                    If TBL6.DefaultView.Count > 0 Then
                        row = TBL6.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_VI = row.Debet
                        Else
                            .Saldo_Period_VI = row.Credit
                        End If
                    End If

                    'pindahkan data periode VII
                    TBL7.DefaultView.RowFilter = filter
                    If TBL7.DefaultView.Count > 0 Then
                        row = TBL7.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_VII = row.Debet
                        Else
                            .Saldo_Period_VII = row.Credit
                        End If
                    End If

                    'pindahkan data periode VIII
                    TBL8.DefaultView.RowFilter = filter
                    If TBL8.DefaultView.Count > 0 Then
                        row = TBL8.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_VIII = row.Debet
                        Else
                            .Saldo_Period_VIII = row.Credit
                        End If
                    End If

                    'pindahkan data periode IX
                    TBL9.DefaultView.RowFilter = filter
                    If TBL9.DefaultView.Count > 0 Then
                        row = TBL9.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_IX = row.Debet
                        Else
                            .Saldo_Period_IX = row.Credit
                        End If
                    End If

                    'pindahkan data periode X
                    TBL10.DefaultView.RowFilter = filter
                    If TBL10.DefaultView.Count > 0 Then
                        row = TBL10.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_X = row.Debet
                        Else
                            .Saldo_Period_X = row.Credit
                        End If
                    End If

                    'pindahkan data periode XI
                    TBL11.DefaultView.RowFilter = filter
                    If TBL11.DefaultView.Count > 0 Then
                        row = TBL11.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_XI = row.Debet
                        Else
                            .Saldo_Period_XI = row.Credit
                        End If
                    End If

                    .Budget_Period_I = 0
                    .Budget_Period_II = 0
                    .Budget_Period_III = 0
                    .Budget_Period_IV = 0
                    .Budget_Period_IX = 0
                    .Budget_Period_V = 0
                    .Budget_Period_VI = 0
                    .Budget_Period_VII = 0
                    .Budget_Period_VIII = 0
                    .Budget_Period_X = 0
                    .Budget_Period_XI = 0
                    .Budget_Period_XII = 0

                    'pindahkan data periode XII
                    TBL12.DefaultView.RowFilter = filter
                    If TBL12.DefaultView.Count > 0 Then
                        row = TBL12.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_XII = row.Debet
                        Else
                            .Saldo_Period_XII = row.Credit
                        End If
                    End If

                    'hitung variance
                    .Variance = 0
                    .Variance_Persen = 0
                End With

                DS.TblCetak.Rows.Add(rowCetak)
            Next

            'Buang semua akun yang nilainya 0 dan akun yang tidak aktif, tergantung pilihan user
            If Me.chkHideZero.Checked = True Then
                i = 0
                While i < DS.TblCetak.Rows.Count
                    rowCetak = DS.TblCetak.Rows(i)
                    If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                    rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 And _
                    rowCetak.Saldo_Period_VII = 0 And rowCetak.Saldo_Period_VIII = 0 And rowCetak.Saldo_Period_IX = 0 And _
                    rowCetak.Saldo_Period_X = 0 And rowCetak.Saldo_Period_XI = 0 And rowCetak.Saldo_Period_XII = 0 And _
                    rowCetak.Level_Perkiraan <> 1 Then
                        DS.TblCetak.Rows.Remove(rowCetak)
                    Else
                        i += 1
                    End If
                End While
            Else
                If Me.chkHideZeroInactiveAccount.Checked = True Then
                    i = 0
                    While i < DS.TblCetak.Rows.Count
                        rowCetak = DS.TblCetak.Rows(i)
                        If rowCetak.Aktif = 0 Then
                            If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                            rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 = 0 And _
                            rowCetak.Saldo_Period_VII = 0 And rowCetak.Saldo_Period_VIII = 0 And rowCetak.Saldo_Period_IX = 0 And _
                            rowCetak.Saldo_Period_X = 0 And rowCetak.Saldo_Period_XI = 0 And rowCetak.Saldo_Period_XII = 0 Then
                                DS.TblCetak.Rows.Remove(rowCetak)
                            Else
                                i += 1
                            End If
                        Else
                            i += 1
                        End If
                    End While
                End If
            End If


            Dim Akun As String = ""
            For ij As Integer = 0 To DS.TblCetak.Rows.Count - 1
                rowCetak = DS.TblCetak.Rows(ij)
                If Akun <> rowCetak.Akun.Substring(0, 1).ToString And rowCetak.Level_Perkiraan = 1 Then
                    Akun = rowCetak.Akun.Substring(0, 1).ToString
                    If ij = 0 Then
                        rowCetak.Group = 1
                    Else
                        rowCetak.Group = 0
                    End If
                Else
                    rowCetak.Group = 1
                End If
            Next

            Me.dtpFrom_I.Value = StartDate1
            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value)
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)

            Me.report_12Months_w_Budget.SetDataSource(DS)
            report_12Months_w_Budget.SetParameterValue(report_12Months_w_Budget.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            report_12Months_w_Budget.SetParameterValue(report_12Months_w_Budget.Parameter_TanggalDari_1.ParameterFieldName, vDtpFrom_1)
            report_12Months_w_Budget.SetParameterValue(report_12Months_w_Budget.Parameter_TanggalDari_2.ParameterFieldName, vDtpFrom_2)
            report_12Months_w_Budget.SetParameterValue(report_12Months_w_Budget.Parameter_Tipe.ParameterFieldName, Tipe)
            report_12Months_w_Budget.SetParameterValue(report_12Months_w_Budget.Parameter_Period_I.ParameterFieldName, "Period I")
            report_12Months_w_Budget.SetParameterValue(report_12Months_w_Budget.Parameter_Period_II.ParameterFieldName, "Period II")

            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report_12Months_w_Budget.SetParameterValue(report_12Months_w_Budget.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.report_12Months_w_Budget
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
            TBL3.Dispose()
            TBL4.Dispose()
            TBL5.Dispose()
            TBL6.Dispose()
        End Try
    End Sub
End Class
