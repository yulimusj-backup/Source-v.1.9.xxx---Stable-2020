Option Explicit On
Option Strict Off

Partial Public Class FrmLaporan_ProfitNLoss
    Dim ReportYTD As New CrpLaporanProfitNLoss_2PeriodYTD

    Public Sub Cetak_P_L_2Period_YTD(ByVal Adjusted As Boolean)
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL_YTD1 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL_YTD2 As New DseLaporanTrialBalance.Trial_BalanceDataTable

        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpTo_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim vDtpTo_2 As Date = Now
        Dim tipe As Double = 0

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL, False)
            Me.IsiTabel(Adjusted, Me.dtpTo_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value), TBL2)
            Me.IsiTabel(Adjusted, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpFrom_I.Value), Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL_YTD1)
            Me.IsiTabel(Adjusted, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpTo_I.Value), Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value), TBL_YTD2)

            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL)
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
                    TBL_YTD1.DefaultView.RowFilter = filter
                    If TBL_YTD1.DefaultView.Count > 0 Then
                        row = TBL_YTD1.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_II = row.Debet
                        Else
                            .Saldo_Period_II = row.Credit
                        End If
                    End If

                    'pindahkan data periode III
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL2.DefaultView.RowFilter = filter
                    If TBL2.DefaultView.Count > 0 Then
                        row = TBL2.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_III = row.Debet
                        Else
                            .Saldo_Period_III = row.Credit
                        End If
                    End If

                    'pindahkan data periode IV
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL_YTD2.DefaultView.RowFilter = filter
                    If TBL_YTD2.DefaultView.Count > 0 Then
                        row = TBL_YTD2.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_IV = row.Debet
                        Else
                            .Saldo_Period_IV = row.Credit
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
                    rowCetak.Saldo_Period_IV = 0 And rowCetak.Level_Perkiraan <> 1 Then
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
                            rowCetak.Saldo_Period_IV = 0 Then
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

            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value)
            vDtpTo_1 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value)
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)
            vDtpTo_2 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_II.Value)

            Me.ReportYTD.SetDataSource(DS)
            Me.ReportYTD.SetParameterValue(ReportYTD.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.ReportYTD.SetParameterValue(ReportYTD.Parameter_TanggalDari_1.ParameterFieldName, vDtpFrom_1)
            Me.ReportYTD.SetParameterValue(ReportYTD.Parameter_TanggalSampai_1.ParameterFieldName, vDtpTo_1)
            Me.ReportYTD.SetParameterValue(ReportYTD.Parameter_TanggalDari_2.ParameterFieldName, vDtpTo_1)
            Me.ReportYTD.SetParameterValue(ReportYTD.Parameter_TanggalSampai_2.ParameterFieldName, vDtpTo_1)
            Me.ReportYTD.SetParameterValue(ReportYTD.Parameter_Tipe.ParameterFieldName, tipe)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.ReportYTD.SetParameterValue(ReportYTD.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.ReportYTD
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
        End Try
    End Sub

    Public Sub Cetak_P_L_6MonthsPeriod(ByVal Adjusted As Boolean)
        Dim TBL1 As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL3 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL4 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL5 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL6 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBLTotal As New DseLaporanTrialBalance.Trial_BalanceDataTable

        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpTo_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim vDtpTo_2 As Date = Now
        Dim tipe As Double = 0
        Dim StartDate As Date = Me.dtpFrom_I.Value.Date
        Dim Bulan1, Bulan2, Bulan3, Bulan4, Bulan5, Bulan6 As Date

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Bulan1 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL1, False)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan2 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL2)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan3 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL3)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan4 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL4)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan5 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL5)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan6 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL6)

            Me.IsiTabel(Adjusted, Bulan1, Me.ObjCD.GetLasDateOfTheMonth(Bulan6), TBLTotal)


            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL1)
            End If

            '2. Tampilkan pada cetakan
            'generate report
            DS.TblCetak.Rows.Clear()
            For i = 0 To TBL1.Rows.Count - 1
                row = TBL1.Rows(i)
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

                    'Hitung total periode I - VI
                    TBLTotal.DefaultView.RowFilter = filter
                    If TBLTotal.DefaultView.Count > 0 Then
                        row = TBLTotal.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Total = row.Debet
                        Else
                            .Total = row.Credit
                        End If
                    End If

                    'hitung variance
                    .Variance = 0
                    .Variance_Persen = 0
                    '.Total = .Saldo_Period_I + .Saldo_Period_II + .Saldo_Period_III + .Saldo_Period_IV + .Saldo_Period_V + .Saldo_Period_VI
                End With

                DS.TblCetak.Rows.Add(rowCetak)
            Next

            'Buang semua akun yang nilainya 0 dan akun yang tidak aktif, tergantung pilihan user
            If Me.chkHideZero.Checked = True Then
                i = 0
                While i < DS.TblCetak.Rows.Count
                    rowCetak = DS.TblCetak.Rows(i)
                    Me.SetSaldoPeriodYgMasihNull(rowCetak)
                    If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                    rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 And _
                    rowCetak.Total = 0 And rowCetak.Level_Perkiraan <> 1 Then
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
                            Me.SetSaldoPeriodYgMasihNull(rowCetak)
                            If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                            rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 And _
                            rowCetak.Total = 0 Then
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

            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value)
            vDtpTo_1 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value)
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)
            vDtpTo_2 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_II.Value)

            Me.Report_6Months.SetDataSource(DS)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_TanggalDari_1.ParameterFieldName, StartDate)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_TanggalDari_2.ParameterFieldName, vDtpTo_1)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_Tipe.ParameterFieldName, tipe)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_bulan1.ParameterFieldName, Bulan1)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_bulan2.ParameterFieldName, Bulan2)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_bulan3.ParameterFieldName, Bulan3)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_bulan4.ParameterFieldName, Bulan4)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_bulan5.ParameterFieldName, Bulan5)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_bulan6.ParameterFieldName, Bulan6)
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_Tipe.ParameterFieldName, tipe)

            Me.dtpFrom_I.Value = StartDate

            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.Report_6Months.SetParameterValue(Report_6Months.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.Report_6Months
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
            TBL1.Dispose()
            TBL3.Dispose()
            TBL4.Dispose()
            TBL5.Dispose()
            TBL6.Dispose()
            TBLTotal.Dispose()
        End Try
    End Sub

    Public Sub Cetak_P_L_QuarterYearly(ByVal Adjusted As Boolean)
        Dim TBL1 As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL3 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL4 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL5 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL6 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL7 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TBL8 As New DseLaporanTrialBalance.Trial_BalanceDataTable

        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpTo_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim vDtpTo_2 As Date = Now
        Dim tipe As Double = 0
        Dim StartDate As Date = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpFrom_I.Value.Date)

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Prev Year
            Me.IsiTabel(Adjusted, ObjCD.GetFisrtDateOfTheYear(StartDate - TimeSpan.FromDays(10)), _
            ObjCD.GetLastDateOfTheYear(StartDate - TimeSpan.FromDays(10)), TBL1)
            '2. 1st QTR
            Me.IsiTabel(Adjusted, StartDate, Me.ObjCD.GetLasDateOfTheMonth(StartDate + TimeSpan.FromDays(2 * 32)), TBL2, False)
            '3 2nd QTR
            Me.IsiTabel(Adjusted, Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(3 * 32)), _
            Me.ObjCD.GetLasDateOfTheMonth(StartDate + TimeSpan.FromDays(5 * 31 + 3)), TBL3)
            '4. 1st semester
            Me.IsiTabel(Adjusted, StartDate, Me.ObjCD.GetLasDateOfTheMonth(StartDate + TimeSpan.FromDays(5 * 31 + 3)), TBL4)
            '5. 3rd QTR
            Me.IsiTabel(Adjusted, Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(6 * 31 + 3)), _
             Me.ObjCD.GetLasDateOfTheMonth(StartDate + TimeSpan.FromDays(8 * 31 + 3)), TBL5)
            '6. 4th QTR
            Me.IsiTabel(Adjusted, Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(9 * 31 + 3)), _
            Me.ObjCD.GetLasDateOfTheMonth(StartDate + TimeSpan.FromDays(11 * 31 + 3)), TBL6)
            '7 2nd semester
            Me.IsiTabel(Adjusted, Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(6 * 31 + 3)), _
            Me.ObjCD.GetLasDateOfTheMonth(StartDate + TimeSpan.FromDays(11 * 31 + 3)), TBL7)
            '8 Full Year
            Me.IsiTabel(Adjusted, StartDate, Me.ObjCD.GetLasDateOfTheMonth(StartDate + TimeSpan.FromDays(11 * 31 + 3)), TBL8)

            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL1)
            End If

            '2. Tampilkan pada cetakan
            'generate report
            DS.TblCetak.Rows.Clear()
            For i = 0 To TBL1.Rows.Count - 1
                row = TBL1.Rows(i)
                rowCetak = DS.TblCetak.NewTblCetakRow()

                With rowCetak
                    'pindahkan data periode I
                    .Akun = row.Akun
                    .Header = row.Header
                    If row.Saldo_Normal = 0 Then
                        .Saldo_Period_I = row.Debet
                        .Saldo_Year_Before = row.Debet
                    Else
                        .Saldo_Period_I = row.Credit
                        .Saldo_Year_Before = row.Credit
                    End If
                    .Total = row.Total
                    .Aktif = row.Aktif
                    .Saldo_Normal = row.Saldo_Normal
                    .Level_Perkiraan = row.Level_Perkiraan
                    .Bold = row.Bold
                    .Urutan = row.Tipe_Perkiraan

                    'pindahkan data periode 1st QTR
                    filter = "[Primary Akun]= " + row.Primary_Akun.ToString
                    TBL2.DefaultView.RowFilter = filter
                    If TBL2.DefaultView.Count > 0 Then
                        row = TBL2.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Q1 = row.Debet
                        Else
                            .Saldo_Q1 = row.Credit
                        End If
                    End If

                    'pindahkan data periode 2nd QTR
                    TBL3.DefaultView.RowFilter = filter
                    If TBL3.DefaultView.Count > 0 Then
                        row = TBL3.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Q2 = row.Debet
                        Else
                            .Saldo_Q2 = row.Credit
                        End If
                    End If

                    'pindahkan data periode 1st semester
                    TBL4.DefaultView.RowFilter = filter
                    If TBL4.DefaultView.Count > 0 Then
                        row = TBL4.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Semester1 = row.Debet
                        Else
                            .Saldo_Semester1 = row.Credit
                        End If
                    End If

                    'pindahkan data periode 3rd QTR
                    TBL5.DefaultView.RowFilter = filter
                    If TBL5.DefaultView.Count > 0 Then
                        row = TBL5.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Q3 = row.Debet
                        Else
                            .Saldo_Q3 = row.Credit
                        End If
                    End If

                    'pindahkan data periode 4th QTR
                    TBL6.DefaultView.RowFilter = filter
                    If TBL6.DefaultView.Count > 0 Then
                        row = TBL6.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Q4 = row.Debet
                        Else
                            .Saldo_Q4 = row.Credit
                        End If
                    End If

                    'pindahkan data periode 2nd semester
                    TBL7.DefaultView.RowFilter = filter
                    If TBL7.DefaultView.Count > 0 Then
                        row = TBL7.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Semester2 = row.Debet
                        Else
                            .Saldo_Semester2 = row.Credit
                        End If
                    End If

                    'pindahkan data periode full year
                    TBL8.DefaultView.RowFilter = filter
                    If TBL8.DefaultView.Count > 0 Then
                        row = TBL8.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Currenct_Year = row.Debet
                        Else
                            .Saldo_Currenct_Year = row.Credit
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
                    If rowCetak.Saldo_Currenct_Year = 0 And rowCetak.Saldo_Q1 = 0 And rowCetak.Saldo_Q2 = 0 And _
                    rowCetak.Saldo_Q3 = 0 And rowCetak.Saldo_Q4 = 0 And rowCetak.Saldo_Semester1 = 0 And _
                    rowCetak.Saldo_Semester2 = 0 And rowCetak.Saldo_Year_Before = 0 And rowCetak.Level_Perkiraan <> 1 Then
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
                            If rowCetak.Saldo_Currenct_Year = 0 And rowCetak.Saldo_Q1 = 0 And rowCetak.Saldo_Q2 = 0 And _
                            rowCetak.Saldo_Q3 = 0 And rowCetak.Saldo_Q4 = 0 And rowCetak.Saldo_Semester1 = 0 And _
                            rowCetak.Saldo_Semester2 = 0 And rowCetak.Saldo_Year_Before = 0 Then
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

            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value)
            vDtpTo_1 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value)
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)
            vDtpTo_2 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_II.Value)

            Me.Report_QTR.SetDataSource(DS)
            Me.Report_QTR.SetParameterValue(Report_QTR.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.Report_QTR.SetParameterValue(Report_QTR.Parameter_TanggalDari_1.ParameterFieldName, StartDate)
            Me.Report_QTR.SetParameterValue(Report_QTR.Parameter_TanggalDari_2.ParameterFieldName, vDtpTo_1)
            Me.Report_QTR.SetParameterValue(Report_QTR.Parameter_Tipe.ParameterFieldName, tipe)
            Me.Report_QTR.SetParameterValue(Report_QTR.Parameter_Year_Before.ParameterFieldName, Format(StartDate - TimeSpan.FromDays(10), "yyyy") + " Total")
            Me.Report_QTR.SetParameterValue(Report_QTR.Parameter_Current_year.ParameterFieldName, Format(StartDate, "yyyy") + " Total")
            Me.Report_QTR.SetParameterValue(Report_QTR.Parameter_Tipe.ParameterFieldName, tipe)

            Me.dtpFrom_I.Value = StartDate

            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.Report_QTR.SetParameterValue(Report_QTR.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.Report_QTR
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
            TBL1.Dispose()
            TBL3.Dispose()
            TBL4.Dispose()
            TBL5.Dispose()
            TBL6.Dispose()
        End Try
    End Sub

    Private Sub SetSaldoPeriodYgMasihNull(ByVal rowCetak As DseLaporanTrialBalance.TblCetakRow)
        Try
            If rowCetak.IsSaldo_Period_INull = True Then
                rowCetak.Saldo_Period_I = 0
            End If
            If rowCetak.IsSaldo_Period_IINull = True Then
                rowCetak.Saldo_Period_II = 0
            End If
            If rowCetak.IsSaldo_Period_IIINull = True Then
                rowCetak.Saldo_Period_III = 0
            End If
            If rowCetak.IsSaldo_Period_IVNull = True Then
                rowCetak.Saldo_Period_IV = 0
            End If
            If rowCetak.IsSaldo_Period_VNull = True Then
                rowCetak.Saldo_Period_V = 0
            End If
            If rowCetak.IsSaldo_Period_VINull = True Then
                rowCetak.Saldo_Period_VI = 0
            End If
            If rowCetak.IsSaldo_Period_VIINull = True Then
                rowCetak.Saldo_Period_VII = 0
            End If
            If rowCetak.IsSaldo_Period_VIIINull = True Then
                rowCetak.Saldo_Period_VIII = 0
            End If
            If rowCetak.IsSaldo_Period_IXNull = True Then
                rowCetak.Saldo_Period_IX = 0
            End If
            If rowCetak.IsSaldo_Period_XNull = True Then
                rowCetak.Saldo_Period_X = 0
            End If
            If rowCetak.IsSaldo_Period_XINull = True Then
                rowCetak.Saldo_Period_XI = 0
            End If
            If rowCetak.IsSaldo_Period_XIINull = True Then
                rowCetak.Saldo_Period_XII = 0
            End If
            If rowCetak.IsTotalNull = True Then
                rowCetak.Total = 0
            End If
            If rowCetak.IsTotal2Null = True Then
                rowCetak.Total2 = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Public Sub Cetak_P_L_12MonthsPeriod(ByVal Adjusted As Boolean)
        Dim TBL1 As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
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
        Dim TBLTotal As New DseLaporanTrialBalance.Trial_BalanceDataTable

        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpTo_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim vDtpTo_2 As Date = Now
        Dim tipe As Double = 0
        Dim StartDate As Date = Me.dtpFrom_I.Value.Date
        'Dim StartDate As Date = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpFrom_I.Value.Date)
        Dim Bulan1, Bulan2, Bulan3, Bulan4, Bulan5, Bulan6, Bulan7, Bulan8, Bulan9, Bulan10, Bulan11, Bulan12 As Date

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Me.dtpFrom_I.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpFrom_I.Value)
            Bulan1 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL1, False)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan2 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL2)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan3 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL3)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan4 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL4)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan5 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL5)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan6 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL6)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan7 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL7)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan8 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL8)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan9 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL9)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan10 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL10)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan11 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL11)
            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value + TimeSpan.FromDays(32))
            Bulan12 = Me.dtpFrom_I.Value.Date
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL12)

            Me.IsiTabel(Adjusted, Bulan1, Me.ObjCD.GetLasDateOfTheMonth(Bulan12), TBLTotal)


            'jika ada kode akun maka buang saja
            If Me.chkShowAccountCode.Checked = False Then
                Me.BuangKodeAkun(TBL1)
            End If

            '2. Tampilkan pada cetakan
            'generate report
            DS.TblCetak.Rows.Clear()
            For i = 0 To TBL1.Rows.Count - 1
                row = TBL1.Rows(i)
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

                    'pindahkan data periode ix
                    TBL9.DefaultView.RowFilter = filter
                    If TBL9.DefaultView.Count > 0 Then
                        row = TBL9.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_IX = row.Debet
                        Else
                            .Saldo_Period_IX = row.Credit
                        End If
                    End If

                    'pindahkan data periode x
                    TBL10.DefaultView.RowFilter = filter
                    If TBL10.DefaultView.Count > 0 Then
                        row = TBL10.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_X = row.Debet
                        Else
                            .Saldo_Period_X = row.Credit
                        End If
                    End If

                    'pindahkan data periode xi
                    TBL11.DefaultView.RowFilter = filter
                    If TBL11.DefaultView.Count > 0 Then
                        row = TBL11.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_XI = row.Debet
                        Else
                            .Saldo_Period_XI = row.Credit
                        End If
                    End If

                    'pindahkan data periode xii
                    TBL12.DefaultView.RowFilter = filter
                    If TBL12.DefaultView.Count > 0 Then
                        row = TBL12.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Saldo_Period_XII = row.Debet
                        Else
                            .Saldo_Period_XII = row.Credit
                        End If
                    End If

                    'Hitung total periode I - XII
                    TBLTotal.DefaultView.RowFilter = filter
                    If TBLTotal.DefaultView.Count > 0 Then
                        row = TBLTotal.DefaultView(0).Row
                        If row.Saldo_Normal = 0 Then
                            .Total = row.Debet
                        Else
                            .Total = row.Credit
                        End If
                    End If

                    'hitung variance
                    .Variance = 0
                    .Variance_Persen = 0
                    '.Total = .Saldo_Period_I + .Saldo_Period_II + .Saldo_Period_III + .Saldo_Period_IV + .Saldo_Period_V + _
                    '         .Saldo_Period_VI + .Saldo_Period_VII + .Saldo_Period_VIII + .Saldo_Period_IX + .Saldo_Period_X + _
                    '         .Saldo_Period_XI + .Saldo_Period_XII
                End With

                DS.TblCetak.Rows.Add(rowCetak)
            Next

            'Buang semua akun yang nilainya 0 dan akun yang tidak aktif, tergantung pilihan user
            If Me.chkHideZero.Checked = True Then
                i = 0
                While i < DS.TblCetak.Rows.Count
                    rowCetak = DS.TblCetak.Rows(i)
                    Me.SetSaldoPeriodYgMasihNull(rowCetak)
                    If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                    rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 And rowCetak.Level_Perkiraan <> 1 And _
                    rowCetak.Saldo_Period_VII = 0 And rowCetak.Saldo_Period_VIII = 0 And rowCetak.Saldo_Period_IX = 0 And _
                    rowCetak.Saldo_Period_X = 0 And rowCetak.Saldo_Period_XI = 0 And rowCetak.Saldo_Period_XII = 0 And rowCetak.Total = 0 Then
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
                            Me.SetSaldoPeriodYgMasihNull(rowCetak)
                            If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Saldo_Period_III = 0 And _
                            rowCetak.Saldo_Period_IV = 0 And rowCetak.Saldo_Period_V = 0 And rowCetak.Saldo_Period_VI = 0 And _
                            rowCetak.Saldo_Period_VII = 0 And rowCetak.Saldo_Period_VIII = 0 And rowCetak.Saldo_Period_IX = 0 And _
                            rowCetak.Saldo_Period_X = 0 And rowCetak.Saldo_Period_XI = 0 And rowCetak.Saldo_Period_XII = 0 And rowCetak.Total = 0 Then
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

            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value)
            vDtpTo_1 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value)
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)
            vDtpTo_2 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_II.Value)
            Me.dtpFrom_I.Value = StartDate

            Me.Report_12Months.SetDataSource(DS)
            Me.Report_12Months.SetParameterValue(Report_12Months.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.Report_12Months.SetParameterValue(Report_12Months.Parameter_TanggalDari_1.ParameterFieldName, StartDate)
            Me.Report_12Months.SetParameterValue(Report_12Months.Parameter_TanggalDari_2.ParameterFieldName, vDtpTo_1)
            Me.Report_12Months.SetParameterValue(Report_12Months.Parameter_Tipe.ParameterFieldName, tipe)
            Me.Report_12Months.SetParameterValue(Report_12Months.Parameter_Tipe.ParameterFieldName, tipe)

            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.Report_12Months.SetParameterValue(Report_12Months.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.Report_12Months
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL1.Dispose()
            TBL2.Dispose()
            TBL3.Dispose()
            TBL4.Dispose()
            TBL5.Dispose()
            TBL6.Dispose()
            TBL7.Dispose()
            TBL8.Dispose()
            TBL9.Dispose()
            TBL10.Dispose()
            TBL11.Dispose()
            TBL12.Dispose()
            TBLTotal.Dispose()
        End Try
    End Sub
End Class
