Option Explicit On
Option Strict Off

Partial Public Class FrmLaporan_ProfitNLoss
    Private Sub Cetak_P_L_ChartMonthly(ByVal Adjusted As Boolean, ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpTo_1 As Date = Now

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Me.IsiTabel(Adjusted, vStartDate, Me.ObjCD.GetLasDateOfTheMonth(vEndDate), TBL)

            '2. Tampilkan pada cetakan
            'generate report
            DS.TblCetak.Rows.Clear()
            For i = 0 To TBL.Rows.Count - 1
                row = TBL.Rows(i)
                rowCetak = DS.TblCetak.NewTblCetakRow()

                With rowCetak
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
                    .Tipe_Perkiraan = row.Tipe_Perkiraan
                End With

                DS.TblCetak.Rows.Add(rowCetak)
            Next

            'Buang semua akun yang nilainya 0 dan akun yang tidak aktif, tergantung pilihan user
            If Me.chkHideZero.Checked = True Then
                i = 0
                While i < DS.TblCetak.Rows.Count
                    rowCetak = DS.TblCetak.Rows(i)
                    If rowCetak.Saldo_Period_I = 0 And rowCetak.Level_Perkiraan <> 1 Then
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
                            If rowCetak.Saldo_Period_I = 0 Then
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

            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(vStartDate)
            vDtpTo_1 = Me.ObjCD.GetLasDateOfTheMonth(vEndDate)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub Cetak_P_L_ChartMonthly(ByVal Adjusted As Boolean)
        Dim StartDate, EndDate As Date
        Dim TBL As DseLaporanTrialBalance.IS_ChartDataTable = Me.DS.IS_Chart
        Dim row As DseLaporanTrialBalance.IS_ChartRow
        Dim Income, COGS, Expense, Profit As Decimal
        Dim DummyEndDate As Date
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Dim TipePerkiraan_Income As ClsMasterAkun.EnumTipeAkun = ClsMasterAkun.EnumTipeAkun.Income
        Dim TipePerkiraan_OtherIncome As ClsMasterAkun.EnumTipeAkun = ClsMasterAkun.EnumTipeAkun.Other_Income
        Dim TipePerkiraan_COGS As ClsMasterAkun.EnumTipeAkun = ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales
        Dim TipePerkiraan_Expense As ClsMasterAkun.EnumTipeAkun = ClsMasterAkun.EnumTipeAkun.Expense
        Dim TipePerkiraan_OtherExpense As ClsMasterAkun.EnumTipeAkun = ClsMasterAkun.EnumTipeAkun.Other_Expense

        Try
            TBL.Rows.Clear()

            StartDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value.Date)
            EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value.Date)

            While StartDate < EndDate
                Income = 0
                COGS = 0
                Expense = 0
                Profit = 0

                'buatkan bulanan
                DummyEndDate = Me.ObjCD.GetLasDateOfTheMonth(StartDate)
                Me.Cetak_P_L_ChartMonthly(Adjusted, StartDate, DummyEndDate)

                'get data untuk bulan itu
                For Each rowCetak In Me.DS.TblCetak.Rows
                    If rowCetak.Header = 0 Then
                        If rowCetak.Tipe_Perkiraan = CInt(TipePerkiraan_Income) Or _
                        rowCetak.Tipe_Perkiraan = CInt(TipePerkiraan_OtherIncome) Then
                            Income += rowCetak.Saldo_Period_I
                        End If
                        If rowCetak.Tipe_Perkiraan = CInt(TipePerkiraan_Expense) Or _
                        rowCetak.Tipe_Perkiraan = CInt(TipePerkiraan_OtherExpense) Then
                            Expense += rowCetak.Saldo_Period_I
                        End If
                        If rowCetak.Tipe_Perkiraan = CInt(TipePerkiraan_COGS) Then
                            COGS += rowCetak.Saldo_Period_I
                        End If
                    End If
                Next

                'masukkan data ke dalam tabel
                Profit = Income - COGS - Expense

                If Me.chkShowProfit.Checked = True Then
                    row = TBL.NewIS_ChartRow()

                    row._Date = DummyEndDate
                    row.Tipe_Data = "Profit"
                    row.Value = Profit

                    TBL.Rows.Add(row)
                End If
                If Me.chkShowCOGS.Checked = True Then
                    row = TBL.NewIS_ChartRow()

                    row._Date = DummyEndDate
                    row.Tipe_Data = "COGS"
                    row.Value = COGS

                    TBL.Rows.Add(row)
                End If
                If Me.chkShowIncome.Checked = True Then
                    row = TBL.NewIS_ChartRow()

                    row._Date = DummyEndDate
                    row.Tipe_Data = "Income"
                    row.Value = Income

                    TBL.Rows.Add(row)
                End If
                If Me.chkShowExpense.Checked = True Then
                    row = TBL.NewIS_ChartRow()

                    row._Date = DummyEndDate
                    row.Tipe_Data = "Expense"
                    row.Value = Expense

                    TBL.Rows.Add(row)
                End If

                'majukan bulannya
                StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            End While

            'tampilkan laporan
            Me.reportChart.SetDataSource(DS)
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.reportChart.SetParameterValue(reportChart.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.reportChart.SetParameterValue(reportChart.Parameter_TanggalDari.ParameterFieldName, Me.dtpFrom_I.Value.Date)
            Me.reportChart.SetParameterValue(reportChart.Parameter_TanggalSampai.ParameterFieldName, Me.dtpTo_I.Value.Date)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.reportChart.SetParameterValue(reportChart.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.reportChart
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
