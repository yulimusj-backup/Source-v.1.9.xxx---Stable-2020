Option Explicit On
Option Strict Off

Public Class FrmLaporanBalanceSheet
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen

    Dim ObjCD As New Core_Module.ClsCekData
    Dim DTA_Beginning_Balance As New DseLaporanTrialBalanceTableAdapters.tbl_beginning_balanceTableAdapter
    Dim TBL_Beginning_Balance As New DseLaporanTrialBalance.tbl_beginning_balanceDataTable
    Dim DTA_TrialBalance As New DseLaporanTrialBalanceTableAdapters.Trial_BalanceTableAdapter

    Dim report As New CrpLaporanBalanceSheetNew
    Dim report_comparison As New CrpLaporanBalanceSheet_ComparisonNew
    Dim report_6Months As New CrpLaporanBalanceSheet_6MonthsNew
    Dim report_6Months_w_Budget As New CrpLaporanBalanceSheet_6Months_wBudgetNew
    Dim report_12Months As New CrpLaporanBalanceSheet_12Months
    Dim report_12Months_w_Budget As New CrpLaporanBalanceSheet_12Months_wBudget

    Dim DS As New DseLaporanTrialBalance
    Dim ObjBudget As New ClsMasterBudget
    Dim ObjJurnal As New ClsTransaksiJournal
    Dim ObjHakAkses As New Core_Module.ClsHakAkses

    Dim Def_Prec As Integer = 0
    Dim ObjCUr As New Core_Module.clsMasterCurrency
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Negatif As Integer = 0
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub FrmLaporanBalanceSheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Def_Prec = Me.ObjCUr.GetDecimal_DefaultCurrency
            Me.cboFormatLaporan.SelectedIndex = 0

            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Today)
            Me.dtpFrom_II.Value = Me.ObjCD.GetFirstDateofTheMonth(Today)

            'jika tidak ada hak akses atas dept, maka kolom dept ditutup
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDepartement") = False Then
            Me.LblDept.Visible = False
            Me.TxtDeptCode.Visible = False
            Me.TxtDeptCode.TabStop = False
            Me.btnFindDept.Visible = False
            'End If

            'jika tidak ada hak akses atas job, maka kolom job ditutup
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterJob") = False Then
            Me.LblJob.Visible = False
            Me.TxtJobCode.Visible = False
            Me.TxtJobCode.TabStop = False
            Me.btnFindJob.Visible = False
            'End If
            Me.chkShowLvl2.Enabled = True
            Me.chkShowLvl3.Enabled = True
            Me.chkShowLvl4.Enabled = True
            Me.chkShowLvl5.Enabled = True
            Me.chkShowLvl6.Enabled = True
            Me.chkShowLvl7.Enabled = True
            Me.chkShowLvl8.Enabled = True
            
            Me.chkUseTotal.Text = "Use Indentation"
            Me.MinimizeBox = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormatLaporan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormatLaporan.SelectedIndexChanged
        Try
            Me.dtpFrom_I.CustomFormat = "MMM yyyy"
            LblPeriod.Text = "Period I"

            If Me.cboFormatLaporan.SelectedIndex = 1 Then
                Me.dtpFrom_II.Enabled = True
                Me.LblPeriodII.Enabled = True
                Me.dtpFrom_II.Visible = True
                Me.LblPeriodII.Visible = True
            Else
                Me.dtpFrom_II.Enabled = False
                Me.LblPeriodII.Enabled = False
                Me.dtpFrom_II.Visible = False
                Me.LblPeriodII.Visible = False
            End If

            'format tahun
            If Me.cboFormatLaporan.SelectedIndex = 5 Or Me.cboFormatLaporan.SelectedIndex = 6 Then
                Me.dtpFrom_I.CustomFormat = "yyyy"
                LblPeriod.Text = "Year"
            ElseIf Me.cboFormatLaporan.SelectedIndex = 4 Or Me.cboFormatLaporan.SelectedIndex = 3 Then
                LblPeriod.Text = "Start Period"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim PrimaryAkunCYE As String = ""
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom_I.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom_I.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpFrom_II.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom_II.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            'tidak bisa buat laporan balance sheet, jika retained earning accoutnya belum diberesin
            'jika akun current year earning tidak ada maka tidak bisa lanjut
            PrimaryAkunCYE = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_CurrenctYearEarning)
            If PrimaryAkunCYE = "" Or PrimaryAkunCYE = "-1" Then
                MsgBox("Please specify current year account in global setting data.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.cboFormatLaporan.SelectedIndex = 0 Then
                Me.CetakBalanceSheet()
            ElseIf Me.cboFormatLaporan.SelectedIndex = 1 Then
                Me.CetakBalanceSheet_Comparison_Period()
            ElseIf Me.cboFormatLaporan.SelectedIndex = 2 Then
                Me.CetakBalanceSheet_Comparison_Budget()
            ElseIf Me.cboFormatLaporan.SelectedIndex = 3 Then
                Me.CetakBalanceSheet_6Months()
            ElseIf Me.cboFormatLaporan.SelectedIndex = 4 Then
                Me.CetakBalanceSheet_6Months_WithBudget()
            ElseIf Me.cboFormatLaporan.SelectedIndex = 5 Then
                Me.CetakBalanceSheet_12Months()
            ElseIf Me.cboFormatLaporan.SelectedIndex = 6 Then
                Me.CetakBalanceSheet_12MonthsWithBudget()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTabel(ByVal vEndDate As Date, _
    ByVal TBL As DseLaporanTrialBalance.Trial_BalanceDataTable)
        vEndDate = Me.ObjCD.GetLasDateOfTheMonth(vEndDate)
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowBeginningBalance As DseLaporanTrialBalance.tbl_beginning_balanceRow
        Dim filter As String = ""
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim CYE As Decimal = 0
        Dim PrimaryAkunCYE As Decimal = 0
        Dim Dummy As Integer = 0
        Dim Space As String = "   "
        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Cari semua akun yang ada dengan begining balance pada awal period (kalo start period Jul 2010 --> begining balance
            '   adalah saldo buku besar pada 30 jun 2010

            'If Me.TxtPrimaryDept.Text <> "-1" And Me.TxtPrimaryJob.Text <> "-1" Then
            '    Me.DTA_Beginning_Balance.Fill(Me.TBL_Beginning_Balance, vEndDate, Me.TxtPrimaryDept.Text, Me.TxtPrimaryJob.Text)
            'ElseIf Me.TxtPrimaryDept.Text <> "-1" And Me.TxtPrimaryJob.Text = "-1" Then
            '    Me.DTA_Beginning_Balance.FillBy_PrimaryDept(Me.TBL_Beginning_Balance, vEndDate, Me.TxtPrimaryDept.Text)
            'ElseIf Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text <> "-1" Then
            '    Me.DTA_Beginning_Balance.FillBy_PrimaryJob(Me.TBL_Beginning_Balance, vEndDate, Me.TxtPrimaryJob.Text)
            'Else
            Me.DTA_Beginning_Balance.FillBy_All(Me.TBL_Beginning_Balance, vEndDate)
            'End If

            '2. Pindahkan ke tabel trial balance
            Me.DTA_TrialBalance.Fill(TBL)

            '2.a. Tambahkan akun current year earning, jika pada bulan tersebut tidak ada jurnal closing
            If Me.ObjJurnal.GetJml(ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, _
            Me.ObjCD.GetFisrtDateOfTheYear(vEndDate), Me.ObjCD.GetLasDateOfTheMonth(vEndDate)) = 0 Then
                CYE = Me.ObjJurnal.GetCurrentYearEarning1(vEndDate, Me.TxtPrimaryDept.Text, Me.TxtPrimaryJob.Text)
                PrimaryAkunCYE = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_CurrenctYearEarning)

                For Each row In TBL.Rows
                    If row.Primary_Akun = PrimaryAkunCYE Then
                        If CYE > 0 Then
                            row.Credit += Math.Abs(CYE)
                        Else
                            row.Debet += Math.Abs(CYE)
                        End If
                    End If
                Next
            End If

            '2.b. Buang akun selain asset, liabilities, equity
            j = 0
            While j < TBL.Rows.Count
                row = TBL.Rows(j)
                If row.Tipe_Perkiraan > 2 Then
                    TBL.Rows.Remove(row)
                Else
                    j += 1
                End If
            End While

            For Each row In TBL.Rows
                filter = "Primary =" + row.Primary_Akun.ToString
                Me.TBL_Beginning_Balance.DefaultView.RowFilter = filter
                If Me.TBL_Beginning_Balance.DefaultView.Count > 0 Then
                    rowBeginningBalance = Me.TBL_Beginning_Balance.DefaultView(0).Row
                    row.Debet += rowBeginningBalance.Debet
                    row.Credit += rowBeginningBalance.Kredit

                    If row.Saldo_Normal = 0 Then
                        row.Debet = row.Beginning_Balance + row.Debet - row.Credit
                        row.Credit = 0
                    Else
                        row.Credit = row.Beginning_Balance + row.Credit - row.Debet
                        row.Debet = 0
                    End If
                End If

                If Me.chkUseTotal.Checked = True Then
                    Dummy = 0
                    While Dummy < row.Level_Perkiraan
                        row.Account_Code = Space + row.Account_Code
                        'row.Description = Space + row.Description
                        Dummy += 1
                    End While
                End If
            Next

            '3. Hitung sum untuk setiap data header 
            Me.UpdateSaldoHeader(TBL)

            '4. Use total
            'If Me.chkUseTotal.Checked = True Then
            '    i = 0
            '    While i < TBL.Rows.Count
            '        row = TBL.Rows(i)
            '        If row.Header = 1 Then
            '            If row.IsTotalNull = True Then
            '                Me.InsertTotal(i, row.Level_Perkiraan, TBL)
            '                row.Total = 0
            '            End If
            '        Else
            '            row.Total = 0
            '        End If
            '        i += 1
            '    End While
            'Else
            For Each row In TBL.Rows
                row.Total = -1
            Next
            'End If

            '5. Hide data 
            If Me.chkShowLvl2.Checked = False Then
                i = 2
            ElseIf Me.chkShowLvl3.Checked = False Then
                i = 3
            ElseIf Me.chkShowLvl4.Checked = False Then
                i = 4
            ElseIf Me.chkShowLvl5.Checked = False Then
                i = 5
            ElseIf Me.chkShowLvl6.Checked = False Then
                i = 6
            ElseIf Me.chkShowLvl7.Checked = False Then
                i = 7
            ElseIf Me.chkShowLvl8.Checked = False Then
                i = 8
            Else
                i = 9
            End If
            j = 0
            While j < TBL.Rows.Count
                row = TBL.Rows(j)
                If row.Level_Perkiraan >= i Then
                    TBL.Rows.Remove(row)
                Else
                    j += 1
                End If
            End While

            '8. tandai yang bold - hanya digunakan pada laporan profit n loss
            For Each row In TBL.Rows
                If row.IsBoldNull = True Then
                    row.Bold = row.Header
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InsertTotal(ByVal vRowIndex As Double, ByVal vLevel As Integer, _
    ByVal vTBL As DseLaporanTrialBalance.Trial_BalanceDataTable)
        Dim TargetIndex, CurrentIndex As Integer
        Dim row, rowTotal As DseLaporanTrialBalance.Trial_BalanceRow
        Dim RetVal As Double = 0
        Dim space As String = "     "

        Try
            CurrentIndex = vRowIndex + 1
            While CurrentIndex <= vTBL.Rows.Count - 1
                row = vTBL.Rows(CurrentIndex)
                If row.Level_Perkiraan <= vLevel Then
                    TargetIndex = CurrentIndex
                    Exit While
                End If

                CurrentIndex += 1
            End While
            If CurrentIndex = vTBL.Rows.Count Then
                TargetIndex = CurrentIndex
            End If

            rowTotal = vTBL.NewTrial_BalanceRow
            row = vTBL.Rows(vRowIndex)

            With rowTotal
                .Aktif = row.Aktif
                If row.Level_Perkiraan = 1 Then
                    .Akun = "Total " + row.Description.Trim
                ElseIf row.Level_Perkiraan = 2 Then
                    .Akun = space + "Total " + row.Description.Trim
                ElseIf row.Level_Perkiraan = 3 Then
                    .Akun = space + space + "Total " + row.Description.Trim
                ElseIf row.Level_Perkiraan = 4 Then
                    .Akun = space + space + space + "Total " + row.Description.Trim
                End If
                .Beginning_Balance = 0
                If row.Saldo_Normal = 0 Then
                    .Debet = row.Debet
                    row.Debet = 0
                    .Credit = 0
                Else
                    .Debet = 0
                    .Credit = row.Credit
                    row.Credit = 0
                End If
                .Account_Code = row.Account_Code
                .Saldo_Normal = row.Saldo_Normal
                .Description = row.Description
                .Header = row.Header
                .Level_Perkiraan = row.Level_Perkiraan
                .Primary_Akun = Me.GetPrimaryAkunMaks(vTBL) + 1
                .Saldo = row.Saldo
                .Total = 1
                .Tipe_Perkiraan = row.Tipe_Perkiraan
                .Urutan = .Tipe_Perkiraan
            End With

            If TargetIndex = vTBL.Rows.Count Then
                vTBL.Rows.Add(rowTotal)
            Else
                vTBL.Rows.InsertAt(rowTotal, TargetIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryAkunMaks(ByVal vTBL As DseLaporanTrialBalance.Trial_BalanceDataTable) As Double
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim Primary As Double = 0

        Try
            For Each row In vTBL.Rows
                If row.Primary_Akun > Primary Then
                    Primary = row.Primary_Akun
                End If
            Next

            Return Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub UpdateSaldoHeader(ByVal vTBL As DseLaporanTrialBalance.Trial_BalanceDataTable)
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim i As Integer = 0
        Try
            For Each row In vTBL.Rows
                If row.Header = 1 Then
                    If row.Saldo_Normal = 0 Then
                        row.Debet = Me.GetSaldoHeader(i, row.Saldo_Normal, row.Level_Perkiraan, vTBL)
                    Else
                        row.Credit = Me.GetSaldoHeader(i, row.Saldo_Normal, row.Level_Perkiraan, vTBL)
                    End If
                End If

                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetSaldoHeader(ByVal vRowIndex As Double, ByVal vDebetKredit As Integer, ByVal vLevel As Integer, _
    ByVal vTBL As DseLaporanTrialBalance.Trial_BalanceDataTable) As Double
        Dim TargetIndex, CurrentIndex As Integer
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim RetVal As Double = 0

        Try
            CurrentIndex = vRowIndex + 1
            While CurrentIndex <= vTBL.Rows.Count - 1
                row = vTBL.Rows(CurrentIndex)
                If row.Level_Perkiraan <= vLevel Then
                    TargetIndex = CurrentIndex - 1
                    Exit While
                ElseIf CurrentIndex = vTBL.Rows.Count - 1 Then
                    TargetIndex = CurrentIndex
                    Exit While
                End If
                CurrentIndex += 1
            End While

            For i As Integer = vRowIndex + 1 To TargetIndex
                row = vTBL.Rows(i)
                If row.Header = 1 And row.Level_Perkiraan = vLevel + 1 Then
                    RetVal += Me.GetSaldoHeader(i, vDebetKredit, row.Level_Perkiraan, vTBL)
                ElseIf row.Level_Perkiraan = vLevel + 1 Then
                    If vDebetKredit = 0 Then
                        RetVal += CDec(row.Debet)
                    Else
                        RetVal += CDec(row.Credit)
                    End If
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub CetakBalanceSheet()
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim vDtpFrom_1 As Date = Now
        Dim TotalAset, TotalHutang, TotalModal As Decimal

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Me.IsiTabel(Me.dtpFrom_I.Value, TBL)
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
                End With

                DS.TblCetak.Rows.Add(rowCetak)

                If row.Primary_Akun = 1 Then
                    TotalAset = row.Debet
                ElseIf row.Primary_Akun = 2 Then
                    TotalHutang = row.Credit
                ElseIf row.Primary_Akun = 3 Then
                    TotalModal = row.Credit
                End If
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

            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value)

            Me.report.SetDataSource(DS)
            Me.report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.report.SetParameterValue(report.Parameter_Tanggal_Dari.ParameterFieldName, vDtpFrom_1)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report.SetParameterValue(report.Parameter_Negatif.ParameterFieldName, Me.Negatif)
            If Me.chkShowAccountCode.Checked = True Then
                Me.report.SetParameterValue(report.Parameter_ShowNoAccount.ParameterFieldName, 1)
            Else
                Me.report.SetParameterValue(report.Parameter_ShowNoAccount.ParameterFieldName, 0)
            End If
            Me.crv1.ReportSource = Me.report
            Me.crv1.Refresh()
            Me.crv1.Show()

            If Math.Round(TotalAset, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data." + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub CetakBalanceSheet_Comparison_Period()
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim Tipe As Double = 0
        Dim TotalAset, TotalHutang, TotalModal As Decimal
        Dim TotalAset1, TotalHutang1, TotalModal1 As Decimal

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Me.IsiTabel(Me.dtpFrom_I.Value, TBL)
            Me.IsiTabel(Me.dtpFrom_II.Value, TBL2)
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

                    'hitung variance
                    .Variance = .Saldo_Period_II - .Saldo_Period_I
                    If .Saldo_Period_I <> 0 Then
                        .Variance_Persen = Math.Round(((.Saldo_Period_II - .Saldo_Period_I) / _
                        .Saldo_Period_I) * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Variance_Persen = 0
                    End If
                End With

                If row.Primary_Akun = 1 Then
                    TotalAset = rowCetak.Saldo_Period_I
                    TotalAset1 = rowCetak.Saldo_Period_II
                ElseIf row.Primary_Akun = 2 Then
                    TotalHutang = rowCetak.Saldo_Period_I
                    TotalHutang1 = rowCetak.Saldo_Period_II
                ElseIf row.Primary_Akun = 3 Then
                    TotalModal = rowCetak.Saldo_Period_I
                    TotalModal1 = rowCetak.Saldo_Period_II
                End If

                DS.TblCetak.Rows.Add(rowCetak)
            Next

            'Buang semua akun yang nilainya 0 dan akun yang tidak aktif, tergantung pilihan user
            If Me.chkHideZero.Checked = True Then
                i = 0
                While i < DS.TblCetak.Rows.Count
                    rowCetak = DS.TblCetak.Rows(i)
                    If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Level_Perkiraan <> 1 Then
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
                            If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 Then
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
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)

            Me.report_comparison.SetDataSource(DS)
            report_comparison.SetParameterValue(report_comparison.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            report_comparison.SetParameterValue(report_comparison.Parameter_TanggalDari_1.ParameterFieldName, vDtpFrom_1)
            report_comparison.SetParameterValue(report_comparison.Parameter_TanggalDari_2.ParameterFieldName, vDtpFrom_2)
            report_comparison.SetParameterValue(report_comparison.Parameter_Tipe.ParameterFieldName, Tipe)
            report_comparison.SetParameterValue(report_comparison.Parameter_Period_I.ParameterFieldName, "Period I")
            report_comparison.SetParameterValue(report_comparison.Parameter_Period_II.ParameterFieldName, "Period II")
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report_comparison.SetParameterValue(report_comparison.Parameter_Negatif.ParameterFieldName, Me.Negatif)
            If Me.chkShowAccountCode.Checked = True Then
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_ShowNoAccount.ParameterFieldName, 1)
            Else
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_ShowNoAccount.ParameterFieldName, 0)
            End If
            Me.crv1.ReportSource = Me.report_comparison
            Me.crv1.Refresh()
            Me.crv1.Show()

            If Math.Round(TotalAset, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.dtpFrom_I.Value, "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset1, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang1, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal1, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.dtpFrom_II.Value, "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset1, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang1, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal1, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
        End Try
    End Sub

    Private Sub CetakBalanceSheet_Comparison_Budget()
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim Tipe As Double = 1
        Dim TotalAset, TotalHutang, TotalModal As Decimal
        Dim TotalAset1, TotalHutang1, TotalModal1 As Decimal

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Me.IsiTabel(Me.dtpFrom_I.Value, TBL)
            Me.ObjBudget.GetBudget_BS(Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), TBL2)
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
                    'hitung variance
                    If .IsSaldo_Period_IINull = True Then
                        .Saldo_Period_II = 0
                    End If
                    .Variance = .Saldo_Period_II - .Saldo_Period_I
                    If .Saldo_Period_I <> 0 Then
                        .Variance_Persen = Math.Round(((.Saldo_Period_II - .Saldo_Period_I) / _
                        .Saldo_Period_I) * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Variance_Persen = 0
                    End If
                End With

                If row.Primary_Akun = 1 Then
                    TotalAset = rowCetak.Saldo_Period_I
                    TotalAset1 = rowCetak.Saldo_Period_II
                ElseIf row.Primary_Akun = 2 Then
                    TotalHutang = rowCetak.Saldo_Period_I
                    TotalHutang1 = rowCetak.Saldo_Period_II
                ElseIf row.Primary_Akun = 3 Then
                    TotalModal = rowCetak.Saldo_Period_I
                    TotalModal1 = rowCetak.Saldo_Period_II
                End If

                DS.TblCetak.Rows.Add(rowCetak)
            Next

            'Buang semua akun yang nilainya 0 dan akun yang tidak aktif, tergantung pilihan user
            If Me.chkHideZero.Checked = True Then
                i = 0
                While i < DS.TblCetak.Rows.Count
                    rowCetak = DS.TblCetak.Rows(i)
                    If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 And rowCetak.Level_Perkiraan <> 1 Then
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
                            If rowCetak.Saldo_Period_I = 0 And rowCetak.Saldo_Period_II = 0 Then
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
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)

            Me.report_comparison.SetDataSource(DS)
            report_comparison.SetParameterValue(report_comparison.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            report_comparison.SetParameterValue(report_comparison.Parameter_TanggalDari_1.ParameterFieldName, vDtpFrom_1)
            report_comparison.SetParameterValue(report_comparison.Parameter_TanggalDari_2.ParameterFieldName, vDtpFrom_2)
            report_comparison.SetParameterValue(report_comparison.Parameter_Tipe.ParameterFieldName, Tipe)
            report_comparison.SetParameterValue(report_comparison.Parameter_Period_I.ParameterFieldName, "Actual")
            report_comparison.SetParameterValue(report_comparison.Parameter_Period_II.ParameterFieldName, "Budget")
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report_comparison.SetParameterValue(report_comparison.Parameter_Negatif.ParameterFieldName, Me.Negatif)
            If Me.chkShowAccountCode.Checked = True Then
                Me.report_comparison.SetParameterValue(report.Parameter_ShowNoAccount.ParameterFieldName, 1)
            Else
                Me.report_comparison.SetParameterValue(report.Parameter_ShowNoAccount.ParameterFieldName, 0)
            End If
            Me.crv1.ReportSource = Me.report_comparison
            Me.crv1.Refresh()
            Me.crv1.Show()

            If Math.Round(TotalAset, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check data on " + Format(Me.dtpFrom_I.Value, "MMM yyyy") + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If

            If Math.Round(TotalAset1, Me.Def_Prec, MidpointRounding.AwayFromZero) <> Math.Round(TotalHutang1, Me.Def_Prec, MidpointRounding.AwayFromZero) + Math.Round(TotalModal1, Me.Def_Prec, MidpointRounding.AwayFromZero) Then
                MsgBox("Unbalance data found, please re-check budget data." + vbCrLf + _
                "Total Asset: " + Me.ObjCD.DigitGrouping(TotalAset1, Me.Def_Prec) + vbCrLf + _
                "Total Liabilities: " + Me.ObjCD.DigitGrouping(TotalHutang1, Me.Def_Prec) + vbCrLf + _
                "Total Equity: " + Me.ObjCD.DigitGrouping(TotalModal1, Me.Def_Prec), MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
        End Try
    End Sub

    Private Sub chkHideZeroInactiveAccount_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHideZeroInactiveAccount.CheckedChanged
        Try
            If Me.chkHideZeroInactiveAccount.Checked = False Then
                Me.chkHideZero.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkHideZero_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHideZero.CheckedChanged
        Try
            If Me.chkHideZero.Checked = True Then
                Me.chkHideZeroInactiveAccount.Checked = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowLvl2.CheckedChanged
        Try
            If Me.chkShowLvl2.Checked = True Then
                Me.chkShowLvl3.Enabled = True
                Me.chkShowLvl4.Enabled = False
                Me.chkShowLvl5.Enabled = False
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl8.Enabled = False
            Else
                Me.chkShowLvl3.Enabled = False
                Me.chkShowLvl4.Enabled = False
                Me.chkShowLvl3.Checked = False
                Me.chkShowLvl4.Checked = False
                Me.chkShowLvl5.Enabled = False
                Me.chkShowLvl5.Checked = False
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl6.Checked = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl7.Checked = False
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowLvl3.CheckedChanged
        Try
            If Me.chkShowLvl3.Checked = True Then
                Me.chkShowLvl4.Enabled = True
                Me.chkShowLvl5.Enabled = False
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl8.Enabled = False
            Else
                Me.chkShowLvl4.Enabled = False
                Me.chkShowLvl4.Checked = False
                Me.chkShowLvl5.Enabled = False
                Me.chkShowLvl5.Checked = False
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl6.Checked = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl7.Checked = False
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub BuangKodeAkun(ByVal TBL As DseLaporanTrialBalance.Trial_BalanceDataTable)
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim Space As String = "     "
        Dim SpaceTotal As String = ""

        Try
            For Each row In TBL.Rows
                If row.Bold = 1 Then
                    If row.Total = 1 Then
                        If row.Level_Perkiraan = 1 Then
                            row.Akun = "Total " + row.Description
                        ElseIf row.Level_Perkiraan = 2 Then
                            row.Akun = Space + "Total " + row.Description
                        ElseIf row.Level_Perkiraan = 3 Then
                            row.Akun = Space + Space + "Total " + row.Description
                        ElseIf row.Level_Perkiraan = 4 Then
                            row.Akun = Space + Space + Space + "Total " + row.Description
                        End If
                    Else
                        If row.Level_Perkiraan = 1 Then
                            row.Akun = row.Description
                        ElseIf row.Level_Perkiraan = 2 Then
                            row.Akun = Space + row.Description
                        ElseIf row.Level_Perkiraan = 3 Then
                            row.Akun = Space + Space + row.Description
                        ElseIf row.Level_Perkiraan = 4 Then
                            row.Akun = Space + Space + Space + row.Description
                        End If
                    End If
                Else
                    If row.Level_Perkiraan = 1 Then
                        row.Akun = row.Description
                    ElseIf row.Level_Perkiraan = 2 Then
                        row.Akun = Space + row.Description
                    ElseIf row.Level_Perkiraan = 3 Then
                        row.Akun = Space + Space + row.Description
                    ElseIf row.Level_Perkiraan = 4 Then
                        row.Akun = Space + Space + Space + row.Description
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDeptCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDeptCode.KeyDown
        Dim i As New Core_Module.FrmMasterDepartementCari
        Dim Obj As New Core_Module.ClsMasterDepartement
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtDeptCode.Text) = True Then
                    Me.TxtPrimaryDept.Text = Obj.GetPrimary(Me.TxtDeptCode.Text)
                    If Me.TxtJobCode.Visible = True Then
                        Me.TxtJobCode.Focus()
                    Else
                        Me.btnShow.Focus()
                    End If
                Else
                    If Me.TxtDeptCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtDeptCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryDepartement > 0 Then
                        Me.TxtPrimaryDept.Text = i.PrimaryDepartement
                        Me.TxtDeptCode.Text = Obj.GetDeptCode(Me.TxtPrimaryDept.Text)
                        If Me.TxtJobCode.Visible = True Then
                            Me.TxtJobCode.Focus()
                        Else
                            Me.btnShow.Focus()
                        End If
                    Else
                        Me.TxtDeptCode.Clear()
                        Me.TxtPrimaryDept.Text = "-1"
                        Me.TxtDeptCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub TxtDeptCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDeptCode.Leave
        Dim i As New Core_Module.FrmMasterDepartementCari
        Dim obj As New Core_Module.ClsMasterDepartement
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtDeptCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtDeptCode.Text) = True Then
                    Me.TxtPrimaryDept.Text = obj.GetPrimary(Me.TxtDeptCode.Text)
                Else
                    If Me.TxtDeptCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtDeptCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryDepartement > 0 Then
                        Me.TxtPrimaryDept.Text = i.PrimaryDepartement
                        Me.TxtDeptCode.Text = obj.GetDeptCode(Me.TxtPrimaryDept.Text)
                        If Me.TxtJobCode.Visible = True Then
                            Me.TxtJobCode.Focus()
                        Else
                            Me.btnShow.Focus()
                        End If
                    Else
                        Me.TxtDeptCode.Clear()
                        Me.TxtPrimaryDept.Text = "-1"
                        Me.TxtDeptCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryDept.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindDept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindDept.Click
        Dim i As New Core_Module.FrmMasterDepartementCari
        Dim obj As New Core_Module.ClsMasterDepartement

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtDeptCode.Text <> "" Then
                i.txtCari.Text = Me.TxtDeptCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryDepartement > 0 Then
                Me.TxtPrimaryDept.Text = i.PrimaryDepartement
                Me.TxtDeptCode.Text = obj.GetDeptCode(Me.TxtPrimaryDept.Text)
                If Me.TxtJobCode.Visible = True Then
                    Me.TxtJobCode.Focus()
                Else
                    Me.btnShow.Focus()
                End If
            Else
                Me.TxtDeptCode.Clear()
                Me.TxtDeptCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtJobCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtJobCode.KeyDown
        Dim i As New Core_Module.FrmMasterJobCari
        Dim Obj As New Core_Module.ClsMasterJob
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtJobCode.Text) = True Then
                    Me.TxtPrimaryJob.Text = Obj.GetPrimary(Me.TxtJobCode.Text)
                    If Me.TxtJobCode.Visible = True Then
                        Me.TxtJobCode.Focus()
                    Else
                        Me.btnShow.Focus()
                    End If
                Else
                    If Me.TxtJobCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtJobCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryJob > 0 Then
                        Me.TxtPrimaryJob.Text = i.PrimaryJob
                        Me.TxtJobCode.Text = Obj.GetJobCode(Me.TxtPrimaryJob.Text)
                        If Me.TxtJobCode.Visible = True Then
                            Me.TxtJobCode.Focus()
                        Else
                            Me.btnShow.Focus()
                        End If
                    Else
                        Me.TxtJobCode.Clear()
                        Me.TxtPrimaryJob.Text = "-1"
                        Me.TxtJobCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub TxtJobCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobCode.Leave
        Dim i As New Core_Module.FrmMasterJobCari
        Dim obj As New Core_Module.ClsMasterJob
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtJobCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtJobCode.Text) = True Then
                    Me.TxtPrimaryJob.Text = obj.GetPrimary(Me.TxtJobCode.Text)
                Else
                    If Me.TxtJobCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtJobCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryJob > 0 Then
                        Me.TxtPrimaryJob.Text = i.PrimaryJob
                        Me.TxtJobCode.Text = obj.GetJobCode(Me.TxtPrimaryJob.Text)
                        If Me.TxtJobCode.Visible = True Then
                            Me.TxtJobCode.Focus()
                        Else
                            Me.btnShow.Focus()
                        End If
                    Else
                        Me.TxtJobCode.Clear()
                        Me.TxtPrimaryJob.Text = "-1"
                        Me.TxtJobCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryJob.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindJob_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindJob.Click
        Dim i As New Core_Module.FrmMasterJobCari
        Dim obj As New Core_Module.ClsMasterJob

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtJobCode.Text <> "" Then
                i.txtCari.Text = Me.TxtJobCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryJob > 0 Then
                Me.TxtPrimaryJob.Text = i.PrimaryJob
                Me.TxtJobCode.Text = obj.GetJobCode(Me.TxtPrimaryJob.Text)
                If Me.TxtJobCode.Visible = True Then
                    Me.TxtJobCode.Focus()
                Else
                    Me.btnShow.Focus()
                End If
            Else
                Me.TxtJobCode.Clear()
                Me.TxtJobCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtJobCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobCode.GotFocus
        Try
            Me.TxtJobCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJobCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobCode.LostFocus
        Try
            Me.TxtJobCode.BackColor = Me.TxtPrimaryDept.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDeptCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDeptCode.GotFocus
        Try
            Me.TxtDeptCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDeptCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDeptCode.LostFocus
        Try
            Me.TxtDeptCode.BackColor = Me.TxtPrimaryDept.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim ij As New FrmLaporan_ProfitNLoss_Setting

        Try
            ij.chkShowLvl2.Checked = False
            ij.chkShowLvl2.Checked = Me.chkShowLvl2.Checked
            ij.chkShowLvl3.Checked = Me.chkShowLvl3.Checked
            ij.chkShowLvl4.Checked = Me.chkShowLvl4.Checked
            ij.chkShowLvl5.Checked = Me.chkShowLvl5.Checked
            ij.chkShowLvl6.Checked = Me.chkShowLvl6.Checked
            ij.chkShowLvl7.Checked = Me.chkShowLvl7.Checked
            ij.chkShowLvl8.Checked = Me.chkShowLvl8.Checked

            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.chkShowLvl2.Checked = ij.chkShowLvl2.Checked
                Me.chkShowLvl3.Checked = ij.chkShowLvl3.Checked
                Me.chkShowLvl4.Checked = ij.chkShowLvl4.Checked
                Me.chkShowLvl5.Checked = ij.chkShowLvl5.Checked
                Me.chkShowLvl6.Checked = ij.chkShowLvl6.Checked
                Me.chkShowLvl7.Checked = ij.chkShowLvl7.Checked
                Me.chkShowLvl8.Checked = ij.chkShowLvl8.Checked
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub chkShowLvl4_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowLvl4.CheckedChanged
        Try
            If Me.chkShowLvl4.Checked = True Then
                Me.chkShowLvl5.Enabled = True
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl8.Enabled = False
            Else
                Me.chkShowLvl5.Enabled = False
                Me.chkShowLvl5.Checked = False
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl6.Checked = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl7.Checked = False
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl5_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowLvl5.CheckedChanged
        Try
            If Me.chkShowLvl5.Checked = True Then
                Me.chkShowLvl6.Enabled = True
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl8.Enabled = False
            Else
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl6.Checked = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl7.Checked = False
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl6_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowLvl6.CheckedChanged
        Try
            If Me.chkShowLvl6.Checked = True Then
                Me.chkShowLvl7.Enabled = True
                Me.chkShowLvl8.Enabled = False
            Else
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl7.Checked = False
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl7_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowLvl7.CheckedChanged
        Try
            If Me.chkShowLvl7.Checked = True Then
                Me.chkShowLvl8.Enabled = True
            Else
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class